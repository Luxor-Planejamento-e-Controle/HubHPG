"""Publica no bucket PRIVADO do Supabase (`hpg-data`) o que o pipeline gerou.

Roda depois de `PGSemanal.py` (aba semanal) e `tools/build_comite.py` (deck do
comitê). O nome no bucket importa: a policy de leitura (sql/hub_schema.sql) usa o
prefixo antes do ponto para decidir quem pode baixar — `semanal.html` exige
`hub_can('semanal')`.

Nada disso pode virar arquivo estático no Netlify: o site é público e o conteúdo
é dado do plantel (nome de animal, comprador, headcount). O hub baixa do bucket
já autenticado e injeta no iframe por srcdoc.

`estado` é caso à parte: é a memória do pipeline (snapshots semanais congelados),
que saiu do Git quando o repo virou público e NÃO se reconstrói, porque as
planilhas do Drive são sobrescritas a cada semana. Vai pro bucket como backup,
sob policy de admin.

Requer a service_role key (ignora RLS) num .env local — NUNCA versionar:

    SUPABASE_URL=https://xxxx.supabase.co
    SUPABASE_SERVICE_ROLE_KEY=eyJ...

Uso:
    python tools/publish_hub.py                 # semanal + comite
    python tools/publish_hub.py comite
    python tools/publish_hub.py estado          # backup da memória do pipeline
    python tools/publish_hub.py --all
"""
import os
import sys
from pathlib import Path

import requests
from dotenv import dotenv_values

ROOT = Path(__file__).resolve().parent.parent
BUCKET = "hpg-data"

# dataset -> (arquivo local, nome no bucket, content-type)
DATASETS = {
    "semanal": (ROOT / "assets/semanal/dashboard.html", "semanal.html",
                "text/html; charset=utf-8"),
    "comite":  (ROOT / "assets/comite/spec.json", "comite.json",
                "application/json"),
    # A auditoria cita animal e comprador — mesma regra dos outros: bucket privado.
    "auditoria": (ROOT / "dashboards/auditoria_semanal.html", "auditoria.html",
                  "text/html; charset=utf-8"),
}
# Memória do pipeline: vários arquivos sob o mesmo prefixo.
ESTADO = [
    (ROOT / "_cache/semanal_snapshots.json", "estado_semanal_snapshots.json"),
    (ROOT / "_cache/headcount_history.json", "estado_headcount_history.json"),
    (ROOT / "_cache/paricoes_extra.json",    "estado_paricoes_extra.json"),
    (ROOT / "_cache/acumulado_piso.json",    "estado_acumulado_piso.json"),
]
# Quem gera cada arquivo, pra mensagem de erro apontar o build certo.
GERADOR = {
    "semanal": "python PGSemanal.py",
    "comite":  "python tools/build_comite.py",
    "auditoria": "python tools/build_auditoria.py",
}
PADRAO = ["semanal", "comite", "auditoria"]


def env():
    """.env local primeiro; variável de ambiente depois.

    Numa Azure Function não há .env — as credenciais chegam como app settings, que o
    runtime expõe como variáveis de ambiente."""
    cfg = dotenv_values(ROOT / ".env")
    url = (cfg.get("SUPABASE_URL") or os.getenv("SUPABASE_URL") or "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY") or os.getenv("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY (no .env da raiz do "
                 "repo ou nas variáveis de ambiente).")
    return url, key


def baixa(url, key, dest: str, alvo: Path) -> bool:
    """Traz um objeto do bucket para o disco. É a volta do `sobe`."""
    r = requests.get(f"{url}/storage/v1/object/{BUCKET}/{dest}",
                     headers={"Authorization": f"Bearer {key}"}, timeout=180)
    if r.status_code == 404:
        print(f"[skip] {dest} ainda não existe no bucket")
        return False
    if r.status_code >= 300:
        print(f"[erro] {dest} -> HTTP {r.status_code}: {r.text[:300]}")
        return False
    alvo.parent.mkdir(parents=True, exist_ok=True)
    alvo.write_bytes(r.content)
    print(f"[ok] {BUCKET}/{dest} ({len(r.content)//1024} KB) -> {alvo.name}")
    return True


def restaurar_estado() -> int:
    """Baixa a memória do pipeline do bucket para o _cache local.

    Existe por causa da execução na nuvem: o disco da Function é descartável e os
    snapshots semanais NÃO se reconstroem — as planilhas do Drive são sobrescritas a
    cada semana. Sem isto, um fechamento na nuvem começaria sem histórico e publicaria
    a primeira semana como se fosse a única (sem diff de saídas, sem Δ, acumulado sem
    piso). Rodar antes do fechamento, sempre."""
    url, key = env()
    return sum(1 for origem, dest in ESTADO if baixa(url, key, dest, origem))


def sobe(url, key, src: Path, dest: str, ctype: str, gerador: str = "") -> bool:
    if not src.exists():
        extra = f" — rode {gerador} antes." if gerador else ""
        print(f"[skip] {src.name} não existe{extra}")
        return False
    body = src.read_bytes()
    r = requests.post(
        f"{url}/storage/v1/object/{BUCKET}/{dest}",
        data=body,
        headers={"Authorization": f"Bearer {key}", "Content-Type": ctype,
                 "x-upsert": "true", "cache-control": "no-store"},
        timeout=180,
    )
    if r.status_code >= 300:
        print(f"[erro] {dest} -> HTTP {r.status_code}: {r.text[:300]}")
        return False
    print(f"[ok] {dest} ({len(body)//1024} KB) -> {BUCKET}/{dest}")
    return True


def main():
    url, key = env()
    alvos = sys.argv[1:] or PADRAO
    if alvos == ["--all"]:
        alvos = list(DATASETS) + ["estado"]

    desconhecido = [a for a in alvos if a not in DATASETS and a != "estado"]
    if desconhecido:
        sys.exit(f"Dataset não publicável: {', '.join(desconhecido)}. "
                 f"Válidos: {', '.join(list(DATASETS) + ['estado'])}")

    falhou = []
    for nome in alvos:
        if nome == "estado":
            print("[aviso] memória do pipeline: backup do que não se reconstrói. "
                  "Só admin lê (policy hpg_estado_read).")
            for src, dest in ESTADO:
                if not sobe(url, key, src, dest, "application/json"):
                    falhou.append(dest)
            continue
        src, dest, ctype = DATASETS[nome]
        print(f"[aviso] {nome} é dado do plantel. Vai pro bucket PRIVADO, "
              f"visível só para quem tem hub_can('{nome}').")
        if not sobe(url, key, src, dest, ctype, GERADOR.get(nome, "")):
            falhou.append(dest)

    if falhou:
        sys.exit(f"Falha ao publicar: {', '.join(falhou)}")


if __name__ == "__main__":
    main()
