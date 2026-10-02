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
import io
import os
import sys
import zipfile
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
    # movimentação do plantel: cota, valor e comprador por animal — privado como
    # o resto. O nome antes do ponto é o que a policy do storage usa pra decidir
    # quem baixa, então 'plantel.json' casa com hub_can('plantel').
    "plantel": (ROOT / "assets/plantel/spec.json", "plantel.json",
                "application/json"),
    # A auditoria cita animal e comprador — mesma regra dos outros: bucket privado.
    "auditoria": (ROOT / "dashboards/auditoria_semanal.html", "auditoria.html",
                  "text/html; charset=utf-8"),
    # 'gastos' NÃO entra aqui, de propósito.
    #
    # O painel de Gastos do Haras é publicado pela rotina `publicar_gastos_haras` do
    # Function App, que lê a planilha do Drive e sobe o HTML direto para este mesmo
    # bucket como `gastos.html`. O arquivo não nasce neste repo.
    #
    # Registrar mesmo assim seria armadilha: `--all` percorre todo o DATASETS, e
    # `sobe()` devolve False para arquivo ausente, o que vira `sys.exit("Falha ao
    # publicar")`. O dataset ficaria quebrando o publish de todo mundo para nunca
    # poder ser publicado daqui.
    #
    # Para republicar à mão, invoque a função no Azure. O acesso segue o mesmo
    # porteiro: a policy usa hub_can(split_part(name,'.',1)), então 'gastos.html'
    # casa com hub_can('gastos'), que está no CHECK do hub_schema.sql.
}
# Memória do pipeline: vários arquivos sob o mesmo prefixo.
ESTADO = [
    (ROOT / "_cache/semanal_snapshots.json", "estado_semanal_snapshots.json"),
    (ROOT / "_cache/headcount_history.json", "estado_headcount_history.json"),
    (ROOT / "_cache/paricoes_extra.json",    "estado_paricoes_extra.json"),
    (ROOT / "_cache/acumulado_piso.json",    "estado_acumulado_piso.json"),
    # confirmação vista só pela planilha de receptoras, registrada na semana em que
    # apareceu e CUMULATIVA (ver _confirmados_por_receptora). Faltava aqui: o ensaio
    # da nuvem de 30/09/2026 perdeu as 5 prenhezes de 18/09 e publicaria acumulado no
    # mês 0 contra 5 — os arquivos do Drive eram byte a byte os mesmos do G:.
    (ROOT / "_cache/confirmados_extra.json", "estado_confirmados_extra.json"),
    # onde cada receptora estava em cada rodada da semana: é o que permite contar a
    # transferência de quem muda mais de uma vez na mesma semana (ver _saltos_internos)
    (ROOT / "_cache/locais_semana.json", "estado_locais_semana.json"),
    # input humano (doadoras ciclando): não sai de planilha nenhuma, então perder o
    # arquivo é perder o dado
    (ROOT / "_cache/semanal_manual.json",    "estado_semanal_manual.json"),
    # Pastas vão zipadas. O arquivo das linhas das fontes (um JSON por semana) e os
    # PDFs de embriões de cada semana, que o dashboard embute: até 30/09/2026 os dois
    # só existiam no disco de quem fechava, e o fechamento na nuvem — disco novo a
    # cada execução — perderia o histórico das linhas e publicaria o dashboard só com
    # os PDFs da semana.
    (ROOT / "_cache/fontes",                 "estado_fontes.zip"),
    (ROOT / "_cache/pdf",                    "estado_pdf.zip"),
]
# Memória do comitê: o que o deck lê e não se refaz a cada clique — o base_bi e os
# parquets mensais do plantel (estoque do mês), o último resumo da aba Plantel (se o
# node falhar), as safras encerradas do comparativo e o cache do Trello. Separada da
# lista do semanal DE PROPÓSITO: um pedido de comitê rodando junto com um de semanal
# subiria os snapshots que restaurou no começo por cima dos que o semanal acabou de
# gravar. O comitê restaura as duas e só publica esta.
ESTADO_COMITE = [
    (ROOT / "bases/base_bi.parquet",          "estado_comite_base_bi.parquet"),
    (ROOT / "_cache/parquet",                 "estado_comite_parquet.zip"),
    (ROOT / "_cache/plantel_hub",             "estado_comite_plantel_hub.zip"),
    (ROOT / "bases/comparativo_fechado.json", "estado_comite_comparativo_fechado.json"),
    (ROOT / "_cache/trello",                  "estado_comite_trello.zip"),
]
_ESTADOS = {"estado": ESTADO, "estado_comite": ESTADO_COMITE}
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
    """Traz um objeto do bucket para o disco. É a volta do `sobe`; `.zip` é pasta
    e volta descompactado dentro de `alvo`."""
    r = requests.get(f"{url}/storage/v1/object/{BUCKET}/{dest}",
                     headers={"Authorization": f"Bearer {key}"}, timeout=180)
    if r.status_code == 404:
        print(f"[skip] {dest} ainda não existe no bucket")
        return False
    if r.status_code >= 300:
        print(f"[erro] {dest} -> HTTP {r.status_code}: {r.text[:300]}")
        return False
    if dest.endswith(".zip"):
        alvo.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(io.BytesIO(r.content)) as z:
            z.extractall(alvo)
            n = len(z.namelist())
        print(f"[ok] {BUCKET}/{dest} ({len(r.content)//1024} KB, {n} arquivo(s)) -> {alvo.name}/")
        return True
    alvo.parent.mkdir(parents=True, exist_ok=True)
    alvo.write_bytes(r.content)
    print(f"[ok] {BUCKET}/{dest} ({len(r.content)//1024} KB) -> {alvo.name}")
    return True


def _zipa(pasta: Path) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(pasta.rglob("*")):
            if f.is_file():
                z.write(f, f.relative_to(pasta).as_posix())
    return buf.getvalue()


def restaurar_estado(comite: bool = False) -> int:
    """Baixa a memória do pipeline do bucket para o _cache local.

    Existe por causa da execução na nuvem: o disco da Function é descartável e os
    snapshots semanais NÃO se reconstroem — as planilhas do Drive são sobrescritas a
    cada semana. Sem isto, um fechamento na nuvem começaria sem histórico e publicaria
    a primeira semana como se fosse a única (sem diff de saídas, sem Δ, acumulado sem
    piso). Rodar antes do fechamento, sempre."""
    url, key = env()
    lista = ESTADO + (ESTADO_COMITE if comite else [])
    return sum(1 for origem, dest in lista if baixa(url, key, dest, origem))


def sobe(url, key, src: Path, dest: str, ctype: str, gerador: str = "") -> bool:
    if not src.exists():
        extra = f" — rode {gerador} antes." if gerador else ""
        print(f"[skip] {src.name} não existe{extra}")
        return False
    body = _zipa(src) if dest.endswith(".zip") else src.read_bytes()
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
        alvos = list(DATASETS) + list(_ESTADOS)

    desconhecido = [a for a in alvos if a not in DATASETS and a not in _ESTADOS]
    if desconhecido:
        sys.exit(f"Dataset não publicável: {', '.join(desconhecido)}. "
                 f"Válidos: {', '.join(list(DATASETS) + list(_ESTADOS))}")

    falhou = []
    for nome in alvos:
        if nome in _ESTADOS:
            print("[aviso] memória do pipeline: backup do que não se reconstrói. "
                  "Só admin lê (policy hpg_estado_read).")
            for src, dest in _ESTADOS[nome]:
                zipado = dest.endswith(".zip")
                if zipado and not src.is_dir():
                    print(f"[skip] {src.name}/ não existe — nada a guardar")
                    continue
                tipo = ("application/zip" if zipado else "application/json"
                        if dest.endswith(".json") else "application/octet-stream")
                if not sobe(url, key, src, dest, tipo):
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
