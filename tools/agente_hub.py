"""Agente que executa os pedidos de atualização feitos pelo hub.

Por que existe: o pipeline lê as planilhas do Google Drive montado em `G:`, que
só existe nesta máquina. O botão no hub não executa nada — ele enfileira um
pedido na tabela `hub_job`, e este agente pega a fila e roda. Mesmo arranjo do
ETL de Indicadores do LuxorP&CHub: o PC precisa estar ligado, e o hub mostra o
estado do pedido enquanto isso.

Uso:
    python tools/agente_hub.py            # processa a fila e sai
    python tools/agente_hub.py --loop     # fica de pé, olhando a fila a cada 5 s
    python tools/agente_hub.py --um       # processa só o pedido mais antigo
    python tools/agente_hub.py --status   # imprime a fila, sem executar

Agendar o `--loop`, com o agendador disparando a cada minuto só como vigia:
    schtasks /create /tn "HPG - Agente do hub" /sc minute /mo 1
             /tr "pythonw C:\\...\\tools\\agente_hub.py --loop"
O schtasks cria a tarefa com "não iniciar nova instância": enquanto o loop vive,
os disparos são ignorados; se ele cair (reboot, rede), o minuto seguinte o
levanta. Tirar o limite de duração (padrão 72 h), senão o Windows mata o loop
no meio de um pedido — Settings.ExecutionTimeLimit = "PT0S" no Set-ScheduledTask.
Mudou este arquivo? Encerrar o pythonw do loop; o vigia sobe o código novo.
Antes a tarefa rodava o modo simples a cada 10 min: em 30/09/2026 a semanal
esperou 7 min na fila para rodar em 34 s.

Duas garantias que o hub sozinho não daria:

  semanal — roda `PGSemanal.py` SEM data e SEM `--forcar`. A janela sai do último
  snapshot congelado (`_janela()`) e o próprio script aborta se existir semana
  posterior já fechada. Ou seja: o botão nunca retroage nem recalcula semana
  passada, porque não tem como pedir isso.

  comitê — crava uma versão do conteúdo ANTES de atualizar. `build_comite.py` só
  lê `comite_conteudo`, então atualizar não apagaria texto de todo jeito; a
  versão cravada é a rede pro caso de alguém rodar isso no meio de uma edição.
"""
from __future__ import annotations

import json
import subprocess
import sys
import time
from datetime import datetime, timezone
from pathlib import Path

import requests
from dotenv import dotenv_values

REPO = Path(__file__).resolve().parent.parent
PY = sys.executable

# Cada tipo é uma sequência de comandos; se um passo falha, o job para nele.
PIPELINES = {
    "semanal": [
        ("fecha a semana", [PY, "PGSemanal.py", "--no-open"]),
        ("monta o dashboard", [PY, "tools/build_semanal.py"]),
        # A auditoria sai do MESMO snapshot que acabou de ser congelado: sem este
        # passo ela seguia mostrando a semana anterior. E `estado` leva os snapshots
        # para o bucket — é a memória do pipeline, que não se reconstrói porque as
        # planilhas do Drive são sobrescritas a cada semana (ver publish_hub.py).
        ("monta a auditoria", [PY, "tools/build_auditoria.py"]),
        ("publica no bucket", [PY, "tools/publish_hub.py", "semanal", "auditoria", "estado"]),
    ],
    # O comitê tem dois pedidos (`detalhe.acao` no hub_job): atualizar o mês no ar
    # e gerar o seguinte. Nenhum remonta os outros meses — vêm do comite.json
    # publicado, como estão. Pedido sem ação (hub de antes de 30/09) é atualizar.
    "comite": [
        ("monta o deck", [PY, "tools/build_comite.py", "--atualizar"]),
        ("publica no bucket", [PY, "tools/publish_hub.py", "comite"]),
    ],
    "comite:novo": [
        ("gera o mês novo", [PY, "tools/build_comite.py", "--novo"]),
        ("publica no bucket", [PY, "tools/publish_hub.py", "comite"]),
    ],
}


def pipeline_de(job: dict):
    """A sequência do pedido. A ação vem do hub, mas só escolhe entre as chaves
    acima: nenhum texto do pedido chega à linha de comando."""
    acao = (job.get("detalhe") or {}).get("acao")
    chave = job["tipo"] if acao in (None, "atualizar") else f"{job['tipo']}:{acao}"
    return PIPELINES.get(chave)

# O script sai com código 0 mesmo quando se recusa a rodar (fonte velha, semana
# posterior já congelada). Recusa não é sucesso: o hub tem de mostrar em
# vermelho, senão o usuário acha que atualizou.
RECUSAS = ("ABORTADO", "ATENCAO: ja existe(m) semana(s) congelada(s)")

LOG_MAX = 20000     # o log vai pro banco e aparece na tela; o fim é o que importa
POLL_S = 5          # o mesmo ritmo com que o hub relê o pedido (assets/hubjob.js)


def env() -> tuple[str, str]:
    cfg = dotenv_values(REPO / ".env")
    url = (cfg.get("SUPABASE_URL") or "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env da raiz.")
    return url, key


def api(url: str, key: str):
    s = requests.Session()
    s.headers.update({"apikey": key, "Authorization": f"Bearer {key}",
                      "Content-Type": "application/json"})
    return s


def agora() -> str:
    return datetime.now(timezone.utc).isoformat()


def fila(s, url) -> list[dict]:
    r = s.get(f"{url}/rest/v1/hub_job?status=in.(fila,rodando)&order=pedido_em", timeout=60)
    r.raise_for_status()
    return r.json()


def atualiza(s, url, jid: int, **campos):
    s.patch(f"{url}/rest/v1/hub_job?id=eq.{jid}", json=campos, timeout=60).raise_for_status()


def crava_versao(s, url, mes: str, rotulo: str) -> str:
    """Congela o conteúdo do mês antes de mexer em qualquer coisa."""
    r = s.post(f"{url}/rest/v1/rpc/comite_cravar_versao",
               json={"p_mes": mes, "p_rotulo": rotulo}, timeout=60)
    if r.status_code >= 300:
        return f"  [versao] não deu pra cravar ({r.status_code}): {r.text[:200]}"
    return f"  [versao] conteúdo de {mes} cravado como v{r.text.strip()}"


def meses_com_conteudo(s, url) -> list[str]:
    r = s.get(f"{url}/rest/v1/comite_conteudo?select=mes&order=mes", timeout=60)
    return [x["mes"] for x in r.json()] if r.status_code < 300 else []


def roda(passos, prefixo: str = "") -> tuple[bool, str]:
    """Executa a sequência. Devolve (deu certo, log)."""
    partes = [prefixo] if prefixo else []
    for titulo, cmd in passos:
        partes.append(f"\n$ {' '.join(Path(c).name if c == PY else c for c in cmd)}")
        try:
            p = subprocess.run(cmd, cwd=REPO, capture_output=True, text=True,
                               encoding="utf-8", errors="replace", timeout=3600)
        except subprocess.TimeoutExpired:
            partes.append(f"[{titulo}] passou de 1h e foi interrompido")
            return False, "\n".join(partes)
        saida = (p.stdout or "") + (p.stderr or "")
        partes.append(saida.strip())
        if p.returncode != 0:
            partes.append(f"[{titulo}] terminou com código {p.returncode}")
            return False, "\n".join(partes)
        recusa = next((m for m in RECUSAS if m in saida), None)
        if recusa:
            partes.append(f"[{titulo}] o pipeline se recusou a rodar — nada foi publicado")
            return False, "\n".join(partes)
    return True, "\n".join(partes)


def processa(s, url, job: dict) -> bool:
    jid, tipo = job["id"], job["tipo"]
    print(f"[job {jid}] {tipo} — pedido por {job.get('pedido_por') or '?'}")
    passos = pipeline_de(job)
    if not passos:
        atualiza(s, url, jid, status="erro", terminado_em=agora(),
                 log=f"pedido que o agente não conhece: {tipo} {job.get('detalhe') or ''}")
        return False
    atualiza(s, url, jid, status="rodando", iniciado_em=agora())

    prefixo = ""
    if tipo == "comite":
        marca = datetime.now().strftime("%d/%m/%Y %H:%M")
        linhas = [crava_versao(s, url, m, f"antes da atualização de {marca}")
                  for m in meses_com_conteudo(s, url)]
        prefixo = "\n".join(linhas)

    ok, log = roda(passos, prefixo)
    atualiza(s, url, jid, status="ok" if ok else "erro", terminado_em=agora(),
             log=log[-LOG_MAX:])
    print(f"[job {jid}] {'ok' if ok else 'ERRO'}")
    return ok


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    url, key = env()
    s = api(url, key)
    args = sys.argv[1:]

    if "--status" in args:
        pend = fila(s, url)
        print(f"{len(pend)} pedido(s) na fila" if pend else "fila vazia")
        for j in pend:
            print(f"  #{j['id']} {j['tipo']:8} {j['status']:8} {j['pedido_em']} {j.get('pedido_por') or ''}")
        return

    if "--loop" in args:
        while True:
            try:
                for j in [j for j in fila(s, url) if j["status"] == "fila"]:
                    processa(s, url, j)
            except requests.RequestException as e:
                # sem rede ou banco fora: o loop não cai por isso, tenta de novo
                print(f"[loop] sem acesso ao banco: {e!r}")
            time.sleep(POLL_S)

    pend = [j for j in fila(s, url) if j["status"] == "fila"]
    if not pend:
        return                      # silencioso: quase sempre é isso
    if "--um" in args:
        pend = pend[:1]
    falhou = sum(0 if processa(s, url, j) else 1 for j in pend)
    if falhou:
        sys.exit(f"{falhou} pedido(s) terminaram em erro — ver o log no hub")


if __name__ == "__main__":
    main()
