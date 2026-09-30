"""Agente que executa, no notebook, os pedidos de atualização feitos pelo hub.

Por que existe: o pipeline lê as planilhas do Google Drive montado em `G:`, que
só existe nesta máquina. O botão no hub não executa nada — ele enfileira um
pedido na tabela `hub_job`, e este agente pega a fila e roda. Mesmo arranjo do
ETL de Indicadores do LuxorP&CHub: o PC precisa estar ligado, e o hub mostra o
estado do pedido enquanto isso.

Desde 30/09/2026 a Azure também executa pedidos (tools/roda_pedido.py, chamado
pelo luxor-planejamento-functions), baixando o Drive pela API — é o que tira o
notebook do caminho. Os passos de cada pedido moram em tools/pipelines_hub.py,
os mesmos para os dois executores, e tomar o pedido é atômico: com os dois de
pé, cada pedido roda uma vez só. Na transição, `--tipos` diz quais tipos o
notebook ainda atende.

Uso:
    python tools/agente_hub.py            # processa a fila e sai
    python tools/agente_hub.py --loop     # fica de pé, olhando a fila a cada 5 s
    python tools/agente_hub.py --loop --tipos comite   # só o comitê
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

import sys
import time

import requests
from dotenv import dotenv_values

from pipelines_hub import REPO, api, fila, processa

POLL_S = 5          # o mesmo ritmo com que o hub relê o pedido (assets/hubjob.js)


def env() -> tuple[str, str]:
    cfg = dotenv_values(REPO / ".env")
    url = (cfg.get("SUPABASE_URL") or "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env da raiz.")
    return url, key


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    url, key = env()
    s = api(url, key)
    args = sys.argv[1:]
    tipos = None
    if "--tipos" in args:
        tipos = set(args[args.index("--tipos") + 1].split(","))
    minha = lambda j: j["status"] == "fila" and (tipos is None or j["tipo"] in tipos)

    if "--status" in args:
        pend = fila(s, url)
        print(f"{len(pend)} pedido(s) na fila" if pend else "fila vazia")
        for j in pend:
            print(f"  #{j['id']} {j['tipo']:8} {j['status']:8} {j['pedido_em']} {j.get('pedido_por') or ''}")
        return

    if "--loop" in args:
        while True:
            try:
                for j in [j for j in fila(s, url) if minha(j)]:
                    processa(s, url, j)
            except requests.RequestException as e:
                # sem rede ou banco fora: o loop não cai por isso, tenta de novo
                print(f"[loop] sem acesso ao banco: {e!r}")
            time.sleep(POLL_S)

    pend = [j for j in fila(s, url) if minha(j)]
    if not pend:
        return                      # silencioso: quase sempre é isso
    if "--um" in args:
        pend = pend[:1]
    falhou = sum(1 for j in pend if processa(s, url, j) is False)
    if falhou:
        sys.exit(f"{falhou} pedido(s) terminaram em erro — ver o log no hub")


if __name__ == "__main__":
    main()
