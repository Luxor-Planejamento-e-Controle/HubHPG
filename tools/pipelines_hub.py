"""O que cada pedido do hub roda, e como um pedido é executado.

Um lugar só para os dois executores: o agente do notebook (tools/agente_hub.py) e a
Azure (tools/roda_pedido.py, chamado pelo luxor-planejamento-functions). Os dois
rodam os MESMOS passos; o que muda é de onde vêm as planilhas (G: no notebook, a
cópia baixada pela API na nuvem) e quem prepara o terreno antes.

Tomar o pedido é atômico: a troca fila → rodando só acontece se o pedido ainda
estiver na fila. Com dois executores de pé — e a fila da Azure entrega mensagem
repetida de vez em quando —, sem isso o mesmo fechamento rodaria duas vezes.
"""
from __future__ import annotations

import subprocess
import sys
from datetime import datetime, timezone
from pathlib import Path

import requests

REPO = Path(__file__).resolve().parent.parent
PY = sys.executable

# Cada pedido é uma sequência de comandos, relativos à raiz do repo; se um passo
# falha, o pedido para nele.
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

# O script sai com código 0 mesmo quando se recusa a rodar (fonte velha, semana
# posterior já congelada). Recusa não é sucesso: o hub tem de mostrar em
# vermelho, senão o usuário acha que atualizou.
RECUSAS = ("ABORTADO", "ATENCAO: ja existe(m) semana(s) congelada(s)")

LOG_MAX = 20000     # o log vai pro banco e aparece na tela; o fim é o que importa


def pipeline_de(job: dict):
    """A sequência do pedido. A ação vem do hub, mas só escolhe entre as chaves
    acima: nenhum texto do pedido chega à linha de comando."""
    acao = (job.get("detalhe") or {}).get("acao")
    chave = job["tipo"] if acao in (None, "atualizar") else f"{job['tipo']}:{acao}"
    return PIPELINES.get(chave)


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


def le(s, url, jid: int) -> dict | None:
    r = s.get(f"{url}/rest/v1/hub_job?id=eq.{jid}", timeout=60)
    r.raise_for_status()
    return (r.json() or [None])[0]


def atualiza(s, url, jid: int, **campos):
    s.patch(f"{url}/rest/v1/hub_job?id=eq.{jid}", json=campos, timeout=60).raise_for_status()


def toma(s, url, jid: int) -> dict | None:
    """fila → rodando, só se ainda estiver na fila. Devolve o pedido, ou None se
    outro executor chegou antes."""
    r = s.patch(f"{url}/rest/v1/hub_job?id=eq.{jid}&status=eq.fila",
                json={"status": "rodando", "iniciado_em": agora()},
                headers={"Prefer": "return=representation"}, timeout=60)
    r.raise_for_status()
    return (r.json() or [None])[0]


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


def roda(passos, raiz: Path = REPO, prefixo: str = "", env: dict | None = None,
         pula=()) -> tuple[bool, str]:
    """Executa a sequência na raiz dada. Devolve (deu certo, log). `pula` são
    títulos de passo que não rodam (o ensaio não publica)."""
    partes = [prefixo] if prefixo else []
    for titulo, cmd in passos:
        if titulo in pula:
            partes.append(f"\n[{titulo}] pulado")
            continue
        partes.append(f"\n$ {' '.join(Path(c).name if c == PY else c for c in cmd)}")
        try:
            p = subprocess.run(cmd, cwd=raiz, capture_output=True, text=True, env=env,
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


def processa(s, url, job: dict, raiz: Path = REPO, env: dict | None = None,
             preparo=None, onde: str = "notebook") -> bool | None:
    """Toma o pedido, prepara (na nuvem: estado e Drive), roda e grava o resultado.
    Devolve None quando outro executor já tinha tomado o pedido."""
    jid, tipo = job["id"], job["tipo"]
    passos = pipeline_de(job)
    if not passos:
        atualiza(s, url, jid, status="erro", terminado_em=agora(),
                 log=f"pedido que o agente não conhece: {tipo} {job.get('detalhe') or ''}")
        return False
    if not toma(s, url, jid):
        print(f"[job {jid}] já estava com outro executor")
        return None
    print(f"[job {jid}] {tipo} — pedido por {job.get('pedido_por') or '?'} — rodando no {onde}")

    linhas = [f"[executor] {onde}"]
    try:
        if preparo:
            linhas += preparo()
        if tipo == "comite":
            marca = datetime.now().strftime("%d/%m/%Y %H:%M")
            linhas += [crava_versao(s, url, m, f"antes da atualização de {marca}")
                       for m in meses_com_conteudo(s, url)]
        ok, log = roda(passos, raiz, "\n".join(linhas), env)
    except Exception as exc:
        # nada pode deixar o pedido preso em "rodando": o hub travaria o botão
        ok, log = False, "\n".join(linhas + [f"[preparo] falhou: {exc!r}"])
    atualiza(s, url, jid, status="ok" if ok else "erro", terminado_em=agora(),
             log=log[-LOG_MAX:])
    print(f"[job {jid}] {'ok' if ok else 'ERRO'}")
    return ok
