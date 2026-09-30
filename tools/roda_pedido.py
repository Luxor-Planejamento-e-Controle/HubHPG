"""Executa UM pedido do hub — é o que a Azure chama.

    python tools/roda_pedido.py <id> --nuvem           # o que a Azure roda
    python tools/roda_pedido.py <id> --nuvem --ensaio  # o mesmo, sem tomar o pedido,
                                                       # sem gravar status e sem publicar

Quem chama é o luxor-planejamento-functions (hpg_hub_job), dentro de uma CÓPIA
gravável deste repo e como subprocesso: o pacote da Function é somente leitura, e
pasta e processo próprios por pedido impedem que dois pedidos na mesma instância
se atropelem. Os passos são os de tools/pipelines_hub.py, os mesmos do notebook.

--nuvem prepara o que no notebook já existe:
  1. restaura do bucket a memória do pipeline (snapshots, linhas das fontes, PDFs de
     embriões) — o disco da Function é novo a cada execução;
  2. espelha do Drive só o que o pipeline abre (tools/sync_drive.py) e aponta
     HPG_DRIVE_ROOT para a cópia.

--ensaio é para conferir o caminho da nuvem sem mexer em nada — rodar SEMPRE numa
cópia do repo, nunca no de trabalho: o fechamento grava o snapshot no _cache de
onde roda.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from pipelines_hub import REPO, api, le, pipeline_de, processa, roda


def _env() -> tuple[str, str]:
    """Na Azure as App Settings viram variáveis de ambiente; no notebook, o .env."""
    from dotenv import dotenv_values
    cfg = dotenv_values(REPO / ".env")
    url = (os.getenv("SUPABASE_URL") or cfg.get("SUPABASE_URL") or "").rstrip("/")
    key = os.getenv("SUPABASE_SERVICE_ROLE_KEY") or cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        raise RuntimeError("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY")
    return url, key


def _prepara_nuvem(env: dict) -> list[str]:
    import publish_hub
    import sync_drive
    linhas = [f"[estado] {publish_hub.restaurar_estado()} arquivo(s) restaurado(s) do bucket"]
    destino = Path(os.getenv("HPG_DRIVE_CACHE") or (REPO / "_drive"))
    raiz = sync_drive.sincronizar(destino, verboso=False)
    env["HPG_DRIVE_ROOT"] = str(raiz)
    n = sum(1 for p in Path(raiz).rglob("*") if p.is_file())
    linhas.append(f"[drive] {n} arquivo(s) espelhado(s) do Drive")
    return linhas


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = sys.argv[1:]
    jid = int(args[0])
    nuvem, ensaio = "--nuvem" in args, "--ensaio" in args
    url, key = _env()
    s = api(url, key)
    job = le(s, url, jid)
    if not job:
        print(f"[job {jid}] não existe")
        return 1
    env = dict(os.environ)
    preparo = (lambda: _prepara_nuvem(env)) if nuvem else None

    if ensaio:
        linhas = preparo() if preparo else []
        ok, log = roda(pipeline_de(job), REPO, "\n".join(linhas), env,
                       pula={"publica no bucket"})
        print(log)
        print(f"[ensaio] {'ok' if ok else 'ERRO'} — nada foi tomado nem publicado")
        return 0 if ok else 1

    if job["status"] != "fila":
        print(f"[job {jid}] está '{job['status']}' — nada a fazer")
        return 0
    ok = processa(s, url, job, REPO, env, preparo, onde="Azure" if nuvem else "notebook")
    return 0 if ok in (True, None) else 1


if __name__ == "__main__":
    sys.exit(main())
