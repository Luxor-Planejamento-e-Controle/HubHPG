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

import contextlib
import os
import shutil
import sys
import time
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


@contextlib.contextmanager
def _trava(pasta: Path):
    """Um espelhamento por vez na mesma pasta. Na Azure o espelho é da instância e
    reaproveitado entre pedidos, e semanal e comitê podem rodar ao mesmo tempo."""
    try:
        import fcntl
    except ImportError:          # Windows: um executor só, ninguém disputa
        yield
        return
    pasta.parent.mkdir(parents=True, exist_ok=True)
    with open(f"{pasta}.lock", "w") as f:
        fcntl.flock(f, fcntl.LOCK_EX)
        try:
            yield
        finally:
            fcntl.flock(f, fcntl.LOCK_UN)


BLOB_CONTA = "https://azblobstoragebz.blob.core.windows.net"
BLOB_CONTAINER = "luxor-planejamento-e-controle"


def _container():
    """O container do Blob. Na Azure, pela identidade do Function App (leitura só neste
    container); no ensaio, pela connection string do usuário."""
    from azure.storage.blob import BlobServiceClient
    conn = os.getenv("AZURE_STORAGE_CONNECTION_STRING")
    if conn:
        svc = BlobServiceClient.from_connection_string(conn)
    else:
        from azure.identity import ManagedIdentityCredential
        svc = BlobServiceClient(BLOB_CONTA, credential=ManagedIdentityCredential())
    return svc.get_container_client(BLOB_CONTAINER)


def _baixa_blob(cont, prefixo: str, destino: Path) -> int:
    """Os blobs sob `prefixo`, com a data do upload como mtime: a regra "base extraída
    depois do fim do mês" do comitê lê a data da base, e o extractor sobe no instante
    em que grava."""
    n = 0
    for b in cont.list_blobs(name_starts_with=prefixo):
        alvo = destino / Path(b.name).name
        alvo.parent.mkdir(parents=True, exist_ok=True)
        with open(alvo, "wb") as f:
            cont.download_blob(b.name).readinto(f)
        quando = b.last_modified.timestamp()
        os.utime(alvo, (quando, quando))
        n += 1
    return n


def _prepara_comite(env: dict) -> list[str]:
    """O que o comitê lê fora do Drive do haras e que no notebook está em outros
    repos: a base e os DRE anuais (cópia que o LxDREdataExtractor sobe) e a foto da
    carteira (ControleInadimplencia.py), baixados do Blob para as pastas que as HPG_*
    do build_comite apontam."""
    from datetime import date
    t = time.time()
    cont = _container()
    nuvem = REPO / "_nuvem"
    raiz = "LuxorControlDatabase"
    n = _baixa_blob(cont, f"{raiz}/excel/DRE_Historico.xlsx", nuvem / "dre")
    hoje = date.today()
    for a in ((hoje.year - 1, hoje.year) if hoje.month <= 2 else (hoje.year,)):
        n += _baixa_blob(cont, f"{raiz}/excel/dre_anual/{a}/",
                         nuvem / "dre_anual" / str(a) / "Fluxo de Caixa e DRE")
    for pasta in ("parquet", "excel"):
        n += _baixa_blob(cont, f"{raiz}/{pasta}/inadimplencia/historico/",
                         nuvem / "inad" / "historico")
    env.update(HPG_DRE_DIR=str(nuvem / "dre"), HPG_DRE_ANUAL_DIR=str(nuvem / "dre_anual"),
               HPG_INAD_DIR=str(nuvem / "inad"))
    return [f"[blob] {n} arquivo(s) do DRE e da cobrança ({time.time() - t:.0f} s)"]


def _prepara_nuvem(env: dict, job: dict) -> list[str]:
    import publish_hub
    import sync_drive
    comite = job["tipo"] == "comite"
    t = time.time()
    n = publish_hub.restaurar_estado(comite=comite)
    # O deck publicado. A auditoria monta a seção do comitê a partir dele, e no
    # notebook ele existe (assets/comite/spec.json, gitignored): sem isto a primeira
    # semanal na Azure (30/09/2026) publicou a auditoria só com a parte da semana.
    url, key = publish_hub.env()
    publish_hub.baixa(url, key, "comite.json", REPO / "assets" / "comite" / "spec.json")
    linhas = [f"[estado] {n} arquivo(s) restaurado(s) do bucket, mais o deck publicado "
              f"({time.time() - t:.0f} s)"]
    # O espelho do Drive fica na instância (HPG_DRIVE_CACHE) e só baixa o que mudou; o
    # pedido trabalha numa cópia própria, feita sob a trava, para outro pedido não
    # trocar uma planilha no meio da leitura. copy2 mantém o mtime, que decide qual
    # cópia é a mais nova.
    t = time.time()
    compartilhado = Path(os.getenv("HPG_DRIVE_CACHE") or (REPO / "_drive"))
    proprio = REPO / "_drive"
    with _trava(compartilhado):
        sync_drive.sincronizar(compartilhado, verboso=False)
        if compartilhado.resolve() != proprio.resolve():
            shutil.copytree(compartilhado, proprio, dirs_exist_ok=True,
                            copy_function=shutil.copy2,
                            ignore=shutil.ignore_patterns("*.baixando"))
    env["HPG_DRIVE_ROOT"] = str(proprio)
    n = sum(1 for p in proprio.rglob("*") if p.is_file())
    linhas.append(f"[drive] {n} arquivo(s) espelhado(s) do Drive ({time.time() - t:.0f} s)")
    if comite:
        linhas += _prepara_comite(env)
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
    preparo = (lambda: _prepara_nuvem(env, job)) if nuvem else None

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
