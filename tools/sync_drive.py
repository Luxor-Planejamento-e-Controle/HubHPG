"""Espelha as planilhas do Google Drive num diretório local, para o pipeline rodar
fora da máquina do Arthur.

O fechamento semanal lê tudo de `G:\\.shortcut-targets-by-id\\...`, que é o mount do
Google Drive for Desktop — não existe numa Azure Function. Em vez de reescrever os
seis resolvedores de fonte para falar API, este módulo baixa o subconjunto que o
pipeline abre e devolve um caminho: `HPG_DRIVE_ROOT` aponta para ele e o resto do
código continua abrindo arquivo por caminho, sem saber da diferença.

Duas coisas que NÃO podem se perder na cópia, porque o pipeline depende delas:

  - a estrutura de pastas ('PLANTEL/Estação 2026-2027/...'), porque a resolução de
    fonte varre as pastas de estação e escolhe a mais recente;
  - o mtime, porque a guarda de fonte velha compara a data do arquivo com a janela
    da semana e ABORTA se a planilha for anterior. Copiar com a data de hoje faria
    toda fonte parecer fresca — exatamente o silêncio que a guarda existe para
    quebrar. Aqui o mtime vem do `modifiedTime` do Drive.

Autenticação: OAuth com refresh token de uma conta que já enxerga a pasta
(HPG_GOOGLE_OAUTH_CLIENT_ID/_SECRET/_REFRESH_TOKEN — gere com
tools/google_auth_drive.py). Service account também serve, se um dia a pasta for
compartilhada com uma; sem acesso, a API devolve 404, não 403, então o erro parece
'arquivo não existe'.

Uso:
    python tools/sync_drive.py [destino]
    # ou, de dentro do pipeline:
    from tools.sync_drive import sincronizar; raiz = sincronizar()
"""
from __future__ import annotations

import io
import json
import os
import sys
from datetime import datetime, timezone
from pathlib import Path

# Atalho "PLANILHAS DE CONTROLE" — a raiz de tudo que o fechamento lê.
RAIZ_DRIVE_ID = os.getenv("HPG_DRIVE_ROOT_ID", "1mBrSeztRwtBnMlkOMnq6aO4LQUkNjiTb")

# Só estas pastas de primeiro nível. O drive inteiro tem muito mais coisa, e baixar
# tudo toda sexta seria pagar banda por arquivo que ninguém abre.
PASTAS = ("PLANTEL", "ATUALIZACAO SEMANAL")

# Só o que o pipeline sabe abrir. '~$' é lock de Excel aberto — o próprio pipeline
# já os ignora, mas não faz sentido baixar.
EXTENSOES = (".xlsx", ".docx")

MIME_PASTA = "application/vnd.google-apps.folder"


ESCOPOS = ["https://www.googleapis.com/auth/drive.readonly"]


def _credenciais():
    """OAuth primeiro, service account depois.

    OAuth (refresh token de uma conta que já enxerga a pasta) é o que funciona aqui:
    a pasta é um atalho para um Shared Drive de terceiro, e compartilhar com uma
    service account depende de quem administra aquele drive. Com refresh token, a
    API vê exatamente o que a pessoa vê.

    Service account fica como alternativa para o dia em que a pasta for compartilhada
    com uma — aí é uma credencial a menos para renovar."""
    cid = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_ID") or os.getenv("GOOGLE_OAUTH_CLIENT_ID")
    seg = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_SECRET") or os.getenv("GOOGLE_OAUTH_CLIENT_SECRET")
    ref = os.getenv("HPG_GOOGLE_OAUTH_REFRESH_TOKEN") or os.getenv("GOOGLE_OAUTH_REFRESH_TOKEN")
    if cid and seg and ref:
        from google.oauth2.credentials import Credentials
        return Credentials(token=None, refresh_token=ref,
                           token_uri="https://oauth2.googleapis.com/token",
                           client_id=cid, client_secret=seg, scopes=ESCOPOS)

    bruto = os.getenv("HPG_GOOGLE_SERVICE_ACCOUNT_JSON") or os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON")
    if bruto:
        from google.oauth2 import service_account
        return service_account.Credentials.from_service_account_info(
            json.loads(bruto), scopes=ESCOPOS)

    sys.exit("Sem credencial do Drive: defina HPG_GOOGLE_OAUTH_CLIENT_ID/"
             "_SECRET/_REFRESH_TOKEN (rode tools/google_auth_drive.py) ou "
             "HPG_GOOGLE_SERVICE_ACCOUNT_JSON.")


def _servico():
    from googleapiclient.discovery import build
    return build("drive", "v3", credentials=_credenciais(), cache_discovery=False)


def _listar(svc, pasta_id: str) -> list[dict]:
    """Filhos de uma pasta. supportsAllDrives/includeItemsFromAllDrives porque a
    origem é um Shared Drive, não um Meu Drive."""
    itens, token = [], None
    while True:
        resp = svc.files().list(
            q=f"'{pasta_id}' in parents and trashed = false",
            pageSize=1000, pageToken=token,
            fields="nextPageToken, files(id, name, mimeType, modifiedTime, size)",
            supportsAllDrives=True, includeItemsFromAllDrives=True,
        ).execute()
        itens += resp.get("files", [])
        token = resp.get("nextPageToken")
        if not token:
            return itens


def _mtime(iso: str) -> float:
    """'2026-08-21T14:03:11.000Z' -> epoch."""
    return datetime.fromisoformat(iso.replace("Z", "+00:00")).replace(
        tzinfo=timezone.utc).timestamp()


def _baixar(svc, arq: dict, destino: Path) -> bool:
    """Baixa se ainda não existe com o mesmo mtime. Devolve True se baixou."""
    from googleapiclient.http import MediaIoBaseDownload

    quando = _mtime(arq["modifiedTime"])
    if destino.exists() and abs(destino.stat().st_mtime - quando) < 1.0:
        return False
    destino.parent.mkdir(parents=True, exist_ok=True)
    buf = io.BytesIO()
    req = svc.files().get_media(fileId=arq["id"], supportsAllDrives=True)
    baixador = MediaIoBaseDownload(buf, req)
    concluido = False
    while not concluido:
        _, concluido = baixador.next_chunk()
    destino.write_bytes(buf.getvalue())
    os.utime(destino, (quando, quando))    # a guarda de fonte velha depende disto
    return True


def _andar(svc, pasta_id: str, destino: Path, baixados: list, pulados: list):
    for item in _listar(svc, pasta_id):
        alvo = destino / item["name"]
        if item["mimeType"] == MIME_PASTA:
            _andar(svc, item["id"], alvo, baixados, pulados)
            continue
        if item["name"].startswith("~$") or not item["name"].lower().endswith(EXTENSOES):
            continue
        if _baixar(svc, item, alvo):
            baixados.append(alvo)
        else:
            pulados.append(alvo)


def sincronizar(destino: Path | None = None, verboso: bool = True) -> Path:
    """Espelha as pastas necessárias e devolve a raiz local (equivalente a DRIVE_ROOT).

    Reaproveita o que já está no destino quando o mtime bate — numa Function o
    diretório sobrevive entre execuções da mesma instância, e uma sexta que só mudou
    duas planilhas não precisa baixar as outras cinco."""
    destino = Path(destino or os.getenv("HPG_DRIVE_CACHE")
                   or (Path(os.getenv("TMPDIR") or os.getenv("TEMP") or "/tmp") / "hpg-drive"))
    svc = _servico()
    filhos = {f["name"]: f for f in _listar(svc, RAIZ_DRIVE_ID)}
    faltando = [p for p in PASTAS if p not in filhos]
    if faltando:
        # não seguir com fonte pela metade: o pipeline concluiria "arquivo não existe"
        # e cairia numa cópia antiga, que é o modo de falhar em silêncio que ele evita
        sys.exit(f"Pastas ausentes na raiz do Drive ({RAIZ_DRIVE_ID}): {faltando}. "
                 f"A credencial enxerga: {sorted(filhos)}")
    baixados, pulados = [], []
    for nome in PASTAS:
        _andar(svc, filhos[nome]["id"], destino / nome, baixados, pulados)
    if verboso:
        print(f"[drive] {len(baixados)} arquivo(s) baixado(s), {len(pulados)} já em cache "
              f"-> {destino}")
        for f in baixados:
            print(f"  + {f.relative_to(destino)}")
    return destino


def main():
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else None
    raiz = sincronizar(destino)
    print(raiz)


if __name__ == "__main__":
    main()
