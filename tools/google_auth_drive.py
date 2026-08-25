"""Gera, uma vez, o refresh token que a Function usa para ler as planilhas no Drive.

Roda LOCALMENTE, na máquina de quem já enxerga a pasta 'PLANILHAS DE CONTROLE' —
abre o navegador, pede consentimento e imprime o refresh token. Esse token é o que
vai para as app settings do Function App; a partir daí a Function lê o Drive com a
mesma visibilidade da pessoa que consentiu.

Por que OAuth e não service account: a pasta é um atalho para um Shared Drive de
terceiro. Compartilhar com uma service account depende de quem administra aquele
drive; um refresh token, não.

O token não expira sozinho, mas morre se a senha da conta mudar, se o consentimento
for revogado, ou se o app OAuth ficar em modo 'Testing' no Google Cloud (aí expira
em 7 dias — publique o app para evitar isso).

Uso:
    set HPG_GOOGLE_OAUTH_CLIENT_ID=...        # do OAuth client (tipo: Desktop app)
    set HPG_GOOGLE_OAUTH_CLIENT_SECRET=...
    python tools/google_auth_drive.py
"""
from __future__ import annotations

import os
import sys

ESCOPOS = ["https://www.googleapis.com/auth/drive.readonly"]


def main():
    cid = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_ID") or os.getenv("GOOGLE_OAUTH_CLIENT_ID")
    seg = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_SECRET") or os.getenv("GOOGLE_OAUTH_CLIENT_SECRET")
    if not cid or not seg:
        sys.exit("Defina HPG_GOOGLE_OAUTH_CLIENT_ID e HPG_GOOGLE_OAUTH_CLIENT_SECRET "
                 "antes de rodar (credenciais do OAuth client no Google Cloud).")

    from google_auth_oauthlib.flow import InstalledAppFlow

    cfg = {"installed": {
        "client_id": cid,
        "client_secret": seg,
        "auth_uri": "https://accounts.google.com/o/oauth2/auth",
        "token_uri": "https://oauth2.googleapis.com/token",
        "redirect_uris": ["http://localhost"],
    }}
    fluxo = InstalledAppFlow.from_client_config(cfg, ESCOPOS)
    # prompt='consent' + access_type='offline': sem os dois o Google devolve só o
    # access token (1h) quando a conta já consentiu antes, e o refresh vem vazio.
    cred = fluxo.run_local_server(port=0, access_type="offline", prompt="consent")

    print("\n=== copie para as app settings do Function App ===")
    print(f"HPG_GOOGLE_OAUTH_CLIENT_ID={cid}")
    print("HPG_GOOGLE_OAUTH_CLIENT_SECRET=<o mesmo que você usou aqui>")
    print(f"HPG_GOOGLE_OAUTH_REFRESH_TOKEN={cred.refresh_token}")

    # conferência imediata: o token serve para o que precisamos?
    from googleapiclient.discovery import build
    svc = build("drive", "v3", credentials=cred, cache_discovery=False)
    raiz = os.getenv("HPG_DRIVE_ROOT_ID", "1mBrSeztRwtBnMlkOMnq6aO4LQUkNjiTb")
    meta = svc.files().get(fileId=raiz, fields="id,name", supportsAllDrives=True).execute()
    print(f"\n[ok] enxerga a raiz: {meta['name']} ({meta['id']})")


if __name__ == "__main__":
    main()
