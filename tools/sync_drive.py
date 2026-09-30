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
import re
import threading
import unicodedata
import os
import sys
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path

# Atalho "PLANILHAS DE CONTROLE" — a raiz de tudo que o fechamento lê.
RAIZ_DRIVE_ID = os.getenv("HPG_DRIVE_ROOT_ID", "1mBrSeztRwtBnMlkOMnq6aO4LQUkNjiTb")

# Só estas pastas de primeiro nível. O drive inteiro tem muito mais coisa, e baixar
# tudo toda sexta seria pagar banda por arquivo que ninguém abre.
#
# REPRODUÇÃO e VENDAS faltavam (lista conferida em 25/09/2026 contra os caminhos que
# o PGSemanalReport de fato abre): a estação de monta — produção, confirmados,
# acumulado — e o EMBRIOES A ENTREGAR moram em REPRODUÇÃO; o mapa de vendas, em
# VENDAS. Nunca apareceu porque a Function morria antes, na credencial do Drive; com
# ela resolvida, o fechamento na nuvem sairia sem produção.
PASTAS = ("PLANTEL", "ATUALIZACAO SEMANAL", "REPRODUÇÃO", "VENDAS")

# E, dentro delas, só o que o pipeline abre. As quatro pastas inteiras somam ~5,7 GB
# (medido em 30/09/2026): PLANEJAMENTO ESTAÇÃO DE MONTA tem 2 GB, MAPAS DE VENDAS
# 2,7 GB — um mapa de 25 MB por semana desde 2020 — e o PLANTEL guarda seis estações.
# Espelhar tudo fazia cada clique do hub esperar meia hora por arquivo que o código
# nunca lê. Regra por pasta:
#   - os arquivos soltos do nível sempre descem (são as planilhas de trabalho);
#   - "estacoes": das subpastas "Estação AAAA-AAAA", só as N mais novas. Os
#     resolvedores varrem todas e ficam com o arquivo mais novo; a anterior entra
#     porque a virada de estação muda a cópia de trabalho de pasta e porque os meses
#     do ano (contagem jan–ago, o controle de cada mês) ficam nela;
#   - "dias": dentro dessas estações, só o modificado nos últimos N dias (e sempre
#     o mais novo) — o mapa é cumulativo, e o comitê só abre o do fechamento do mês;
#   - "sub": subpastas nomeadas que descem, com a regra delas. Subpasta que não
#     está aqui não desce.
ALVOS = {
    "PLANTEL": {"estacoes": 2},
    "ATUALIZACAO SEMANAL": {},
    "REPRODUÇÃO": {"sub": {"ESTAÇÃO DE MONTA": {"estacoes": 2}}},
    "VENDAS": {"sub": {"MAPAS DE VENDAS": {"estacoes": 2, "dias": 70},
                       "SAIDA DE ANIMAIS VENDIDOS": {}}},
}

# Só o que o pipeline sabe abrir. '~$' é lock de Excel aberto — o próprio pipeline
# já os ignora, mas não faz sentido baixar.
EXTENSOES = (".xlsx", ".docx")

MIME_PASTA = "application/vnd.google-apps.folder"


ESCOPOS = ["https://www.googleapis.com/auth/drive.readonly"]


def _credenciais():
    """A conta de serviço do HPG; OAuth de reserva.

    Desde 30/09/2026 a pasta está compartilhada, como leitor, com uma conta de
    serviço do próprio HPG (HPG_GOOGLE_SERVICE_ACCOUNT_JSON). OAuth (refresh token de
    quem enxerga a pasta, gerado com tools/google_auth_drive.py) fica para o dia em
    que o compartilhamento não for possível.

    Só nomes HPG_*. Até 30/09 caía também em GOOGLE_SERVICE_ACCOUNT_JSON, e no
    Function App esse nome é o da conta das automações da Controladoria: a sexta na
    nuvem usava a credencial de outro time, levava 404 na pasta do haras e morria sem
    dizer por quê."""
    cid = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_ID")
    seg = os.getenv("HPG_GOOGLE_OAUTH_CLIENT_SECRET")
    ref = os.getenv("HPG_GOOGLE_OAUTH_REFRESH_TOKEN")
    bruto = os.getenv("HPG_GOOGLE_SERVICE_ACCOUNT_JSON")
    if bruto:
        from google.oauth2 import service_account
        return service_account.Credentials.from_service_account_info(
            json.loads(bruto), scopes=ESCOPOS)
    if cid and seg and ref:
        from google.oauth2.credentials import Credentials
        return Credentials(token=None, refresh_token=ref,
                           token_uri="https://oauth2.googleapis.com/token",
                           client_id=cid, client_secret=seg, scopes=ESCOPOS)
    # RuntimeError, não sys.exit: numa Function o SystemExit derruba o worker e o log
    # só mostra "The client reset the request stream", sem este texto
    raise RuntimeError("Sem credencial do Drive: defina HPG_GOOGLE_SERVICE_ACCOUNT_JSON "
                       "ou HPG_GOOGLE_OAUTH_CLIENT_ID/_SECRET/_REFRESH_TOKEN "
                       "(tools/google_auth_drive.py).")


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
    # grava ao lado e troca de uma vez: o espelho da nuvem é reaproveitado entre
    # pedidos, e ninguém pode abrir uma planilha pela metade
    tmp = destino.with_name(destino.name + ".baixando")
    tmp.write_bytes(buf.getvalue())
    os.utime(tmp, (quando, quando))        # a guarda de fonte velha depende disto
    os.replace(tmp, destino)
    return True


# NFC dos dois lados: 'REPRODUÇÃO' pode vir da API decomposto (C + cedilha
# combinante) e aí não bateria com o literal do código — nem na busca das pastas
# nem no glob("Estação *") dos resolvedores, que compara byte a byte no Linux.
_nfc = lambda t: unicodedata.normalize("NFC", t)
RE_ESTACAO = re.compile(r"^estação \d{4}", re.IGNORECASE)


def _andar(svc, pasta_id: str, destino: Path, regra: dict, plano: list):
    """Desce uma pasta segundo a regra dela em ALVOS (ver lá) e anota em `plano` o
    que tem de estar no espelho — o download vem depois, em paralelo."""
    itens = _listar(svc, pasta_id)
    arquivos = [i for i in itens if i["mimeType"] != MIME_PASTA
                and not i["name"].startswith("~$") and i["name"].lower().endswith(EXTENSOES)]
    pastas = [i for i in itens if i["mimeType"] == MIME_PASTA]
    dias = regra.get("dias")
    if dias and arquivos:
        corte = datetime.now(timezone.utc).timestamp() - dias * 86400
        mais_novo = max(arquivos, key=lambda a: a["modifiedTime"])
        arquivos = [a for a in arquivos if a is mais_novo or _mtime(a["modifiedTime"]) >= corte]
    for arq in arquivos:
        plano.append((arq, destino / _nfc(arq["name"])))
    n = regra.get("estacoes")
    estacoes = sorted((p for p in pastas if RE_ESTACAO.match(_nfc(p["name"]))),
                      key=lambda p: _nfc(p["name"]), reverse=True)[:n] if n else []
    for p in estacoes:
        _andar(svc, p["id"], destino / _nfc(p["name"]), {"dias": dias} if dias else {}, plano)
    subs = regra.get("sub", {})
    for p in pastas:
        if _nfc(p["name"]) in subs:
            _andar(svc, p["id"], destino / _nfc(p["name"]), subs[_nfc(p["name"])], plano)


def sincronizar(destino: Path | None = None, verboso: bool = True) -> Path:
    """Espelha as pastas necessárias e devolve a raiz local (equivalente a DRIVE_ROOT).

    Reaproveita o que já está no destino quando o mtime bate — numa Function o
    diretório sobrevive entre execuções da mesma instância, e uma sexta que só mudou
    duas planilhas não precisa baixar as outras cinco."""
    destino = Path(destino or os.getenv("HPG_DRIVE_CACHE")
                   or (Path(os.getenv("TMPDIR") or os.getenv("TEMP") or "/tmp") / "hpg-drive"))
    svc = _servico()
    filhos = {_nfc(f["name"]): f for f in _listar(svc, RAIZ_DRIVE_ID)}
    faltando = [p for p in PASTAS if _nfc(p) not in filhos]
    if faltando:
        # não seguir com fonte pela metade: o pipeline concluiria "arquivo não existe"
        # e cairia numa cópia antiga, que é o modo de falhar em silêncio que ele evita.
        # Sem acesso à raiz a API devolve lista vazia, então é aqui que cai credencial
        # sem compartilhamento.
        raise RuntimeError(f"Pastas ausentes na raiz do Drive ({RAIZ_DRIVE_ID}): {faltando}. "
                           f"A credencial enxerga: {sorted(filhos)}")
    plano = []
    for nome in PASTAS:
        _andar(svc, filhos[_nfc(nome)]["id"], destino / nome, ALVOS[nome], plano)
    vistos = {alvo for _, alvo in plano}
    # Em paralelo: na primeira execução da Azure (30/09/2026) o espelho em série levou
    # uns 4 min dos 4,5 do pedido, quase tudo nos três mapas de vendas de 25 MB. O
    # cliente da API não é thread-safe, então cada thread abre o seu.
    local = threading.local()

    def baixa(par):
        if not hasattr(local, "svc"):
            local.svc = _servico()
        arq, alvo = par
        return alvo if _baixar(local.svc, arq, alvo) else None

    with ThreadPoolExecutor(max_workers=4) as ex:
        baixados = [a for a in ex.map(baixa, plano) if a is not None]
    # espelho: o que saiu do Drive (renomeado, apagado, estação que ficou velha) sai
    # daqui também — senão um resolvedor por mtime podia ficar com a cópia morta
    removidos = [f for f in destino.rglob("*") if f.is_file() and f not in vistos]
    for f in removidos:
        f.unlink()
    for d in sorted((d for d in destino.rglob("*") if d.is_dir()), reverse=True):
        if not any(d.iterdir()):
            d.rmdir()
    if verboso:
        print(f"[drive] {len(baixados)} arquivo(s) baixado(s), "
              f"{len(vistos) - len(baixados)} já em cache, {len(removidos)} removido(s) "
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
