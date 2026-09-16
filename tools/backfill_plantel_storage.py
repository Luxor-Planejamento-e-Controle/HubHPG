"""Copia os snapshots que já estão na tabela `plantel_snapshot` para arquivos no
bucket privado — um por mês, `plantel.<AAAA-MM>.json`.

Roda UMA vez, na virada do modelo de leitura da aba Plantel (16/09/2026). Depois
disso quem grava o arquivo é a própria tela, ao importar o mês.

Por que existe: a aba passou a ler objeto estático por URL assinada, que é o
caminho rápido do hub (o mesmo do semanal, do comitê e do indicadores no P&C).
Mês sem arquivo continua funcionando — a tela cai pra tabela —, mas pela rota
lenta. Este script tira todo mundo da rota lenta de uma vez.

O nome usa PONTO como separador de propósito: a policy de leitura do bucket é
`hub_can(split_part(name, '.', 1))`, então `plantel.2026-08.json` continua
casando com `hub_can('plantel')`.

Requer a service_role key (ignora RLS) no .env da raiz — NUNCA versionar:

    SUPABASE_URL=https://xxxx.supabase.co
    SUPABASE_SERVICE_ROLE_KEY=eyJ...

Uso:
    python tools/backfill_plantel_storage.py            # todos os meses
    python tools/backfill_plantel_storage.py 2026-08    # só um mês
    python tools/backfill_plantel_storage.py --dry-run  # só lista o que faria
"""
import json
import sys
from pathlib import Path

import requests
from dotenv import dotenv_values

ROOT = Path(__file__).resolve().parent.parent
BUCKET = "hpg-data"
TABELA = "plantel_snapshot"


def env():
    cfg = dotenv_values(ROOT / ".env")
    url = (cfg.get("SUPABASE_URL") or "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env da raiz.")
    return url, key


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    dry = "--dry-run" in sys.argv
    url, key = env()
    h = {"apikey": key, "Authorization": f"Bearer {key}"}

    # só as colunas que o arquivo carrega: `atribuicao` morreu com a atribuição
    # por mapa (o dono sai do sufixo, ver donoDaLinha no plantel.js)
    q = f"{url}/rest/v1/{TABELA}?select=mes,arquivo,linhas,log&order=mes"
    if args:
        q += f"&mes=in.({','.join(args)})"
    r = requests.get(q, headers=h, timeout=120)
    r.raise_for_status()
    linhas = r.json()
    if not linhas:
        print("Nenhum snapshot na tabela para copiar.")
        return 0

    for row in linhas:
        mes = row["mes"]
        nome = f"plantel.{mes}.json"
        corpo = json.dumps(row, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
        n = len((row.get("linhas") or {}).get("rows") or [])
        if dry:
            print(f"  [dry-run] {nome}: {n} linhas, {len(corpo)/1024:.0f} KB")
            continue
        # x-upsert: reimportar o mesmo mês sobrescreve em vez de dar 409
        up = requests.post(
            f"{url}/storage/v1/object/{BUCKET}/{nome}",
            headers={**h, "Content-Type": "application/json", "x-upsert": "true"},
            data=corpo, timeout=180)
        if up.status_code >= 300:
            print(f"  ✗ {nome}: {up.status_code} {up.text[:160]}")
            continue
        print(f"  ✓ {nome}: {n} linhas, {len(corpo)/1024:.0f} KB")

    print(f"\n{len(linhas)} mês(es) processado(s) em {BUCKET}.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
