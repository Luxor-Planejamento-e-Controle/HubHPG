"""Migração ÚNICA: _docs/comite_conteudo.json -> tabela comite_conteudo no Supabase.

Sobe as fotos pro bucket privado comite-fotos (mesma redução que build_comite.py
já fazia pro embutir no spec.json: lado máx 760px, JPEG qualidade 68) e reescreve
"arquivos" pra apontar pro path no bucket em vez do caminho local.

Depois de rodar uma vez, comite_conteudo.json vira só backup/seed inicial —
edição de verdade passa a ser pelo hub. Idempotente: roda de novo sem duplicar
(upsert por mês; upload já teria sobrescrito o mesmo path no bucket).

Uso: python tools/migrar_comite_conteudo.py
"""
from __future__ import annotations

import io
import json
import sys
from pathlib import Path

import requests
from dotenv import dotenv_values
from PIL import Image

ROOT = Path(__file__).resolve().parent.parent
CONTEUDO = ROOT / "_docs" / "comite_conteudo.json"
FOTOS_DIR = ROOT / "assets" / "comite"
FOTO_LADO_MAX = 760
FOTO_QUALIDADE = 68
MESES_MIGRAR = ("2026-06", "2026-07")


def env():
    cfg = dotenv_values(ROOT / ".env")
    url = cfg.get("SUPABASE_URL", "").rstrip("/")
    key = cfg.get("SUPABASE_SERVICE_ROLE_KEY")
    if not url or not key:
        sys.exit("Faltam SUPABASE_URL / SUPABASE_SERVICE_ROLE_KEY no .env")
    return url, key


def _reduz(caminho: Path) -> bytes | None:
    try:
        im = Image.open(caminho)
        im.thumbnail((FOTO_LADO_MAX, FOTO_LADO_MAX))
        if im.mode != "RGB":
            im = im.convert("RGB")
        buf = io.BytesIO()
        im.save(buf, "JPEG", quality=FOTO_QUALIDADE, optimize=True)
        return buf.getvalue()
    except Exception as exc:
        print(f"  [foto] {caminho.name}: {exc!r} — pulando")
        return None


def _upload(url, key, mes, nome_arquivo, dados):
    """path no bucket: <mes>/<nome>. Sobrescreve se já existir (upsert)."""
    path = f"{mes}/{nome_arquivo}"
    r = requests.post(
        f"{url}/storage/v1/object/comite-fotos/{path}",
        headers={
            "apikey": key, "Authorization": f"Bearer {key}",
            "Content-Type": "image/jpeg", "x-upsert": "true",
        },
        data=dados,
    )
    if r.status_code not in (200, 201):
        print(f"  [upload] {path} falhou: {r.status_code} {r.text[:200]}")
        return None
    return path


def migra_fotos(url, key, mes, fotos):
    """Aceita os dois formatos que existem hoje: lista de caminhos (legado, jun/26)
    ou lista de {tema, arquivos} (jul/26+). Devolve no MESMO formato de entrada,
    só com os paths trocados pro bucket."""
    if not fotos:
        return fotos
    if isinstance(fotos[0], str):
        novos = []
        for f in fotos:
            caminho = FOTOS_DIR / f
            if not caminho.exists():
                print(f"  [foto] {caminho} não existe — pulando")
                continue
            dados = _reduz(caminho)
            if dados is None:
                continue
            path = _upload(url, key, mes, caminho.name, dados)
            if path:
                novos.append(path)
        return novos
    novos = []
    for grupo in fotos:
        arquivos_novos = []
        for f in grupo.get("arquivos") or []:
            caminho = FOTOS_DIR / f
            if not caminho.exists():
                print(f"  [foto] {caminho} não existe — pulando")
                continue
            dados = _reduz(caminho)
            if dados is None:
                continue
            path = _upload(url, key, mes, caminho.name, dados)
            if path:
                arquivos_novos.append(path)
        novos.append({"tema": grupo.get("tema", ""), "arquivos": arquivos_novos})
    return novos


def main():
    url, key = env()
    todos = json.loads(CONTEUDO.read_text(encoding="utf-8"))
    for mes in MESES_MIGRAR:
        c = todos.get(mes)
        if not c:
            print(f"[{mes}] não existe no JSON, pulando")
            continue
        print(f"[{mes}] migrando...")
        fotos_novas = migra_fotos(url, key, mes, c.get("fotos") or [])
        n_fotos = (sum(len(g["arquivos"]) for g in fotos_novas)
                   if fotos_novas and isinstance(fotos_novas[0], dict)
                   else len(fotos_novas))
        print(f"  [foto] {n_fotos} arquivo(s) no bucket comite-fotos/{mes}/")

        linha = {
            "mes": mes,
            "comentarios": c.get("comentarios") or [],
            "exposicoes": c.get("exposicoes") or {},
            "manejo": c.get("manejo") or [],
            "fotos": fotos_novas,
        }
        r = requests.post(
            f"{url}/rest/v1/comite_conteudo",
            headers={
                "apikey": key, "Authorization": f"Bearer {key}",
                "Content-Type": "application/json", "Prefer": "resolution=merge-duplicates",
            },
            json=[linha],
        )
        if r.status_code in (200, 201):
            print(f"  [ok] {mes} gravado em comite_conteudo")
        else:
            print(f"  [erro] {mes}: {r.status_code} {r.text[:300]}")


if __name__ == "__main__":
    main()
