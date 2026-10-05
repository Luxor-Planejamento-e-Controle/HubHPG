"""Atualiza os animais do SGPG (tabela sgpg_animal) com o plantel do haras.

O prêmio registrado no SGPG aponta para um animal desta tabela, e a tela de registro
só oferece os animais que estão no plantel. A lista é o roster do controle mensal
mais recente (CONTROLE_DE_PLANTEL, aba PLANTEL), pela mesma regra da contagem do
fechamento semanal (PGSemanalReport._plantel_por_status): sem embrião e receptora,
uma linha por animal mesmo quando o controle repete por cotista.

Roda no fim de cada fechamento semanal (tools/pipelines_hub.py) e pode rodar à mão.
Quem some do roster fica na tabela com no_plantel=false: o prêmio antigo continua
apontando para ele. O nome mostrado (`nome`) nasce do nome do plantel e a
sincronização não o sobrescreve depois.

Falhar aqui não derruba o fechamento: o aviso sai no log do pedido e a lista do SGPG
fica como estava até a próxima semana.

    python tools/sync_sgpg_animais.py
"""
from __future__ import annotations

import re
import sys
import unicodedata
from pathlib import Path

import requests

REPO = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(REPO / "scripts"))
sys.path.insert(0, str(REPO / "tools"))

PARTICULAS = {"DA", "DE", "DO", "DAS", "DOS", "E", "DI"}


def chave(nome: str) -> str:
    """Identidade do animal: nome sem cotista, maiúsculo, sem acento, espaço simples."""
    t = unicodedata.normalize("NFD", nome).encode("ascii", "ignore").decode()
    return " ".join(t.upper().split())


def _palavra(p: str, primeira: bool) -> str:
    if not p:
        return p
    if any(c.isdigit() for c in p) or not re.search(r"[AEIOU]", p):
        return p.upper()                       # MH2, N19, PG, MH, X
    if p in PARTICULAS and not primeira:
        return p.lower()
    return p[0] + p[1:].lower()


def nome_de_exibicao(nome: str) -> str:
    """'NINA DO ATLANTICO SUL DA PAO GRANDE' -> 'Nina do Atlantico Sul da Pao Grande'.
    O controle não tem acento; o nome do SGPG pode ser corrigido depois."""
    return " ".join("-".join(_palavra(p, i == 0) for p in w.split("-"))
                    for i, w in enumerate(nome.split()))


def linhas_do_roster() -> tuple[list[dict], str]:
    from PGSemanalReport import _plantel_por_status
    r = _plantel_por_status()
    vistos, linhas = set(), []
    for x in r["linhas"]:
        k = chave(x["nome"])
        if k in vistos:        # mesmo nome, filiação diferente: fica o primeiro
            continue
        vistos.add(k)
        linhas.append({"chave": k, "nome_plantel": x["nome"], "nome": nome_de_exibicao(x["nome"]),
                       "categoria": x.get("categoria"), "mae": x.get("mae"), "pai": x.get("pai"),
                       "local": x.get("local"), "status_plantel": x.get("status_plantel")})
    return linhas, r["fonte"]


def main() -> int:
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    try:
        import publish_hub
        url, key = publish_hub.env()
        linhas, fonte = linhas_do_roster()
        if not linhas:
            print("[sgpg] roster vazio — lista de animais do SGPG não foi mexida")
            return 0
        resp = requests.post(f"{url}/rest/v1/rpc/sgpg_sincroniza_animais",
                             headers={"apikey": key, "Authorization": f"Bearer {key}",
                                      "Content-Type": "application/json"},
                             json={"linhas": linhas, "origem": fonte}, timeout=60)
        resp.raise_for_status()
        print(f"[sgpg] {len(linhas)} animais do plantel na lista do SGPG ({fonte})")
    except Exception as exc:
        print(f"[aviso] lista de animais do SGPG não atualizada: {exc!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
