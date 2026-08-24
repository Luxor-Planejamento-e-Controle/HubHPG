"""Aplica na `main` a mesma proteção do LuxorP&CHub.

Não roda enquanto o repo for PRIVADO: branch protection e ruleset exigem GitHub
Pro/Team em repo privado — a API devolve 403 pedindo upgrade. Ordem obrigatória:

    1. git push --force-with-lease origin main   (história já purgada)
    2. tornar o repo público
    3. python tools/proteger_main.py

O bypass de admin é DE PROPÓSITO (mesma decisão do P&C): `enforce_admins:false`
na proteção clássica e `RepositoryRole 5` no ruleset. Consequência assumida —
push direto do admin não passa por PR nem pela Action `guarda`, então os hooks
locais (`python tools/install_hooks.py`) são a única barreira nesse caminho.

Uso: python tools/proteger_main.py [--repo OWNER/NOME]
"""
import json
import subprocess
import time
import sys

REPO = "Luxor-Planejamento-e-Controle/HubHPG"

PROTECAO = {
    "required_status_checks": {"strict": False, "contexts": ["guarda"]},
    "enforce_admins": False,
    "required_pull_request_reviews": {
        "dismiss_stale_reviews": True,
        "require_code_owner_reviews": True,
        "required_approving_review_count": 1,
        "require_last_push_approval": False,
    },
    "restrictions": None,
    "required_linear_history": False,
    "allow_force_pushes": False,
    "allow_deletions": False,
    "block_creations": False,
    "required_conversation_resolution": False,
}

RULESET = {
    "name": "protege main",
    "target": "branch",
    "enforcement": "active",
    "bypass_actors": [{"actor_id": 5, "actor_type": "RepositoryRole", "bypass_mode": "always"}],
    "conditions": {"ref_name": {"include": ["~DEFAULT_BRANCH"], "exclude": []}},
    "rules": [
        {"type": "deletion"},
        {"type": "non_fast_forward"},
        {"type": "pull_request", "parameters": {
            "required_approving_review_count": 1,
            "dismiss_stale_reviews_on_push": True,
            "require_code_owner_review": False,
            "require_last_push_approval": False,
            "required_review_thread_resolution": False,
            "allowed_merge_methods": ["merge", "squash"],
        }},
    ],
}


# Trocar a visibilidade deixa o repo LOCKED por alguns minutos enquanto o GitHub
# reindexa, e a API responde 403 "Repository has been locked". Nao e erro de
# configuracao — e so cedo. Sem retry, quem roda logo depois de tornar publico ve
# um 403 cru e precisa lembrar de rodar de novo.
TENTATIVAS = 6
ESPERA_S = 20


def gh(metodo, caminho, corpo=None, tentar_de_novo=True):
    cmd = ["gh", "api", "-X", metodo, caminho]
    entrada = None
    if corpo is not None:
        cmd += ["--input", "-"]
        entrada = json.dumps(corpo)
    for n in range(1, TENTATIVAS + 1):
        r = subprocess.run(cmd, input=entrada, capture_output=True, text=True)
        saida = (r.stdout or "") + (r.stderr or "")
        if r.returncode == 0:
            return True, saida
        if not (tentar_de_novo and "has been locked" in saida and n < TENTATIVAS):
            return False, saida
        print(f"  repo travado pelo GitHub (troca de visibilidade recente); "
              f"tentativa {n}/{TENTATIVAS}, esperando {ESPERA_S}s...")
        time.sleep(ESPERA_S)
    return False, saida


def main():
    repo = REPO
    if "--repo" in sys.argv:
        repo = sys.argv[sys.argv.index("--repo") + 1]

    ok, saida = gh("GET", f"repos/{repo}")
    if not ok:
        sys.exit(f"Não consegui ler o repo {repo}: {saida[:200]}")
    if json.loads(saida).get("private"):
        sys.exit("Repo ainda é PRIVADO. Branch protection e ruleset exigem GitHub Pro "
                 "nesse caso — a API devolve 403. Torne público primeiro (e só depois "
                 "de conferir que a história purgada já foi empurrada).")

    ok, saida = gh("PUT", f"repos/{repo}/branches/main/protection", PROTECAO)
    print(("[ok] " if ok else "[erro] ") + "branch protection" + ("" if ok else f": {saida[:200]}"))

    ok2, saida2 = gh("GET", f"repos/{repo}/rulesets")
    existe = ok2 and any(r.get("name") == RULESET["name"] for r in json.loads(saida2 or "[]"))
    if existe:
        print("[skip] ruleset 'protege main' já existe")
    else:
        ok3, saida3 = gh("POST", f"repos/{repo}/rulesets", RULESET)
        print(("[ok] " if ok3 else "[erro] ") + "ruleset" + ("" if ok3 else f": {saida3[:200]}"))


if __name__ == "__main__":
    main()
