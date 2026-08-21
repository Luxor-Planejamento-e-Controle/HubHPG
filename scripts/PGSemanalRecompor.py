"""PGSemanalRecompor — reaplica as regras ATUAIS sobre as semanas já congeladas.

Quando uma regra de contagem muda, as semanas anteriores ficam com o número velho e o
dashboard passa a mostrar séries que se contradizem. Re-rodar o fechamento daquela
semana NÃO resolve: `PGSemanal.py 14/08` lê as planilhas de HOJE, e headcount,
receptoras e pendentes são retrato do momento, não da janela — em 21/08/2026 isso
trocaria o headcount de 203 (correto para 14/08) pelo de agora.

A saída é o próprio snapshot: ele guarda as LINHAS (`detalhe.terceiros_vendidos`,
`detalhe.terceiros_sociedade`), não só os totais. Então as regras são reaplicadas em
cima do que foi congelado, sem abrir planilha nenhuma.

Regras reaplicadas aqui — as duas que se perderam na migração para o STATUS PLANTEL:
  - REPOSIÇÃO fora dos vendidos pendentes (sai para repor outro animal, não vendido)
  - pendência documental ('FALTANDO EXAME') fora da sociedade pendente

Uso:
    python scripts/PGSemanalRecompor.py            # mostra o que mudaria
    python scripts/PGSemanalRecompor.py --aplicar  # grava
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
import PGSemanalReport as R


def _reposicoes() -> set:
    """Núcleos de nome marcados REPOSICAO no 'Animais para sair'."""
    try:
        wb = R._load(R._latest_animais_sair())
    except FileNotFoundError:
        print("  ! 'Animais para sair' ausente — sem como reconhecer reposição")
        return set()
    out = set()
    for i, r in enumerate(wb["ANIMAIS VENDIDOS"].iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        if R._norm(r[6]) == "REPOSICAO":
            out.add(R._nucleo_nome(r[1]))
    wb.close()
    return out


def _recompoe(snap: dict, reposicoes: set) -> dict | None:
    """Novos totais de pendentes para um snapshot. None = nada a mudar."""
    det = snap.get("detalhe") or {}
    vend, soc = det.get("terceiros_vendidos"), det.get("terceiros_sociedade")
    if vend is None or soc is None:
        return None
    # embrião guardado com 'Pronto - Nasce na PG' não é pendência de saída: nasce aqui.
    # A regra é posterior a esses snapshots, então a linha ficou lá dentro.
    def _fica(x):
        if x.get("especie") != "EMBRIAO":
            return True
        return R.EMB_STATUS_PENDENTE in R._norm(x.get("obs"))
    vend_novo = [x for x in vend
                 if R._nucleo_nome(x.get("nome")) not in reposicoes and _fica(x)]
    soc_novo = [x for x in soc
                if not any(b in R._norm(x.get("obs")) for b in R.OBS_BLOQUEIA_SAIDA)
                and _fica(x)]
    emb = lambda L: [x for x in L if x.get("especie") == "EMBRIAO"]
    ani = lambda L: [x for x in L if x.get("especie") != "EMBRIAO"]
    novo = {
        "vendidos_pendentes": len(vend_novo),
        "vendidos_pendentes_animais": len(ani(vend_novo)),
        "vendidos_pendentes_embrioes": len(emb(vend_novo)),
        "sociedade_pendentes": len(soc_novo),
        "sociedade_pendentes_animais": len(ani(soc_novo)),
        "sociedade_pendentes_embrioes": len(emb(soc_novo)),
        "terceiros_propriedade": len(vend_novo),
        "terceiros_animais": len(ani(vend_novo)),
        "terceiros_embrioes": len(emb(vend_novo)),
    }
    atual = snap.get("terceiros") or {}
    if all(atual.get(k) == v for k, v in novo.items()):
        return None
    return {"terceiros": novo, "vend": vend_novo, "soc": soc_novo}


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    aplicar = "--aplicar" in sys.argv[1:]
    hist = R._load_hist()
    if not hist:
        sys.exit("Nenhum snapshot em _cache/semanal_snapshots.json")
    reposicoes = _reposicoes()
    print(f"reposições reconhecidas: {sorted(reposicoes) or '(nenhuma)'}\n")

    mudou = 0
    for wid in sorted(hist):
        res = _recompoe(hist[wid], reposicoes)
        if res is None:
            print(f"  {wid}  sem mudança")
            continue
        mudou += 1
        antes, depois = hist[wid].get("terceiros") or {}, res["terceiros"]
        print(f"  {wid}  RECOMPOSTO")
        for k, v in depois.items():
            if antes.get(k) != v:
                print(f"      {k:32} {antes.get(k)} -> {v}")
        if aplicar:
            # marca o que foi reaplicado: o recompositor só REMOVE sob regra mais
            # estrita. Embrião de venda que passou a contar depois daquela semana nunca
            # foi coletado e não há de onde tirar — a semana fica assim.
            hist[wid]["regras_reaplicadas"] = ["reposicao_fora_de_vendidos",
                                               "bloqueio_documental_fora_de_sociedade",
                                               "embriao_so_aguardando_entrega"]
            hist[wid]["terceiros"] = {**(hist[wid].get("terceiros") or {}), **depois}
            hist[wid]["detalhe"]["terceiros_vendidos"] = res["vend"]
            hist[wid]["detalhe"]["terceiros_sociedade"] = res["soc"]
            hist[wid]["detalhe"]["pendentes_saida"] = res["vend"] + res["soc"]

    if not aplicar:
        print(f"\n{mudou} semana(s) mudariam. Rode com --aplicar para gravar.")
        return
    R.HIST_SNAPSHOTS.write_text(json.dumps(hist, ensure_ascii=False, indent=2),
                                encoding="utf-8")
    print(f"\n{mudou} semana(s) recompostas -> {R.HIST_SNAPSHOTS.name}")


if __name__ == "__main__":
    main()
