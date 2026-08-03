"""Monta a especificação do Comitê Mensal HPG a partir das bases reais.

O deck é a saída, não a fonte: cada slide vira um objeto {t: <tipo>, ...} num
spec JSON, e daí saem as DUAS renderizações — o HTML (hub/comite.html) e o PPTX
(botão "Exportar PPTX", via pptxgen no navegador). Fonte única, dois formatos.

Mapa completo de slide × fonte: `_docs/COMITE_MAPEAMENTO.md`.

O que este build já preenche com dado real:
  S04/S07  DRE Haras competência, mês e YTD   -> DRE 2026 HPG - HARAS.xlsx
  S09      Investimentos em animais/produtos  -> mesma pasta, aba Investimentos
  S10      Haras caixa                        -> aba Real x Orçado (Caixa)
  S11      Estoque em equinos                 -> bases/base_bi.parquet
  S12      Movimentação do plantel            -> PlantelHPG/mov_cascata.parquet
  S13/S14  Casa/FPG, mês e YTD                -> DRE 2026 FPG - CASA.xlsx

O resto entra como slide `pendente`, que diz na cara qual é a fonte e por que
ainda não tem dado. Slide sem fonte NÃO recebe número inventado.

Uso:
    python hub/tools/build_comite.py            # mês que a planilha do DRE está
    python hub/tools/build_comite.py 06/2026    # mês específico
"""
import json
import sys
from datetime import date
from pathlib import Path

import numpy as np
import pandas as pd

HUB = Path(__file__).resolve().parent.parent
REPO = HUB.parent
OUT = HUB / "assets" / "comite"

DRE_DIR = Path(r"G:/Drives compartilhados/Luxor Controladoria/Ambiente de testes/DRE Data")
PLANTEL_DIR = Path(r"C:/Users/Arthur/repos/LuxorMonthlyP-CRoutines/PlantelHPG")
BASE_BI = REPO / "bases" / "base_bi.parquet"

MESES = ["Janeiro", "Fevereiro", "Março", "Abril", "Maio", "Junho",
         "Julho", "Agosto", "Setembro", "Outubro", "Novembro", "Dezembro"]
MES_ABR = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
           "Jul", "Ago", "Set", "Out", "Nov", "Dez"]

avisos = []


def aviso(msg):
    avisos.append(msg)
    print(f"  [aviso] {msg}")


# ------------------------------------------------------------------ helpers
def _json_default(o):
    if isinstance(o, np.generic):
        return o.item()
    raise TypeError(f"tipo não serializável: {type(o).__name__}")


def num(v):
    """Célula -> float ou None. Texto ('N/A', '-') vira None."""
    if v is None or isinstance(v, str):
        return None
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return None if pd.isna(f) else f


def pendente(n, titulo, sub, fonte, motivo):
    return {"t": "pendente", "n": n, "titulo": titulo, "sub": sub,
            "fonte": fonte, "motivo": motivo}


# ------------------------------------------------- leitor das abas Real x Orçado
# As três abas (Haras competência, Haras caixa, Casa) têm a MESMA forma:
#   r1: B = rótulo do mês ("Fevereiro 2026"), F = rótulo do YTD
#   r2: B..E = Orçado/Realizado/∆k/∆% do MÊS · F..I = idem do YTD
#   r4+: A = natureza. Negrito em A marca total; sem negrito é filho.
def le_real_x_orcado(caminho: Path, aba: str):
    import openpyxl
    wb = openpyxl.load_workbook(caminho, data_only=True)   # com estilo: o negrito é a hierarquia
    if aba not in wb.sheetnames:
        wb.close()
        raise KeyError(f"aba '{aba}' não existe em {caminho.name}: {wb.sheetnames}")
    ws = wb[aba]
    rot_mes = str(ws.cell(1, 2).value or "").strip()
    linhas = []
    for i in range(4, ws.max_row + 1):
        c = ws.cell(i, 1)
        nome = str(c.value).strip() if c.value is not None else ""
        if not nome:
            continue
        linhas.append({
            "nome": nome,
            "total": bool(c.font.b),
            "mes": [num(ws.cell(i, j).value) for j in (2, 3, 4, 5)],
            "ytd": [num(ws.cell(i, j).value) for j in (6, 7, 8, 9)],
        })
    wb.close()
    return rot_mes, linhas


def slide_dre(n, titulo, sub, linhas, campo):
    return {"t": "dre", "n": n, "titulo": titulo, "sub": sub,
            "linhas": [{"nome": l["nome"], "total": l["total"], "v": l[campo]} for l in linhas]}


# --------------------------------------------------------------- Seção 01
def secao_financeiro(mes_ref, spec):
    haras = DRE_DIR / "DRE 2026 HPG - HARAS.xlsx"
    casa = DRE_DIR / "DRE 2026 FPG - CASA.xlsx"
    rotulo = f"{MESES[mes_ref.month - 1]}/{str(mes_ref.year)[2:]}"

    if haras.exists():
        rot, comp = le_real_x_orcado(haras, "Real x Orçado (Comp)")
        if rot and rot.split()[0].lower() != MESES[mes_ref.month - 1].lower():
            aviso(f"a planilha do DRE Haras está em '{rot}', não em {rotulo} — "
                  f"S04/S07 saem com o mês da planilha")
        spec.append(slide_dre(4, f"RESUMO FINANCEIRO — HARAS COMPETÊNCIA — ORÇADO X REALIZADO {rot.upper()}",
                              f"DRE 2026 | HPG · Competência mensal · Fonte: aba Real x Orçado (Comp)",
                              comp, "mes"))
        spec.append(pendente(5, f"ANÁLISE DE CUSTOS — {rot.upper()}",
                             "DRE Haras · Custos indiretos de produção",
                             "DRE 2026 HPG - HARAS.xlsx, aba DRE-Compet (col 30=Orçado, 31=Realizado)",
                             "abertura por natureza de custo ainda não extraída"))
        spec.append(pendente(6, f"ANÁLISE DE DESPESAS — {rot.upper()}",
                             "DRE Haras · Despesas do mês",
                             "DRE 2026 HPG - HARAS.xlsx, aba DRE-Compet",
                             "abertura por natureza de despesa ainda não extraída"))
        spec.append(slide_dre(7, f"HARAS COMPETÊNCIA — ACUMULADO {mes_ref.year} (YTD)",
                              "DRE 2026 | HPG · Acumulado no ano · Fonte: aba Real x Orçado (Comp)",
                              comp, "ytd"))
        spec.append(pendente(8, "COMENTÁRIOS — VARIAÇÕES YTD",
                             "Análise mês a mês por categoria",
                             "COMENTARIOS_DRE_HARAS.docx",
                             "texto escrito por pessoa — não sai de base"))
        spec.append(slide_investimentos(haras, mes_ref))
        rot_cx, caixa = le_real_x_orcado(haras, "Real x Orçado (Caixa)")
        spec.append(slide_dre(10, f"HARAS CAIXA — ORÇADO X REALIZADO {rot_cx.upper()}",
                              "FC 2026 | HPG · Caixa mensal · Fonte: aba Real x Orçado (Caixa)",
                              caixa, "mes"))
    else:
        aviso(f"DRE do Haras não encontrado em {DRE_DIR} — S04–S10 ficam pendentes")
        for n, t in ((4, "RESUMO FINANCEIRO — HARAS COMPETÊNCIA"), (5, "ANÁLISE DE CUSTOS"),
                     (6, "ANÁLISE DE DESPESAS"), (7, "HARAS COMPETÊNCIA — ACUMULADO (YTD)"),
                     (8, "COMENTÁRIOS — VARIAÇÕES YTD"), (9, "INVESTIMENTOS"),
                     (10, "HARAS CAIXA — ORÇADO X REALIZADO")):
            spec.append(pendente(n, t, "", "DRE 2026 HPG - HARAS.xlsx", "arquivo não encontrado no Drive"))

    spec.append(slide_estoque(mes_ref))
    spec.append(slide_movimentacao(mes_ref))

    if casa.exists():
        rot_c, linhas_c = le_real_x_orcado(casa, "Real x Orçado")
        spec.append(slide_dre(13, f"RESUMO FINANCEIRO — CASA/FPG — ORÇADO X REALIZADO {rot_c.upper()}",
                              "FPG | Casa · Competência mensal · Fonte: aba Real x Orçado",
                              linhas_c, "mes"))
        spec.append(slide_dre(14, f"CASA/FPG — ORÇADO X REALIZADO ACUMULADO {mes_ref.year}",
                              "FPG | Casa · Acumulado no ano", linhas_c, "ytd"))
    else:
        aviso(f"DRE da Casa não encontrado — S13/S14 ficam pendentes")
        spec.append(pendente(13, "RESUMO FINANCEIRO — CASA/FPG", "", "DRE 2026 FPG - CASA.xlsx",
                             "arquivo não encontrado no Drive"))
        spec.append(pendente(14, "CASA/FPG — ACUMULADO", "", "DRE 2026 FPG - CASA.xlsx",
                             "arquivo não encontrado no Drive"))


def slide_investimentos(haras: Path, mes_ref: date):
    """S09 — só COMPRA DE ANIMAIS E PRODUTOS. Obra/infraestrutura fica de fora
    (regra do guia). A aba é uma lista com cabeçalhos:
      'INVESTIMENTOS - <MÊS>/<ANO>'  -> abre o mês
      'INFRAESTRUTURA' / 'COMPRA DE ANIMAIS E PRODUTOS' -> abre o bloco
      demais linhas: A=fornecedor, B=descrição, C=valor
    """
    import openpyxl
    wb = openpyxl.load_workbook(haras, data_only=True, read_only=True)
    ws = wb["Investimentos"]
    meses, atual, bloco = [], None, None
    for r in ws.iter_rows(values_only=True):
        a = str(r[0]).strip().upper() if r[0] is not None else ""
        b = str(r[1]).strip() if len(r) > 1 and r[1] is not None else ""
        v = num(r[2]) if len(r) > 2 else None
        if a.startswith("INVESTIMENTOS -"):
            nome = a.split("-", 1)[1].strip().split("/")[0].title()
            atual = {"mes": nome, "total": 0.0, "itens": []}
            meses.append(atual)
            bloco = None
            continue
        if atual is None:
            continue
        if a in ("INFRAESTRUTURA", "COMPRA DE ANIMAIS E PRODUTOS", "MÁQUINAS E EQUIPAMENTOS",
                 "INSTALAÇÕES", "FORMAÇÃO DE PASTAGEM"):
            bloco = a
            if bloco == "COMPRA DE ANIMAIS E PRODUTOS" and v:
                atual["total"] = v
            continue
        if bloco == "COMPRA DE ANIMAIS E PRODUTOS" and v is not None:
            atual["itens"].append({"desc": b or a.title(), "valor": v})
    wb.close()
    # mês sem compra aparece zerado — some da lista esconderia a informação
    for m in meses:
        if not m["itens"]:
            m["itens"] = [{"desc": "Sem compra de animais e produtos registrada no mês", "valor": 0.0}]
    return {"t": "lista_mes", "n": 9, "titulo": f"INVESTIMENTOS — COMENTÁRIOS {mes_ref.year}",
            "sub": "Compra de animais e produtos · acumulado do ano",
            "meses": meses}


# S11 — estoque em equinos. Regra do guia: status PLANTEL e sufixo EXATO
# 'DA PAO GRANDE' ou 'OUTRO' (variação com percentual duplicaria o animal).
SUFIXOS_S11 = ("DA PAO GRANDE", "OUTRO")


def slide_estoque(mes_ref: date):
    alvo = f"{mes_ref.year}-{mes_ref.month:02d}"
    if not BASE_BI.exists():
        return pendente(11, "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE", "",
                        "bases/base_bi.parquet", "rode python scripts/PGBaseBI.py")
    d = pd.read_parquet(BASE_BI)
    d["mes"] = pd.to_datetime(d["mes_referencia"]).dt.strftime("%Y-%m")
    if alvo not in set(d["mes"]):
        ult = sorted(d["mes"].unique())[-1]
        aviso(f"base_bi não tem {alvo} (vai até {ult}) — S11 fica pendente")
        return pendente(11, "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE", "",
                        "bases/base_bi.parquet", f"a base vai até {ult}; sem o mês {alvo}")
    m = d[(d["mes"] == alvo) & (d["status_plantel"] == "PLANTEL")
          & (d["sufixo_grupo"].isin(SUFIXOS_S11))]
    patrim = float(m["patrimonio_proporcional"].sum())
    avaliados = m["valor_100"].notna().sum()
    medio = float(m["valor_100"].mean()) if avaliados else 0.0
    cat = m["categoria"].value_counts()
    rows = [[k.title(), int(v), f"{v / len(m) * 100:.0f}%"] for k, v in cat.items()]
    return {"t": "kpis_tabela", "n": 11,
            "titulo": "ESTOQUE EM EQUINOS — FAZENDA PAO GRANDE",
            "sub": (f"Composição patrimonial do plantel · {MESES[mes_ref.month-1].upper()} {mes_ref.year}"
                    f" · {len(m)} animais · Status PLANTEL · Sufixo: Da PG / Outros"),
            "kpis": [
                {"v": f"{len(m)}", "l": "Animais Ativos", "s": "DA PAO GRANDE + OUTROS"},
                {"v": brl_curto(patrim), "l": "Patrimônio HPG", "s": "soma do patrimônio proporcional"},
                {"v": brl_curto(medio), "l": "Valor Médio", "s": f"{avaliados} animais avaliados"},
            ],
            "tabela": {"cols": ["CATEGORIA", "Nº", "%"], "rows": rows, "num": [1, 2]}}


MOV_LINHAS = [("saldo_ini", "Saldo Inicial"), ("compra", "(+) Compras"),
              ("producao", "(+) Prod. Emb."), ("venda", "(-) Baixa Vendas"),
              ("morte", "(-) Baixa Mortes"), ("doacao", "(-) Doações"),
              ("reaval", "(±) Reavaliação"), ("saiu_controle", "(-) Saiu do Controle"),
              ("saldo_fim", "Saldo Final")]


def slide_movimentacao(mes_ref: date):
    f = PLANTEL_DIR / "mov_cascata.parquet"
    if not f.exists():
        return pendente(12, f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {mes_ref.year}", "",
                        "LuxorMonthlyP-CRoutines/PlantelHPG/mov_cascata.parquet",
                        "rode o LxMovimentacao.py no repo LuxorMonthlyP-CRoutines")
    c = pd.read_parquet(f)
    ano = c[c["mes"].str.startswith(str(mes_ref.year))].sort_values("mes")
    ano = ano[ano["mes"] <= f"{mes_ref.year}-{mes_ref.month:02d}"]
    if ano.empty:
        return pendente(12, f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {mes_ref.year}", "",
                        f.name, f"sem meses de {mes_ref.year} na cascata")
    cols = ["TÍTULO"] + [MES_ABR[int(m.split("-")[1]) - 1].upper() for m in ano["mes"]]
    rows = [[rot] + [ano[k].tolist()[i] for i in range(len(ano))] for k, rot in MOV_LINHAS]
    ult = ano.iloc[-1]
    return {"t": "matriz", "n": 12,
            "titulo": f"RESUMO DA MOVIMENTAÇÃO DO PLANTEL — {mes_ref.year}",
            "sub": "Saldo mensal · compras, produções, vendas e baixas",
            "kpis": [
                {"v": brl_curto(ult["producao"]), "l": "Produção Emb.", "s": cols[-1]},
                {"v": brl_curto(ult["venda"]), "l": "Baixa Vendas", "s": cols[-1]},
                {"v": brl_curto(ult["morte"] + ult["doacao"]), "l": "Mortes/Doações", "s": cols[-1]},
                {"v": brl_curto(ult["saldo_fim"]), "l": "Saldo Final", "s": "Haras PG"},
            ],
            "cols": cols, "rows": rows, "moeda": True}


def brl_curto(v):
    if v is None:
        return "—"
    s = "-" if v < 0 else ""
    a = abs(v)
    if a >= 1e6:
        return f"{s}R$ {a/1e6:.1f}M".replace(".", ",")
    if a >= 1e3:
        return f"{s}R$ {a/1e3:.0f}k"
    return f"{s}R$ {a:.0f}"


# ------------------------------------------------------- Seções 02 a 05 (pendentes)
def secoes_pendentes(spec, mes_ref):
    est = "ESTACAO DE MONTA.xlsx (REPRODUÇÃO/ESTAÇÃO DE MONTA no Drive)"
    emb = "EMBRIOES A ENTREGAR - A RECEBER.xlsx"
    spec += [
        divisor(2, "ESTAÇÃO DE MONTA", "Embriões · Doadoras · Garanhões"),
        pendente(16, "ESTAÇÃO DE MONTA — EMBRIÕES E PRENHEZES", "Funil de prenhez da estação",
                 f"{est}, aba ESTAÇÃO (K=lavado, M=15d, N=30d, O=45d, P=60d, Q=aborto, AJ=estação)",
                 "extração ainda não feita"),
        pendente(17, "ESTAÇÃO DE MONTA — GARANHÕES", "Lavados, confirmados e índice por garanhão",
                 f"{est}, aba GARANHOES", "extração ainda não feita"),
        pendente(18, "ESTAÇÃO DE MONTA — COMPARATIVO COM ANOS ANTERIORES",
                 "Embriões confirmados por mês, por estação",
                 f"{est}, aba COMPARATIVO",
                 "a aba do arquivo em cache está em 20/21–23/24; confirmar a fonte do histórico atual"),
        pendente(19, "ESTAÇÃO DE MONTA — DOADORAS TIME A", "Meta × realizado por doadora",
                 f"{est}, abas REC. EMBR. (TIME) + PLANEJAMENTO (meta/real)", "extração ainda não feita"),
        pendente(20, "ESTAÇÃO DE MONTA — DOADORAS TIME B", "Meta × realizado por doadora",
                 f"{est}, abas REC. EMBR. + PLANEJAMENTO", "extração ainda não feita"),
        pendente(21, "ESTAÇÃO DE MONTA — COBERTURAS DISPONÍVEIS", "Saldo por garanhão de fora",
                 "COBERTURAS_CAVALOS_FORA.xlsx, aba Planilha2",
                 "arquivo não localizado no repo nem no Drive"),

        divisor(3, "EXPOSIÇÕES", "Programação e resultados"),
        pendente(23, "EXPOSIÇÕES — PROGRAMAÇÃO", "Calendário de participações previstas",
                 "digitado", "sem planilha por trás — é texto"),
        pendente(24, "RESULTADOS DAS EXPOSIÇÕES", "Animais, títulos e colocações",
                 "digitado após cada evento", "sem planilha por trás — é texto"),

        divisor(4, "VENDAS", "Pipeline e contratos"),
        pendente(29, "VENDAS — RESULTADO ACUMULADO", "Mês, YTD, meta anual e saldo para meta",
                 "PG_Mapa Vendas.xlsx, aba MAPA VENDAS (col 7=valor, 14=vendedor, 21=ano, 22=mês)",
                 "extração ainda não feita; a meta anual é parâmetro, não sai de planilha"),
        pendente(30, "VENDAS — DETALHAMENTO POR MÊS E EVENTO", "Filtro: vendedor = CARLA, sem cancelados",
                 "PG_Mapa Vendas.xlsx, aba MAPA VENDAS", "extração ainda não feita"),
        pendente(31, "VENDAS — INADIMPLÊNCIAS E RECEBÍVEIS", "Posição do mês",
                 "controle-de-inadimplencia -> dashboard_conferencia.html",
                 "hoje é print; dá pra embutir o HTML que o ControleInadimplencia.py já gera"),
        pendente(32, "VENDAS — EMBRIÕES VENDIDOS A FAZER (QUITADO / PAGANDO)", "",
                 f"{emb}, aba ENTREGAR — status_embrião='A fazer' e pgto Pagando/Quitado",
                 "extração ainda não feita"),
        pendente(33, "VENDAS — EMBRIÕES VENDIDOS A FAZER (PGTO PAUSADO / APÓS CONF.)", "",
                 f"{emb}, aba ENTREGAR — status_embrião='A fazer' e pgto Pausado/Após conf",
                 "extração ainda não feita"),
        pendente(34, "VENDAS — EMBRIÕES DE DIREITO / REPOSIÇÃO", "",
                 f"{emb}, aba ENTREGAR — status_embrião='Reposição' ou 'A fazer' com pgto Direito/Troca",
                 "extração ainda não feita"),
        pendente(35, "ESTAÇÃO DE MONTA — EMBRIÕES COMPRADOS A RECEBER", "",
                 f"{emb}, aba RECEBER — status_embrião='A fazer'", "extração ainda não feita"),

        divisor(5, "DECISÕES E MANEJO", "Plantel · Obras · Casa"),
        pendente(37, "PLANTEL — PAO GRANDE, ARRENDAMENTO E SÓCIOS", "Total de animais sob responsabilidade da PG",
                 "CONTROLE PLANTEL.xlsx, aba CONTAGEM (o PGSemanalReport.py já lê)",
                 "extração ainda não feita"),
        pendente(38, "MANEJO — PONTOS DE MELHORIA E DECISÕES", "Histórico de intervenções",
                 "digitado", "sem planilha por trás — é texto"),
        pendente(39, "MANEJO — FOTOS E REGISTROS", "Fotos do mês", "fotos", "upload manual"),
    ]


def divisor(n, titulo, sub):
    return {"t": "divisor", "n": n, "titulo": titulo, "sub": sub}


# ------------------------------------------------------------------- main
def build(mes_ref: date):
    spec = [
        {"t": "capa", "titulo": "RELATÓRIO DE DESEMPENHO ESTRATÉGICO",
         "mes": f"{MESES[mes_ref.month - 1].upper()} / {mes_ref.year}",
         "org": "HARAS PAO GRANDE"},
        {"t": "agenda", "titulo": "AGENDA",
         "sub": f"RELATÓRIO DESEMPENHO ESTRATÉGICO — {MES_ABR[mes_ref.month-1].upper()}/{str(mes_ref.year)[2:]}",
         "itens": [
             {"n": "01", "titulo": "FINANCEIRO", "sub": "DRE Haras · Caixa · Plantel"},
             {"n": "02", "titulo": "ESTAÇÃO DE MONTA", "sub": "Embriões · Doadoras · Garanhões"},
             {"n": "03", "titulo": "EXPOSIÇÕES", "sub": "Programação e resultados"},
             {"n": "04", "titulo": "VENDAS", "sub": "Pipeline e contratos"},
             {"n": "05", "titulo": "DECISÕES E MANEJO", "sub": "Plantel · Obras · Casa"},
         ]},
        divisor(1, "FINANCEIRO", f"DRE Haras · Caixa · Plantel | {MESES[mes_ref.month-1].upper()} {mes_ref.year}"),
    ]
    secao_financeiro(mes_ref, spec)
    secoes_pendentes(spec, mes_ref)
    spec.append({"t": "encerramento", "titulo": "HARAS PAO GRANDE"})

    payload = {
        "mes": f"{mes_ref.year}-{mes_ref.month:02d}",
        "mesLabel": f"{MESES[mes_ref.month - 1]} {mes_ref.year}",
        "avisos": avisos,
        "slides": spec,
    }
    OUT.mkdir(parents=True, exist_ok=True)
    js = json.dumps(payload, ensure_ascii=False, separators=(",", ":"), default=_json_default)
    (OUT / "spec.json").write_text(js, encoding="utf-8")
    (OUT / "spec.js").write_text(f"window.COMITE_SPEC = {js};\n", encoding="utf-8")
    prontos = sum(1 for s in spec if s["t"] not in ("pendente",))
    print(f"[comite] {len(spec)} slides ({prontos} com conteúdo, "
          f"{len(spec) - prontos} pendentes) · {len(js)//1024} KB -> assets/comite/spec.js")


def parse_mes(arg):
    try:
        mm, aaaa = arg.split("/")
        return date(int(aaaa), int(mm), 1)
    except Exception:
        sys.exit(f"mês inválido: {arg!r} — use MM/AAAA (ex.: 06/2026)")


if __name__ == "__main__":
    if len(sys.argv) > 1:
        build(parse_mes(sys.argv[1]))
    else:
        # sem argumento: usa o mês em que a planilha do DRE está
        try:
            import openpyxl
            wb = openpyxl.load_workbook(DRE_DIR / "DRE 2026 HPG - HARAS.xlsx",
                                        data_only=True, read_only=True)
            rot = str(wb["Real x Orçado (Comp)"].cell(1, 2).value or "").strip()
            wb.close()
            nome, ano = rot.split()
            build(date(int(ano), MESES.index(nome.capitalize()) + 1, 1))
        except Exception as e:
            sys.exit(f"não consegui descobrir o mês pela planilha ({e}). "
                     f"Passe explícito: python hub/tools/build_comite.py MM/AAAA")
