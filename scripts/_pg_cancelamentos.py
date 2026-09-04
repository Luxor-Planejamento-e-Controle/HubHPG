"""Cancelamentos lançados em MOVIMENTAÇÕES: o que foi cancelado e pra que lado.

Cancelamento não é um evento a mais na lista — ele DESFAZ outro. Ler a linha como
texto solto erra o sinal: "VENDA CANCELADA" faz a cota VOLTAR pra Pao Grande
(entra patrimônio), "COMPRA CANCELADA" faz a cota IR A ZERO (sai o animal que
nunca foi nosso), e as duas frases têm a mesma palavra-chave.

O que motivou: em 31/08/2026 a RELIQUIA DA TERRA BRAVA teve a venda pro Cicero
cancelada e a cota devolvida a 50%, mas STATUS ficou 'VENDIDO E ENTREGUE' e
CONDIÇÃO ATUAL 'SAIU DO HARAS'. O cancelamento entrou só na cota, então ela segue
fora do headcount enquanto, pelo cancelamento, voltou a ser nossa. Sem alguém
cruzando as duas coisas, isso não aparece em relatório nenhum — nem no nosso nem
no do haras.

As regras saíram das 58 ocorrências com 'CANCEL' do controle de plantel (2023-2026),
não de suposição. A variedade que elas cobrem:

    VENDA CANCELADA - PASSOU DE 50% PARA 100%
    VOLTOU PARA O PLANTEL - A VENDA DE 50% DA POTRA FOI CANCELADA
    MUDOU A COTA - ESTAVA 45% PASSOU PARA 70% - VENDA 25% PARA SILVIO LUCIO CANCELADA
    MUDOU A COTA PARA 56,25% - LUIZ ANTONIO CANCELOU A COMPRA DE 6,25%   <- volta pra nós
    COMPRA CANCELADA - ZEROU A COTA E O VALOR NO PLANTEL                 <- sai
    CANCELAMENTO DE 25% DE JULIANO MARQUES - VOLTOU A COTA DE 0% PARA 25%
    VENDA PARA ANTIGA SOCIA (50% X) CANCELADA - VENDIDO 100% PARA Y      <- cancela E vende

Uso:
    from _pg_cancelamentos import interpreta, pendencias
    c = interpreta("VENDA CANCELADA - PASSOU DE 50% PARA 100%")
    c["sentido"]      # 'volta'
    c["cota_para"]    # 1.0
"""

from __future__ import annotations

import re
import unicodedata

# ------------------------------------------------------------------
# normalização (o mesmo espírito do _norm do pipeline: sem acento, maiúsculo)
# ------------------------------------------------------------------
def _norm(s) -> str:
    if s is None:
        return ""
    t = unicodedata.normalize("NFKD", str(s))
    t = "".join(c for c in t if not unicodedata.combining(c))
    return " ".join(t.upper().split())


# 'CANCELADA', 'CANCELADO', 'CANCELAMENTO', 'CANCELOU' — sempre o radical
RX_CANCEL = re.compile(r"CANCEL")

# "cancelou a compra" é TERCEIRO desistindo de comprar de nós: a cota volta pra
# casa, igual a uma venda cancelada. Vem antes de COMPRA CANCELADA na ordem de
# teste porque as duas casam a palavra COMPRA.
RX_TERCEIRO_DESISTIU = re.compile(r"CANCEL\w*\s+A\s+COMPRA|COMPRA\s+DE\s+[\d,.]+%?\s+CANCEL")
RX_COMPRA_CANCELADA = re.compile(r"COMPRA\s+CANCELAD")
RX_VENDA_CANCELADA = re.compile(r"VENDA[^.]{0,80}?CANCELAD|CANCELAD\w*\s+A\s+VENDA"
                                r"|VENDA\s+CANCELAD|CANCELAMENTO\s+DE")

# volta explícita ao plantel, em todas as formas que aparecem na planilha
RX_VOLTOU = re.compile(r"VOLTOU (PARA|A SER)|VOLTOU \d|ENTROU NO PLANTEL"
                       r"|INSERIDO NO PLANTEL|PASSOU PARA PLANTEL|VOLTOU O VALOR")
RX_ZEROU = re.compile(r"ZEROU (A COTA|A %|O VALOR)")

# cota antes -> depois. Duas formas, e a ordem importa: "ESTAVA/ERA x ... PARA y"
# ganha de "DE x PARA y", senão "ESTAVA COMO 48,75% - VENDA DE 1,25% PARA EDELVO
# LARA CANCELADA - PASSOU PARA 50%" lia a cota VENDIDA (1,25%) como cota anterior.
# O "DE x PARA y" exige x e y colados, sem frase no meio, pelo mesmo motivo.
RX_ESTAVA_PARA = re.compile(r"(?:ESTAVA(?:\s+COMO|\s+EM)?|ERA)\s*([\d]+(?:[.,]\d+)?)\s*%?"
                            r"[^%]{0,60}?PARA\s*([\d]+(?:[.,]\d+)?)\s*%")
RX_DE_PARA = re.compile(r"DE\s*([\d]+(?:[.,]\d+)?)\s*%?\s*"
                        r"(?:PASSOU\s+)?PARA\s*([\d]+(?:[.,]\d+)?)\s*%")
RX_SO_PARA = re.compile(r"(?:PARA|ALTERADA PARA|MUDOU A COTA PARA)\s*([\d]+(?:[.,]\d+)?)\s*%")

# uma venda NOVA na mesma linha do cancelamento (o caso OARA): "... CANCELADA -
# VENDIDO 100% PARA RICARDO ...". Não é volta líquida, é troca de comprador.
RX_VENDA_NOVA = re.compile(r"CANCELAD\w*[^.]{0,40}?\bVENDID")

# nome do comprador/vendedor. Para de ler no primeiro token que não é nome —
# CANCELADA, FOI, E, NA, DE... — senão sai "RAPHAEL BERTOLINI CANCELADA".
RX_CONTRAPARTE = re.compile(r"(?:PARA|DE)\s+(?:O\s+|A\s+)?"
                            r"((?:[A-Z][A-Z'`\-]{2,})(?:\s+(?:DA|DE|DO|DOS|DAS)\s+[A-Z][A-Z'`\-]{2,}"
                            r"|\s+[A-Z][A-Z'`\-]{2,}){0,3})")
_CORTA_NOME = {"CANCELADA", "CANCELADO", "CANCELAMENTO", "CANCELOU", "FOI", "E", "NA", "NO",
               "COM", "POR", "QUE", "PASSOU", "MUDOU", "VOLTOU", "TOTAL", "VENDIDO",
               "VENDIDA", "ALTEROU", "ALTERADA", "ZEROU", "PLANTEL", "HARAS", "COTA",
               "VALOR", "PORCENTAGEM", "AVALIACAO", "ULTIMA", "SEMANA", "LEILAO"}

_PALAVRAS_NAO_NOME = {"O PLANTEL", "PLANTEL", "A COTA", "COTA", "O VALOR", "VALOR",
                      "O HARAS", "HARAS", "A VENDA", "VENDA", "A COMPRA", "COMPRA",
                      "A ULTIMA AVALIACAO", "ULTIMA AVALIACAO", "A PORCENTAGEM"}


def _num(txt: str | None):
    if txt is None:
        return None
    v = float(txt.replace(".", "").replace(",", ".")) if "," in txt else float(txt)
    # a planilha escreve tanto 50 quanto 50,00 para 50% — cota fica em fração
    return round(v / 100.0, 6)


def _contraparte(oc: str) -> str | None:
    for m in RX_CONTRAPARTE.finditer(oc):
        toks = []
        for w in m.group(1).split():
            if w in _CORTA_NOME:
                break
            toks.append(w)
        nome = " ".join(toks).strip(" -,.")
        if not nome or nome in _PALAVRAS_NAO_NOME or len(nome) < 6 or " " not in nome:
            continue
        return nome
    return None


def interpreta(ocorrencia) -> dict | None:
    """Lê UMA ocorrência. None quando não é cancelamento.

    Devolve:
      tipo         'venda' | 'compra'      — o que foi cancelado
      sentido      'volta' | 'sai'         — pra onde vai a cota
      cota_de/para fração (0.5 = 50%), quando a frase declara
      contraparte  nome do comprador/vendedor, quando dá pra ler
      volta_ao_plantel / zerou            — o que a frase afirma explicitamente
      venda_nova   True quando cancela e revende na MESMA linha (não é volta líquida)
    """
    oc = _norm(ocorrencia)
    if not RX_CANCEL.search(oc):
        return None

    if RX_TERCEIRO_DESISTIU.search(oc):
        tipo, sentido = "venda", "volta"      # o comprador desistiu: cota volta pra casa
    elif RX_COMPRA_CANCELADA.search(oc):
        tipo, sentido = "compra", "sai"       # nossa compra caiu: o animal não é nosso
    elif RX_VENDA_CANCELADA.search(oc):
        tipo, sentido = "venda", "volta"
    else:
        tipo, sentido = None, None            # cancelamento de outra coisa (ver testes)

    de = para = None
    m = RX_ESTAVA_PARA.search(oc) or RX_DE_PARA.search(oc)
    if m:
        de, para = _num(m.group(1)), _num(m.group(2))
    else:
        m2 = RX_SO_PARA.search(oc)
        if m2:
            para = _num(m2.group(1))

    zerou = bool(RX_ZEROU.search(oc))
    if zerou and para is None:
        para = 0.0
    venda_nova = bool(RX_VENDA_NOVA.search(oc))
    if venda_nova:
        sentido = "sai"                       # cancelou pra vender pra outro

    return {"tipo": tipo, "sentido": sentido, "cota_de": de, "cota_para": para,
            "contraparte": _contraparte(oc), "volta_ao_plantel": bool(RX_VOLTOU.search(oc)),
            "zerou": zerou, "venda_nova": venda_nova, "texto": str(ocorrencia).strip()}


# ------------------------------------------------------------------
# pendência = cancelamento que o cadastro não acompanhou
# ------------------------------------------------------------------
_STATUS_SAIU = ("VENDIDO E ENTREGUE", "DOADO", "OBITO", "DE TERCEIRO")


def pendencias(linhas, log, ate=None) -> list:
    """Animais cuja cota VOLTOU por cancelamento mas cujo cadastro segue dizendo
    que eles saíram. É o caso da RELIQUIA: cota 50% de volta em 31/08/2026 com
    STATUS 'VENDIDO E ENTREGUE' e CONDIÇÃO 'SAIU DO HARAS'.

    `linhas`: [{'nome','status','local','condicao','cota'}] do controle de plantel.
    `log`   : [{'produto','data','ocorrencia'}] da aba MOVIMENTAÇÕES.
    `ate`   : opcional, ignora cancelamento lançado depois dessa data (date).

    Devolve uma linha por animal, com o cancelamento mais recente que o explica.
    """
    ultimo = {}
    for x in log or []:
        c = interpreta(x.get("ocorrencia"))
        if not c or c["sentido"] != "volta" or c["venda_nova"]:
            continue
        d = x.get("data")
        if ate is not None and d is not None and d > ate:
            continue
        k = _norm(x.get("produto"))
        if k and (k not in ultimo or (d or "") >= (ultimo[k].get("data") or "")):
            ultimo[k] = {"data": d, **c}

    out = []
    for l in linhas or []:
        k = _norm(l.get("nome"))
        c = ultimo.get(k)
        if not c:
            continue
        status = _norm(l.get("status"))
        cota = l.get("cota")
        # Só o STATUS tira do headcount. 'CONDIÇÃO ATUAL = SAIU DO HARAS' com
        # STATUS='PLANTEL' e cota viva é animal no sócio, que CONTA (bucket SOCIO)
        # — MORFEU e MELISSA caíam aqui como pendência sem ser: o texto da
        # condição é histórico, não estado.
        if status in _STATUS_SAIU and bool(cota):
            out.append({"animal": l.get("nome"), "local": l.get("local"),
                        "status": l.get("status"), "condicao": l.get("condicao"),
                        "cota": cota, "cancelado_em": c.get("data"),
                        "contraparte": c.get("contraparte"), "ocorrencia": c["texto"]})
    return out
