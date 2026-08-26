"""Gera a auditoria de fontes a partir do SNAPSHOT, não da memória de quem escreve.

A primeira versão desta página foi digitada à mão. Envelheceu em silêncio duas
vezes — publicava "vendidos 5" depois de a regra mudar para 6 — que é exatamente
a doença que o resto do pipeline passou a acusar. Página que audita dado não pode
ser a única coisa sem fonte.

Lê o que já existe: `_cache/semanal_snapshots.json` (o congelado), `bases/
semanal_docx.json` (o relatório oficial) e a mesma comparação do fechamento
(scripts/PGSemanalPrecisao._linhas). O que é metadado estável — de qual arquivo e
aba sai cada número, e a regra — vive em FONTES aqui embaixo.

Saída: dashboards/auditoria_semanal.html (gitignored — tem nome de animal).

Uso:
    python tools/build_auditoria.py            # semana mais recente
    python tools/build_auditoria.py 21/08/2026
"""
from __future__ import annotations

import html
import json
import sys
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "scripts"))

import PGSemanalReport as R          # noqa: E402
from PGSemanalPrecisao import _eq, _linhas   # noqa: E402

SAIDA = ROOT / "dashboards" / "auditoria_semanal.html"
SPEC_COMITE = ROOT / "assets" / "comite" / "spec.json"

# Uma fonte mensal com mais de 45 dias já não descreve o mês que o deck publica.
# Não é erro — o haras pode não ter atualizado —, mas tem de aparecer.
DIAS_FONTE_VELHA = 45

# Metadado ESTÁVEL por indicador: (seção, fonte·aba, regra).
# Só muda quando a origem muda — e aí muda aqui, num lugar só.
FONTES = {
    "Acumulado estação": ("1 · Produção",
        "EMBRIÕES E MATRIZES · EMBRIÕES PAO GRANDE + SOCIOS\nESTACAO DE MONTA · ESTAÇÃO\nCONTROLE PLANTEL · PLANTEL",
        "Vivos no grupo + parições da safra. A soma funciona porque a planilha do grupo apaga a linha de quem pariu; parição não lançada em lugar nenhum some da conta. Parição conhecida só pelo roster entra aqui, com registro cumulativo. O valor nunca cai — há piso por safra."),
    "Acumulado estação (safra nova)": ("1 · Produção",
        "EMBRIÕES E MATRIZES · EMBRIÕES PAO GRANDE + SOCIOS\nESTACAO DE MONTA · ESTAÇÃO",
        "Mesma regra da safra que fecha, aplicada na que começa — o relatório publica as duas durante a transição. Enquanto não há uma linha lançada para a safra nova em nenhuma das duas planilhas, o número é 0, que é o '--' do relatório."),
    "Confirmados semana": ("1 · Produção", "ESTACAO DE MONTA · ESTAÇÃO",
        "Coluna +/- = OK. A contagem da semana é a diferença do conjunto contra o snapshot anterior."),
    "Nascimentos": ("1 · Produção", "CONTROLE_DE_PLANTEL mensal · PLANTEL",
        "Linha do roster com NASCIMENTO na janela — por data e filiação (MAE/PAI), nunca por nome: o potro entra no roster ora com nome próprio ('PRINCIPE MN DA PAO GRANDE'), ora com o cruzamento ('MACHO LIBRA x OLIMPO'), e detectar por texto perdia o primeiro caso. A parição lançada na aba ESTAÇÃO fica como conferência: divergir é lançamento faltando em um dos dois."),
    "Acumulado no mês": ("1 · Produção", "ESTACAO DE MONTA · ESTAÇÃO",
        "Confirmados com data dentro do mês da referência."),
    "Abortos / óbitos": ("1 · Produção",
        "ESTACAO DE MONTA · ESTAÇÃO\nCONTROLE_DE_PLANTEL mensal · CONFIRMAÇÕES, ABORTOS, MORTES",
        "Aborto de embrião confirmado e óbito de potro vêm da estação; óbito de animal já no plantel só existe no controle mensal, por palavra-chave."),
    "Receptoras total": ("2 · Receptoras", "ARRENDAMENTOS E RECEPTORAS · ANIMAIS",
        "LOCAL em Pao Grande ou Arrendamento Cesar Furtado, STATUS prenha ou vazia."),
    "Receptoras prenhas": ("2 · Receptoras", "ARRENDAMENTOS E RECEPTORAS · ANIMAIS",
        "STATUS começando com PRENHA."),
    "Receptoras vazias": ("2 · Receptoras", "ARRENDAMENTOS E RECEPTORAS · ANIMAIS",
        "STATUS começando com VAZIA."),
    "Índice eficiência": ("2 · Receptoras",
        "ARRENDAMENTOS E RECEPTORAS · ANIMAIS\nCONTROLE PLANTEL · PLANTEL",
        "Vazias ÷ doadoras contadas (CATEGORIA = DOADORA). O relatório já oscilou entre divisor fixo 10 e o contado."),
    "Headcount total": ("3 · Headcount",
        "CONTROLE PLANTEL · PLANTEL\nARRENDAMENTOS E RECEPTORAS · ANIMAIS",
        "Animais por LOCAL mais receptoras do mesmo local. Reproduz o COUNTIF da aba CONTAGEM, que fica só como conferência."),
    "Fazenda Pao Grande": ("3 · Headcount", "CONTROLE PLANTEL · PLANTEL", "Animais mais receptoras da fazenda."),
    "Arrendamento": ("3 · Headcount", "CONTROLE PLANTEL · PLANTEL", "Animais mais receptoras do arrendamento."),
    "Centro de Treinamento": ("3 · Headcount", "CONTROLE PLANTEL · PLANTEL",
        "LOCAL OUTROS no roster. Mato Grosso fica fora da contagem por decisão de negócio."),
    "Sócios": ("3 · Headcount", "CONTROLE PLANTEL · PLANTEL", "LOCAL SOCIO no roster."),
    "Δ headcount": ("3 · Headcount", "snapshot local · headcount_history.json",
        "Variação do número de ANIMAIS, que é o que o relatório publica no '+02 / -01'. Receptora não entra aqui: ela é contada à parte, e receptora que vai pro sócio sai da contagem sem aparecer nesse Δ. A variação do total (animais + receptoras) fica em delta, separada."),
    "Saídas semana": ("5 · Saídas", "CONTROLE_DE_PLANTEL mensal · SAIDAS-ENTRADAS",
        "Sentido pelo prefixo da classificação. Classificação desconhecida avisa, nunca vira zero calado."),
    "Entradas semana": ("5 · Saídas",
        "CONTROLE_DE_PLANTEL mensal · SAIDAS-ENTRADAS",
        "Só animal que CHEGA de fora. Potro nascido não entra aqui — é produção, já contada em Nascimentos —, mas conta no headcount e no Δ."),
    "Transferências internas": ("5 · Saídas", "ARRENDAMENTOS E RECEPTORAS · ANIMAIS",
        "Diferença do LOCAL contra o snapshot da semana anterior. A aba MOVIMENTAÇÕES não serve: a última transferência lançada lá é de setembro/2025."),
    "Vendidos pendentes": ("5 · Saídas",
        "CONTROLE_DE_PLANTEL mensal · PLANTEL\nEMBRIOES A ENTREGAR · ENTREGAR",
        "STATUS PLANTEL = VENDIDO PENDENTE SAIDA mais embrião de cota integral aguardando entrega. Reposição conta (sai para repor outro animal, mas está pendente)."),
    "Sociedade pendentes": ("5 · Saídas",
        "Animais para sair · ANIMAIS VENDIDOS",
        "ANIMAIS tipo SOCIEDADE. Embrião de cota parcial não entra aqui: tem indicador "
        "próprio, porque o relatório soma embrião nesta linha em algumas semanas e em "
        "outras não. Animal com pendência documental fica fora. O Animais para sair está "
        "congelado em 24/07, então o animal só sai da conta quando some do roster mensal — "
        "sociedade nunca recebe marca de STATUS PLANTEL, exigir marca derrubava todos por "
        "construção. A comparação usa a abertura do relatório quando ela existe."),
    "Total terceiros": ("4 · Terceiros",
        "CONTROLE_DE_PLANTEL mensal · PLANTEL\nEMBRIOES A ENTREGAR · ENTREGAR",
        "É o pendente de saída, como diz o rótulo do próprio relatório."),
    "Outros terceiros": ("4 · Terceiros", "—",
        "Não existe fonte: nenhuma planilha marca animal de terceiro em cavalgada ou treino. Fica em branco por ausência de informação, nunca zero."),
}

# Indicadores que o dashboard publica mas o relatório não tem linha para comparar.
SEM_CONTRAPARTE = [
    ("1 · Produção", "Aberto · PG / sócio / vendido",
     lambda s: " · ".join(str(s.get("acumulado_estacao_split", {}).get(k, 0))
                          for k in ("pg", "socio", "vendido")),
     "EMBRIÕES E MATRIZES · EMBRIÕES PAO GRANDE + SOCIOS",
     "Fatia por cota e STATUS. Parição cuja cota não foi recuperada no arquivo da semana anterior fica fora, e a soma sai abaixo do acumulado."),
    ("2 · Receptoras", "Doadoras ciclando",
     lambda s: (s.get("receptoras") or {}).get("doadoras_ciclando"),
     "_cache/semanal_manual.json (input humano)",
     "Égua disponível hormonalmente para doar óvulo. NÃO existe em planilha: é avaliação do veterinário. O roster e o PLANEJAMENTO da estação só têm cadastro e logística; a aba MATRIZES tem data de ovulação/coleta, que é evento passado. Preenchido à mão por semana, sem herdar a semana anterior. O relatório oficial também não publica."),
    ("2 · Receptoras", "Doadoras (estação)",
     lambda s: (s.get("receptoras") or {}).get("doadoras"),
     "CONTROLE PLANTEL · PLANTEL",
     "CATEGORIA = DOADORA no roster. O relatório não publica o número, só o usa como divisor do índice."),
    ("5 · Saídas", "Embriões em sociedade aguardando entrega",
     lambda s: (s.get("terceiros") or {}).get("sociedade_pendentes_embrioes"),
     "EMBRIOES A ENTREGAR · ENTREGAR",
     "Cota PG parcial com status PRONTO - AGUARDANDO ENTREGA. PRONTO - NASCE NA PG fica de fora. O relatório não tem linha fixa para isto — às vezes soma na sociedade pendente, às vezes omite."),
    ("4 · Terceiros", "Doadoras terceiros",
     lambda s: (s.get("terceiros") or {}).get("doadoras_terceiros"),
     "CONTROLE_DE_PLANTEL mensal · PLANTEL",
     "DE TERCEIRO + CATEGORIA = DOADORA + local na propriedade. As marcadas em OUTROS não estão aqui."),
]

ORDEM = ["1 · Produção", "2 · Receptoras", "3 · Headcount", "4 · Terceiros", "5 · Saídas"]

# Nome que FONTES escreve -> rótulo com que o pipeline registrou a fonte. É o que
# permite trocar "CONTROLE PLANTEL" pelo caminho real do arquivo lido na semana.
# Fonte que não é planilha do Drive (snapshot local, input humano) não entra.
ROTULO_DA_FONTE = {
    "EMBRIÕES E MATRIZES": "acumulado na estação",
    "ESTACAO DE MONTA": "estacao de monta",
    "CONTROLE PLANTEL": "roster do plantel",
    "ARRENDAMENTOS E RECEPTORAS": "receptoras",
    "CONTROLE_DE_PLANTEL mensal": "controle mensal",
    "Animais para sair": "sociedade pendente",
    "EMBRIOES A ENTREGAR": "embrioes a entregar",
}


def _fmt(v):
    if v is None or v == "":
        return "--"
    if isinstance(v, float):
        return f"{v:.1f}".replace(".", ",")
    return str(v)


def _chip(ok):
    return ('<span class="chip ok">bate</span>' if ok
            else '<span class="chip bad">não bate</span>')


def _porque(lab, calc, alvo, snap, dx, hist, semana):
    """Frase que EXPLICA a divergencia, montada com os numeros reais.

    Sem isto a auditoria diz 'nao bate' e para ali — quem le tem de refazer a
    investigacao toda vez. O texto de regra e estavel; este e o desta semana.
    """
    if lab == "Δ headcount":
        hc = snap.get("headcount") or {}
        ant = max((w for w in sorted(hist) if R._is_iso(w) and w < semana), default=None)
        nosso_ant = (hist[ant].get("headcount") or {}).get("total") if ant else None
        return (f"O Δ do relatório conta ANIMAIS ({dx['headcount'].get('delta_txt')}), não o "
                f"total: receptora que vai pro sócio sai da contagem e não aparece ali. "
                f"Nesta semana os animais foram {hc.get('delta_animais'):+d} e as receptoras "
                f"{hc.get('delta_receptoras'):+d}, então o total foi de {nosso_ant} para "
                f"{hc.get('total')} ({hc.get('delta'):+d}). As duas contas estão certas — "
                f"medem coisas diferentes.")

    t = snap.get("terceiros") or {}
    if lab == "Vendidos pendentes":
        return (f"Nosso {calc} = {t.get('vendidos_pendentes_animais')} animais e "
                f"{t.get('vendidos_pendentes_embrioes')} embrião(ões).")
    if lab == "Sociedade pendentes":
        e = t.get("sociedade_pendentes_embrioes")
        return (f"Contagem de ANIMAIS dos dois lados. Os {e} embrião(ões) em sociedade "
                f"ficam no indicador ao lado — o relatório ora soma embrião nesta linha "
                f"(14/08: '04 = 01 animal e 03 embriões'), ora não (21/08: '01'), com os "
                f"mesmos embriões abertos nas duas semanas.")
    if lab == "Total terceiros":
        return (f"É o mesmo conjunto dos vendidos pendentes: "
                f"{t.get('terceiros_animais')} animais e {t.get('terceiros_embrioes')} embrião(ões).")
    if lab in ("Vendidos pendentes", "Sociedade pendentes", "Total terceiros"):
        return ""
    if lab == "Índice eficiência":
        r = snap.get("receptoras") or {}
        return (f"Nosso {calc} = {r.get('vazias')} vazias ÷ {r.get('doadoras')} doadoras contadas.")
    if lab == "Entradas semana":
        n = (snap.get("saidas") or {}).get("entradas_nascimento")
        return (f"{n} nascimento(s) na semana ficam FORA desta linha (entram no "
                f"headcount e no Δ, não nas entradas)." if n else "")
    return ""


def _linhas_da_semana(semana):
    hist, docx = R._load_hist(), R._load_docx_ref()
    if semana not in hist:
        sys.exit(f"Sem snapshot para {semana}. Rode o fechamento primeiro.")
    snap = hist[semana]
    dx = docx.get(semana)
    comparaveis = []
    if dx and dx.get("ref_confere") is not False:
        comparaveis = _linhas(snap, dx)
    return snap, dx, comparaveis


def _secao_comite() -> str:
    """Tabelas do comitê a partir de assets/comite/spec.json."""
    if not SPEC_COMITE.exists():
        return ('<section><h2>Comitê mensal</h2><div class="nota">'
                'spec.json não existe — rode <code>python tools/build_comite.py</code>.'
                '</div></section>')
    spec = json.loads(SPEC_COMITE.read_text(encoding="utf-8"))
    mes = spec.get("padrao")
    deck = (spec.get("decks") or {}).get(mes) or []
    fontes = spec.get("fontes") or {}

    linhas_f, velhas = [], 0
    for rotulo in sorted(fontes):
        f = fontes[rotulo]
        quando = f.get("modificado")
        idade = ""
        if quando:
            dias = (datetime.now() - datetime.fromisoformat(quando)).days
            idade = f"{dias} dia{'s' if dias != 1 else ''}"
            if dias > DIAS_FONTE_VELHA:
                velhas += 1
                idade = f'<span class="chip bad">{idade}</span>'
        linhas_f.append(
            f'<tr><td>{html.escape(rotulo)}</td>'
            f'<td class="src">{html.escape(f.get("arquivo") or "—")}<br>'
            f'<span class="cam">{html.escape(f.get("caminho") or f.get("pasta") or "")}</span></td>'
            f'<td>{html.escape((quando or "—").replace("T", " "))}</td>'
            f'<td class="num">{idade}</td></tr>')

    linhas_s, pendentes = [], 0
    for sl in deck:
        tipo = sl.get("t")
        if tipo in ("capa", "agenda", "divisor", "encerramento"):
            continue          # moldura do deck, não tem fonte de dado
        if tipo == "pendente":
            pendentes += 1
            situacao = '<span class="chip bad">pendente</span>'
            obs = (f'<b>{html.escape(sl.get("motivo") or "")}</b><br>'
                   f'{html.escape(sl.get("fonte") or "")}')
        else:
            situacao = '<span class="chip ok">com dado</span>'
            obs = html.escape(sl.get("sub") or "")
        linhas_s.append(
            f'<tr><td class="num">{sl.get("n")}</td>'
            f'<td>{html.escape(sl.get("titulo") or "")}</td>'
            f'<td><span class="aba">{html.escape(tipo or "")}</span></td>'
            f'<td>{situacao}</td><td class="obs">{obs}</td></tr>')

    com_dado = len(linhas_s) - pendentes
    rotulo_mes = (spec.get("labels") or {}).get(mes) or (mes or "")
    return f"""<section>
  <h2>Comitê mensal &middot; {html.escape(rotulo_mes)}</h2>
  <div class="tiles">
    <div class="tile ok"><span class="num">{com_dado}</span><span class="lab">slides com dado, de {len(linhas_s)} com fonte</span></div>
    <div class="tile bad"><span class="num">{pendentes}</span><span class="lab">pendentes — falta conteúdo ou fonte</span></div>
    <div class="tile neu"><span class="num">{len(fontes)}</span><span class="lab">arquivos lidos no build{(" · " + str(velhas) + " com mais de " + str(DIAS_FONTE_VELHA) + " dias") if velhas else ""}</span></div>
  </div>
  <div class="scroll"><table>
    <thead><tr><th>Rótulo</th><th>Arquivo &middot; pasta</th><th>Modificado</th><th>Idade</th></tr></thead>
    <tbody>
{chr(10).join(linhas_f)}
    </tbody>
  </table></div>
  <div class="scroll" style="margin-top:18px"><table>
    <thead><tr><th>#</th><th>Slide</th><th>Tipo</th><th>Situação</th><th>Fonte declarada / detalhe</th></tr></thead>
    <tbody>
{chr(10).join(linhas_s)}
    </tbody>
  </table></div>
  <div class="nota">Lida de <code>assets/comite/spec.json</code>: o que o build de fato
  abriu, não o que está escrito em <code>_docs/COMITE_MAPEAMENTO.md</code>. A regra de
  cada slide (aba, filtro) continua lá — é texto estável.</div>
</section>"""


def build(semana=None):
    hist = R._load_hist()
    semanas = sorted(w for w in hist if R._is_iso(w))
    semana = semana or semanas[-1]
    snap, dx, comparaveis = _linhas_da_semana(semana)

    # gravado por _registra_caminhos; snapshot antigo não tem, e aí a coluna fica
    # como era antes — só o nome do arquivo
    caminhos = snap.get("fontes_caminhos") or {}
    fora_de_lugar = set(snap.get("fontes_fora_de_lugar") or [])

    por_secao = {s: [] for s in ORDEM}
    n_ok = n_tot = 0
    for lab, calc, alvo in comparaveis:
        secao, fonte, regra = FONTES.get(lab, ("5 · Saídas", "—", ""))
        ok = _eq(calc, alvo)
        n_tot += 1
        n_ok += 1 if ok else 0
        texto = regra
        if not ok:
            porque = _porque(lab, calc, alvo, snap, dx, hist, semana)
            if porque:
                texto = f'<b>{html.escape(porque)}</b><br>{regra}'
        por_secao[secao].append((lab, _fmt(calc), _fmt(alvo), _chip(ok), fonte, texto))
    for secao, lab, fn, fonte, regra in SEM_CONTRAPARTE:
        por_secao[secao].append((lab, _fmt(fn(snap)), "—",
                                 '<span class="chip neu">sem linha</span>', fonte, regra))

    avisos = []
    if dx:
        if dx.get("ref_confere") is False:
            avisos.append(f"O relatório desta semana declara <code>{html.escape(str(dx.get('semana_txt')))}</code> "
                          "— foi reaproveitado como rascunho de outra semana, então não serve de alvo.")
        h = dx["headcount"]
        if h.get("coerente") is False:
            avisos.append(f"O relatório não fecha consigo mesmo: os locais somam {h['soma_locais']} "
                          f"e o total declarado é {h['total']}.")
    sp = snap.get("acumulado_estacao_split") or {}
    soma = sum(v for v in sp.values() if isinstance(v, int))
    ac = snap.get("acumulado_estacao")
    if ac and soma and soma != ac:
        avisos.append(f"O split PG/sócio/vendido soma {soma} e o acumulado é {ac} — "
                      f"a fatia de {ac - soma} parição(ões) não foi recuperada.")

    linhas_html = []
    for secao in ORDEM:
        if not por_secao[secao]:
            continue
        linhas_html.append(f'<tr class="grp"><td colspan="6">{secao}</td></tr>')
        for lab, calc, alvo, chip, fonte, regra in por_secao[secao]:
            def _uma(p):
                nome = p.split(" · ")[0]
                out = html.escape(nome)
                if " · " in p:
                    out += f'<br><span class="aba">{html.escape(p.split(" · ", 1)[1])}</span>'
                # caminho do arquivo que a semana leu de verdade; sem registro
                # (fonte que não é planilha, ou snapshot antigo) fica só o nome
                rot = ROTULO_DA_FONTE.get(nome, "")
                cam = caminhos.get(rot)
                if cam:
                    out += f'<br><span class="cam">{html.escape(cam)}</span>'
                    # pasta de divulgação usada como fonte tem de aparecer, não ficar
                    # implícita no caminho para quem souber ler
                    if rot in fora_de_lugar:
                        out += '<br><span class="chip bad">pasta de divulgação</span>'
                return out
            fonte_html = "<br>".join(_uma(p) for p in fonte.split("\n"))
            linhas_html.append(
                f'<tr><td>{html.escape(lab)}</td><td class="num">{calc}</td>'
                f'<td class="num">{alvo}</td><td>{chip}</td>'
                f'<td class="src">{fonte_html}</td><td class="obs">{regra}</td></tr>')

    avisos_html = ""
    if avisos:
        avisos_html = ('<div class="nota"><b>O relatório desta semana:</b><ul>'
                       + "".join(f"<li>{a}</li>" for a in avisos) + "</ul></div>")

    gerado = datetime.now().strftime("%d/%m/%Y %H:%M")
    # o snapshot nao guarda o inicio da janela; ela vai do fechamento anterior + 1 dia
    anteriores = [w for w in sorted(hist) if R._is_iso(w) and w < semana]
    if anteriores:
        d = datetime.strptime(anteriores[-1], "%Y-%m-%d") + timedelta(days=1)
        ini = d.strftime("%d/%m")
    else:
        ini = "?"
    janela = f"{ini} – {datetime.strptime(semana, '%Y-%m-%d').strftime('%d/%m/%Y')}"
    doc = TEMPLATE.format(
        semana=semana, janela=html.escape(janela), gerado=gerado,
        n_ok=n_ok, n_nao=n_tot - n_ok, n_tot=n_tot,
        n_extra=len(SEM_CONTRAPARTE), linhas="\n".join(linhas_html), avisos=avisos_html,
        comite=_secao_comite())
    SAIDA.parent.mkdir(parents=True, exist_ok=True)
    SAIDA.write_text(doc, encoding="utf-8")
    print(f"[auditoria] semana {semana} · {n_ok}/{n_tot} batem · "
          f"{len(SEM_CONTRAPARTE)} sem contraparte -> {SAIDA.relative_to(ROOT)}")


TEMPLATE = """<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Auditoria de Fontes · {semana}</title>
<style>
:root{{--ground:#F4F6F3;--surface:#fff;--surface-alt:#ECEFEA;--line:#D5DBD3;--line-soft:#E4E8E2;
--ink:#1B211D;--ink-soft:#4E5852;--ink-mute:#77827B;--accent:#8C4A2F;--accent-soft:#F0E2DA;
--ok:#2E6F4E;--ok-bg:#E2EFE7;--bad:#9C352F;--bad-bg:#F5E1DF;--zebra:#FAFBF9;
--serif:"Iowan Old Style","Palatino Linotype",Palatino,Georgia,serif;
--sans:system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
--mono:ui-monospace,"SF Mono",Menlo,Consolas,monospace}}
@media(prefers-color-scheme:dark){{:root:not([data-theme=light]){{--ground:#121614;--surface:#1A1F1C;
--surface-alt:#222824;--line:#333B36;--line-soft:#272E2A;--ink:#E6EAE5;--ink-soft:#AFB8B1;
--ink-mute:#838D86;--accent:#CE8462;--accent-soft:#35231B;--ok:#6FBE8F;--ok-bg:#1B2E23;
--bad:#E08078;--bad-bg:#33201E;--zebra:#1D2320}}}}
:root[data-theme=dark]{{--ground:#121614;--surface:#1A1F1C;--surface-alt:#222824;--line:#333B36;
--line-soft:#272E2A;--ink:#E6EAE5;--ink-soft:#AFB8B1;--ink-mute:#838D86;--accent:#CE8462;
--accent-soft:#35231B;--ok:#6FBE8F;--ok-bg:#1B2E23;--bad:#E08078;--bad-bg:#33201E;--zebra:#1D2320}}
*{{box-sizing:border-box}}
body{{background:var(--ground);color:var(--ink);font-family:var(--sans);font-size:16px;
line-height:1.55;margin:0;padding:0 20px 64px;-webkit-font-smoothing:antialiased}}
.wrap{{max-width:1280px;margin:0 auto}}
header{{display:flex;flex-direction:column;gap:10px;padding:44px 0 22px;border-bottom:2px solid var(--ink)}}
.eyebrow{{font-family:var(--mono);font-size:11px;letter-spacing:.16em;text-transform:uppercase;color:var(--accent)}}
h1{{font-family:var(--serif);font-size:clamp(1.9rem,5vw,2.7rem);line-height:1.08;font-weight:600;margin:0}}
.stamp{{display:flex;flex-wrap:wrap;gap:6px 22px;font-family:var(--mono);font-size:12px;color:var(--ink-mute)}}
.stamp b{{color:var(--ink-soft);font-weight:600}}
section{{padding-top:36px}}
h2{{font-family:var(--serif);font-size:1.45rem;font-weight:600;margin:0 0 14px;
padding-bottom:8px;border-bottom:1px solid var(--line)}}
.tiles{{display:grid;grid-template-columns:repeat(auto-fit,minmax(148px,1fr));gap:12px;margin-bottom:20px}}
.tile{{background:var(--surface);border:1px solid var(--line);border-top:3px solid var(--h,var(--ink-mute));
padding:14px 16px 16px;display:flex;flex-direction:column;gap:2px}}
.tile .num{{font-family:var(--mono);font-size:1.9rem;font-weight:600;line-height:1.1;color:var(--h,var(--ink));
font-variant-numeric:tabular-nums}}
.tile .lab{{font-size:.8rem;color:var(--ink-soft);line-height:1.4}}
.tile.ok{{--h:var(--ok)}}.tile.bad{{--h:var(--bad)}}.tile.neu{{--h:var(--accent)}}
.scroll{{overflow-x:auto;border:1px solid var(--line);background:var(--surface)}}
table{{border-collapse:collapse;width:100%;font-size:.86rem;min-width:1080px}}
thead th{{text-align:left;font-family:var(--mono);font-size:10.5px;letter-spacing:.1em;text-transform:uppercase;
color:var(--ink-mute);font-weight:600;padding:10px 13px;background:var(--surface-alt);
border-bottom:1px solid var(--line);white-space:nowrap}}
tbody td{{padding:9px 13px;border-bottom:1px solid var(--line-soft);vertical-align:top}}
tbody tr:nth-child(even){{background:var(--zebra)}}
td.num{{font-family:var(--mono);font-variant-numeric:tabular-nums;white-space:nowrap}}
td.src{{font-family:var(--mono);font-size:.77rem;line-height:1.5;white-space:nowrap}}
td.src .aba{{color:var(--ink-mute)}}
td.src .cam{{display:block;color:var(--ink-mute);font-size:.72rem;white-space:normal;
word-break:break-word;max-width:34ch;margin-top:2px}}
td.obs{{color:var(--ink-soft);font-size:.84rem;min-width:260px}}
.grp td{{background:var(--surface-alt);font-family:var(--mono);font-size:10px;letter-spacing:.12em;
text-transform:uppercase;color:var(--accent);padding:7px 13px;font-weight:600}}
tbody tr.grp:nth-child(even){{background:var(--surface-alt)}}
.chip{{display:inline-block;font-family:var(--mono);font-size:10px;letter-spacing:.06em;text-transform:uppercase;
padding:3px 7px;border-radius:2px;white-space:nowrap;font-weight:600}}
.chip.ok{{color:var(--ok);background:var(--ok-bg)}}
.chip.bad{{color:var(--bad);background:var(--bad-bg)}}
.chip.neu{{color:var(--ink-mute);background:var(--surface-alt)}}
.nota{{border-left:3px solid var(--accent);background:var(--accent-soft);padding:12px 16px;
font-size:.88rem;color:var(--ink-soft);margin-top:18px}}
.nota b{{color:var(--ink)}} .nota ul{{margin:6px 0 0;padding-left:18px}}
code{{font-family:var(--mono);font-size:.88em}}
footer{{margin-top:44px;padding-top:16px;border-top:1px solid var(--line);font-family:var(--mono);
font-size:11.5px;color:var(--ink-mute)}}
</style></head><body><div class="wrap">
<header>
  <span class="eyebrow">Haras Pao Grande &middot; fechamento semanal e comitê mensal</span>
  <h1>Auditoria de Fontes</h1>
  <div class="stamp">
    <span>Semana <b>{semana}</b></span>
    <span>Janela <b>{janela}</b></span>
    <span>Gerada em <b>{gerado}</b></span>
  </div>
</header>
<section>
  <h2>Fechamento semanal</h2>
  <div class="tiles">
    <div class="tile ok"><span class="num">{n_ok}</span><span class="lab">batem, de {n_tot} comparáveis</span></div>
    <div class="tile bad"><span class="num">{n_nao}</span><span class="lab">não batem</span></div>
    <div class="tile neu"><span class="num">{n_extra}</span><span class="lab">sem linha no relatório</span></div>
  </div>
  <div class="scroll"><table>
    <thead><tr><th>Indicador</th><th>Calc.</th><th>Relat.</th><th>Situação</th>
    <th>Fonte · aba</th><th>Regra</th></tr></thead>
    <tbody>
{linhas}
    </tbody>
  </table></div>
  {avisos}
</section>
{comite}
<footer>Gerada por tools/build_auditoria.py a partir do snapshot congelado e do spec do
comitê — não digitada à mão.</footer>
</div></body></html>
"""


def main():
    arg = next((a for a in sys.argv[1:] if not a.startswith("--")), None)
    semana = R._parse_d(arg).isoformat() if arg else None
    build(semana)


if __name__ == "__main__":
    main()
