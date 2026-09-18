"""
PGSemanalReport — gera a ATUALIZAÇÃO SEMANAL do Haras Pão Grande a partir das
fontes REORGANIZADAS do Drive (layout 2026-07).

Substitui o pipeline antigo (PGSemanalExtractor/_pg_semanal), que apontava pra
paths e modelos de planilha que não existem mais.

Fontes (ver memória project-hpg-semanal):
  - Headcount   : ATUALIZACAO SEMANAL/CONTROLE PLANTEL.xlsx  aba CONTAGEM (pré-agregada)
  - Receptoras  : idem, aba 'RECEPTORAS ' (censo; filtro do report a definir)
  - Produção    : REPRODUÇÃO/ESTAÇÃO DE MONTA/Estação 2025-2026/{YYMMDD} ESTACAO DE MONTA.xlsx
                  aba ESTAÇÃO (confirmado = coluna '+ / -' == 'OK'; data = TE + 15 dias)
  - Movimentação: PLANILHAS SEMANAIS/SAIDA E ENTRADA DE ANIMAIS - MODELO ENVIAR NO GRUPO.xlsx
  - Comerciais  : REPRODUÇÃO/EMBRIOES A ENTREGAR - A RECEBER.xlsx  abas PAINEL/ENTREGAR/RECEBER

Uso:
    python PGSemanalReport.py                 # semana termina hoje
    python PGSemanalReport.py 17/07/2026       # semana termina nessa data
    python PGSemanalReport.py 03/07/2026 17/07/2026   # janela explícita

Saída:
    - imprime as 5 seções no console, comparadas aos alvos do docx 17-07 (quando aplicável)
    - grava semanal_data.json (consumido pelo dashboard HTML)
"""

from __future__ import annotations

import json
import re
import sys
import unicodedata
from collections import Counter
from dataclasses import dataclass, field, asdict
from datetime import date, datetime, timedelta
from pathlib import Path

import openpyxl

from _pg_cancelamentos import pendencias as _cancel_pendencias
from _pg_common import DRIVE_ROOT, ensure_cache

BASE_DIR = Path(__file__).resolve().parent.parent  # raiz do projeto (scripts/ fica 1 nível abaixo)

# ------------------------------------------------------------------
# Localização das fontes (layout novo do Drive)
# ------------------------------------------------------------------
# ATUALIZACAO SEMANAL é a pasta de SAÍDA do fechamento — é onde o relatório em Word é
# publicado. Puxar dado de lá é ler do lugar errado: as três planilhas que moram ali
# pertencem, por assunto, a PLANTEL, REPRODUÇÃO e VENDAS.
#
# Foi feita a varredura por substituto de CONTEÚDO e não existe: o acumulado de 61 não
# aparece em nenhuma aba da estação de monta (PLANEJAMENTO e REC. EMBR. dão 56, RESUMO
# está com as fórmulas em #REF!) nem no controle mensal (aba EMBRIOES PG é financeira;
# os 65 embriões do roster estão com a coluna SAFRA vazia); e pendência de sociedade
# não é marcada em lugar nenhum — o MAPA VENDAS é histórico de vendas, sem status.
#
# Então o que falta é MUDAR OS ARQUIVOS DE PASTA, não trocar de fonte. Cada caminho
# abaixo é uma LISTA em ordem de preferência: a pasta canônica primeiro, ATUALIZACAO
# SEMANAL só como último recurso e com aviso no run. No dia em que o haras mover os
# arquivos, o pipeline segue sozinho — e enquanto não mover, o run diz o que falta.
FALLBACK_DIR = DRIVE_ROOT / "ATUALIZACAO SEMANAL"

# Roster do plantel: assunto é PLANTEL.
CONTROLE_PLANTEL_DIRS = (DRIVE_ROOT / "PLANTEL", FALLBACK_DIR)
CONTROLE_PLANTEL_GLOB = "CONTROLE PLANTEL.xlsx"
# "EMBRIÕES E MATRIZES - MODELO ENVIAR NO GRUPO 3 <DD-MM>.xlsx" — fonte do acumulado na
# estação. Assunto é REPRODUÇÃO; já morou em PLANILHAS SEMANAIS. Sufixo de data muda
# (e mente: a cópia viva se chama "30-05" e é de agosto), então resolve por glob+mtime.
EMB_MATRIZES_DIRS = (
    DRIVE_ROOT / "REPRODUÇÃO" / "ESTAÇÃO DE MONTA",
    DRIVE_ROOT / "REPRODUÇÃO",
    FALLBACK_DIR,
)
EMB_MATRIZES_GLOB = "EMBRI*E MATRIZES*.xlsx"
# Animais para sair (hoje só sociedade pendente) — assunto é VENDAS. A pasta
# PLANILHAS PARA O EDUARDO saiu da lista: não tem nenhum arquivo desse nome.
ANIMAIS_SAIR_DIRS = (
    DRIVE_ROOT / "VENDAS" / "SAIDA DE ANIMAIS VENDIDOS",
    FALLBACK_DIR,
)
ANIMAIS_SAIR_GLOB = "Animais para sair*.xlsx"

EMB_COMERCIAIS = DRIVE_ROOT / "REPRODUÇÃO" / "EMBRIOES A ENTREGAR - A RECEBER.xlsx"
ESTACAO_MONTA_BASE = DRIVE_ROOT / "REPRODUÇÃO" / "ESTAÇÃO DE MONTA"
# O plantel e as receptoras vivem em PLANTEL/Estação <ano>-<ano>, e a copia de
# trabalho MUDA DE PASTA quando a estacao vira: em 21/08/2026 os arquivos "EDITAR
# SETEMBRO" passaram para "Estação 2026-2027", porque setembro abre estacao nova.
# Olhar so a pasta da estacao corrente fez o pipeline concluir que os arquivos tinham
# sido apagados e cair numa copia congelada de 05/08. Varremos TODAS as pastas de
# estacao e ficamos com o mais recente — aqui frescor e o que importa, e a guarda de
# fonte velha cobre o resto.
PLANTEL_DIR_BASE = DRIVE_ROOT / "PLANTEL"
PLANTEL_ESTACAO_GLOB = "Estação *"
RECEPTORAS_DIR = PLANTEL_DIR_BASE / "Estação 2025-2026"   # so p/ mensagens de erro
# Mapa de Vendas — quem consome é o deck do comitê (tools/build_comite.py),
# não o fechamento semanal. Cheguei a tratar como constante morta por não achar
# uso neste módulo; o uso está no outro.
MAPA_VENDAS_DIR = DRIVE_ROOT / "VENDAS" / "MAPAS DE VENDAS" / "Estação 2025-2026"
# CONTROLE_DE_PLANTEL mensal (STATUS PLANTEL, SAIDAS-ENTRADAS, MOVIMENTAÇÕES)
CONTROLE_MENSAL_DIR = RECEPTORAS_DIR

HIST_HEADCOUNT = BASE_DIR / "_cache" / "headcount_history.json"
HIST_SNAPSHOTS = BASE_DIR / "_cache" / "semanal_snapshots.json"

# Virada de estacao ANCORADA NA DATA da liberacao em que a safra nova passa a
# valer: a interseason (as duas safras convivendo no relatorio) acaba na PRIMEIRA
# liberacao de setembro — 04/09/2026, combinado com o Arthur nessa data. Era uma
# constante editada a mao (mexida em 31/08/2026), o que fazia a safra do
# relatorio depender de quando alguem editou este arquivo, e nao da semana que
# esta sendo fechada.
VIRADAS_DE_SAFRA = (
    (date(2026, 9, 4), "2026/2027"),
    (date(2025, 9, 5), "2025/2026"),
)


def safra_vigente(quando: date | None = None) -> str:
    d = quando or date.today()
    for inicio, safra in VIRADAS_DE_SAFRA:
        if d >= inicio:
            return safra
    return VIRADAS_DE_SAFRA[-1][1]


SAFRA_ATUAL = safra_vigente()
# SAFRA_PROXIMA=None ate a proxima virada comecar a ser anunciada — com None, o
# card 'acumulado_estacao_proxima' some sozinho (ver skip em
# PGSemanalDashboard.py) em vez de nascer zerado.
SAFRA_PROXIMA = None


def _rotulo_safra(safra: str | None) -> str | None:
    """'2025/2026' -> '25/26', que e como o relatorio escreve. None (sem safra
    de transicao) passa direto — quem le isso trata None como 'sem rotulo'."""
    if not safra:
        return None
    a, b = safra.split("/")
    return f"{a[-2:]}/{b[-2:]}"
BASES_DIR = BASE_DIR / "bases"
JSON_OUT = BASES_DIR / "semanal_data.json"

# Alvo da validação em tela = a liberação do haras DAQUELA semana (o docx cuja
# data de referência é o fim da janela). Era um dicionário fixo com os números
# de 17-07-26, de quando aquele era o único relatório parseado: em 04/09/2026 a
# tela cobrava headcount 206 e receptoras 63/36/27 de julho contra o cálculo de
# setembro, e todo item aparecia divergente sem ter divergência nenhuma. Semana
# ainda não liberada não tem alvo — melhor sem comparação do que comparando com
# a semana errada.
def _docx_proximo(rep):
    """Relatório publicado com data um pouco depois do nosso fechamento.

    O haras nem sempre publica no dia do fechamento: a semana fechada em
    10/09/2026 (quinta) saiu num arquivo datado 11/09. Sem isto o placar dizia
    'sem relatório oficial nessa data' e a semana ficava sem conferência nenhuma.
    Aceita até 3 dias DEPOIS — nunca antes, que seria comparar com a semana
    passada."""
    from datetime import timedelta
    ref = date.fromisoformat(rep.semana_atual)
    for dias in (1, 2, 3):
        w = (rep.docx_ref or {}).get((ref + timedelta(days=dias)).isoformat())
        if w:
            print(f"  [placar] usando o relatório de {(ref + timedelta(days=dias)).strftime('%d/%m')} "
                  f"para conferir a semana que fechou em {ref.strftime('%d/%m')}")
            return w
    return None


def _alvos(rep) -> dict:
    w = (rep.docx_ref or {}).get(rep.semana_atual) or _docx_proximo(rep)
    if not w:
        return {}
    pr, rc, hc, sa = (w.get(k) or {} for k in ("producao", "receptoras", "headcount", "saidas"))
    return {
        "acumulado_estacao": pr.get("acumulado_estacao"),
        "confirmados_semana": pr.get("confirmados_semana"),
        "nascimentos": pr.get("nascimentos"),
        "abortos_obitos": pr.get("abortos_obitos"),
        "receptoras_total": rc.get("total"),
        "receptoras_prenhas": rc.get("prenhas"),
        "receptoras_vazias": rc.get("vazias"),
        "headcount_total": hc.get("total"),
        "headcount_fpg": hc.get("fazenda_pg"),
        "headcount_arr": hc.get("arrendamento"),
        "headcount_cte": hc.get("cte"),
        "headcount_soc": hc.get("socio"),
        "saidas_semana": sa.get("saidas_semana"),
        "vendidos_pendentes": sa.get("vendidos_pendentes"),
        "sociedade_pendentes": sa.get("sociedade_pendentes"),
    }


# ------------------------------------------------------------------
# Helpers
# ------------------------------------------------------------------
def _norm(v) -> str:
    return str(v).strip().upper() if v is not None else ""


def _s(v):
    if v is None:
        return None
    s = str(v).strip()
    return s or None


def _dt(v):
    if isinstance(v, datetime):
        return v.date()
    return None


def _to_num(v):
    try:
        return float(v) if v is not None else None
    except (TypeError, ValueError):
        return None


def _load(path: Path):
    cached = ensure_cache(path)
    return openpyxl.load_workbook(cached, data_only=True, read_only=True)


# {rotulo: Path} das fontes efetivamente abertas neste run — base do aviso de fonte
# velha. Sem isso, apagar a copia de trabalho faz o pipeline cair numa versao
# congelada semanas atras SEM UM RUIDO: em 21/08/2026 os dois arquivos "EDITAR
# SETEMBRO" foram apagados no meio do dia e o fechamento passou a ler receptoras de
# 05/08 e roster mensal de 13/08, publicando headcount 205 no lugar de 202.
_FONTES_USADAS: dict = {}


# Pasta de DIVULGAÇÃO, não de dado: é onde o haras deixa o que foi enviado ao grupo,
# e serve para conferir o calculado contra o publicado. Três arquivos vivos só existem
# lá hoje, então ler de lá é a situação atual — mas tem de doer, não passar batido.
FONTES_FORA_DE_LUGAR: set = set()


def _registra_fonte(rotulo: str, f: Path) -> Path:
    _FONTES_USADAS[rotulo] = f
    try:
        if FALLBACK_DIR in Path(f).parents:
            if rotulo not in FONTES_FORA_DE_LUGAR:
                print(f"  [fonte] {rotulo}: lendo de '{FALLBACK_DIR.name}', que é pasta de "
                      f"divulgação, não de dado ({Path(f).name}). Fora dela a cópia mais "
                      f"nova é de 2025 — o arquivo precisa ser movido na origem.")
            FONTES_FORA_DE_LUGAR.add(rotulo)
    except Exception:
        pass
    return f


def caminho_curto(f) -> str:
    """Caminho legível: relativo a 'PLANILHAS DE CONTROLE' quando vem do Drive.

    A raiz é um atalho (`G:\\.shortcut-targets-by-id\\1mBrSez...`) — mostrar isso
    inteiro só polui. O que identifica a fonte é a pasta dentro do Drive, porque o
    mesmo nome de arquivo existe em mais de uma (a estação de monta muda de pasta a
    cada safra). Fora do Drive, devolve o caminho como está."""
    p = Path(f)
    try:
        return str(p.relative_to(DRIVE_ROOT)).replace("\\", "/")
    except ValueError:
        try:
            return str(p.relative_to(BASE_DIR)).replace("\\", "/")
        except ValueError:
            return str(p)


# so o orquestrador libera, via --forcar
PERMITIR_FONTE_VELHA = False
# Fontes de ESTADO: descrevem quem esta onde AGORA e por isso mudam toda semana.
# Paradas, a fonte esta perdida (mudou de pasta, foi renomeada, apagada) e o
# fechamento nao pode publicar o retrato de outra semana como se fosse esta.
#
# Estacao de monta e 'EMBRIÕES E MATRIZES' NAO entram aqui, e isso foi corrigido em
# 10/09/2026: elas sao fonte de EVENTO — sem IA, confirmacao, parição ou aborto na
# semana, nao ha por que salvar o arquivo, e mtime parado significa "nada
# aconteceu", nao "dado faltando". Naquele fechamento as duas estavam em 04/09 e a
# liberacao do haras dizia exatamente a mesma coisa: confirmados '--',
# nascimentos '--', abortos '--', acumulado parado em 01. O bloqueio travou um
# fechamento correto. Para essas o aviso sai e o run segue — quem confirma que
# nada aconteceu e o confronto com a liberacao, no placar.
FONTES_SEMANAIS = ("receptoras", "controle mensal", "roster do plantel")


def _avisar_fontes_velhas(ini: date, fim: date):
    """Fonte salva antes do inicio da janela nao pode descrever esta semana.

    BLOQUEIA antes de congelar o snapshot: avisar depois nao serve, porque o numero
    errado ja foi publicado. Em 21/08/2026 isso aconteceu duas vezes seguidas."""
    velhas = []
    for rotulo, f in sorted(_FONTES_USADAS.items()):
        try:
            m = datetime.fromtimestamp(f.stat().st_mtime).date()
        except OSError:
            continue
        if m < ini:
            velhas.append((rotulo, f.name, m))
    if not velhas:
        return
    bloqueiam = [v for v in velhas if v[0] in FONTES_SEMANAIS]
    print(f"  [fontes] !! {len(velhas)} fonte(s) mais VELHAS que a janela "
          f"({ini.strftime('%d/%m')}-{fim.strftime('%d/%m')}) — o que sai delas nao "
          f"descreve esta semana:")
    for rotulo, nome, m in velhas:
        marca = ("  <- estado da semana, BLOQUEIA" if rotulo in FONTES_SEMANAIS
                 else "  (fonte de evento: parada = nada aconteceu; conferir no placar)")
        print(f"    - {rotulo}: {nome} (salvo em {m.strftime('%d/%m/%Y')}){marca}")
    if not bloqueiam:
        print("    Nenhuma delas descreve estado — segue, e o placar confronta com a liberação.")
        return
    print("    Conferir se a copia de trabalho mudou de pasta, foi apagada ou renomeada.")
    if not PERMITIR_FONTE_VELHA:
        raise RuntimeError(
            "fonte(s) semanal(is) mais velha(s) que a janela: "
            + "; ".join(f"{r} ({n}, {m:%d/%m})" for r, n, m in bloqueiam)
            + " — snapshot NAO congelado. Use --forcar para gravar assim mesmo.")


def _latest_by_mtime(folder: Path, pattern: str) -> Path:
    """Arquivo mais recente por MTIME (data de modificação real). Inclui as cópias
    'EDITAR ...' — o operador trabalha nelas (versão viva), então são as MAIS frescas.
    Os arquivos {YYMMDD} congelados são snapshots antigos."""
    cands = [f for f in folder.glob(pattern) if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(f"Nenhum arquivo {pattern} em {folder}")
    return max(cands, key=lambda f: f.stat().st_mtime)


# Arquivos ainda resolvidos na pasta de saída — o run avisa no fim.
_NA_PASTA_DE_SAIDA: list = []


def _tem_aba(f: Path, aba: str) -> bool:
    """Só os nomes das abas — não carrega célula nenhuma nem passa pelo cache."""
    try:
        wb = openpyxl.load_workbook(f, read_only=True)
    except Exception:
        return False
    try:
        return aba in wb.sheetnames
    finally:
        wb.close()


def _resolver(pattern: str, dirs, rotulo: str, requer_aba: str | None = None) -> Path:
    """Primeiro diretório da lista que tenha o arquivo; dentro dele, o mais recente por
    mtime. A ordem é intencional (pasta canônica antes do fallback), então NÃO compare
    mtime entre diretórios: a cópia velha no lugar certo ganha da nova no lugar errado
    — é assim que a migração acontece sozinha quando alguém move o arquivo.

    `requer_aba` existe porque casar o nome não basta. O glob do Windows é
    case-insensitive e 'Animais para sair*.xlsx' casa o legado 'ANIMAIS PARA SAIR OU
    BUSCAR - ATUALIZADA 02-01-25.xlsx', de 2025 e com outro layout — a pasta canônica
    tinha um homônimo velho. Candidato sem a aba esperada é descartado, com aviso: um
    arquivo com o nome certo e a estrutura errada não pode virar fonte em silêncio."""
    tentadas, recusados = [], []
    for d in dirs:
        tentadas.append(str(d))
        if not d.exists():
            continue
        cands = sorted((f for f in d.glob(pattern) if not f.name.startswith("~$")),
                       key=lambda f: f.stat().st_mtime, reverse=True)
        for f in cands:
            if requer_aba and not _tem_aba(f, requer_aba):
                recusados.append(f"{f.name} (sem aba {requer_aba!r})")
                continue
            if recusados:
                print(f"  [fontes] {rotulo}: ignorado(s) {'; '.join(recusados)}")
            # o roster é resolvido 5x no run (headcount, doadoras, conferência...);
            # o aviso é sobre o ARQUIVO, então registra uma vez só
            if d == FALLBACK_DIR and (rotulo, f.name) not in {
                    (r, n) for r, n, _ in _NA_PASTA_DE_SAIDA}:
                _NA_PASTA_DE_SAIDA.append((rotulo, f.name, dirs[0]))
            return _registra_fonte(rotulo, f)
    detalhe = (" | recusados: " + "; ".join(recusados)) if recusados else ""
    raise FileNotFoundError(
        f"Nenhum {pattern!r} ({rotulo}) em: " + " | ".join(tentadas) + detalhe)


def _avisar_pasta_de_saida():
    """Fecha o run listando o que ainda sai da pasta de publicação e pra onde deveria
    ir. Sem isso a dependência some de vista e ninguém move o arquivo."""
    if not _NA_PASTA_DE_SAIDA:
        return
    print(f"  [fontes] {len(_NA_PASTA_DE_SAIDA)} arquivo(s) ainda lidos de "
          f"{FALLBACK_DIR.name} (pasta de publicação, não de dado):")
    for rotulo, nome, destino in _NA_PASTA_DE_SAIDA:
        print(f"    - {rotulo}: {nome}  ->  mover para {destino.name}")


def _controle_plantel() -> Path:
    return _resolver(CONTROLE_PLANTEL_GLOB, CONTROLE_PLANTEL_DIRS, "roster do plantel",
                     requer_aba="PLANTEL")


def _estacao_dirs() -> list:
    """Pastas de estacao do PLANTEL, da mais nova para a mais velha."""
    if not PLANTEL_DIR_BASE.exists():
        return []
    return sorted((d for d in PLANTEL_DIR_BASE.glob(PLANTEL_ESTACAO_GLOB) if d.is_dir()),
                  reverse=True)


def _latest_no_plantel(pattern: str, rotulo: str) -> Path:
    """Arquivo que casa o padrao em QUALQUER pasta de estacao, preferindo a CÓPIA
    DE TRABALHO.

    O mesmo mês vive em duas cópias irmãs: '..._AGO_26.xlsx' é o FECHAMENTO de
    agosto, congelado, e '..._EDITAR OUTUBRO_..._AGO_26.xlsx' é onde o haras
    lança o que acontece agora. Escolher por mtime pega quem salvou por último e
    troca de fonte no meio do caminho: em 04/09/2026 a semana fechava 24/24 pela
    cópia de trabalho e, depois de alguém salvar o fechamento de agosto, a MESMA
    semana passou a 16/24 — sem NASDAQ (01/09), sem a saída da POTRA MORENA
    (01/09) nem da ADRENALINA (04/09), porque nenhuma delas é evento de agosto.
    Semanal precisa do estado de HOJE, então cópia de trabalho manda; o
    fechamento mensal é o oposto e por isso o módulo de plantel ignora as
    'EDITAR' (ver tools/seed_plantel_hub.py).
    """
    cands = [f for d in _estacao_dirs() for f in d.glob(pattern)
             if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(
            f"Nenhum {pattern!r} ({rotulo}) em: "
            + " | ".join(str(d) for d in _estacao_dirs()))
    trabalho = [f for f in cands if "EDITAR" in f.name.upper()]
    return _registra_fonte(rotulo, max(trabalho or cands, key=lambda f: f.stat().st_mtime))


def _latest_estacao_master() -> Path:
    """Master da estação de monta, na pasta 'Estação <ano>-<ano>' certa.

    Era fixo em 'Estação 2025-2026'. Na virada de safra o haras passa a atualizar o
    arquivo em 'Estação 2026-2027', e esse caminho fixo nunca ia ver — igual ao bug
    já corrigido pro roster/receptoras (ver comentário de PLANTEL_ESTACAO_GLOB).
    Mesma solução: varre TODAS as pastas de safra e fica com a mais nova por mtime."""
    if not ESTACAO_MONTA_BASE.exists():
        raise FileNotFoundError(f"Pasta não encontrada: {ESTACAO_MONTA_BASE}")
    dirs = sorted((d for d in ESTACAO_MONTA_BASE.glob("Estação *") if d.is_dir()),
                  reverse=True)
    cands = [f for d in dirs for f in d.glob("*ESTACAO DE MONTA.xlsx")
             if not f.name.startswith("~$")]
    if not cands:
        raise FileNotFoundError(
            "Nenhum '*ESTACAO DE MONTA.xlsx' (estacao de monta) em: "
            + " | ".join(str(d) for d in dirs))
    return _registra_fonte("estacao de monta", max(cands, key=lambda f: f.stat().st_mtime))


def _latest_by_yymmdd(folder: Path, pattern: str, rotulo: str | None = None) -> Path:
    """Mais recente por mtime, registrado sob `rotulo`.

    O rótulo era adivinhado pelo padrão — RECEPTORAS virava 'receptoras', o resto
    virava 'controle mensal'. O comitê chama esta mesma função para o mapa de vendas,
    e o registro passava a dizer que o controle mensal era o mapa de vendas. Rótulo
    errado não muda número, mas vai direto para a auditoria, que serve exatamente
    para dizer de onde veio cada dado."""
    if rotulo is None:
        rotulo = "receptoras" if "RECEPTORAS" in pattern.upper() else "controle mensal"
    return _registra_fonte(rotulo, _latest_by_mtime(folder, pattern))


# ------------------------------------------------------------------
# Estrutura do relatório
# ------------------------------------------------------------------
@dataclass
class Report:
    semana_inicio: str
    semana_fim: str
    fontes: dict = field(default_factory=dict)
    fontes_caminhos: dict = field(default_factory=dict)  # rótulo -> caminho no Drive
    fontes_fora_de_lugar: list = field(default_factory=list)  # lidas da pasta de divulgação
    producao: dict = field(default_factory=dict)
    receptoras: dict = field(default_factory=dict)
    headcount: dict = field(default_factory=dict)
    terceiros: dict = field(default_factory=dict)
    saidas: dict = field(default_factory=dict)
    detalhe: dict = field(default_factory=dict)  # tabelas de detalhe p/ o dashboard
    eventos: dict = field(default_factory=dict)  # listas COMPLETAS datadas (filtro client-side)
    calendario: list = field(default_factory=list)  # semanas travadas p/ o seletor
    semana_atual: str = ""  # id (segunda) da semana selecionada por padrão
    snapshots: dict = field(default_factory=dict)  # snapshot por semana (histórico local)
    roster: list = field(default_factory=list)  # nomes do plantel (p/ diff saídas/entradas)
    receptoras_locais: dict = field(default_factory=dict)  # {animal: LOCAL} p/ diff de transferência
    populacao: list = field(default_factory=list)  # plantel + receptoras contadas (p/ diff saídas/entradas)
    saidas_planilha: dict | None = None  # aba SAIDAS-ENTRADAS, quando o haras preencher
    confirmed: list = field(default_factory=list)  # embriões confirmados (+/-=OK) p/ diff semanal
    docx_ref: dict = field(default_factory=dict)  # números dos relatórios oficiais (SÓ p/ validar)


# ------------------------------------------------------------------
# Seção 1 — PRODUÇÃO (master da estação de monta)
# ------------------------------------------------------------------
# Abas da planilha do grupo que compõem o acumulado.
#   (aba, fatia fixa ou None = deduzir do STATUS)
# As COLUNAS são resolvidas pelo CABEÇALHO (linha 3), nunca por índice fixo. Em
# 07/08/2026 a coluna STATUS foi apagada da aba de PAO GRANDE, tudo à direita
# andou uma casa e o índice de ESTAÇÃO passou a cair em COTA PG: nenhuma linha
# batia a safra, a fatia "pg" zerou e o acumulado caiu de 61 pra 27 sem erro
# nenhum. Cabeçalho por nome faz a planilha poder mexer coluna sem quebrar.
ABAS_ACUMULADO = (
    ("EMBRIÕES PAO GRANDE", "pg"),
    ("EMBRIOES SOCIOS - VENDIDOS", None),
)
HDR_ROW_ACUMULADO = 3  # linha do cabeçalho nas duas abas


def _sem_acento(s: str) -> str:
    return "".join(c for c in unicodedata.normalize("NFD", str(s))
                   if unicodedata.category(c) != "Mn")


def _col_idx(hdr, *nomes) -> int:
    """Índice da coluna pelo nome do cabeçalho (sem acento, case-insensitive).
    Explode se não achar — coluna que sumiu tem de virar erro, não zero calado."""
    alvo = {_sem_acento(n).strip().upper() for n in nomes}
    for i, h in enumerate(hdr):
        if h is not None and _sem_acento(h).strip().upper() in alvo:
            return i
    achados = [str(h) for h in hdr if h is not None]
    raise KeyError(f"coluna {nomes[0]!r} não está no cabeçalho: {achados}")


def _latest_emb_matrizes() -> Path:
    return _resolver(EMB_MATRIZES_GLOB, EMB_MATRIZES_DIRS, "acumulado na estação",
                     requer_aba="EMBRIÕES PAO GRANDE")


def _acumulado_grupo(safra: str = SAFRA_ATUAL) -> dict:
    """Embriões vivos da safra na planilha 'EMBRIÕES E MATRIZES' — a que o haras
    manda no grupo e usa como fonte oficial do acumulado.

    Ela é um retrato do que está EM PÉ: a linha é apagada quando o embrião sai da
    conta, seja por parição (vira potro) ou por aborto. Por isso o acumulado da
    estação = linhas vivas + parições da safra. Aborto NÃO volta — é baixa
    definitiva, confirmado com o haras em 2026-07-31 (os 3 abortos da safra também
    sumiram da planilha e não entram no número oficial).

    A aba ICSI fica de fora: hoje só tem safra 2023/2024."""
    src = _latest_emb_matrizes()
    wb = _load(src)
    out = {"pg": 0, "socio": 0, "vendido": 0, "fonte": src.name,
           "chaves": set(), "linhas": []}
    for aba, fatia_fixa in ABAS_ACUMULADO:
        if aba not in wb.sheetnames:
            print(f"  [acumulado] aba {aba!r} não existe em {src.name} — fatia zerada")
            continue
        iest = ist = idoa = igar = irec = icot = isoc = None
        for i, r in enumerate(wb[aba].iter_rows(values_only=True), start=1):
            if i == HDR_ROW_ACUMULADO:
                iest = _col_idx(r, "ESTACAO")
                ist = None if fatia_fixa else _col_idx(r, "STATUS")
                # as duas abas têm as mesmas colunas em posições diferentes
                # (RECEPTORA é a 5ª numa e a 4ª na outra) — resolver por nome
                idoa, igar, irec = (_col_idx(r, "DOADORA"), _col_idx(r, "GARANHÃO"),
                                    _col_idx(r, "RECEPTORA"))
                # cota e sócio não existem nas duas abas; ausentes viram None
                def _opt(*nomes):
                    try:
                        return _col_idx(r, *nomes)
                    except KeyError:
                        return None
                icot, isoc = _opt("COTA PG"), _opt("COMPRADOR/SOCIO")
                continue
            if i <= HDR_ROW_ACUMULADO or r[1] is None or not str(r[1]).strip():
                continue
            if _s(r[iest]) != safra:
                continue
            fatia = fatia_fixa or ("vendido" if _norm(r[ist]) == "VENDIDO" else "socio")
            out[fatia] += 1
            chave = _chave_embriao(r[idoa], r[igar], r[irec])
            out["chaves"].add(chave)
            # a LINHA, não só a chave: quando o embrião pare, o haras apaga a linha e a
            # cota vai com ela. Guardada aqui, a fatia da parição é recuperável.
            out["linhas"].append({
                "chave": list(chave), "aba": aba, "fatia": fatia,
                "doadora": _s(r[idoa]), "garanhao": _s(r[igar]),
                "receptora": _s(r[irec]), "cota": (r[icot] if icot is not None else None),
                "socio": (_s(r[isoc]) if isoc is not None else None),
            })
    wb.close()
    out["total"] = out["pg"] + out["socio"] + out["vendido"]
    return out


# Animal bloqueado por pendência documental continua "pendente de saída" no sentido
# literal, mas o relatório NÃO o conta — 14/08 e 21/08 dizem "01 animal" onde temos
# MUSICA e NOBRE, e NOBRE é o único com essa observação nas duas semanas.
OBS_BLOQUEIA_SAIDA = ("FALTANDO EXAME",)


# Cotista é só CARLA ou EDUARDO. Entre parênteses também aparecem
# '(AGUARDANDO FICAR PRONTO)' e '(MAE X PAI)', que NÃO podem ser apagados.
RX_COTISTA = re.compile(r"\s*\((?:CARLA|EDUARDO)\)\s*")


def _sem_cotista(n) -> str:
    """'PARIS DA PAO GRANDE (EDUARDO)' -> 'PARIS DA PAO GRANDE'.

    O roster mensal repete a mesma linha por cotista, com o nome do cotista no
    nome. Sem tirar isso, um animal conta duas ou três vezes.

    O cotista NÃO está sempre no fim: em produto de embrião ele vem no meio,
    seguido de data e receptora — 'MACHO PROFESSORA RIOMINAS DA CACHOEIRA X
    FEITICO BANDEIRANTE (EDUARDO) 17/08/2024 RECEP 46'. Tirando só o parêntese
    final, as duas linhas (CARLA e EDUARDO) sobreviviam ao dedup e o potro
    entrava DUAS vezes no headcount. Achado em 04/09/2026 pelo Arthur, na
    conferência da lista em Excel."""
    t = RX_COTISTA.sub(" ", _norm(n))
    t = re.sub(r"\s*\([^)]*\)\s*$", "", t)
    return " ".join(t.split()).strip()


def _chave_animal(nome, mae, pai) -> tuple:
    """Identidade do animal no roster mensal: nome sem cotista + filiação.

    Nome sozinho junta dois potros distintos chamados `MACHO`. Com a data de
    nascimento na chave, `XARDA DO SALTO` viraria dois animais por causa de um dia
    de diferença digitado errado (20/09 e 21/09 de 2021)."""
    return (_sem_cotista(nome), _norm(mae), _norm(pai))


def _nucleo_nome(n) -> str:
    """Nome comparável entre planilhas. O roster chama o mesmo animal de
    'POTRA MORENA L2 X DAMASCO DA PAO GRANDE - 07/03/2025 RECEP 07 V' e o Animais para
    sair de 'FEMEA MORENA L2 X DAMASCO DA PAO GRANDE': muda o prefixo de sexo e sobra
    a data/receptora no fim. Sem normalizar, a reposição não é reconhecida e entra na
    conta de vendidos."""
    t = _norm(n)
    t = re.sub(r"^(POTRA|POTRO|FEMEA|FEMA|MACHO)\s+", "", t)
    t = re.sub(r"\s*-\s*\d.*$", "", t)
    t = re.sub(r"\s+RECEP.*$", "", t)
    return t.strip()


def _chave_embriao(doadora, garanhao, receptora) -> tuple:
    """Identidade do embrião entre a planilha do grupo e a aba ESTAÇÃO. Receptora
    normalizada porque vem '309' numa e 309.0 na outra."""
    def _rec(v):
        try:
            return str(int(float(v)))
        except (TypeError, ValueError):
            return _norm(v)
    return (_norm(doadora), _norm(garanhao), _rec(receptora))


def _mortes_do_plantel(ini: date, fim: date) -> list:
    """Óbitos registrados na aba 'CONFIRMAÇÕES, ABORTOS, MORTES' do controle mensal.

    Reunião 2026-07-31: a estação de monta só registra aborto/absorção, morte de
    receptora prenha e potro que nasce e morre em seguida. Óbito de animal já no
    plantel só aparece aqui. Colunas: B produto, C data, D observação (texto livre,
    por isso a classificação é por palavra-chave).
    """
    try:
        src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    except FileNotFoundError:
        return []
    wb = _load(src)
    aba = "CONFIRMAÇÕES, ABORTOS, MORTES"
    if aba not in wb.sheetnames:
        wb.close()
        return []
    achados = []
    for i, r in enumerate(wb[aba].iter_rows(values_only=True), start=1):
        if i < 3 or r[1] is None:
            continue
        d = _dt(r[2])
        if not d or not (ini <= d <= fim):
            continue
        obs = _norm(r[3])
        if "ABORT" in obs:            # aborto já vem da estação; aqui só morte
            continue
        if any(p in obs for p in ("MORREU", "MORTE", "OBITO", "MORTO")):
            achados.append({"animal": _s(r[1]), "data": d.isoformat(),
                            "ocorrencia": _s(r[3]), "origem": "plantel"})
    wb.close()
    return achados


def _acumulado_planejamento(wb) -> int:
    """Total da estação de monta = soma de 'TOTAL EMBRIÕES' REAL da aba PLANEJAMENTO,
    linhas de doadora (col0 numérica). É referência de conferência do acumulado.

    A coluna é achada pelo par de cabeçalhos — linha 2 diz PLAN/REAL e linha 3 diz o
    nome — e não por índice fixo. Estava travada em idx8, que na planilha da safra
    26/27 é o PLAN de 'VENDIDOS A ENTREGAR': somava 10 (2+1+0,5+1,5+2+0,5+0,5+2, as
    metas de venda) e o fechamento passava semanas acusando 'acumulado 1 vs 10 na
    estação de monta' como se a fonte divergisse. O REAL de TOTAL EMBRIÕES é 1 —
    igual ao nosso acumulado e ao que o haras publica."""
    if "PLANEJAMENTO" not in wb.sheetnames:
        return 0
    ws = wb["PLANEJAMENTO"]
    linhas = list(ws.iter_rows(values_only=True))
    col = None
    if len(linhas) >= 3:
        planreal, nomes = linhas[1], linhas[2]
        for i, nome in enumerate(nomes):
            if _norm(nome).startswith("TOTAL EMBRI") and _norm(planreal[i]) == "REAL":
                col = i
                break
    if col is None:
        print("  [acumulado] coluna 'TOTAL EMBRIÕES (REAL)' não encontrada na aba "
              "PLANEJAMENTO — conferência da estação de monta fica de fora")
        return 0
    total = 0.0
    for i, row in enumerate(linhas, start=1):
        if i < 4 or row[0] is None:
            continue
        if not str(row[0]).strip().isdigit():
            continue
        v = _to_num(row[col])
        if v:
            total += v
    return int(round(total))


def build_producao(rep: Report, ini: date, fim: date):
    master = _latest_estacao_master()
    rep.fontes["estacao_master"] = master.name
    wb = _load(master)
    ws = wb["ESTAÇÃO"]
    embrioes = []
    _SAFRA_PARICAO_PENDENTE.clear()
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[0] is None:
            continue
        safra_linha = _s(r[35])
        # Receptora -> safra do embrião CONFIRMADO cuja parição ainda não foi
        # lançada, de QUALQUER safra. Preenchido antes do filtro de safra porque é
        # justamente a linha de fora da safra vigente que interessa: sem isto,
        # _paricoes_do_roster carimba a safra do calendário no potro que o roster
        # entrega e a parição vai parar na estação errada. Em 17/09/2026 os potros
        # das recep 453 (IA 04/10/2025) e 440 (IA 11/10/2025) — embriões da safra
        # 2025/2026, já dentro do acumulado 61 dela — subiram o acumulado da 26/27
        # de 1 para 3 contra 1 publicado pelo haras.
        if _norm(r[17]) == "OK" and not _dt(r[23]) and _s(r[11]):
            rec = _norm(_s(r[11]))
            ia_linha = _dt(r[7])
            antes = _SAFRA_PARICAO_PENDENTE.get(rec)
            # mesma receptora serve várias estações: vale a cobrição mais recente
            if not antes or (ia_linha and (not antes[1] or ia_linha > antes[1])):
                _SAFRA_PARICAO_PENDENTE[rec] = (safra_linha, ia_linha)
        if safra_linha not in (SAFRA_ATUAL, SAFRA_PROXIMA):
            continue
        ia = _dt(r[7])
        te = _dt(r[9])
        # CONFIRMAÇÃO É A PARTIR DOS 60 DIAS (regra do Arthur, 18/09/2026). A aba tem
        # a série de diagnósticos — 15D, 30D, 45D, 60D (idx 12..15) — e a coluna
        # consolidada '+ / -' (idx17). Só o consolidado era lido, e isso perde o
        # embrião cujo exame de 60 dias já saiu positivo e o '+ / -' ainda não foi
        # fechado: ele ficava fora do card na semana em que de fato confirmou.
        # 'NAO OK' no consolidado é baixa posterior (perda depois dos 60 dias) e
        # manda, mesmo com 60D positivo — são as 3 linhas da safra 25/26 nessa
        # situação, que não podem voltar para a conta.
        consolidado = _norm(r[17])
        confirmado = consolidado == "OK" or (_s(r[15]) == "+" and consolidado != "NAO OK")
        # QUANDO ele conta como confirmado: no primeiro diagnóstico positivo, que é
        # o exame de 15 dias contado da TE/coleta. Era IA + 60 dias, e isso jogava
        # o embrião dois meses pra frente: em 04/09/2026 o JAVA DA PAO GRANDE x
        # XODÓ PORTEIRA AZUL (TE 19/08, LAVADO '+', 15D '+', +/-='OK') é o
        # "confirmado na semana" da liberação do haras, e caía em 09/10 aqui — a
        # semana fechava com 0 contra 1 deles. Sem TE lançada, IA+15 aproxima (a
        # coleta vem ~7 dias depois da IA).
        base_conf = te or ia
        data_conf = (base_conf + timedelta(days=15)) if base_conf else None
        cotas = _to_num(r[5])
        # split PG / sócio / vendido (defs do comitê: total produzido aberto nessas 3 fatias)
        comprador = _s(r[34])
        if cotas is not None and cotas == 0:
            fatia = "vendido"
        elif (cotas is not None and 0 < cotas < 1) or _s(r[6]):
            fatia = "socio"
        else:
            fatia = "pg"
        embrioes.append({
            "doadora": _s(r[2]), "garanhao": _s(r[3]), "local": _s(r[4]),
            "cotas_pg": cotas, "socio": _s(r[6]), "fatia": fatia,
            "data_ia": ia.isoformat() if ia else None,
            "receptora": _s(r[11]),
            "confirmado": confirmado,
            "data_confirmacao": data_conf.isoformat() if data_conf else None,
            "sexo_potro": _s(r[24]), "nome_potro": _s(r[25]),
            "data_paricao": _dt(r[23]).isoformat() if _dt(r[23]) else None,
            "data_aborto": _dt(r[21]).isoformat() if _dt(r[21]) else None,
            "data_obito": _dt(r[28]).isoformat() if _dt(r[28]) else None,
            "status": _s(r[32]), "categoria": _s(r[33]), "comprador": comprador,
            "safra": safra_linha,
        })

    # o resto de build_producao e da safra corrente; a proxima entra so no acumulado
    embrioes_proxima = [e for e in embrioes if e["safra"] == SAFRA_PROXIMA]
    embrioes = [e for e in embrioes if e["safra"] == SAFRA_ATUAL]

    _EMBRIAO_POR_RECEP.clear()
    for e in embrioes:
        if e.get("receptora"):
            _EMBRIAO_POR_RECEP[_chave_recep(e["receptora"])] = e

    def _in_week(iso):
        return bool(iso and ini <= date.fromisoformat(iso) <= fim)

    def _in_month(iso):
        if not iso:
            return False
        d = date.fromisoformat(iso)
        return d.year == fim.year and d.month == fim.month

    def _split(items):
        return {f: sum(1 for e in items if e["fatia"] == f) for f in ("pg", "socio", "vendido")}

    # ACUMULADO NA ESTAÇÃO (2026-07-31, reunião com o haras):
    #   linhas vivas da planilha "EMBRIÕES E MATRIZES" (PG + sócios) + parições da safra.
    # A planilha do grupo é o retrato do que está em pé — some a linha quando pare ou
    # aborta —, então as parições voltam pra conta e os abortos não. A estação de monta
    # sozinha não serve: só enxerga o que passou pela FPG (56 vs 61 hoje). Fica como
    # conferência em `acumulado_estacao_monta`.
    # ACUMULADO: confirmados da safra na própria aba ESTAÇÃO. Ela não apaga linha —
    # parição e aborto ficam registrados na mesma linha —, então a contagem já é
    # cumulativa. O grupo (EMBRIÕES E MATRIZES) segue sendo lido só para o split por
    # fatia e para a lista arquivada.
    grupo = _acumulado_grupo()
    rep.fontes["embrioes_matrizes"] = grupo["fonte"]
    _LINHAS_BRUTAS["grupo"] = grupo["linhas"]
    acumulado_planejamento = _acumulado_planejamento(wb)

    # "Confirmados semana" é diff contra o snapshot anterior (ver
    # _compute_confirmados_diff) e tem de enxergar as DUAS safras — o relatório
    # oficial conta confirmação nova independente de qual estação ela é (foi
    # assim que a confirmação do Java x Xodó, 1ª da safra 26/27, virou "01" pra
    # ele). Até 28/08/2026 só entrava `embrioes` (SAFRA_ATUAL, filtrado ali em
    # cima) — uma confirmação nova de safra NOVA nunca aparecia aqui, mesmo já
    # lançada na fonte, e o indicador ficava preso em 0 na transição de safra.
    confirmados = [e for e in embrioes + embrioes_proxima if e["confirmado"]]
    for e in confirmados:
        e["key"] = f"{e['doadora']}|{e['garanhao']}|{e['receptora']}|{e['data_ia']}"
    rep.confirmed = confirmados
    na_semana = []          # confirmados_semana e acumulado_mes: preenchidos por _compute_confirmados_diff
    nascimentos = [e for e in embrioes if _in_week(e["data_paricao"])]
    # aborto = embrião confirmado que não nasceu (data aborto na semana);
    # óbito = nasceu e morreu (data óbito na semana). (absorção = perda pré-60d, não confirmada)
    abortos = [e for e in embrioes if e["confirmado"] and _in_week(e["data_aborto"])]
    obitos = [e for e in embrioes if _in_week(e["data_obito"])]
    # óbito de animal já no plantel não passa pela estação — vem do controle mensal
    mortes_plantel = _mortes_do_plantel(ini, fim)
    if mortes_plantel:
        print(f"  [produção] +{len(mortes_plantel)} óbito(s) da aba "
              f"'CONFIRMAÇÕES, ABORTOS, MORTES' do controle mensal")

    # Parições da safra inteira (não só da semana) — voltam pro acumulado, PORQUE a
    # planilha do grupo apaga a linha de quem pariu. Quando ela ainda não foi apagada,
    # somar as duas conta o mesmo embrião duas vezes: em 17/08/2026 a parição de JAVA
    # DA PAO GRANDE x QUEBRUTO (recep 309) foi lançada na estação às 17:05 e a planilha
    # do grupo estava salva desde 14/08 com a linha viva — o acumulado pulou de 61 pra
    # 62 numa semana sem nenhuma produção nova. Descontar é o certo: enquanto o embrião
    # está nas duas listas, ele é UM.
    paridos_safra = [e for e in embrioes if e["data_paricao"]]
    ainda_no_grupo = [e for e in paridos_safra
                      if _chave_embriao(e["doadora"], e["garanhao"], e["receptora"])
                      in grupo["chaves"]]
    paridos_novos = [e for e in paridos_safra if e not in ainda_no_grupo]
    if ainda_no_grupo:
        print(f"  [acumulado] {len(ainda_no_grupo)} parição(ões) ainda listadas como "
              f"vivas em {grupo['fonte']} — contadas UMA vez, não duas:")
        for e in ainda_no_grupo:
            print(f"    - {e['doadora']} x {e['garanhao']} (recep {e['receptora']}), "
                  f"pariu {e['data_paricao']}")
    # Confirmados da safra na aba ESTAÇÃO = acumulado. Aborto não volta (baixa
    # definitiva, decidido com o haras em 2026-07-31), então sai da conta.
    confirmados_safra = [e for e in embrioes if e["confirmado"]]
    abortados = [e for e in confirmados_safra if e["data_aborto"]]
    acumulado = len(confirmados_safra) - len(abortados)
    split = {"pg": grupo["pg"], "socio": grupo["socio"], "vendido": grupo["vendido"]}
    for e in paridos_novos:
        split[e["fatia"]] = split.get(e["fatia"], 0) + 1
    if acumulado != acumulado_planejamento:
        print(f"  [produção] acumulado {acumulado} "
              f"(planilha do grupo {grupo['total']} + {len(paridos_novos)} parições) "
              f"vs {acumulado_planejamento} na estação de monta")

    # Mesma regra da safra corrente, aplicada na que comeca: vivos na planilha do
    # grupo + paricoes que ja sairam de la. Sem SAFRA_PROXIMA (fora de transicao),
    # nem le a planilha do grupo de novo — o card correspondente ja nasce None e
    # some no dashboard (skip por falta de rotulo).
    if SAFRA_PROXIMA:
        grupo_prox = _acumulado_grupo(SAFRA_PROXIMA)
        paridos_prox = [e for e in embrioes_proxima if e["data_paricao"]
                        and _chave_embriao(e["doadora"], e["garanhao"], e["receptora"])
                        not in grupo_prox["chaves"]]
        acumulado_prox = grupo_prox["total"] + len(paridos_prox)
        if acumulado_prox:
            print(f"  [acumulado] safra {SAFRA_PROXIMA} ja tem {acumulado_prox} "
                  f"(grupo {grupo_prox['total']} + {len(paridos_prox)} parições)")
    else:
        acumulado_prox = None

    rep.producao = {
        "acumulado_estacao": acumulado,
        "acumulado_estacao_proxima": acumulado_prox,
        "safra_atual": SAFRA_ATUAL,
        "safra_proxima": SAFRA_PROXIMA,
        "safra_atual_rotulo": _rotulo_safra(SAFRA_ATUAL),
        "safra_proxima_rotulo": _rotulo_safra(SAFRA_PROXIMA),
        "acumulado_estacao_split": split,
        "acumulado_estacao_monta": acumulado_planejamento,   # conferência
        "acumulado_grupo_vivos": grupo["total"],
        "acumulado_paricoes_safra": len(paridos_novos),
        "acumulado_paricoes_ainda_no_grupo": len(ainda_no_grupo),
        "acumulado_estacao_split_confirmados": _split(confirmados),  # split antigo (estação)
        "confirmados_semana": None,   # _compute_confirmados_diff
        "acumulado_mes": None,        # _compute_confirmados_diff
        "nascimentos": None,          # _nascimentos_do_roster, abaixo
        "abortos_obitos": len(abortos) + len(obitos) + len(mortes_plantel),
    }
    def _publica_nascimentos():
        rep.producao["nascimentos"] = len(rep.detalhe.get("nascimentos_semana") or [])

    def _produto(e):   # nome do animal nascido (ou descrição sexo — doadora × garanhão)
        base = f"{e.get('doadora') or ''} × {e.get('garanhao') or ''}".strip(" ×")
        sx = {"M": "Macho", "F": "Fêmea"}.get((e.get("sexo_potro") or "").upper(), e.get("sexo_potro") or "")
        return e.get("nome_potro") or (f"{sx} — {base}" if base else sx) or "--"
    rep.detalhe["confirmados_semana"] = na_semana
    _SOCIO_POR_RECEP.clear()
    for e in embrioes:
        soc = _limpa_socio(e.get("comprador")) or _limpa_socio(e.get("socio"))
        if soc and e.get("receptora"):
            _SOCIO_POR_RECEP[_norm(e["receptora"])] = soc
    # Nascimento vem do roster mensal (data + filiação). A parição lançada na aba
    # ESTAÇÃO fica como conferência: divergir significa lançamento faltando em um dos
    # dois lados, e isso tem de aparecer em vez de escolher um número calado.
    nasc_roster = _nascimentos_do_roster(ini, fim)
    rep.producao["nascimentos_estacao"] = len(nascimentos)
    if len(nasc_roster) != len(nascimentos):
        print(f"  [nascimentos] roster mensal {len(nasc_roster)} x aba ESTAÇÃO "
              f"{len(nascimentos)} — publicando o roster (é onde o potro entra); "
              f"a diferença é parição não lançada em um dos dois")
    rep.detalhe["nascimentos_semana"] = nasc_roster
    rep.detalhe["nascimentos_estacao"] = [
        dict(e, produto=_produto(e),
             socio=_limpa_socio(e.get("comprador")) or _limpa_socio(e.get("socio")))
        for e in nascimentos]
    _publica_nascimentos()
    rep.detalhe["abortos_obitos_semana"] = abortos + obitos + mortes_plantel
    rep.detalhe["embrioes_confirmados"] = confirmados
    # listas COMPLETAS datadas (o dashboard filtra por semana no cliente)
    rep.eventos["confirmados"] = [dict(e, data=e["data_confirmacao"]) for e in confirmados if e["data_confirmacao"]]
    rep.eventos["nascimentos"] = [dict(e, data=e["data_paricao"]) for e in embrioes if e["data_paricao"]]
    rep.eventos["abortos_obitos"] = [
        dict(e, data=(e["data_aborto"] or e["data_obito"]))
        for e in embrioes if (e["confirmado"] and e["data_aborto"]) or e["data_obito"]
    ]
    wb.close()


# ------------------------------------------------------------------
# Seção 2 — RECEPTORAS (censo + candidatos de filtro)
# ------------------------------------------------------------------
RECEPTORAS_LOCAIS_ATIVOS = ("PAO GRANDE", "ARRENDAMENTO CESAR FURTADO")


def build_receptoras(rep: Report):
    """Rebanho ATIVO = aba ANIMAIS do PLANTEL ARRENDAMENTOS E RECEPTORAS, filtrando
    LOCAL em (PAO GRANDE, ARRENDAMENTO CESAR FURTADO) e STATUS prenha/vazia.
    NÃO usa mais a aba 'ATUALIZAÇÃO SEMANAL' (aba a ser aposentada). Validado vs
    docx 24/07: prenhas 34 / vazias 28 / total 62 — bate EXATO.
    Colunas ANIMAIS (linha 3 header, dados r4+): 1 ANIMAL, 2 STATUS, 3 LOCAL."""
    src = _latest_no_plantel("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx", "receptoras")
    rep.fontes["receptoras"] = src.name
    wb = _load(src)
    ws = wb["ANIMAIS"]
    pren = vaz = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        if _norm(r[3]) not in RECEPTORAS_LOCAIS_ATIVOS:
            continue
        st = _norm(r[2])
        if st.startswith("PRENHA"):
            pren += 1
        elif st.startswith("VAZIA"):
            vaz += 1
    wb.close()
    doadoras_plantel = _count_doadoras()             # linhas da aba PLANEJAMENTO
    doadoras = DOADORAS_INDICE or doadoras_plantel
    rep.receptoras = {
        "total": pren + vaz,
        "prenhas": pren,
        "vazias": vaz,
        # Índice = vazias ÷ doadoras (linhas da aba PLANEJAMENTO). Confirmado
        # contra os dois pontos do histórico (21/08: doadoras=12, ciclando=10,
        # índice=2,5 — só 30÷12 fecha; 30÷10 dá 3,0) depois de eu ter "corrigido"
        # isto errado em 28/08/2026 pra usar ciclando, sem checar contra semana
        # nenhuma. Revertido. 'doadoras_ciclando' é indicador PRÓPRIO no relatório
        # (card ao lado), não entra nesta conta.
        "doadoras": doadoras,
        "doadoras_fonte": "fixo" if DOADORAS_INDICE else "planejamento",
        "doadoras_plantel": doadoras_plantel,
        "indice_eficiencia": round(vaz / doadoras, 1) if doadoras else None,
        # preenchido em bases/semanal_manual.json (ver _manual): sem fonte de dado
        "doadoras_ciclando": None,
    }
    if DOADORAS_INDICE and DOADORAS_INDICE != doadoras_plantel:
        print(f"  [receptoras] índice usa {DOADORAS_INDICE} doadoras (fixo); "
              f"PLANEJAMENTO tem {doadoras_plantel}")


# Denominador do índice de eficiência.
# Era FIXO em 10 pela reunião de 2026-07-31, porque nenhuma leitura do plantel dava
# esse número. Em 07/08/2026 o relatório oficial trocou o divisor: 29 vazias com
# índice 2,4 = 29/12, e 12 é exatamente CATEGORIA='DOADORA' no plantel (nas 4 semanas
# anteriores o divisor implícito no docx era 10). Ou seja, o haras passou a usar o
# contado — então paramos de fixar e contamos.
# None = usar as doadoras contadas no plantel. Só voltar a pôr número aqui se o
# haras decidir travar o divisor de novo.
DOADORAS_INDICE = None


def _count_doadoras() -> int:
    """Doadoras = linhas da aba PLANEJAMENTO (ESTAÇÃO DE MONTA), não CATEGORIA
    'DOADORA' no plantel.

    Achado em 28/08/2026, pelo próprio haras: 'não devemos puxar as doadoras
    pelo plantel, mas sim pela estação de monta'. Faz sentido — PLANEJAMENTO é
    quem está sendo trabalhada na safra corrente; o plantel é cadastro histórico
    e trazia gente que já vendeu, já saiu ou nunca fez parte deste time (RELIQUIA,
    CANCAO, INUSITADA, BISCA, MEIRELLES, XARDA, XICA apareciam lá e NENHUMA está
    em PLANEJAMENTO). As duas listas não têm quase nada em comum — não é ajuste
    fino, é fonte errada desde o início. PLANEJAMENTO tem exatamente 12 linhas
    (header na linha 3, dado 4 em diante), batendo com o índice oficial (30÷12
    = 2,5) sem precisar de nenhuma exclusão adicional."""
    master = _latest_estacao_master()
    wb = _load(master)
    ws = wb["PLANEJAMENTO"]
    n = 0
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None or not str(r[1]).strip():
            continue
        n += 1
    wb.close()
    return n


# ------------------------------------------------------------------
# Seção 3 — HEADCOUNT (calculado do roster, espelhando a aba CONTAGEM)
# ------------------------------------------------------------------
# A CONTAGEM é COUNTIF puro sobre o LOCAL do roster:
#     C3 =COUNTIF(PLANTEL!E:E;"FAZENDA PAO GRANDE")      D3 = 25   (digitado)
#     C4 =COUNTIF(PLANTEL!E:E;"ARRENDAMENTO CESAR FURTADO")  D4 = 37   (digitado)
#     C5 =COUNTIF(PLANTEL!E:E;"OUTROS")                  D5 = 0     (rótulo "CTE")
#     C6 =COUNTIF(PLANTEL!E:E;"SOCIO")                   D6 = 0
#     E7 =SUM(E3:E6)
# Reproduzimos a MESMA regra (conta linha por LOCAL, sem filtrar STATUS — por isso
# os vendidos pendentes de saída entram) e o MESMO conjunto de buckets. Só muda de
# onde vem o número das receptoras: contado da fonte de receptoras (a mesma da
# seção 2) em vez de digitado na mão. Confere — PAO GRANDE 25 e ARRENDAMENTO CESAR
# FURTADO 37, idênticos ao que estava fixo em D3/D4.
#
# MATO GROSSO fica FORA do headcount por decisão de negócio (confirmado em
# 2026-07-31): os animais de lá não entram na contagem do plantel, e é por isso
# que a CONTAGEM nunca teve linha para eles.
#
# LOCAL vazio não entra (o COUNTIF também ignora), o que descarta de graça a
# pseudo-linha "RECEPTORAS 67" que mora no meio do roster.

# LOCAL no roster -> chave no relatório (rótulo que a CONTAGEM usa)
HEADCOUNT_BUCKETS = {
    "FAZENDA PAO GRANDE": ("FAZENDA", "fazenda_pg"),
    "ARRENDAMENTO CESAR FURTADO": ("ARRENDAMENTO", "arrendamento"),
    "OUTROS": ("CTE", "cte"),
    "SOCIO": ("SOCIO", "socio"),
}
# LOCAL que existe no roster mas não conta no headcount.
HEADCOUNT_LOCAIS_FORA = ("MATO GROSSO",)
# LOCAL da fonte de receptoras -> LOCAL do roster
RECEPTORAS_PARA_BUCKET = {
    "PAO GRANDE": "FAZENDA PAO GRANDE",
    "ARRENDAMENTO CESAR FURTADO": "ARRENDAMENTO CESAR FURTADO",
}


def _slug_local(local: str) -> str:
    return local.lower().replace(" ", "_")


def _receptoras_por_local(src: Path | None = None) -> dict:
    """{LOCAL do roster: nº de receptoras}. Mesma regra da seção 2 (prenha/vazia)."""
    if src is None:
        src = _latest_no_plantel("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx", "receptoras")
    wb = _load(src)
    ws = wb["ANIMAIS"]
    out = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        loc = _norm(r[3])
        if loc not in RECEPTORAS_LOCAIS_ATIVOS:
            continue
        st = _norm(r[2])
        if st.startswith("PRENHA") or st.startswith("VAZIA"):
            out[RECEPTORAS_PARA_BUCKET[loc]] = out.get(RECEPTORAS_PARA_BUCKET[loc], 0) + 1
    wb.close()
    return out


def headcount_de(plantel_src: Path | None = None,
                 receptoras_src: Path | None = None,
                 quando: date | None = None) -> tuple[dict, list[str]]:
    """Headcount por bucket a partir de UM par de arquivos (roster + receptoras).

    Separado de `build_headcount` porque o comitê precisa do mesmo número para um
    mês PASSADO — e reimplementar a contagem lá seria ter duas versões da regra
    (é a mesma razão pela qual tools/excel_headcount.py chama estas funções em
    vez de refazer os filtros). Sem argumento, é o estado de hoje.

    Devolve (headcount, LOCAIS desconhecidos fora da contagem)."""
    # 1) animais por LOCAL, do MESMO roster que o resto do fechamento usa.
    #    Relia a planilha aqui dentro e contava linha a linha sem filtro nenhum —
    #    funcionava porque a planilha semanal já vinha curada. Com o roster mensal
    #    isso contaria vendido, morto, embrião e a linha repetida por cotista.
    animais: dict[str, int] = {}
    for linha in _plantel_por_status(plantel_src, quando)["linhas"]:
        local = _norm(linha["local"])
        if not local:
            continue
        animais[local] = animais.get(local, 0) + 1

    # 2) receptoras: contadas da fonte de receptoras
    receptoras = _receptoras_por_local(receptoras_src)

    # 3) monta os buckets. LOCAL fora da contagem é ignorado; LOCAL desconhecido
    #    também não entra (a CONTAGEM não o teria), mas vira aviso — assim um
    #    local novo no roster aparece pra alguém decidir, em vez de sumir calado.
    detalhe, chaves, fora = {}, {}, {}
    for local in sorted(set(animais) | set(receptoras)):
        a = animais.get(local, 0)
        rc = receptoras.get(local, 0)
        if local in HEADCOUNT_LOCAIS_FORA or local not in HEADCOUNT_BUCKETS:
            fora[local] = a + rc
            continue
        rotulo, chave = HEADCOUNT_BUCKETS[local]
        detalhe[rotulo] = {"animais": a, "receptoras": rc, "total": a + rc}
        chaves[chave] = a + rc

    total = sum(v["total"] for v in detalhe.values())
    detalhe["TOTAL GERAL"] = {
        "animais": sum(v["animais"] for k, v in detalhe.items() if k != "TOTAL GERAL"),
        "receptoras": sum(v["receptoras"] for k, v in detalhe.items() if k != "TOTAL GERAL"),
        "total": total,
    }

    # A conferência contra a aba CONTAGEM saiu: aquela aba é um COUNTIF dentro do
    # arquivo de DIVULGAÇÃO, e conferir o cálculo contra o que foi divulgado não
    # confere nada — é o próprio número que se quer auditar. Quem confere o headcount
    # é o relatório oficial, no placar do fechamento.
    desconhecidos = [l for l in fora if l not in HEADCOUNT_LOCAIS_FORA]
    return ({"total": total, **chaves, "detalhe": detalhe, "fora_da_contagem": fora},
            desconhecidos)


def build_headcount(rep: Report):
    rep.headcount, desconhecidos = headcount_de()
    if desconhecidos:
        fora = rep.headcount["fora_da_contagem"]
        print("  [headcount] LOCAL novo no roster, FORA da contagem — conferir:")
        for l in desconhecidos:
            print(f"    - {l}: {fora[l]}")


# ------------------------------------------------------------------
# Seção 5 — SAÍDAS / ENTRADAS / TRANSFERÊNCIAS  (CONTROLE_DE_PLANTEL aba MOVIMENTAÇÕES)
# ------------------------------------------------------------------
def _categorize_mov(obs: str):
    if "MUDOU O LOCAL PARA FAZENDA PAO GRANDE" in obs or "MUDOU O LOCAL PARA ARRENDAMENTO" in obs:
        return "TRANSFERENCIA"
    if "SAIU DO HARAS" in obs:
        return "SAIDA"
    # Doação também é saída física, e não vem escrita como "SAIU DO HARAS".
    if "DOADO" in obs or "DOACAO" in obs:
        return "SAIDA"
    if "CHEGOU NO HARAS" in obs:
        return "ENTRADA"
    return None


# Aba SAIDAS-ENTRADAS do controle mensal (reunião 2026-07-31): fonte oficial de
# saída/entrada. O haras começou a preencher em 12/08/2026.
# Colunas: B animal, C local saída, D local entrada, E data, F classificação.
# Δ combinado: entrada = nascimento + compra; saída = venda + morte.
#
# A classificação vem digitada com o sentido colado no motivo — a primeira linha real
# foi 'SAIDA-SOCIO'. Só as palavras de motivo abaixo não bastavam: 'SAIDA-SOCIO' não
# casava com nada, caía no `else None` e era descartada em silêncio. Como a aba já
# tinha linha, ela vencia como fonte oficial e a saída da semana virava 0 — foi assim
# que LINDEZA DA PAO GRANDE (12/08/2026) sumiu do fechamento de 14/08.
CLASSIF_ENTRADA = ("NASCIMENTO", "COMPRA")
CLASSIF_SAIDA = ("VENDA", "MORTE", "SOCIO")
# SAIDA-SOCIO só fica fora do Δ quando quem sai é RECEPTORA: ela só é contada em
# PAO GRANDE/ARRENDAMENTO, então ir pro sócio a tira de lá mas não mexe no Δ de
# ANIMAIS (é outro total, tratado à parte — ver _refina_afeta_headcount). Era regra
# geral pra qualquer 'SOCIO' e estava errada: em 28/08/2026 a GIM MATIZA (GARANHÃO,
# destino nomeado 'VALTER LIMA') saiu de verdade, física — Δ oficial contou as DUAS
# saídas da semana (-02), não só a venda. Physical is physical: se o bicho deixou a
# fazenda, conta, sócio ou não.
CLASSIF_FORA_DO_DELTA = ("SOCIO",)


def _classificar_se(classif: str, animal: str = ""):
    """(sentido, afeta_headcount) da classificação da aba SAIDAS-ENTRADAS.
    (None, None) = vocabulário desconhecido; quem chama tem de avisar, não engolir."""
    def _sai():
        eh_socio_exempt = (any(c in classif for c in CLASSIF_FORA_DO_DELTA)
                           and _norm(animal).startswith("RECEPTORA"))
        return "SAIDA", not eh_socio_exempt
    if classif.startswith("ENTRADA"):
        return "ENTRADA", True
    if classif.startswith("SAIDA"):
        return _sai()
    if any(c in classif for c in CLASSIF_ENTRADA):
        return "ENTRADA", True
    if any(c in classif for c in CLASSIF_SAIDA):
        return _sai()
    return None, None


def _saidas_entradas_planilha(wb, ini: date, fim: date):
    """Lê a aba SAIDAS-ENTRADAS. Devolve None se ela não existir ou estiver vazia,
    pra que o cálculo caia na diferença de roster (comportamento atual)."""
    if "SAIDAS-ENTRADAS" not in wb.sheetnames:
        return None
    evs = {"SAIDA": [], "ENTRADA": []}
    achou = False
    desconhecidas = []
    for i, r in enumerate(wb["SAIDAS-ENTRADAS"].iter_rows(values_only=True), start=1):
        if i < 3 or r[1] is None or not str(r[1]).strip():
            continue
        achou = True
        d = _dt(r[4])
        if not d or not (ini <= d <= fim):
            continue
        classif = _norm(r[5])
        alvo, afeta = _classificar_se(classif, _s(r[1]))
        if alvo is None:
            # classificação nova na planilha: avisar, nunca virar zero calado
            desconhecidas.append(f"linha {i} ({_s(r[1])}): {_s(r[5])!r}")
            continue
        evs[alvo].append({"animal": _s(r[1]), "data": d.isoformat(),
                          "classificacao": _s(r[5]), "afeta_headcount": afeta,
                          "local_saida": _s(r[2]), "local_entrada": _s(r[3])})
    if desconhecidas:
        print(f"  [SAIDAS-ENTRADAS] classificação não reconhecida em "
              f"{len(desconhecidas)} linha(s) da janela — NÃO entraram na conta: "
              + "; ".join(desconhecidas))
    return evs if achou else None


# TRANSFERÊNCIA INTERNA = animal que trocou de LOCAL entre os dois locais próprios
# (FPG <-> arrendamento) na aba ANIMAIS do PLANTEL ARRENDAMENTOS E RECEPTORAS.
# A aba MOVIMENTAÇÕES do controle mensal NÃO serve: a última transferência lançada
# lá é de setembro/2025. Em 07/08/2026 o relatório oficial disse 14 transferências e
# o diff de LOCAL dá exatamente 14 (5 arrendamento->FPG, 9 FPG->arrendamento).
# Igual a saídas/entradas, é diff: precisa do mapa da semana anterior. Sem ele
# (primeira captura), fica em branco — nunca zero.
def _receptoras_arquivos() -> list:
    """Arquivos de receptoras, do mais recente pro mais antigo (por mtime)."""
    cands = [f for d in _estacao_dirs()
             for f in d.glob("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx")
             if not f.name.startswith("~$")]
    return sorted(cands, key=lambda f: f.stat().st_mtime, reverse=True)


def _receptoras_info(src: Path | None = None) -> dict:
    """{ANIMAL: {local, status, embriao, obs}} da aba ANIMAIS — TODAS as linhas,
    inclusive fora dos nossos locais, pra saber pra onde o animal foi."""
    if src is None:
        src = _latest_no_plantel("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx", "receptoras")
    wb = _load(src)
    ws = wb["ANIMAIS"]
    out = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 4 or r[1] is None:
            continue
        out[_norm(r[1])] = {"local": _norm(r[3]), "status": _s(r[2]),
                            "embriao": _s(r[4]), "obs": _s(r[5])}
    wb.close()
    return out


def _receptoras_locais(src: Path | None = None) -> dict:
    """{ANIMAL: LOCAL} da aba ANIMAIS, só quem está num local nosso."""
    return {k: v["local"] for k, v in _receptoras_info(src).items()
            if v["local"] in RECEPTORAS_LOCAIS_ATIVOS}


def _transferencias_internas(rep: Report) -> list | None:
    """Diff do LOCAL vs o mapa da semana anterior. None = sem base de comparação."""
    hist = _load_hist()
    prev = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("receptoras_locais"):
            prev = hist[wid]["receptoras_locais"]
    cur = rep.receptoras_locais or {}
    if not cur:
        return None
    if not prev:
        # Bootstrap: nenhuma semana anterior guardou o mapa (o campo é novo). Cai no
        # arquivo de receptoras anterior, que é o retrato de onde os animais estavam.
        anteriores = _receptoras_arquivos()[1:]
        if not anteriores:
            print("  [transferências] sem mapa de LOCAL da semana anterior e sem "
                  "arquivo anterior de receptoras — fica EM BRANCO, não zero")
            return None
        prev = _receptoras_locais(anteriores[0])
        print(f"  [transferências] primeira semana com mapa de LOCAL: comparando com "
              f"{anteriores[0].name} (da próxima em diante, compara com o snapshot)")
    transf = [{"animal": k, "tipo": "RECEPTORA", "local_saida": prev[k],
               "local_entrada": cur[k]}
              for k in cur if k in prev and prev[k] != cur[k]]
    transf += _transferencias_de_animais(rep)
    return transf


# Transferência interna é FPG <-> ARRENDAMENTO, nos dois sentidos (regra do Arthur,
# 10/09/2026). Qualquer outra troca de LOCAL é entrada, saída ou mudança de bucket,
# não transferência.
LOCAIS_INTERNOS = ("FAZENDA PAO GRANDE", "ARRENDAMENTO CESAR FURTADO")


def _transferencias_de_animais(rep: Report) -> list:
    """Animal do plantel que mudou de LOCAL entre a semana passada e esta.

    O detector olhava só o mapa de RECEPTORAS, então transferência de animal
    passava batido: em 10/09/2026 o LEGADO DA PAO GRANDE saiu de FAZENDA PAO
    GRANDE para ARRENDAMENTO CESAR FURTADO, o haras publicou "Transferências
    internas: 01" e o nosso card ficou 0 — o Arrendamento até subia para 44, sem
    nada explicando de onde vinha.

    A comparação sai do arquivo de linhas arquivado (ver _arquivo_anterior), que
    guarda o LOCAL de cada animal do roster — é exatamente para isto que ele
    existe."""
    ant = {_norm(x.get("nome")): _norm(x.get("local"))
           for x in (_arquivo_anterior(rep.semana_atual).get("roster") or [])}
    if not ant:
        return []
    # as linhas do roster ainda nao estao em _LINHAS_BRUTAS aqui — este passo roda
    # antes do build do headcount —, entao le direto (o workbook ja esta em cache)
    linhas = _LINHAS_BRUTAS.get("roster") or _plantel_por_status()["linhas"]
    out = []
    for l in linhas:
        k, agora = _norm(l.get("nome")), _norm(l.get("local"))
        antes = ant.get(k)
        if not antes or antes == agora:
            continue
        if antes in LOCAIS_INTERNOS and agora in LOCAIS_INTERNOS:
            out.append({"animal": l.get("nome"), "tipo": _s(l.get("categoria")),
                        "local_saida": antes, "local_entrada": agora})
    if out:
        print(f"  [transferências] {len(out)} animal(is) mudaram de LOCAL entre "
              f"Fazenda e Arrendamento: "
              + "; ".join(f"{x['animal']} ({x['local_saida']} -> {x['local_entrada']})"
                          for x in out))
    return out


def build_movimentacao(rep: Report, ini: date, fim: date):
    src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    rep.fontes["controle_plantel_mensal"] = src.name
    wb = _load(src)
    rep.saidas_planilha = _saidas_entradas_planilha(wb, ini, fim)
    ws = wb["MOVIMENTAÇÕES"]
    evs = {"SAIDA": [], "ENTRADA": [], "TRANSFERENCIA": []}
    ultimo = None  # último lançamento DATADO da aba, de qualquer tipo
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[3] is None:
            continue
        d = _dt(r[3])
        if not d:
            continue
        if ultimo is None or d > ultimo:
            ultimo = d
        tipo = _categorize_mov(str(r[4] or "").upper())
        if tipo is None:
            continue
        evs[tipo].append({"animal": _s(r[2]), "data": d.isoformat(), "ocorrencia": _s(r[4])})
    inw = lambda x: ini <= date.fromisoformat(x["data"]) <= fim
    # A aba só mede a semana se ela chegou na semana. Parada antes da janela, contar
    # zero afirmaria "não houve movimentação" quando o que existe é falta de
    # lançamento — foi o que fez transferências virar 0 com 14 no relatório oficial.
    rep.saidas = {
        "saidas_semana": sum(1 for x in evs["SAIDA"] if inw(x)),
        "entradas_semana": sum(1 for x in evs["ENTRADA"] if inw(x)),
    }
    # transferências não vêm daqui (ver comentário acima de _receptoras_locais)
    rep.receptoras_locais = _receptoras_locais()
    transf = _transferencias_internas(rep)
    rep.saidas["transferencias_semana"] = len(transf) if transf is not None else None
    rep.detalhe["transferencias_internas"] = transf
    rep.saidas["movimentacao_ultimo_lancamento"] = ultimo.isoformat() if ultimo else None
    if ultimo is None or ultimo < ini:
        rep.saidas["movimentacao_defasada"] = True
        print(f"  [movimentação] {src.name}: último lançamento datado é "
              f"{ultimo.strftime('%d/%m/%Y') if ultimo else 'nenhum'}, antes da janela "
              f"({ini.strftime('%d/%m')}) — a aba MOVIMENTAÇÕES não mede esta semana")
    # listas COMPLETAS datadas p/ filtro client-side
    rep.eventos["saidas"] = evs["SAIDA"]
    rep.eventos["entradas"] = evs["ENTRADA"]
    rep.eventos["transferencias"] = evs["TRANSFERENCIA"]
    wb.close()


# ------------------------------------------------------------------
# Seção 4b — PENDENTES DE SAÍDA / TERCEIROS  (CONTROLE PLANTEL aba PLANTEL)
# ------------------------------------------------------------------
# Reunião com o haras (2026-07-31): vendidos pendentes e terceiros passam a sair do
# roster, pela coluna STATUS PLANTEL.
#
# O haras CUMPRIU — só que no CONTROLE_DE_PLANTEL_PAO_GRANDE mensal (pasta PLANTEL),
# não na cópia semanal. Vocabulário de STATUS PLANTEL em 14/08/2026:
#   semanal (CONTROLE PLANTEL.xlsx) : PLANTEL 143, VENDIDO 6
#   mensal  (CONTROLE_DE_PLANTEL...) : PLANTEL 224, VENDIDO E ENTREGUE 108, OBITO 26,
#                                      DOADO 16, DE TERCEIRO 8, VENDIDO PENDENTE SAIDA 6,
#                                      VENDIDO 3
# Lendo o arquivo errado, os dois indicadores morriam em silêncio: vendidos caía no
# 'Animais para sair', congelado em 24/07 (2 em vez de 6, faltando PATRIMONIO, PODIO e
# PAETE) e terceiros saía 2 em vez de 8. O relatório oficial de 14/08 traz 08 terceiros
# e 06 animais vendidos pendentes — os números do MENSAL.
STATUS_VENDIDO_PENDENTE = "VENDIDO PENDENTE"      # rótulo, para mensagem
# Casar por SUBSTRING exata deixava de fora a grafia errada: em 10/09/2026 o
# PROSPERO DA PAO GRANDE estava como 'VENDIDO PENDENDE DE SAIDA' (typo) e ficava
# fora da lista de vendidos pendentes, mesmo já contando no headcount pelo
# prefixo. Mesma tolerância nos dois lugares — ver PREFIXO_VENDIDO_PENDENTE.
def _e_vendido_pendente(status_plantel, local=None) -> bool:
    """Venda fechada e o animal ainda não entregue ao comprador.

    NÃO filtra por LOCAL, ao contrário da sociedade pendente: para a VENDA o destino
    é o comprador, e estar em SOCIO não quer dizer entregue — a MELISSA DA PAO GRANDE
    está lá e o haras a conta entre os 05 animais. Para a SOCIEDADE é o oposto: ir
    para SOCIO É a saída (ver a marca de OBS em _status_plantel_mensal).

    `local` fica no parâmetro porque quem chama já o tem em mãos e a assimetria acima
    é fácil de esquecer — documentada aqui em vez de virar filtro escondido."""
    return PREFIXO_VENDIDO_PENDENTE in _norm(status_plantel)

STATUS_TERCEIRO = "TERCEIRO"
# Sociedade pendente de animal: até 28/08/2026 não tinha marca nenhuma (comentário
# antigo em build_pendentes: "sociedade nunca recebe marca"), então soc_animais
# ficava sempre vazio. O haras passou a marcar na coluna OBS (Y, índice 24) — mesma
# ideia do STATUS_VENDIDO_PENDENTE, coluna diferente.
STATUS_SOCIEDADE_PENDENTE = "EM SOCIEDADE PENDENTE DE SAIDA"
COL_MENSAL_OBS = 24

# Status que significam "o animal AINDA ESTÁ AQUI".
#
# Vendido conta: a venda foi fechada mas o animal não saiu da fazenda, e enquanto
# não sai ele é headcount. O que tira da conta é a ENTREGA — por isso 'VENDIDO E
# ENTREGUE' fica de fora, junto de OBITO e DE TERCEIRO.
#
# DOADO NÃO conta. Em 04/09/2026 eu tinha incluído (o 61 de sócios divulgado
# batia com 53 PLANTEL + 8 DOADO), e estava errado: os 8 doados de 31/08/2026
# estavam nos sócios e o Eduardo pediu para serem RETIRADOS do plantel — a
# doação é a formalização dessa saída. A lista que o haras mantém para a
# atualização semanal também não os tem. Com eles fora, a contagem vai para 184,
# que é o número daquela lista.
#
# A comparação é EXATA, não por substring: 'VENDIDO' como pedaço de texto casaria
# com 'VENDIDO E ENTREGUE' e traria de volta justamente quem já foi embora.
# 'VENDIDO PENDENTE ...' é venda fechada com animal AINDA AQUI, e a planilha
# escreve isso de três formas: 'VENDIDO PENDENTE SAIDA', 'VENDIDO PENDENTE DE
# SAIDA' e 'VENDIDO PENDENDE DE SAIDA' (typo). A lista exata só tinha a primeira,
# e em 10/09/2026 isso tirou MELISSA, MUSICA e PROSPERO do headcount de uma vez —
# 3 animais que não saíram do lugar, só mudaram de status. Prefixo resolve as três
# e não pega 'VENDIDO E ENTREGUE', que é quem de fato foi embora.
STATUS_NO_PLANTEL = ("PLANTEL", "VENDIDO", "VENDIDO PENDENTE SAIDA")
PREFIXO_VENDIDO_PENDENTE = "VENDIDO PENDEN"


def _status_conta(status) -> bool:
    st = _norm(status)
    return st in STATUS_NO_PLANTEL or st.startswith(PREFIXO_VENDIDO_PENDENTE)
# MARRETADA, autorizada pelo Arthur em 04/09/2026: animal que saiu do plantel e
# cuja linha no controle mensal ainda não foi atualizada. A lista que o haras usa
# na atualização semanal já não o tem. Cada fechamento imprime um aviso, e a
# entrada sai daqui no dia em que a origem for corrigida — override permanente
# esconde erro de cadastro em vez de resolver.
# Cada entrada guarda A PARTIR DE QUANDO vale. O fechamento semanal é o estado de
# HOJE e usa todas, sempre — é o comportamento de sempre e não muda. A data existe
# para quem reconstrói um mês PASSADO (o comitê): uma saída de 04/09/2026 não pode
# apagar animal do fechamento de junho, onde a linha estava certa.
FORA_NA_MAO = {
    "MINEIRO DA PAO GRANDE": (
        date(2026, 9, 4),
        "saiu do plantel (Ana, 04/09/2026); linha ainda diz PLANTEL / ARRENDAMENTO"),
}


def _overrides_validos(quando: date | None = None) -> dict:
    """{NOME: motivo}. `quando=None` (padrão) = todos, que é o que o semanal quer."""
    return {_norm(k): motivo for k, (desde, motivo) in FORA_NA_MAO.items()
            if quando is None or quando >= desde}
# Categoria que não é animal do headcount: embrião não nasceu; receptora é contada
# pela planilha de receptoras, e somar aqui duplicaria.
CATEGORIAS_FORA_DO_HEADCOUNT = ("EMBRIAO", "RECEPTORA")
# Colunas do mensal usadas só aqui. COTAS (%) é a fatia que a PG ainda tem; CONDICAO
# ATUAL registra a movimentação física do animal.
COL_MENSAL_COTAS = 16
COL_MENSAL_CONDICAO = 18
CONDICAO_SAIU = "SAIU DO HARAS"
# LOCAL que significa "está aqui". 'OUTROS' fica de fora de propósito: apesar de a
# CONTAGEM usar esse rótulo para o Centro de Treinamento, na prática ele é o destino
# de quem sai ("MUDOU O LOCAL PARA OUTROS" na aba MOVIMENTAÇÕES). Se o haras confirmar
# que OUTROS é só o CTE, é aqui que se acrescenta.
LOCAIS_NA_PROPRIEDADE = ("FAZENDA PAO GRANDE", "ARRENDAMENTO CESAR FURTADO")
# Layout da aba PLANTEL nos dois arquivos (0-based). O mensal é o mesmo roster com 3
# colunas a mais na frente e cabeçalho 3 linhas abaixo — por isso não dá pra apontar
# o mesmo leitor pros dois sem parametrizar.
PLANTEL_LAYOUT_SEMANAL = {"linha1": 2, "nome": 0, "categoria": 2, "status": 3, "local": 4}
PLANTEL_LAYOUT_MENSAL = {"linha1": 5, "nome": 3, "categoria": 5, "status": 6, "local": 7}


ROSTER_FONTE = "controle_mensal"      # gravado no snapshot; ver _conferir_delta


# "ainda está aqui" para o embrião = onde a receptora dele está. LOCAL SOCIO ou
# COMPRADOR significa que já saiu.
LOCAIS_RECEPTORA_NA_PG = ("PAO GRANDE", "ARRENDAMENTO CESAR FURTADO")


# Embrião de sociedade: 100% do sócio. Na aba ESTAÇÃO isso é COTAS EMBRIÃO vazia
# (ou zero) com SÓCIO EMBRIÃO = 1. Cota parcial (0,25 / 0,5) é embrião da PG COM
# sócio, que não é pendência de saída — a planilha do grupo dizia o mesmo separando
# em duas abas.
def _e_sociedade(cota, socio) -> bool:
    c = _to_num(cota)
    s = _to_num(socio)
    return (c is None or c == 0) and s is not None and s >= 1


def _embrioes_sociedade_pendentes() -> list:
    """Embriões em sociedade que ainda estão na PG.

    Regra do haras: a `EMBRIOES A ENTREGAR` NÃO serve aqui — ali só tem embrião
    vendido e não gestado, sem sociedade e sem vendido já gestado. O que vale é
    esta aba, com STATUS SOCIO e o embrião ainda em terra nossa."""
    locais = _receptoras_locais()
    wb = _load(_latest_estacao_master())
    ws = wb["ESTAÇÃO"]
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[0] is None or _s(r[35]) != SAFRA_ATUAL:
            continue
        if not _e_sociedade(r[5], r[6]):
            continue
        if _norm(r[17]) != "OK":            # não confirmado, não é pendência
            continue
        if r[21] is not None or r[23] is not None:   # abortou ou já pariu
            continue
        local = _norm(locais.get(_norm(r[11])))
        if local not in LOCAIS_RECEPTORA_NA_PG:      # já está com o sócio: saiu
            continue
        out.append({
            "nome": f"{_s(r[2])} x {_s(r[3])}",
            "receptora": _s(r[11]),
            "local": local,
            "estacao": _s(r[35]),
            "socio": _s(r[34]),
            "tipo": "SOCIEDADE",
            "especie": "EMBRIAO",
        })
    wb.close()
    return out


def _fora_linha(acc: list, r, L, motivo: str):
    """Linha que o roster descartou, com o motivo. Serve pra auditar a contagem
    (ver tools/excel_headcount.py) sem reescrever os filtros em outro lugar — regra
    duplicada é regra que sai de sincronia."""
    acc.append({
        "nome": _s(r[L["nome"]]), "categoria": _s(r[L["categoria"]]),
        "status": _s(r[L["status"]]), "local": _s(r[L["local"]]),
        "cota": r[COL_MENSAL_COTAS] if len(r) > COL_MENSAL_COTAS else None,
        "condicao": _s(r[COL_MENSAL_CONDICAO]) if len(r) > COL_MENSAL_CONDICAO else None,
        "obs": _s(r[COL_MENSAL_OBS]) if len(r) > COL_MENSAL_OBS else None,
        "motivo": motivo})


def _plantel_por_status(src: Path | None = None, quando: date | None = None) -> dict:
    """Roster do plantel a partir do CONTROLE_DE_PLANTEL mensal, na pasta PLANTEL.

    `src` e `quando` servem pra reconstruir um mês PASSADO (o comitê pede o
    fechamento do mês do deck, não o estado de hoje). Sem eles — que é como o
    fechamento semanal chama —, nada muda: arquivo mais recente e todos os
    overrides de FORA_NA_MAO.

    Antes vinha do `CONTROLE PLANTEL.xlsx` da pasta ATUALIZACAO SEMANAL, que é de
    divulgação. As regras estão no cabeçalho do commit e resumidas abaixo; cada uma
    saiu de comparar nome a nome os dois rosters, não de suposição.

    O nome do animal NÃO é identidade aqui: o mensal batiza o potro
    (`PRINCIPE MN DA PAO GRANDE`) enquanto o semanal o descrevia pelo cruzamento
    (`MACHO LIBRA DA PAO GRANDE X OLIMPO DO MH`). Quem identifica é nome+MAE+PAI."""
    if src is None:
        src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    wb = _load(src)
    ws = wb["PLANTEL"]
    L = PLANTEL_LAYOUT_MENSAL
    overrides = _overrides_validos(quando)
    vistos, linhas, descartadas = {}, [], []
    fora = {"status": 0, "categoria": 0, "duplicado": 0}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < L["linha1"] or r[L["nome"]] is None:
            continue
        nome = _s(r[L["nome"]])
        if not nome:
            continue
        if not _status_conta(r[L["status"]]):
            fora["status"] += 1
            _fora_linha(descartadas, r, L, "status fora do plantel")
            continue
        if _norm(_sem_cotista(nome)) in overrides:
            fora["na_mao"] = fora.get("na_mao", 0) + 1
            _fora_linha(descartadas, r, L, "override manual (ver FORA_NA_MAO)")
            continue
        categoria = _norm(r[L["categoria"]])
        if categoria in CATEGORIAS_FORA_DO_HEADCOUNT:
            fora["categoria"] += 1
            _fora_linha(descartadas, r, L, "embrião/receptora (conta em outra lista)")
            continue
        # Saiu do haras E a PG não tem cota nenhuma: acabou. Uma coisa só não basta —
        # animal no sócio segue no plantel enquanto a PG tem parte dele (29 estão
        # nessa situação), e vendido de cota zero que ainda não saiu continua aqui
        # (os 5 'vendido pendente saída', o PRADO). Juntas, as duas dizem que o
        # animal não é mais da casa nem está mais nela.
        cota = r[COL_MENSAL_COTAS] if len(r) > COL_MENSAL_COTAS else None
        condicao = _norm(r[COL_MENSAL_CONDICAO]) if len(r) > COL_MENSAL_CONDICAO else ""
        if condicao == CONDICAO_SAIU and (cota is None or not cota):
            fora["saiu_sem_cota"] = fora.get("saiu_sem_cota", 0) + 1
            _fora_linha(descartadas, r, L, "saiu do haras e sem cota")
            continue
        # REVERTIDO em 04/09/2026. Havia aqui uma exclusão de "compra não entregue"
        # baseada na OBS "o vendedor entregará", disparada pela ELEITA DA PAO
        # GRANDE. Leitura errada: ELEITA é DA PAO GRANDE e o STATUS='VENDIDO' é
        # venda NOSSA — o vendedor somos nós, o animal está na fazenda e sai da
        # conta só quando for entregue (aí o status vira 'VENDIDO E ENTREGUE').
        # Com a exclusão, FPG saía 88 contra 89 divulgados; sem ela, 71 animais +
        # 18 receptoras = 89, exato.
        mae = _s(r[COL_MENSAL_MAE]) if len(r) > COL_MENSAL_MAE else None
        pai = _s(r[COL_MENSAL_PAI]) if len(r) > COL_MENSAL_PAI else None
        chave = _chave_animal(nome, mae, pai)
        if chave in vistos:
            fora["duplicado"] += 1
            _fora_linha(descartadas, r, L, "linha repetida por cotista")
            continue
        # o nome sem cotista é o que vai para o roster: com o cotista, o mesmo animal
        # apareceria como saída de uma semana e entrada na outra ao mudar de sócio
        limpo = _sem_cotista(nome)
        vistos[chave] = limpo
        linhas.append({"nome": limpo, "categoria": _s(r[L["categoria"]]),
                       "status_plantel": _s(r[L["status"]]), "local": _s(r[L["local"]]),
                       "mae": mae, "pai": pai})
    wb.close()
    if fora.get("na_mao"):
        print(f"  [marretada] {fora['na_mao']} linha(s) fora da contagem por override "
              f"manual — corrigir na origem e apagar de FORA_NA_MAO:")
        for k, motivo in overrides.items():
            print(f"    - {k}: {motivo}")
    compra_pend = fora.get("compra_nao_entregue", 0)
    extra = f", {compra_pend} compra(s) ainda não entregue(s)" if compra_pend else ""
    print(f"  [roster] {len(vistos)} animais em {src.name} "
          f"(fora: {fora['status']} por status, {fora['categoria']} embrião/receptora, "
          f"{fora['duplicado']} linha(s) repetida(s) por cotista{extra})")
    return {"roster": sorted(set(vistos.values())), "linhas": linhas,
            "descartadas": descartadas,
            "fonte": src.name, "roster_fonte": ROSTER_FONTE}


# CONTROLE_DE_PLANTEL mensal, aba PLANTEL: colunas que não estão no layout mínimo.
# MAE/PAI/NASCIMENTO são o que permite achar o potro sem depender do nome dele.
COL_MENSAL_MAE = 8
COL_MENSAL_PAI = 9
COL_MENSAL_NASCIMENTO = 10


def _nascimentos_do_roster(ini: date, fim: date) -> list:
    """Nascimentos da janela pela coluna NASCIMENTO do roster mensal.

    Por data e filiação, nunca por nome: o potro entra no roster com nome próprio
    (`PRINCIPE MN DA PAO GRANDE`) ou com o cruzamento (`MACHO LIBRA x OLIMPO`),
    e as duas formas convivem. Data e MAE/PAI existem nas duas."""
    src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    wb = _load(src)
    ws = wb["PLANTEL"]
    L = PLANTEL_LAYOUT_MENSAL
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < L["linha1"] or r[L["nome"]] is None:
            continue
        d = _dt(r[COL_MENSAL_NASCIMENTO]) if len(r) > COL_MENSAL_NASCIMENTO else None
        if not d or not (ini <= d <= fim):
            continue
        nome = _s(r[L["nome"]])
        m = RE_PRODUTO.search(_norm(nome))
        out.append({
            "produto": nome,
            "mae": _s(r[COL_MENSAL_MAE]), "pai": _s(r[COL_MENSAL_PAI]),
            "receptora": m.group(1) if m else None,
            "socio": _limpa_socio(r[COL_MENSAL_NOME_SOCIO])
                     if len(r) > COL_MENSAL_NOME_SOCIO else None,
            "data": d.isoformat(),
            "local": _s(r[L["local"]]),
        })
    wb.close()
    return out


def _status_plantel_mensal() -> dict:
    """Vendidos pendentes e terceiros pela coluna STATUS PLANTEL do CONTROLE_DE_PLANTEL
    mensal. Devolve as listas cruas; quem chama decide o que é zero e o que é ausência."""
    src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    wb = _load(src)
    ws = wb["PLANTEL"]
    L = PLANTEL_LAYOUT_MENSAL
    vendidos_pend, terceiros, soc_pend, marcado = [], [], [], False
    # A égua vendida em cotas tem UMA linha por cotista — PEDRITA DA PAO GRANDE
    # aparece como (CARLA) e (EDUARDO), e as duas estão 'VENDIDO PENDENTE DE SAIDA'
    # na fazenda. É um animal só: contar as duas dava 6 onde o haras publica 05.
    # Mesma deduplicação que o roster já faz em _plantel_por_status.
    vend_vistos = set()
    # NOME SOCIO / COTAS (%) ficam fora do layout mínimo porque só este trecho usa.
    # Indexado pela receptora do fim do nome: é o que o roster semanal e o mensal
    # têm em comum (o mensal escreve a data no meio e acentua o garanhão).
    socio_por_recep = {}
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < L["linha1"] or r[L["nome"]] is None:
            continue
        nome = _s(r[L["nome"]])
        if not nome:
            continue
        status_plantel = _norm(r[L["status"]])
        categoria, local = _norm(r[L["categoria"]]), _s(r[L["local"]])
        m_rec = RE_PRODUTO.search(_norm(nome))
        soc = _limpa_socio(r[COL_MENSAL_NOME_SOCIO]) if len(r) > COL_MENSAL_NOME_SOCIO else None
        if m_rec and soc:
            socio_por_recep[m_rec.group(1)] = soc
        if _e_vendido_pendente(status_plantel, local) and _sem_cotista(nome) not in vend_vistos:
            marcado = True
            vend_vistos.add(_sem_cotista(nome))
            vendidos_pend.append({"nome": _sem_cotista(nome), "local": local, "cota": None,
                                  "comprador": None, "tipo": "VENDA",
                                  "obs": _s(r[L["status"]]), "reposicao": False,
                                  "categoria": categoria,
                                  "especie": "EMBRIAO" if categoria == "EMBRIAO" else None})
        if STATUS_TERCEIRO in status_plantel:
            marcado = True
            terceiros.append({"nome": nome, "local": local, "categoria": categoria,
                              "status_plantel": _s(r[L["status"]]),
                              "especie": "EMBRIAO" if categoria == "EMBRIAO" else None})
        obs = _norm(r[COL_MENSAL_OBS]) if len(r) > COL_MENSAL_OBS else ""
        # Mesma regra dos vendidos: pendente de saída é quem ainda está aqui. A
        # GABRIELA ELFAR e o ORFEU MH2 seguem marcados na OBS, mas o LOCAL já é SOCIO
        # — são justamente as duas saídas desta semana. Contá-los punha 3 animais onde
        # o haras publica 01 (a LIBRA DA PAO GRANDE, essa sim ainda na fazenda).
        if STATUS_SOCIEDADE_PENDENTE in obs and _norm(local) in LOCAIS_NA_PROPRIEDADE:
            soc_pend.append({"nome": nome, "local": local, "categoria": categoria,
                             "obs": _s(r[COL_MENSAL_OBS]),
                             "especie": "EMBRIAO" if categoria == "EMBRIAO" else None})
    wb.close()
    return {"fonte": src.name, "marcado": marcado,
            "vendidos_pendentes": vendidos_pend, "terceiros": terceiros,
            "sociedade_pendentes": soc_pend, "socio_por_recep": socio_por_recep}


# Marca DIRETA na ESTAÇÃO — ajuste combinado com o haras em 28/08/2026, ainda sem
# nenhuma linha lançada nos dois arquivos daquela data. Coluna OBSERVAÇÃO (AF,
# índice 31) = 'pendente de saída' quando STATUS (AG, índice 32) é VENDIDO ou
# SOCIO. Vira A fonte assim que tiver linha marcada — mais direta que os dois
# jeitos indiretos abaixo (ENTREGAR pra venda, COTAS/SÓCIO EMBRIÃO pra
# sociedade), que ficam só de fallback enquanto a coluna está vazia.
COL_ESTACAO_OBS = 31
COL_ESTACAO_STATUS = 32
STATUS_ESTACAO_VENDA = "VENDIDO"
STATUS_ESTACAO_SOCIO = "SOCIO"


def _embrioes_pendentes_estacao() -> list:
    """Embriões marcados pendentes de saída direto na ESTAÇÃO. Vazio até o haras
    começar a marcar — quem chama trata lista vazia como 'ainda sem marca',
    não como 'não há pendência'."""
    # AJUSTE NA MARRETADA (28/08/2026) — só nesta exibição do hub, pra liberar o
    # relatório sem esperar o haras corrigir a fonte: LIBRA DA PAO GRANDE x LATINO
    # DA PAO GRANDE está com STATUS='SOCIO' na ESTAÇÃO, mas o comprador (LAEL) é
    # 100% dono, não é sociedade de verdade — STATUS mal marcado na origem. TIRAR
    # esta exclusão assim que o STATUS da linha virar VENDIDO na planilha (aí ela
    # entra sozinha em 'vendidos', sem precisar disto aqui).
    EXCECAO_TEMP_NAO_SOCIEDADE = {("LIBRA DA PAO GRANDE", "LATINO DA PAO GRANDE")}

    locais = _receptoras_locais()
    master = _latest_estacao_master()
    wb = _load(master)
    ws = wb["ESTAÇÃO"]
    out = []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i < 3 or r[2] is None:
            continue
        obs = _norm(r[COL_ESTACAO_OBS]) if len(r) > COL_ESTACAO_OBS else ""
        if "PENDENTE" not in obs or "SAIDA" not in obs:
            continue
        if (_s(r[2]), _s(r[3])) in EXCECAO_TEMP_NAO_SOCIEDADE:
            continue
        status = _norm(r[COL_ESTACAO_STATUS]) if len(r) > COL_ESTACAO_STATUS else ""
        if status == STATUS_ESTACAO_VENDA:
            tipo = "VENDA"
        elif status == STATUS_ESTACAO_SOCIO:
            tipo = "SOCIEDADE"
        else:
            print(f"  [embriões] {_s(r[2])} x {_s(r[3])}: OBS marca pendente de "
                  f"saída mas STATUS ({_s(r[COL_ESTACAO_STATUS])!r}) não é "
                  f"VENDIDO nem SOCIO — fora da conta, conferir")
            continue
        # categoria/local no mesmo formato das outras linhas da tabela (animal e
        # sociedade-por-OBS) — sem isto a linha aparecia com LOCAL/CATEGORIA em
        # branco no dashboard, e foi isso que gerou a dúvida do haras em 28/08/2026:
        # a marca nova já tinha pego a linha certa, só faltava exibir completo.
        local = _norm(locais.get(_norm(r[11]))) if len(r) > 11 else None
        out.append({
            "nome": f"{_s(r[2])} x {_s(r[3])}", "local": local, "categoria": "EMBRIAO",
            "cota": None, "comprador": _s(r[34]) if len(r) > 34 else None,
            "tipo": tipo, "obs": _s(r[COL_ESTACAO_OBS]), "reposicao": False,
            "especie": "EMBRIAO",
        })
    wb.close()
    return out


# Embrião comercial pendente de saída: aba ENTREGAR do "EMBRIOES A ENTREGAR - A
# RECEBER". 'Cota PG' < 1 = sociedade; = 1 = venda 100%.
#
# PRONTO tem DOIS estados e só um deles é pendência de saída:
#   PRONTO - AGUARDANDO ENTREGA  (4)  o embrião vai embora  -> PENDENTE
#   PRONTO - NASCE NA PG         (3)  o produto nasce aqui  -> NÃO é pendência
# Casar só o prefixo 'PRONTO' misturava os dois e inflava o card de sociedade em 3.
# Os demais estados — A FAZER, ENTREGUE, NASCIDO, CANCELADO, REPOSIÇÃO — já ficavam
# de fora. A aba EMBRIOES VENDIDOS do "Animais para sair" está vazia e não é usada.
EMB_STATUS_PENDENTE = "AGUARDANDO ENTREGA"
# Registro de embrião que JÁ entrou como venda pendente. A venda é lançada numa
# semana e a entrega demora: o corte por data da venda (que existe para não arrastar
# as ~25 linhas 'A fazer' de 2022-2025) tirava o embrião do card na semana seguinte.
# Em 18/09/2026 foi o NATUREZA DA PAO GRANDE x LEGITIMO ELFAR DA MONTE BRANCO,
# vendido em 07/09 e publicado no fechamento de 10/09, que sumiu sozinho. Pendência é
# ESTADO: uma vez dentro, fica até o status virar terminal.
VENDIDOS_EMB_EXTRA = BASE_DIR / "_cache" / "vendidos_embrioes_extra.json"
# Status que encerram a pendência — entregue, nascido, cancelado ou virou reposição.
EMB_STATUS_TERMINAL = ("ENTREGUE", "NASCIDO", "CANCELADO", "REPOSICAO")


def _embrioes_pendentes(ini: date | None = None, fim: date | None = None) -> list:
    """Embriões prontos e aguardando entrega, tipo SOCIEDADE (cota parcial) ou VENDA.

    Conta também a venda LANÇADA NESTA SEMANA, mesmo com 'Status embrião' ainda em
    'A fazer': vendido e não entregue é pendente por definição. O corte é a data da
    venda dentro da janela — sem ele, ou a linha nova some (foi o caso do NATUREZA
    DA PAO GRANDE x LEGITIMO ELFAR, vendido em 07/09/2026) ou entrariam as 25
    linhas 'A fazer' da planilha inteira, vendas de 2022 a 2025."""
    reg = {}
    if VENDIDOS_EMB_EXTRA.exists():
        try:
            reg = json.loads(VENDIDOS_EMB_EXTRA.read_text(encoding="utf-8"))
        except Exception:
            reg = {}
    elif fim:
        # BOOTSTRAP: o registro é novo, mas o card já publicava embrião pendente antes
        # dele existir. Semeia do arquivo de linhas da semana anterior para não perder
        # quem já estava lá — sem isto o NATUREZA x LEGITIMO ELFAR, publicado em
        # 10/09/2026, sumiria justamente na semana em que o registro foi criado.
        anterior = _arquivo_anterior(fim.isoformat()).get("pendentes_saida") or []
        for x in anterior:
            if x.get("especie") == "EMBRIAO" and x.get("tipo") == "VENDA":
                reg[_norm(x.get("nome"))] = {"comprador": x.get("comprador"),
                                             "data_venda": None, "origem": "snapshot"}
        if reg:
            print(f"  [embriões] registro de pendentes semeado com {len(reg)} venda(s) "
                  f"já publicada(s): " + "; ".join(sorted(reg)))
    wb = _load(EMB_COMERCIAIS)
    ws = wb["ENTREGAR"]
    out, cols, ficam, novos, mantidos = [], None, [], [], []
    for i, r in enumerate(ws.iter_rows(values_only=True), start=1):
        if i == 3:
            cols = {n: _col_idx(r, n) for n in
                    ("ID Embrião", "Doadora", "Garanhão", "Comprador", "Cota PG",
                     "Status embrião", "Observação", "Data venda")}
            continue
        # Linha sem ID ainda é venda: o ID entra depois. Exigir ID descartava a
        # venda recém-digitada — em 10/09/2026 o embrião NATUREZA DA PAO GRANDE x
        # LEGITIMO ELFAR (venda de 07/09, VENDA DIRETA) não chegava nem a ser
        # avaliado. Identidade cai para doadora x garanhão quando falta o ID.
        if cols is None or (r[cols["ID Embrião"]] is None and r[cols["Doadora"]] is None):
            continue
        status = _norm(r[cols["Status embrião"]])
        # Status vazio numa linha com venda lançada = venda nova, ainda sem
        # tratamento: conta como pendente e sai avisada, para não ficar invisível
        # esperando alguém preencher a coluna.
        sem_status = not status and r[cols["Doadora"]] is not None
        venda_na_semana = False
        if ini and fim and cols.get("Data venda") is not None:
            dv = _dt(r[cols["Data venda"]])
            venda_na_semana = bool(dv and ini <= dv <= fim)
        # Já registrado como pendente numa semana anterior continua pendente até a
        # entrega — é o que faz a venda sobreviver à semana em que foi lançada.
        chave_reg = _norm(f'{_s(r[cols["Doadora"]])} x {_s(r[cols["Garanhão"]])}')
        terminal = any(t in status for t in EMB_STATUS_TERMINAL)
        se_mantem = chave_reg in reg and not terminal
        if se_mantem:
            mantidos.append(f'{_s(r[cols["Doadora"]])} x {_s(r[cols["Garanhão"]])}'
                            f' [{_s(r[cols["Status embrião"]]) or "sem status"}]')
        if chave_reg in reg and terminal:
            reg.pop(chave_reg, None)
        if (EMB_STATUS_PENDENTE not in status and not sem_status
                and not venda_na_semana and not se_mantem):
            if status.startswith("PRONTO"):
                ficam.append(f'{_s(r[cols["ID Embrião"]])} ({_s(r[cols["Status embrião"]])})')
            continue
        if sem_status or venda_na_semana:
            novos.append(f'{_s(r[cols["Doadora"]])} x {_s(r[cols["Garanhão"]])}'
                         + (f' [{_s(r[cols["Status embrião"]])}]' if status else ' [sem status]'))
        cota = r[cols["Cota PG"]]
        # Cota ZERADA é 100% vendido (regra do Arthur, 11/09/2026), não sociedade:
        # a coluna guarda a fatia que fica com a PG, e zero quer dizer que não
        # sobrou nada. Só cota ENTRE 0 e 1 é sociedade.
        try:
            parcial = cota is not None and 0 < float(cota) < 1
        except (TypeError, ValueError):
            parcial = False
        out.append({
            "nome": f'{_s(r[cols["Doadora"]])} x {_s(r[cols["Garanhão"]])}',
            "id": _s(r[cols["ID Embrião"]]), "local": None, "cota": cota,
            "comprador": _s(r[cols["Comprador"]]),
            # Venda LANÇADA NA SEMANA é venda, mesmo com cota parcial: a abertura
            # publicada pelo haras em 11/09/2026 põe 2 embriões em VENDIDOS, e os
            # únicos candidatos são o LIBRA x OLIMPO (cota 100%, 'Pronto -
            # Aguardando Entrega') e o NATUREZA x LEGITIMO (cota 50%, vendido em
            # 07/09). Cota parcial só manda para SOCIEDADE quando a venda é antiga
            # — aí o que sobrou é a sociedade, não o movimento da semana.
            # `se_mantem` entra junto com `venda_na_semana`: o que nasceu VENDA não
            # vira sociedade só por a semana da venda ter passado. O NATUREZA (cota
            # 50%, vendido 07/09/2026) é venda pendente enquanto não for entregue.
            "tipo": ("SOCIEDADE" if (parcial and not venda_na_semana and not se_mantem)
                     else "VENDA"),
            "obs": _s(r[cols["Status embrião"]]), "reposicao": False,
            "especie": "EMBRIAO",
        })
        # Só VENDA entra no registro. Registrar tudo que passou do filtro fazia o
        # embrião de SOCIEDADE virar venda na rodada seguinte (ele voltava por
        # `se_mantem`), e os vendidos pendentes saltaram de 7 para 10 numa semana em
        # que nada mudou na planilha.
        if out[-1]["tipo"] == "VENDA" and not terminal:
            reg.setdefault(chave_reg, {"comprador": _s(r[cols["Comprador"]]),
                                       "data_venda": _s(r[cols["Data venda"]])})
    if ficam:
        print(f"  [embriões] {len(ficam)} pronto(s) que NÃO saem, fora da pendência: "
              + "; ".join(ficam))
    if novos:
        print(f"  [embriões] {len(novos)} venda(s) da semana contada(s) como pendente "
              f"mesmo sem 'Pronto - Aguardando Entrega': " + "; ".join(novos))
    if mantidos:
        print(f"  [embriões] {len(mantidos)} venda(s) de semana anterior ainda "
              f"pendente(s) de entrega: " + "; ".join(mantidos))
    VENDIDOS_EMB_EXTRA.parent.mkdir(parents=True, exist_ok=True)
    VENDIDOS_EMB_EXTRA.write_text(json.dumps(reg, ensure_ascii=False, indent=2),
                                  encoding="utf-8")
    wb.close()
    return out


def build_pendentes(rep: Report):
    plantel = _plantel_por_status()
    rep.roster = plantel["roster"]
    _LINHAS_BRUTAS["roster"] = plantel["linhas"]
    # as descartadas vêm junto: é nelas que está a CAUSA de quem sumiu do roster
    # (ver _saidas_por_mudanca_de_local)
    _LINHAS_BRUTAS["roster_descartadas"] = plantel["descartadas"]
    rep.fontes["roster_plantel"] = plantel["fonte"]

    # O "Animais para sair" saiu do pipeline: mora na pasta de divulgação (o que foi
    # enviado ao grupo), está congelado em 24/07/2026 e só alimentava a marcação de
    # reposição, que já era apenas um aviso. Vendido pendente vem do STATUS PLANTEL do
    # controle mensal; sociedade, da aba ESTAÇÃO.
    pend = []
    pend_emb = _embrioes_pendentes(date.fromisoformat(rep.semana_inicio),
                                   date.fromisoformat(rep.semana_fim))

    # VENDIDOS PENDENTES: fonte é o STATUS PLANTEL do controle mensal. O "Animais para
    # sair" só entra se ninguém estiver marcado lá — e aí com aviso, porque ele está
    # congelado desde 24/07/2026 e perde os marcados depois disso.
    mensal = _status_plantel_mensal()
    rep.fontes["status_plantel"] = mensal["fonte"]
    _SOCIO_ROSTER.clear()
    _SOCIO_ROSTER.update(mensal.get("socio_por_recep") or {})
    mensal_terceiros_embrioes = [t for t in mensal["terceiros"]
                                 if t.get("especie") == "EMBRIAO"]
    # ERRADO até 28/08/2026: usava DE TERCEIRO + CATEGORIA=EMBRIAO do roster mensal,
    # porque coincidiu em número (2) com o relatório de uma semana ("07 = 05 animais
    # e 02 embriões"). O haras confirmou direto: essas 2 linhas (HORAH..., BEGONIA...)
    # NÃO são pendência de saída — são embrião de TERCEIRO na propriedade, outra
    # coisa (mesmo raciocínio de 'terceiros = vendidos' logo abaixo, que já tinha
    # sido corrigido por coincidência igual). O embrião de venda pendente de
    # verdade é o de cota 100% em EMB_COMERCIAIS/ENTREGAR, 'Pronto - Aguardando
    # Entrega' — é o que _embrioes_pendentes() já lia, sem ninguém consumir.
    emb_estacao = _embrioes_pendentes_estacao()

    def _chave_emb(x):
        """doadora x garanhão, sem a cauda de data/receptora.

        A mesma prenhez aparece com nomes diferentes nas duas fontes: na comercial
        é 'ADRENALINA DA PAO GRANDE x XODO PORTEIRA AZUL' e no roster é a mesma
        coisa mais '14/03/2026 RECEP 532'. Deduplicar pelo nome cru contava o
        embrião duas vezes."""
        n = _norm(x.get("nome"))
        n = re.split(r"\s+\d{1,2}/\d{1,2}/\d{2,4}|\s+RECEP", n)[0]
        return " ".join(n.split())

    def _uniao_emb(principal, complemento):
        vistos = {_chave_emb(x) for x in principal}
        return principal + [x for x in complemento if _chave_emb(x) not in vistos]

    # FONTE do embrião pendente é a planilha comercial (EMB_COMERCIAIS/ENTREGAR,
    # coluna 'Status embrião') — decisão do Arthur em 11/09/2026. A marca da ESTAÇÃO
    # entra como COMPLEMENTO, não como substituta: antes ela vencia sempre que
    # existisse, e escondia o que só a comercial tem — o NATUREZA DA PAO GRANDE x
    # LEGITIMO ELFAR, vendido em 07/09, ficava invisível nos dois cards.
    vend_embrioes = _uniao_emb([e for e in pend_emb if e["tipo"] == "VENDA"],
                               [e for e in emb_estacao if e["tipo"] == "VENDA"])
    # REPOSIÇÃO não é venda pendente: o animal está saindo para repor outro, não para
    # um comprador. O STATUS PLANTEL não tem essa marca — ela vive na coluna de obs do
    # Animais para sair —, então cruzamos os dois pelo núcleo do nome. Essa regra
    # existia antes da migração para o STATUS PLANTEL e se perdeu no caminho: era o que
    # separava os nossos 5 dos 4 do relatório.
    # REPOSIÇÃO CONTA. Cheguei a excluir, apoiado num rascunho de relatorio que dizia
    # "04 animais"; a conferencia por movimentacao derrubou isso — o STATUS PLANTEL tem
    # 5 marcados e o relatorio fechado diz "05 animais", reposicao inclusa. Fica so o
    # aviso, porque a natureza da saida e diferente e alguem pode querer separar.
    reposicoes = {_nucleo_nome(x["nome"]) for x in pend if x["reposicao"]}
    if mensal["vendidos_pendentes"]:
        marcados = mensal["vendidos_pendentes"]
        repostos = [x for x in marcados if _nucleo_nome(x["nome"]) in reposicoes]
        if repostos:
            print("  [pendentes] entre os vendidos pendentes ha reposição (sai para "
                  "repor outro animal, nao para comprador): "
                  + "; ".join(x["nome"] for x in repostos))
        vendidos = marcados + vend_embrioes
        fonte_vendidos = "status_plantel"
    else:
        vendidos = [p for p in pend if p["tipo"] == "VENDA" and not p["reposicao"]]
        fonte_vendidos = "animais_para_sair"
        print(f"  [terceiros] nenhum '{STATUS_VENDIDO_PENDENTE}' no STATUS PLANTEL de "
              f"{mensal['fonte']}; vendidos pendentes caindo no Animais para sair "
              f"(congelado em 24/07/2026)")

    # SOCIEDADE pendente = animais + embriões (regra do relatório desde 07/08/2026).
    # Até 28/08/2026 não havia marca viva pra animal em sociedade — só o "Animais
    # para sair", congelado em 24/07 — e soc_animais ficava sempre vazio (o
    # relatório de 21/08 confirmava: "01 (embrião)", nenhum animal). O haras passou
    # a marcar direto na coluna OBS do roster mensal com a frase
    # STATUS_SOCIEDADE_PENDENTE, mesma ideia do VENDIDO PENDENTE SAIDA — então a
    # fonte agora é viva e o teste vira leitura direta, igual aos vendidos.
    # A marca de sociedade vive na OBS do roster mensal, e a linha marcada pode ser
    # de EMBRIÃO — foi o caso em 10/09/2026 com 'ADRENALINA DA PAO GRANDE X XODO
    # PORTEIRA AZUL 14/03/2026 RECEP 532' (CATEGORIA=EMBRIAO), que entrava no card
    # de ANIMAIS: 4 animais / 0 embriões, contra "03" do relatório, que conta
    # animal. Separar por categoria acerta os dois cards e mantém o total.
    soc_animais = [x for x in mensal["sociedade_pendentes"]
                   if _norm(x.get("categoria")) != "EMBRIAO"]
    soc_emb_roster = [x for x in mensal["sociedade_pendentes"]
                      if _norm(x.get("categoria")) == "EMBRIAO"]
    # Embrião de sociedade: preferência é a marca direta na ESTAÇÃO (ver
    # _embrioes_pendentes_estacao, ajuste de 28/08/2026). Sem marca ainda, cai no
    # jeito indireto anterior — aba de sócios do grupo (COTAS/SÓCIO EMBRIÃO,
    # sem parto nem aborto).
    # SOCIEDADE continua saindo da marca (ESTAÇÃO, senão a indireta pela aba de
    # sócios). A planilha comercial tem 3 embriões de cota parcial em 'Pronto -
    # Aguardando Entrega' e a liberação do haras conta 3 — só os ANIMAIS —, então
    # puxar sociedade de lá inflava o card sem respaldo. A comercial manda só na
    # VENDA, que foi o que o Arthur pediu (o NATUREZA).
    if emb_estacao:
        soc_embrioes = [e for e in emb_estacao if e["tipo"] == "SOCIEDADE"]
    else:
        soc_embrioes = _embrioes_sociedade_pendentes()
    # embrião marcado na OBS do roster entra junto, sem duplicar — pela chave sem
    # data/receptora (_chave_emb), senão a mesma prenhez entra duas vezes com nomes
    # diferentes, que foi o que aconteceu com o ADRENALINA x XODO em 11/09/2026
    soc_embrioes = _uniao_emb(soc_embrioes, soc_emb_roster)
    sociedade = soc_animais + soc_embrioes
    rep.fontes["embrioes_pendentes"] = EMB_COMERCIAIS.name

    # TERCEIROS NA PROPRIEDADE = o que é de terceiro e ainda está aqui, ou seja, os
    # PENDENTES DE SAÍDA — é o que o próprio relatório diz no rótulo da linha:
    # "Total terceiros: 08 (vendidos pendentes)".
    #
    # NÃO é a contagem de STATUS PLANTEL 'DE TERCEIRO'. Aquilo são doadoras, matriz,
    # receptora e embriões de terceiro que passaram pela estação: em 14/08/2026 dava 8
    # e batia com o relatório por coincidência de número, não por ser a mesma coisa.
    # Nenhuma das 8 linhas está no roster do plantel, e 7 das 8 estão fora da fazenda.
    #
    # "NA PROPRIEDADE" é presença física — embrião não ocupa espaço, então não conta
    # aqui mesmo sendo 'vendido pendente'. Em 14/08 os dois totais empatavam (8 e 8,
    # sem embrião no lote), o que escondia a diferença; em 28/08 o relatório abriu os
    # dois: linha 4 "05 (vendidos pendentes)" só animal, linha 5 "07 (05 animais e 02
    # embriões)" com embrião. `terceiros = vendidos` tratava as duas linhas como a
    # mesma contagem e ficou errado assim que apareceu embrião no lote.
    terc_embrioes = [t for t in vendidos if t.get("especie") == "EMBRIAO"]
    terc_animais = [t for t in vendidos if t.get("especie") != "EMBRIAO"]
    terceiros = terc_animais

    # DOADORAS DE TERCEIROS: doadora de terceiro que está NA PROPRIEDADE — fazenda ou
    # arrendamento. As marcadas em LOCAL 'OUTROS' não contam: 'OUTROS' é para onde o
    # animal vai quando deixa o haras (a aba MOVIMENTAÇÕES registra "MUDOU O LOCAL PARA
    # OUTROS" nas saídas), e nenhuma delas aparece no roster do plantel. Contá-las dava
    # 4 onde o relatório escreve "--".
    de_terceiro = mensal["terceiros"]
    doadoras_terc = [t for t in de_terceiro
                     if t["categoria"] == "DOADORA"
                     and _norm(t["local"]) in LOCAIS_NA_PROPRIEDADE]
    fora = [t for t in de_terceiro
            if t["categoria"] == "DOADORA" and _norm(t["local"]) not in LOCAIS_NA_PROPRIEDADE]
    if fora:
        print(f"  [terceiros] {len(fora)} doadora(s) de terceiro fora da propriedade, "
              f"não contadas: " + "; ".join(f"{t['nome']} ({t['local']})" for t in fora))
    if not mensal["marcado"]:
        print(f"  [terceiros] nenhuma linha marcada com {STATUS_TERCEIRO} ou "
              f"{STATUS_VENDIDO_PENDENTE} em {mensal['fonte']}")

    rep.terceiros.update({
        "vendidos_pendentes": len(vendidos),
        "vendidos_pendentes_animais": len(terc_animais),
        "vendidos_pendentes_embrioes": len(terc_embrioes),
        "vendidos_pendentes_fonte": fonte_vendidos,
        "sociedade_pendentes": len(sociedade),
        "sociedade_pendentes_animais": len(soc_animais),
        "sociedade_pendentes_embrioes": len(soc_embrioes),
        "terceiros_propriedade": len(terceiros),
        "terceiros_animais": len(terc_animais),
        "terceiros_embrioes": len(terc_embrioes),
        "doadoras_terceiros": len(doadoras_terc) if doadoras_terc else None,
        "outros_terceiros": None,
    })
    rep.detalhe["doadoras_terceiros"] = doadoras_terc
    rep.detalhe["terceiros_propriedade"] = terceiros
    rep.detalhe["terceiros_vendidos"] = vendidos          # vendidos pendentes (KPI seção 5)
    # A lista de vendidos pendentes vai logo abaixo da seção 4 (é o mesmo conjunto do
    # "Total terceiros", já que o relatório oficial escreve "05 (vendidos pendentes)"
    # — as duas linhas SÃO a mesma coisa). Embrião pendente de venda tem lista própria
    # na seção 5, porque não é "terceiro na propriedade": embrião não ocupa espaço.
    rep.detalhe["terceiros_vendidos_embrioes"] = terc_embrioes
    # ANIMAIS e EMBRIÕES em listas separadas: são dois cards distintos no relatório
    # ("Animais em sociedade pendentes de saída" e "Embriões em sociedade aguardando
    # entrega"), e juntar os dois fazia a tabela de animais mostrar 4 com um EMBRIAO
    # no meio (ADRENALINA x XODO RECEP 532), contra os 3 da liberação.
    rep.detalhe["terceiros_sociedade"] = soc_animais
    rep.detalhe["terceiros_sociedade_embrioes"] = soc_embrioes
    # lista completa da seção 5 = o que os dois KPIs contam. Era `pend + pend_emb` (só
    # o "Animais para sair"), então os marcados no STATUS PLANTEL não apareciam.
    rep.detalhe["pendentes_saida"] = vendidos + sociedade


def _latest_animais_sair() -> Path:
    """'Animais para sair*.xlsx' — hoje só sociedade pendente (vendidos migraram pro
    STATUS PLANTEL). Pasta canônica é VENDAS/SAIDA DE ANIMAIS VENDIDOS."""
    return _resolver(ANIMAIS_SAIR_GLOB, ANIMAIS_SAIR_DIRS, "sociedade pendente",
                     requer_aba="ANIMAIS VENDIDOS")


# ------------------------------------------------------------------
# Δ headcount vs run anterior (histórico local leve, automático)
# ------------------------------------------------------------------
def build_headcount_delta(rep: Report, fim: date):
    HIST_HEADCOUNT.parent.mkdir(parents=True, exist_ok=True)
    hist = {}
    if HIST_HEADCOUNT.exists():
        try:
            hist = json.loads(HIST_HEADCOUNT.read_text(encoding="utf-8"))
        except Exception:
            hist = {}
    total = rep.headcount.get("total")
    # run anterior = maior data < fim
    prev = None
    for k in sorted(hist):
        if k < fim.isoformat():
            prev = hist[k]
    if prev is not None and total is not None:
        rep.headcount["delta"] = total - prev.get("total", total)
    else:
        rep.headcount["delta"] = None

    # Δ do TOTAL nao e o mesmo que o Δ do relatorio.
    # O relatorio escreve '+02 / -01' contando so ANIMAIS: 2 potros nascidos, 1
    # vendido. As receptoras que foram pro socio sairam da contagem (-2) e nao
    # aparecem ali. Resultado em 21/08/2026: total 203 -> 202 = -1, animais
    # 143 -> 144 = +1. As duas contas estao certas, medem coisas diferentes — e
    # comparar a nossa do total contra a dele de animais dava divergencia falsa.
    # Nas semanas sem movimento de receptora as duas coincidem, e por isso o
    # problema so apareceu quando duas receptoras sairam na mesma semana.
    det = rep.headcount.get("detalhe") or {}
    tg = det.get("TOTAL GERAL") or {}
    ani, rec = tg.get("animais"), tg.get("receptoras")
    pa, pr = prev.get("animais") if prev else None, prev.get("receptoras") if prev else None
    rep.headcount["delta_animais"] = (ani - pa) if None not in (ani, pa) else None
    rep.headcount["delta_receptoras"] = (rec - pr) if None not in (rec, pr) else None
    atual = {"total": total, "fpg": rep.headcount.get("fazenda_pg"),
             "arr": rep.headcount.get("arrendamento"),
             "cte": rep.headcount.get("cte"), "soc": rep.headcount.get("socio"),
             # abertura animais/receptoras: base do Δ de animais, que e o que o
             # relatorio publica
             "animais": tg.get("animais"), "receptoras": tg.get("receptoras")}
    # CONTAGEM idêntica à da semana passada, local por local, quase sempre significa
    # que a aba não foi atualizada — não que nada mudou. Em 31/07/2026 isso aconteceu:
    # o snapshot repetiu 205 de 24/07, Δ saiu 0, e o relatório oficial dizia 204 / -01.
    if prev is not None and atual == prev:
        print(f"  [headcount] CONTAGEM idêntica à de {max(k for k in hist if k < fim.isoformat())} "
              f"em todos os locais ({total} total) — conferir se a aba foi atualizada; "
              f"Δ desta semana sai 0 por isso")
    # grava snapshot desta run (idempotente por data)
    hist[fim.isoformat()] = atual
    HIST_HEADCOUNT.write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")


# ------------------------------------------------------------------
# Seção 4 — TERCEIROS / COMERCIAIS (embriões a entregar / a receber)
# ------------------------------------------------------------------
def build_comerciais(rep: Report):
    wb = _load(EMB_COMERCIAIS)
    painel = {}
    ws = wb["PAINEL"]
    grid = list(ws.iter_rows(values_only=True))
    for r in grid:
        cells = [_s(c) for c in r]
        for j, c in enumerate(cells):
            if c and c.upper() in ("A FAZER", "EM ANDAMENTO", "TOTAL VENDIDOS", "TOTAL COMPRADOS"):
                # o número costuma estar mais à direita na mesma linha
                nums = [x for x in r if isinstance(x, (int, float))]
                if nums:
                    painel[c] = nums[-1]
    rep.terceiros = {"painel": painel}
    # tabelas detalhe
    def _tab(sheet, hdr_row=3):
        ws = wb[sheet]
        data = list(ws.iter_rows(values_only=True))
        hdr = [_s(c) for c in data[hdr_row - 1]]
        rows = []
        for r in data[hdr_row:]:
            if all(c is None for c in r):
                continue
            if _s(r[0]) is None and _s(r[1]) is None:
                continue
            rows.append({(hdr[i] or f"c{i}"): _s(v) for i, v in enumerate(r) if hdr[i]})
        return rows
    rep.detalhe["a_entregar"] = _tab("ENTREGAR")
    rep.detalhe["a_receber"] = _tab("RECEBER")
    wb.close()


# ------------------------------------------------------------------
# Orquestração
# ------------------------------------------------------------------
def _calendario_dos_snapshots(hist: dict) -> list:
    """Seletor = semanas com snapshot (docx-semente ou capturadas pelo script).
    Chave = data de referência (fim da semana). Janela = docx anterior+1 .. este."""
    keys = sorted(k for k in hist if _is_iso(k))
    semanas = []
    for i, wid in enumerate(keys):
        ref = date.fromisoformat(wid)
        ini = (date.fromisoformat(keys[i - 1]) + timedelta(days=1)) if i > 0 else (ref - timedelta(days=7))
        semanas.append({
            "id": wid,
            "ini": ini.isoformat(),
            "fim": wid,
            "iso": ref.isocalendar()[1],
            "source": hist[wid].get("source"),
        })
    return semanas


def _is_iso(s: str) -> bool:
    try:
        date.fromisoformat(s)
        return True
    except (ValueError, TypeError):
        return False


def build_report(ini: date, fim: date) -> Report:
    rep = Report(semana_inicio=ini.isoformat(), semana_fim=fim.isoformat())
    # ANTES dos builds: build_movimentacao -> _transferencias_internas compara com o
    # snapshot da semana anterior usando `wid < rep.semana_atual`. Atribuído depois,
    # semana_atual era "" ali, nenhuma semana passava no teste e o diff caía sempre no
    # bootstrap contra o arquivo anterior de receptoras — em 14/08/2026 isso recontou
    # as 14 transferências de 07/08 (o diff real dos snapshots é 0).
    rep.semana_atual = fim.isoformat()           # semana de referência = data do fechamento
    build_producao(rep, ini, fim)
    build_receptoras(rep)
    build_headcount(rep)
    build_headcount_delta(rep, fim)
    build_movimentacao(rep, ini, fim)
    build_comerciais(rep)
    build_pendentes(rep)
    rep.docx_ref = _load_docx_ref()               # relatórios oficiais (validação + seed do 1º caso)
    rep.detalhe["cancelamentos_pendentes"] = _cancelamentos_pendentes(rep)
    _compute_movimento(rep)                       # saídas/entradas = diff da população contada
    _paricoes_do_roster(rep)                      # potro no roster sem parição na ESTAÇÃO
    _registra_caminhos(rep)                       # pasta de cada fonte, p/ auditoria
    _aplica_manual(rep)                           # campos sem fonte de planilha
    # ANTES do piso: o acumulado da safra soma as confirmações que só a planilha de
    # receptoras tem, e o piso precisa gravar o número já com elas — senão a semana
    # seguinte publica o valor sem esse pedaço.
    _compute_confirmados_diff(rep)                # confirmados na semana = diff de confirmados (forward)
    _acumulado_nunca_cai(rep)                     # agregador da safra, nao cai
    # UMA vez, no fim: chamado no meio do caminho ele via as entradas ainda sem os
    # nascimentos e acusava movimentacao fantasma que se resolvia duas linhas depois
    _conferir_delta(rep)
    _avisar_pasta_de_saida()
    _avisar_fontes_velhas(ini, fim)                # BLOQUEIA se a fonte for velha
    _arquivar_linhas(rep)                         # historico linha a linha
    _persist_snapshot(rep)                        # congela snapshot CALCULADO desta semana
    rep.calendario = _calendario_dos_snapshots(rep.snapshots)
    return rep


def _load_hist() -> dict:
    if HIST_SNAPSHOTS.exists():
        try:
            return json.loads(HIST_SNAPSHOTS.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _populacao_contada(roster, receptoras_locais) -> list:
    """Conjunto que o headcount conta: animais do plantel + receptoras nos NOSSOS
    locais. Receptora que vai pro sócio sai da contagem (CONTAGEM só tem receptora
    na fazenda e no arrendamento), então tem de contar como saída."""
    return sorted(set(roster or []) | set(receptoras_locais or {}))


def _descreve_mov(nome: str, info_ant: dict, info_atual: dict) -> dict:
    """Uma saída/entrada com contexto: '309' virou 'receptora 309, prenha de JAVA x
    QUEBRUTO, Pao Grande -> sócio'."""
    ant, atual = info_ant.get(nome), info_atual.get(nome)
    ref = atual or ant or {}
    return {
        "animal": nome,
        "tipo": "RECEPTORA" if (ant or atual) else "ANIMAL",
        "local_saida": (ant or {}).get("local"),
        "local_entrada": (atual or {}).get("local"),
        "status": ref.get("status"),
        "embriao": ref.get("embriao"),
        "obs": ref.get("obs"),
    }


def _roster_congelado(semana: str) -> list:
    """Roster da última semana congelada antes desta, RENORMALIZADO com as regras
    de hoje.

    O snapshot guarda o nome como a normalização daquele dia produziu. Quando a
    regra muda, o mesmo animal parece outro: em 04/09/2026 a correção do cotista
    no meio do nome transformou 'MACHO ... (CARLA) 17/08/2024 RECEP 46' em
    'MACHO ... 17/08/2024 RECEP 46', e o diff leu isso como potro NOVO — virou
    parição da safra e somou +1 no acumulado da estação, que o piso então travou
    em 2 contra 1 divulgado. Renormalizar na leitura faz o diff comparar maçã
    com maçã mesmo depois de mexer na normalização."""
    hist = _load_hist()
    for wid in sorted(hist, reverse=True):
        if wid < semana and hist[wid].get("roster"):
            return sorted({_sem_cotista(x) for x in hist[wid]["roster"]})
    return []


def _mapa_receptoras_anterior(semana: str) -> dict:
    """Mapa {receptora: local} da ultima semana congelada antes desta."""
    hist = _load_hist()
    prev = {}
    for wid in sorted(hist):
        if wid < semana and hist[wid].get("receptoras_locais"):
            prev = hist[wid]["receptoras_locais"]
    return prev


def _avisa_troca_de_fonte(rep: Report):
    """Rodar a MESMA semana em cima de arquivo diferente muda o número sem que
    nenhuma regra tenha mudado.

    O controle de plantel vive em cópias irmãs do mesmo mês — '..._AGO_26.xlsx' e
    '..._EDITAR OUTUBRO_..._AGO_26.xlsx' —, e a escolha é por mtime: quem salvou
    por último ganha. Em 04/09/2026 isso aconteceu no meio do fechamento: a cópia
    EDITAR OUTUBRO reproduzia a semana divulgada (24/24 no placar) e, depois de a
    Ana salvar a outra, a mesma semana passou a fechar com 6 saídas em vez de 8,
    sem NASDAQ e com 2 pendentes a menos — 16/24. Sem aviso, isso vira número
    publicado.
    """
    hist = _load_hist()
    ant = (hist.get(rep.semana_atual) or {}).get("fontes_caminhos") or {}
    if not ant:
        return
    mudou = [(r, ant[r], rep.fontes_caminhos.get(r))
             for r in ant if rep.fontes_caminhos.get(r) and rep.fontes_caminhos[r] != ant[r]]
    if mudou:
        print(f"  [fonte] ATENÇÃO: a semana {rep.semana_atual} já foi fechada com OUTRO "
              f"arquivo. Conferir qual é o certo ANTES de publicar:")
        for r, antes, agora in mudou:
            print(f"    - {r}: {antes}")
            print(f"      agora: {agora}")


def _cancelamentos_pendentes(rep: Report) -> list:
    """Cancelamento que a planilha registrou na cota mas não no cadastro.

    Ver scripts/_pg_cancelamentos.py: "VENDA CANCELADA" devolve a cota pra Pao
    Grande, e enquanto STATUS/CONDIÇÃO seguirem dizendo que o animal saiu, ele
    fica fora do headcount mesmo tendo voltado a ser nosso. Não dá pra corrigir
    daqui — o cadastro é do haras —, então isto vira aviso e tabela, todo
    fechamento, até a linha ser arrumada na origem.
    """
    src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    wb = _load(src)
    L = PLANTEL_LAYOUT_MENSAL
    linhas = []
    for i, r in enumerate(wb["PLANTEL"].iter_rows(values_only=True), start=1):
        if i < L["linha1"] or r[L["nome"]] is None:
            continue
        linhas.append({
            "nome": _s(r[L["nome"]]), "status": _s(r[L["status"]]), "local": _s(r[L["local"]]),
            "condicao": _s(r[COL_MENSAL_CONDICAO]) if len(r) > COL_MENSAL_CONDICAO else None,
            "cota": r[COL_MENSAL_COTAS] if len(r) > COL_MENSAL_COTAS else None})
    log = []
    for i, r in enumerate(wb["MOVIMENTAÇÕES"].iter_rows(values_only=True), start=1):
        if i < 3 or r[2] is None or len(r) < 5:
            continue
        d = _dt(r[3])
        log.append({"produto": _s(r[2]), "data": d.isoformat() if d else None,
                    "ocorrencia": _s(r[4])})
    wb.close()
    pend = _cancel_pendencias(linhas, log, ate=rep.semana_fim)
    if pend:
        print(f"  [cancelamento] {len(pend)} animal(is) com venda cancelada e cota de volta, "
              f"mas cadastro ainda dizendo que saiu — FORA do headcount até a origem corrigir:")
        for x in pend:
            print(f"    - {x['animal']} (cota {x['cota']}, LOCAL {x['local']}, "
                  f"STATUS {x['status']}, cancelado em {x['cancelado_em']})")
    return pend


def _refina_afeta_headcount(rep: Report):
    """Uma saída só mexe no headcount se o animal deixou de estar num LOCAL contado.

    Era decidido pela CLASSE do lançamento (SAIDA-SOCIO exceto receptora), regra de
    quando SOCIO/DOADO não contavam. Com o bucket SOCIO valendo, isso passou a
    contar como queda quem continua na conta: em 04/09/2026 a abertura saiu +0/-5
    com 8 lançamentos, enquanto a liberação diz -06 — e -06 é o certo.

    Quem manda é o retrato de agora, no controle mensal e no arquivo de receptoras:
      - animal: sai da conta se o LOCAL não é bucket de headcount OU o status não é
        de quem está aqui (QUANTICO, POTENTE, PAETE, PATRIMONIO e POTRA MORENA
        ficaram LOCAL='OUTROS' com status 'VENDIDO E ENTREGUE' -> saíram);
      - receptora: só é contada em PAO GRANDE e ARRENDAMENTO, então ir pro sócio a
        tira da conta (RECEPTORA 397 -> saiu), e ela nunca aparece no diff do
        roster, que é só do plantel;
      - OUSADO e ADRENALINA foram pro sócio mas seguem LOCAL='SOCIO'/'PLANTEL':
        trocaram de bucket, não saíram da conta.
    """
    src = _latest_no_plantel("*CONTROLE_DE_PLANTEL_PAO_GRANDE_*.xlsx", "controle mensal")
    wb = _load(src)
    L = PLANTEL_LAYOUT_MENSAL
    local_de, status_de = {}, {}
    for i, r in enumerate(wb["PLANTEL"].iter_rows(values_only=True), start=1):
        if i < L["linha1"] or r[L["nome"]] is None:
            continue
        k = _norm(_sem_cotista(_s(r[L["nome"]])))
        local_de.setdefault(k, _norm(r[L["local"]]))
        status_de.setdefault(k, _norm(r[L["status"]]))
    wb.close()
    rec_local = {}
    try:
        wb2 = _load(_latest_no_plantel("*PLANTEL ARRENDAMENTOS E RECEPTORAS.xlsx", "receptoras"))
        for i, r in enumerate(wb2["ANIMAIS"].iter_rows(values_only=True), start=1):
            if i < 4 or r[1] is None:
                continue
            rec_local.setdefault(_norm(r[1]), _norm(r[3]))
        wb2.close()
    except FileNotFoundError:
        pass

    def _fora_da_contagem(nome: str) -> bool:
        n = _norm(_sem_cotista(nome))
        if n.startswith("RECEPTORA"):
            loc = rec_local.get(n.replace("RECEPTORA ", ""), "")
            return loc not in RECEPTORAS_LOCAIS_ATIVOS
        loc, st = local_de.get(n), status_de.get(n)
        if loc is None:                      # sumiu do controle: saiu
            return True
        return loc not in HEADCOUNT_BUCKETS or not _status_conta(st)

    # Mexer no total é MUDAR DE ESTADO entre contado e não contado. Só "onde ele
    # está agora" não basta: o NASDAQ DA PAO GRANDE chegou do sócio em 01/09/2026
    # ("CHEGOU NO HARAS - ESTAVA NO SOCIO"), e ele já era headcount desde sempre —
    # estava no bucket SOCIO desde 02/09/2025. A chegada troca o bucket (SOCIO ->
    # FAZENDA PAO GRANDE), não soma +1, e também não é transferência interna, que é
    # só FPG <-> ARRENDAMENTO. Contando como entrada, a abertura saía +1/-6 numa
    # semana que é +0/-6.
    roster_ant = set()
    rec_ant = _mapa_receptoras_anterior(rep.semana_atual)
    hist = _load_hist()
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("roster"):
            roster_ant = {_norm(x) for x in _roster_congelado(rep.semana_atual)}

    def _era_contado(nome: str) -> bool:
        n = _norm(_sem_cotista(nome))
        if n.startswith("RECEPTORA"):
            return rec_ant.get(n.replace("RECEPTORA ", ""), "") in RECEPTORAS_LOCAIS_ATIVOS
        return n in roster_ant

    for e in rep.detalhe.get("saidas_diff") or []:
        nome = str(e.get("animal") or "")
        e["afeta_headcount"] = _era_contado(nome) and _fora_da_contagem(nome)
    for e in rep.detalhe.get("entradas_diff") or []:
        nome = str(e.get("animal") or "")
        e["afeta_headcount"] = (not _era_contado(nome)) and not _fora_da_contagem(nome)
        if not e["afeta_headcount"] and not _fora_da_contagem(nome):
            e["trocou_de_bucket"] = True
    rep.saidas["saidas_no_headcount"] = sum(
        1 for x in (rep.detalhe.get("saidas_diff") or []) if x.get("afeta_headcount"))
    rep.saidas["entradas_no_headcount"] = sum(
        1 for x in (rep.detalhe.get("entradas_diff") or []) if x.get("afeta_headcount"))


# Uma frase só para as duas formas de saída vista pelo roster. O QUE aconteceu já vai
# em `classificacao` (SAIDA-SOCIO, SAIDA-VENDIDO E ENTREGUE) e em `de`/`para`; a OBS
# diz apenas o que falta na origem, sem repetir a causa em texto corrido.
OBS_SEM_LANCAMENTO = "não consta o movimento da aba SAIDAS-ENTRADAS"


def _saidas_por_mudanca_de_local(rep: Report, ja_lancadas: list) -> list:
    """Saída que o ROSTER mostra e a aba SAIDAS-ENTRADAS não registrou.

    Compara o LOCAL de cada animal com o do arquivo de linhas da semana anterior.
    Sair da propriedade para SOCIO é saída; o resto não entra aqui — transferência
    entre FAZENDA e ARRENDAMENTO é interna, e doação para MATO GROSSO não é saída
    (31/08/2026: 14 doações, ninguém mudou de lugar).

    Só complementa: quem já tem lançamento na aba fica com o lançamento, para o mesmo
    animal não ser contado duas vezes."""
    ant = _arquivo_anterior(rep.semana_atual).get("roster") or []
    if not ant:
        return []
    antes = {_norm(l.get("nome")): _norm(l.get("local")) for l in ant}
    lancados = {_norm(e.get("animal")) for e in ja_lancadas}
    novas = []
    for linha in _LINHAS_BRUTAS.get("roster") or []:
        nome, local = _norm(linha.get("nome")), _norm(linha.get("local"))
        origem = antes.get(nome)
        if not origem or nome in lancados:
            continue
        if origem in LOCAIS_NA_PROPRIEDADE and local == "SOCIO":
            novas.append({
                "animal": _s(linha.get("nome")), "classificacao": "SAIDA-SOCIO",
                "de": _s(origem), "para": _s(local), "fonte": "roster",
                # continua na contagem: mudou de bucket, não saiu do headcount
                "afeta_headcount": False,
                "obs": OBS_SEM_LANCAMENTO,
            })

    # SAIU DO ROSTER por entrega/baixa: some da contagem sem passar pela aba. Em
    # 18/09/2026 a MELISSA DA PAO GRANDE virou 'VENDIDO E ENTREGUE' e o headcount caiu
    # de 186 para 185 com o movimento dizendo +0/-0 — o Δ não fechava e a saída não
    # aparecia em lugar nenhum. Esta, ao contrário da que vai pro sócio, SAI da
    # contagem (afeta_headcount=True).
    agora = {_norm(l.get("nome")) for l in (_LINHAS_BRUTAS.get("roster") or [])}
    descartadas = {_norm(l.get("nome")): l
                   for l in (_LINHAS_BRUTAS.get("roster_descartadas") or [])}
    for nome in sorted(set(antes) - agora - lancados):
        linha = descartadas.get(nome)
        if not linha or linha.get("motivo") != "status fora do plantel":
            continue      # sumiço sem causa na planilha continua virando aviso
        # Quem já estava FORA da propriedade não saiu nesta semana: a saída física
        # aconteceu antes e o que mudou agora foi só o status. A MELISSA DA PAO GRANDE
        # estava no sócio desde antes de as saídas passarem a ser registradas assim, e
        # virou 'VENDIDO E ENTREGUE' em 18/09/2026 — o relatório do haras faz a mesma
        # distinção: Δ '-01' (ela sai da CONTAGEM) e 'Saídas na semana: 02', sem ela.
        na_propriedade = _norm(antes.get(nome)) in LOCAIS_NA_PROPRIEDADE
        novas.append({
            "animal": _s(linha.get("nome")),
            "classificacao": f'SAIDA-{_norm(linha.get("status")) or "BAIXA"}',
            "de": _s(antes.get(nome)), "para": _s(linha.get("local")), "fonte": "roster",
            # sai da contagem nos dois casos — é o Δ que muda
            "afeta_headcount": True,
            "saida_da_semana": na_propriedade,
            "obs": OBS_SEM_LANCAMENTO if na_propriedade else
                   "baixa de status; a saída física é anterior a esta semana",
        })
    if novas:
        print(f"  [saídas] {len(novas)} saída(s) vistas só pelo roster — a aba "
              f"SAIDAS-ENTRADAS não foi preenchida para elas:")
        for n in novas:
            if n["afeta_headcount"]:
                print(f"    - {n['animal']}: {n['classificacao']} (sai da contagem)")
            else:
                print(f"    - {n['animal']}: {n['de']} -> {n['para']} "
                      f"(muda de bucket, segue contado)")
    return novas


def _compute_movimento(rep: Report):
    """Saídas/entradas na semana.

    Ordem de preferência:
      1. aba SAIDAS-ENTRADAS do controle mensal, classificada — entrada = nascimento
         ou compra, saída = venda ou morte. É a fonte oficial quando o haras preencher.
      2. diff da POPULAÇÃO CONTADA vs o snapshot anterior.
      3. bootstrap do relatório oficial em Word, na primeira captura.

    O diff era só do roster do plantel e por isso vivia dando 0: o plantel (aba
    PLANTEL) não mexe quando a movimentação é de receptora. Em 07/08/2026 a única
    saída da semana foi a receptora 309 indo pro sócio — diff de roster: 0; diff da
    população contada: 1 saída, 0 entradas, que é o '+00 / -01' do relatório e fecha
    com o Δ do headcount (204 -> 203).
    """
    if rep.saidas_planilha:
        ent, sai = rep.saidas_planilha["ENTRADA"], rep.saidas_planilha["SAIDA"]
        # A aba SAIDAS-ENTRADAS registra a VENDA quando ela é fechada, não quando o
        # animal fisicamente sai — e o STATUS PLANTEL pode continuar 'VENDIDO
        # PENDENTE SAIDA' depois disso. Achado em 28/08/2026: INUSITADA DA PAO GRANDE
        # tinha SAIDA-VENDA lançada nesta semana E status ainda pendente — contada
        # como saída E como pendente ao mesmo tempo. O relatório oficial só conta
        # como saída quando ela DE FATO sai (03 saídas, não 04) — então quem ainda
        # está pendente sai da conta de saídas e fica só na de pendentes.
        pendentes_nomes = {_norm(p["nome"]) for p in rep.detalhe.get("terceiros_propriedade") or []}
        sai_pendente = [e for e in sai if _norm(e.get("animal")) in pendentes_nomes]
        if sai_pendente:
            print(f"  [saídas] {len(sai_pendente)} lançamento(s) de venda com STATUS "
                  f"ainda 'VENDIDO PENDENTE SAIDA' — venda fechada mas animal não "
                  f"saiu de fato, fora da conta de saídas: "
                  + "; ".join(e["animal"] for e in sai_pendente))
            sai = [e for e in sai if _norm(e.get("animal")) not in pendentes_nomes]
        # Doação NÃO é saída: em 31/08/2026 entrou um lote de 14 doações pro Mato
        # Grosso e nenhum daqueles animais mudou de lugar — trocou a titularidade e
        # zerou a cota, e 8 deles seguem com LOCAL='SOCIO', contados no headcount
        # (ver STATUS_NO_PLANTEL). Contar como saída punha 15 saídas na semana onde
        # a fonte oficial registra 7.
        # FALLBACK: saída que o ROSTER mostra e a aba não registrou. O relatório de
        # 18/09/2026 publica 'Saídas na semana: 02' — GABRIELA ELFAR e ORFEU MH2 DA
        # PAO GRANDE, os dois pro sócio — e a aba SAIDAS-ENTRADAS não tem nenhuma das
        # duas linhas; o roster tem, as duas trocaram FAZENDA PAO GRANDE por SOCIO.
        # Ficar em 0 porque ninguém preencheu a aba é perder movimentação que a fonte
        # mostra. Elas entram marcadas com a origem, e a falta do lançamento vira
        # aviso em vez de sumir.
        sai = sai + _saidas_por_mudanca_de_local(rep, sai)
        # `saida_da_semana=False` é baixa de status de quem já tinha saído: conta no Δ
        # (deixa a contagem) e fica fora da lista de saídas da semana.
        rep.saidas["saidas_semana"] = sum(1 for x in sai if x.get("saida_da_semana", True))
        rep.saidas["entradas_semana"] = len(ent)
        rep.saidas["fonte"] = "SAIDAS-ENTRADAS"
        # O Δ do headcount só pode ser conferido contra quem entra/sai da CONTAGEM:
        # saída pro sócio deixa a fazenda e continua contada (ver CLASSIF_FORA_DO_DELTA).
        rep.saidas["saidas_no_headcount"] = sum(1 for x in sai if x.get("afeta_headcount"))
        rep.saidas["entradas_no_headcount"] = sum(1 for x in ent if x.get("afeta_headcount"))
        rep.detalhe["saidas_diff"] = sai
        rep.detalhe["entradas_diff"] = ent
        _refina_afeta_headcount(rep)
        return

    rep.populacao = _populacao_contada(rep.roster, rep.receptoras_locais)
    hist = _load_hist()
    prev_pop = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("populacao"):
            prev_pop = hist[wid]["populacao"]
    if prev_pop is None:
        # Bootstrap: nenhuma semana anterior guardou a população (campo novo). Monta a
        # anterior com o roster daquele snapshot + o arquivo de receptoras anterior.
        prev_snap = None
        for wid in sorted(hist):
            if wid < rep.semana_atual and hist[wid].get("roster"):
                prev_snap = hist[wid]
        anteriores = _receptoras_arquivos()[1:]
        if prev_snap and anteriores:
            prev_pop = _populacao_contada(prev_snap["roster"],
                                          _receptoras_locais(anteriores[0]))
            print(f"  [saídas/entradas] primeira semana com população guardada: "
                  f"receptoras da semana anterior vindas de {anteriores[0].name}")

    cur = set(rep.populacao)
    if prev_pop and cur:
        rep.saidas["fonte"] = "diff_populacao"
        prev = set(prev_pop)
        saidas = sorted(prev - cur)
        entradas = sorted(cur - prev)
        rep.saidas["saidas_semana"] = len(saidas)
        rep.saidas["entradas_semana"] = len(entradas)
        # nome sozinho não diz nada — receptora é número ("309"). Anexa o que ela é,
        # de onde saiu, pra onde foi e a observação da planilha.
        info_atual = _receptoras_info()
        info_ant = _receptoras_info(_receptoras_arquivos()[1]) if len(_receptoras_arquivos()) > 1 else {}
        rep.detalhe["saidas_diff"] = [_descreve_mov(n, info_ant, info_atual) for n in saidas]
        rep.detalhe["entradas_diff"] = [_descreve_mov(n, info_ant, info_atual) for n in entradas]
    else:
        # BOOTSTRAP: 1ª captura, sem semana anterior p/ diff → semeia do relatório oficial
        dx = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("saidas", {})
        rep.saidas["saidas_semana"] = dx.get("saidas_semana")
        rep.saidas["entradas_semana"] = dx.get("entradas")
        rep.saidas["fonte"] = "docx"
        rep.saidas["_seed"] = "docx" if dx else None
    _conferir_delta(rep)


# receptora -> a LINHA do embrião na aba ESTAÇÃO, confirmada ou não. A prenhez que
# aparece só na planilha de receptoras tem linha aqui (faltou o diagnóstico, não a
# linha), e é daqui que saem doadora × garanhão, cota, sócio e IA para ela ser
# publicada no MESMO formato das outras — ver _confirmados_por_receptora.
_EMBRIAO_POR_RECEP: dict = {}

# receptora -> "50% FULANO", montado em build_producao a partir da ESTACAO DE MONTA
_SOCIO_POR_RECEP: dict = {}
# receptora -> "50% FULANO" pela coluna NOME SOCIO do roster mensal (fonte primária:
# todo potro nascido entra no roster, mesmo antes de a parição ir para a estação)
_SOCIO_ROSTER: dict = {}
# receptora -> (safra, data_ia) do embrião confirmado com parição PENDENTE, em
# qualquer safra. Montado em build_producao, lido por _paricoes_do_roster para
# saber a que estação pertence o potro que só o roster conhece.
_SAFRA_PARICAO_PENDENTE: dict = {}


def _limpa_socio(v) -> str | None:
    """'50% ELIANE ANDRADE/vendido' -> '50% ELIANE ANDRADE'.

    O sufixo '/vendido' e marca de controle da planilha, nao parte do nome."""
    t = _s(v)
    if not t:
        return None
    t = re.sub(r"\s*/\s*vendid[oa]\s*$", "", t, flags=re.IGNORECASE).strip()
    return t or None


# CONTROLE_DE_PLANTEL mensal, aba PLANTEL: 'NOME SOCIO' ('50% RENATA CAZZANI DE
# CARVALHO'). Fora do PLANTEL_LAYOUT_MENSAL de proposito — só as parições usam.
COL_MENSAL_NOME_SOCIO = 17

# Dado que NENHUMA planilha tem e que muda toda semana. Fica versionado, por
# semana, para congelar no snapshot e aparecer na auditoria como o que é: input
# humano. Não usar os overrides do dashboard para isto — eles vivem no localStorage
# de um navegador só.
MANUAL = BASE_DIR / "_cache" / "semanal_manual.json"


def _manual(semana: str) -> dict:
    """Campos manuais da semana. Semana sem entrada devolve {} — e o campo fica
    vazio no dashboard, nunca herdado da semana anterior: 'ciclando' de outra semana
    é um número errado com cara de certo."""
    if not MANUAL.exists():
        return {}
    try:
        return (json.loads(MANUAL.read_text(encoding="utf-8")) or {}).get(semana) or {}
    except Exception as exc:
        print(f"  [manual] {MANUAL.name} ilegível ({exc!r}) — campos manuais vazios")
        return {}


PARICOES_EXTRA = BASE_DIR / "_cache" / "paricoes_extra.json"
# assinatura de nome de produto no roster: "... RECEP 309", "MACHO ... RECEP 258"
RE_PRODUTO = re.compile(r"RECEP\w*\s*([A-Z0-9]+)\s*$")


# Arquivo das LINHAS das fontes, uma pasta por semana. FICA FORA do
# semanal_snapshots.json de proposito: aquele JSON e embutido no dashboard, e linha a
# linha ele pesaria centenas de KB por semana. Aqui e so historico consultavel.
FONTES_DIR = BASE_DIR / "_cache" / "fontes"
# As linhas brutas ficam AQUI, nao em rep.detalhe: o semanal_data.json e embutido
# inteiro no dashboard, e linha a linha engordaria o HTML sem servir a ninguem lendo.
_LINHAS_BRUTAS: dict = {}


def _arquivar_linhas(rep: Report):
    """Congela as linhas lidas nesta semana: sem isso, linha apagada na origem leva a
    informacao embora — foi assim que a cota das parições de 21/08/2026 se perdeu."""
    FONTES_DIR.mkdir(parents=True, exist_ok=True)
    dados = {
        "semana": rep.semana_atual,
        "fontes": dict(rep.fontes),
        "grupo_embrioes": _LINHAS_BRUTAS.get("grupo") or [],
        "receptoras": [{"animal": k, **v} for k, v in sorted(_receptoras_info().items())],
        "roster": _LINHAS_BRUTAS.get("roster") or [],
        "pendentes_saida": rep.detalhe.get("pendentes_saida") or [],
        "terceiros_de_terceiro": rep.detalhe.get("terceiros_propriedade") or [],
    }
    alvo = FONTES_DIR / f"{rep.semana_atual}.json"
    alvo.write_text(json.dumps(dados, ensure_ascii=False, indent=1), encoding="utf-8")
    n = sum(len(v) for v in dados.values() if isinstance(v, list))
    print(f"  [fontes] {n} linhas arquivadas em _cache/fontes/{alvo.name}")


def _linhas_da_semana(wid: str) -> dict:
    f = FONTES_DIR / f"{wid}.json"
    if not f.exists():
        return {}
    try:
        return json.loads(f.read_text(encoding="utf-8"))
    except Exception:
        return {}


def _arquivo_anterior(semana: str) -> dict:
    """Ultimo arquivo de linhas antes desta semana."""
    if not FONTES_DIR.exists():
        return {}
    ant = sorted(f.stem for f in FONTES_DIR.glob("*.json") if f.stem < semana)
    return _linhas_da_semana(ant[-1]) if ant else {}


ACUMULADO_PISO = BASE_DIR / "_cache" / "acumulado_piso.json"


def _acumulado_nunca_cai(rep: Report):
    """O acumulado da estacao e um agregador: soma ocorrencias ate o fim da estacao e
    nao volta atras. A formula 'vivos no grupo + parições' e derivada, e por isso fragil
    — linha apagada na origem sem parição lançada fazia o numero CAIR (61 -> 60 em
    21/08/2026). Guardamos o maior valor ja visto na safra e usamos como piso, avisando
    quando a derivacao vem abaixo dele."""
    ac = rep.producao.get("acumulado_estacao")
    if ac is None:
        return
    piso = {}
    if ACUMULADO_PISO.exists():
        try:
            piso = json.loads(ACUMULADO_PISO.read_text(encoding="utf-8"))
        except Exception:
            piso = {}
    ant = piso.get(SAFRA_ATUAL)
    if ant is not None and ac < ant:
        print(f"  [acumulado] a derivacao deu {ac}, abaixo do maior valor ja registrado "
              f"na safra ({ant}). Acumulado nao cai — publicando {ant} e mantendo a "
              f"diferenca visivel. Alguma linha saiu da planilha do grupo sem parição "
              f"nem aborto lançado.")
        rep.producao["acumulado_estacao_derivado"] = ac
        rep.producao["acumulado_estacao"] = ant
        ac = ant
    if ant is not None and ac > ant:
        # Subida de piso fica registrada: em 04/09/2026 o piso da safra 26/27 foi
        # gravado como 2 numa rodada intermediária (fonte trocando de cópia no meio
        # do fechamento) e passou a publicar 2 contra 1 da derivação e da liberação,
        # sem nada no log explicando de onde vinha.
        print(f"  [acumulado] piso da safra {SAFRA_ATUAL} subiu de {ant} para {ac}")
    if ant is None or ac > ant:
        piso[SAFRA_ATUAL] = ac
        ACUMULADO_PISO.parent.mkdir(parents=True, exist_ok=True)
        ACUMULADO_PISO.write_text(json.dumps(piso, ensure_ascii=False, indent=2),
                                  encoding="utf-8")


def _produto_do_roster(nome: str) -> str:
    """'MACHO FACEIRA MAPEJO X IMPERIO SAPECADO RECEP 258' -> 'Macho — Faceira
    Mapejo × Imperio Sapecado'. O roster grava em caixa alta com a receptora
    colada; a tabela do dashboard mostra lado a lado com os que vem da ESTACAO,
    entao o formato tem de ser o mesmo."""
    t = re.sub(r"\s+RECEP\w*\s*[A-Z0-9]+\s*$", "", _norm(nome)).strip()
    sexo = ""
    for pref, rot in (("FEMEA", "Fêmea"), ("FEMA", "Fêmea"), ("MACHO", "Macho"),
                      ("POTRA", "Fêmea"), ("POTRO", "Macho")):
        if t.startswith(pref + " "):
            sexo, t = rot, t[len(pref) + 1:]
            break
    t = t.title().replace(" X ", " × ").replace(" Da ", " da ").replace(" De ", " de ")
    return f"{sexo} — {t}" if sexo else t


def _safra_da_paricao(receptora) -> str | None:
    """A que estação pertence o potro que só o roster conhece: a safra do EMBRIÃO
    correspondente na aba ESTAÇÃO, casado pela receptora.

    Devolve None quando a ESTAÇÃO não tem embrião pendente naquela receptora — e
    None aqui quer dizer "não sei", nunca "é da safra vigente". A diferença importa
    na reavaliação do registro em disco: parição antiga cuja linha já foi lançada
    sai do mapa, e tratá-la como desconhecida é o que a impede de ser promovida
    para a safra corrente."""
    achado = _SAFRA_PARICAO_PENDENTE.get(_norm(receptora)) if receptora else None
    return achado[0] if achado else None


def _paricoes_do_roster(rep: Report):
    """Parição que o roster conhece e a aba ESTAÇÃO não.

    O roster é prova de que o potro nasceu: ele entra lá com nome de produto. A aba
    ESTAÇÃO é a fonte oficial, mas só vê o que passou pela estação de monta — em
    21/08/2026 nasceram DOIS potros e só um tinha parição lançada. O que faltava
    derrubava DUAS linhas de uma vez, porque o acumulado é 'vivos no grupo + parições':
    a planilha do grupo já tinha apagado a linha do embrião (ele pariu), mas sem
    parição correspondente o número não voltava. Resultado: nascimentos 1 em vez de 2
    E acumulado 60 em vez de 61, pela mesma causa.

    O registro é CUMULATIVO em disco. Se contasse só os novos da semana, o acumulado
    subiria nesta semana e cairia na próxima, quando o potro deixa de ser novidade no
    roster.
    """
    # NOME SOCIO do roster mensal manda: e a coluna que o relatorio copia, e cobre
    # tambem o potro cuja paricao ainda nao foi lancada na estacao (recep 258).
    for row in rep.detalhe.get("nascimentos_semana") or []:
        do_roster = _SOCIO_ROSTER.get(_norm(row.get("receptora")))
        if do_roster:
            row["socio"] = do_roster

    hist = _load_hist()
    prev = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("roster"):
            prev = _roster_congelado(rep.semana_atual)
    reg = {}
    if PARICOES_EXTRA.exists():
        try:
            bruto = json.loads(PARICOES_EXTRA.read_text(encoding="utf-8"))
        except Exception:
            bruto = {}
        # Chave normalizada com as regras de HOJE, e a entrada mais ANTIGA vence.
        # O registro é cumulativo em disco, então mexer na normalização do nome
        # fazia o MESMO potro entrar duas vezes: em 04/09/2026 o 'MACHO ...
        # (EDUARDO) 17/08/2024 RECEP 46', gravado em 28/08 na safra 25/26,
        # reapareceu como 'MACHO ... 17/08/2024 RECEP 46' depois da correção do
        # cotista no meio do nome — virou parição da safra 26/27 e somou +1 no
        # acumulado, que o piso travou em 2 contra 1 divulgado.
        for k, v in sorted(bruto.items(), key=lambda kv: kv[1].get("semana") or ""):
            reg.setdefault(_sem_cotista(k), v)
    if prev and rep.roster:
        # receptoras das parições que a ESTAÇÃO já entregou — evita contar duas vezes
        na_estacao = {_norm(e.get("receptora")) for e in rep.detalhe.get("nascimentos_semana", [])}
        for nome in sorted(set(rep.roster) - set(prev)):
            m = RE_PRODUTO.search(_norm(nome))
            if not m or m.group(1) in na_estacao or nome in reg:
                continue
            reg[nome] = {"receptora": m.group(1), "semana": rep.semana_atual,
                         "safra": _safra_da_paricao(m.group(1)) or SAFRA_ATUAL}
        # A safra de quem JÁ estava no registro também é reavaliada: o carimbo
        # antigo era sempre a safra do calendário, e o registro é cumulativo — sem
        # isto o erro de uma semana fica em disco para sempre, inflando o acumulado
        # de toda semana seguinte.
        for k, v in reg.items():
            certa = _safra_da_paricao(v.get("receptora"))
            if certa and v.get("safra") != certa:
                print(f"  [nascimentos] parição de {k} estava na safra "
                      f"{v.get('safra')} e o embrião é da {certa} — corrigido")
                v["safra"] = certa
        PARICOES_EXTRA.parent.mkdir(parents=True, exist_ok=True)
        PARICOES_EXTRA.write_text(json.dumps(reg, ensure_ascii=False, indent=2),
                                  encoding="utf-8")

    # Parição desta semana é nascimento desta semana em qualquer estação — o
    # acumulado é que é por safra. Por isso `desta` sai de `reg`, não de `da_safra`:
    # potro da safra passada que nasce agora tem de aparecer em "Nascimentos".
    desta = [k for k, v in reg.items() if v["semana"] == rep.semana_atual]

    # O acumulado da estação só se move com EMBRIÃO CONFIRMADO NOVO (regra do Arthur,
    # 17/09/2026). Parição não é confirmação: o embrião que pariu já entrou na conta
    # quando foi confirmado, e a aba ESTAÇÃO não apaga a linha. Então a parição que o
    # roster entrega só soma no caso que criou esta função — embrião que a ESTAÇÃO
    # não conhece, em 21/08/2026. Se a receptora casa com linha confirmada lá, ele já
    # está contado e somar é contar duas vezes.
    ja_na_estacao = [k for k, v in reg.items() if v.get("safra") == SAFRA_ATUAL
                     and _SAFRA_PARICAO_PENDENTE.get(_norm(v.get("receptora")))]
    if ja_na_estacao:
        print(f"  [acumulado] {len(ja_na_estacao)} parição(ões) do roster fora do "
              f"acumulado: o embrião já está confirmado na aba ESTAÇÃO e contado lá:")
        for k in ja_na_estacao:
            print(f"    - {k}")
    da_safra = {k: v for k, v in reg.items()
                if v.get("safra") == SAFRA_ATUAL and k not in set(ja_na_estacao)}
    if not da_safra and not desta:
        return
    rep.producao["acumulado_estacao"] = (rep.producao.get("acumulado_estacao") or 0) + len(da_safra)
    rep.producao["acumulado_paricoes_so_no_roster"] = len(da_safra)

    # FATIA: a cota do embrião vive na planilha do grupo e vai embora quando a linha é
    # apagada na parição. O arquivo de linhas da semana anterior ainda tem — é o que
    # permite fechar o split PG/sócio/vendido em vez de deixá-lo abaixo do acumulado.
    ant = _arquivo_anterior(rep.semana_atual).get("grupo_embrioes") or []
    por_recep = {_norm(l.get("receptora")): l for l in ant}
    split = rep.producao.get("acumulado_estacao_split") or {}
    sem_fatia = []
    for k, v in da_safra.items():
        linha = por_recep.get(_norm(v.get("receptora")))
        if linha and linha.get("fatia"):
            split[linha["fatia"]] = split.get(linha["fatia"], 0) + 1
            v["fatia"] = linha["fatia"]
            v["cota"] = linha.get("cota")
        else:
            sem_fatia.append(k)
    rep.producao["acumulado_estacao_split"] = split
    if sem_fatia:
        print(f"  [nascimentos] fatia nao recuperada para {len(sem_fatia)} parição(ões) "
              f"— o split fica abaixo do acumulado. Sem arquivo de linhas da semana "
              f"anterior nao ha de onde tirar a cota:")
        for k in sem_fatia:
            print(f"    - {k}")
    if desta:
        # `desta` são as parições registradas nesta semana, de QUALQUER safra: um
        # potro concebido na estação passada que nasce agora é nascimento desta
        # semana, mesmo sem tocar no acumulado da safra corrente. Por isso as linhas
        # abaixo leem `reg`, e não `da_safra`.
        def _socio_da_recep(rec):
            r = _norm(rec)
            if r in _SOCIO_ROSTER:
                return _SOCIO_ROSTER[r]
            if r in _SOCIO_POR_RECEP:
                return _SOCIO_POR_RECEP[r]
            linha = por_recep.get(r) or {}
            return _limpa_socio(linha.get("socio"))

        sem_socio = [k for k in desta if not _socio_da_recep(reg[k]["receptora"])]
        if sem_socio:
            print(f"  [nascimentos] sem sócio na estação nem no arquivo do grupo "
                  f"({len(sem_socio)}) — o relatório publica esse nome, aqui fica vazio:")
            for k in sem_socio:
                print(f"    - {k} (recep {reg[k]['receptora']})")
        rep.detalhe["nascimentos_so_roster"] = [
            {"produto": _produto_do_roster(k), "receptora": reg[k]["receptora"],
             "socio": _socio_da_recep(reg[k]["receptora"]),
             "origem": "roster"} for k in desta]

        # O potro desta lista deveria já estar em `nascimentos_semana`, que sai da
        # coluna NASCIMENTO do roster. Quando ela vem VAZIA, não está: em 17/09/2026
        # as duas parições (recep 453 e 440) entraram no roster sem data, somaram no
        # acumulado da safra e o card "Nascimentos na semana" publicou 0. Parição
        # contada de um lado e invisível do outro é o pior dos dois mundos — entra
        # aqui, com a origem marcada, em vez de sumir calada.
        ja_publicadas = {_norm(n.get("receptora"))
                         for n in (rep.detalhe.get("nascimentos_semana") or [])}
        por_nome = {l.get("nome"): l for l in (_LINHAS_BRUTAS.get("roster") or [])}
        sem_data = [k for k in desta if _norm(reg[k]["receptora"]) not in ja_publicadas]
        if sem_data:
            print(f"  [nascimentos] {len(sem_data)} parição(ões) sem data na coluna "
                  f"NASCIMENTO do roster — contadas assim mesmo (a linha existe, a "
                  f"data não); pedir a data ao haras:")
            lista = rep.detalhe.setdefault("nascimentos_semana", [])
            for k in sem_data:
                linha = por_nome.get(k) or {}
                print(f"    - {k}")
                lista.append({
                    "produto": k,
                    "mae": linha.get("mae"), "pai": linha.get("pai"),
                    "receptora": reg[k]["receptora"],
                    "socio": _socio_da_recep(reg[k]["receptora"]),
                    "data": None, "local": linha.get("local"),
                    "origem": "roster (sem data)",
                })
            rep.producao["nascimentos"] = len(lista)

        # Nascimento NÃO é entrada — entrada e saída no Δ do headcount são
        # FÍSICAS: animal que chega de fora ou que sai da fazenda de verdade.
        # Potro nascido aqui não "entra" de lugar nenhum, já está na conta.
        # Errado antes: somava a `entradas_no_headcount`, dando '+02' quando o
        # relatório de 28/08/2026 diz '+00' — confirmado direto: 2 saídas físicas
        # na semana (GIM MATIZA + PODIO), nenhuma entrada. Fica só registrado
        # aqui pra quem precisar da abertura de nascimento-só-por-roster.
        rep.saidas["entradas_nascimento"] = len(desta)
    if da_safra:
        print(f"  [nascimentos] {len(da_safra)} parição(ões) da safra conhecidas só "
              f"pelo roster, somadas ao acumulado (a aba ESTAÇÃO não as tem):")
        for k in sorted(da_safra):
            marca = "  <- nesta semana" if k in desta else ""
            print(f"    - {k} (recep {reg[k]['receptora']}){marca}")


def _registra_caminhos(rep: Report):
    """rep.fontes tem o NOME do arquivo; aqui vai o caminho, para a auditoria dizer
    em qual pasta clicar. Um dict à parte para não mexer no formato de rep.fontes,
    que os snapshots antigos já gravaram."""
    rep.fontes_caminhos = {r: caminho_curto(f) for r, f in _FONTES_USADAS.items()}
    _avisa_troca_de_fonte(rep)
    rep.fontes_fora_de_lugar = sorted(FONTES_FORA_DE_LUGAR)


def _aplica_manual(rep: Report):
    """Traz os campos manuais da semana para o relatório.

    Roda depois de a semana estar definida, porque o arquivo é indexado por semana —
    e de propósito NÃO cai para a semana anterior quando falta: número de outra
    semana passaria por atual sem ninguém notar."""
    m = _manual(rep.semana_atual)
    ciclando = m.get("doadoras_ciclando")
    rep.receptoras["doadoras_ciclando"] = ciclando
    if ciclando is None:
        print(f"  [manual] doadoras ciclando não preenchida para {rep.semana_atual} "
              f"— escreva em {MANUAL.name}; o card fica vazio")


def _conferir_delta(rep: Report):
    """Δ do headcount = entradas - saídas. As duas contas vêm de fontes diferentes
    (CONTAGEM vs diff da população), então uma confere a outra. Divergir significa
    movimentação que não passou pelas planilhas — tem de aparecer, não sumir."""
    ent, sai = rep.saidas.get("entradas_semana"), rep.saidas.get("saidas_semana")
    ent = rep.saidas.get("entradas_no_headcount", ent)
    sai = rep.saidas.get("saidas_no_headcount", sai)

    # O badge '+02 / -01' do dashboard tem de dizer a MESMA coisa que o relatorio, e
    # o relatorio abre o movimento de ANIMAIS: potro que nasceu entrou, animal
    # vendido saiu. Receptora nao entra nessa abertura — ela e contada a parte, e
    # receptora que vai pro socio sai da contagem sem aparecer ali.
    # A conta vem do roster, que e a lista de animais: quem entrou e quem saiu dele.
    # Usar o movimento que afeta o headcount dava '+2 / -3' em 21/08/2026, somando as
    # duas receptoras, contra o '+02 / -01' do relatorio.
    hist = _load_hist()
    prev_roster = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("roster"):
            prev_roster = _roster_congelado(rep.semana_atual)
    # Fonte do roster mudou de uma semana para a outra? O diff não vale: o mensal
    # batiza o potro e o semanal o descrevia pelo cruzamento, então TODO potro
    # apareceria como uma saída mais uma entrada. Pular é o certo — inventar
    # movimentação é pior que não ter diff nesta semana.
    prev_fonte = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("roster"):
            prev_fonte = hist[wid].get("roster_fonte")
    if prev_roster and prev_fonte != ROSTER_FONTE:
        print(f"  [roster] fonte mudou de {prev_fonte or 'controle_semanal'} para "
              f"{ROSTER_FONTE} — diff de roster PULADO nesta semana (nome de potro é "
              f"diferente nas duas planilhas). A partir da próxima o diff volta.")
        prev_roster = None
    if prev_roster and rep.roster:
        entraram = set(rep.roster) - set(prev_roster)
        sairam = set(prev_roster) - set(rep.roster)

        # O diff CRU do roster nao serve: renome aparece como uma saida mais uma
        # entrada. Em 21/08/2026 'PERSIA ING DA PAO GRANDE' virou 'PERSIA DA PAO
        # GRANDE' e o badge saiu '+3 / -2' no lugar de '+02 / -01'. Contamos por
        # CAUSA CONHECIDA — nascimento entra, saida lancada sai — e o que sobra do
        # diff vira aviso, em vez de virar numero.
        nasc = rep.producao.get("nascimentos") or 0
        lancadas = {_norm(e.get("animal")) for e in (rep.detalhe.get("saidas_diff") or [])}
        saiu_com_lancamento = {n for n in sairam if _norm(n) in lancadas}
        if rep.saidas.get("fonte") == "SAIDAS-ENTRADAS":
            # A aba oficial é a fonte da abertura quando existe: o diff do roster é
            # só do PLANTEL, então perde receptora (a 397 foi pro sócio e saiu da
            # conta) e conta como entrada apenas nascimento (o NASDAQ, ENTRADA-SOCIO,
            # ficava de fora). Deu +0/-5 numa semana de +1/-6.
            # NASCIMENTO ENTRA NO Δ. O '+' do relatório é a abertura de ANIMAIS, e o
            # potro é animal novo na conta: 17/07 (+01, 1 nascimento), 24/07 (+01, 1),
            # 21/08 (+02, 2) e 18/09/2026 (+02, 2) — quatro semanas, mesma regra.
            # Havia aqui o oposto, tirado de 28/08/2026, onde o relatório publicou
            # '+00' — mas naquela semana ele também publicou nascimentos '--', isto é,
            # não reconheceu as duas parições que só o roster tinha. A exceção era o
            # relatório omitindo o nascimento, não o nascimento ficando fora do Δ.
            # Isso NÃO faz do nascimento uma entrada física: ele continua fora de
            # `entradas_semana` (o animal não chegou de lugar nenhum), só aparece na
            # abertura do Δ, que é o que o haras publica.
            nasc = rep.producao.get("nascimentos") or 0
            rep.headcount["delta_entradas"] = (rep.saidas.get("entradas_no_headcount") or 0) + nasc
            rep.headcount["delta_saidas"] = rep.saidas.get("saidas_no_headcount") or 0
        else:
            rep.headcount["delta_entradas"] = nasc
            rep.headcount["delta_saidas"] = len(saiu_com_lancamento)

        # Quem entrou no roster por ter nascido: produto pelo nome, ou nome batizado
        # que está na lista de nascimentos da semana. É o que explica a variação do
        # total sem ser entrada física (ver o aviso do Δ, no fim desta função).
        nomes_nasc = {_norm(n.get("produto"))
                      for n in (rep.detalhe.get("nascimentos_semana") or [])}
        rep.headcount["delta_nascimentos"] = len(
            [n for n in entraram if RE_PRODUTO.search(_norm(n)) or _norm(n) in nomes_nasc])

        sem_causa_saida = sorted(sairam - saiu_com_lancamento)
        sem_causa_entrada = sorted(
            n for n in entraram if not RE_PRODUTO.search(_norm(n)))
        if sem_causa_saida or sem_causa_entrada:
            print(f"  [roster] movimento sem causa lançada — provavel renome, "
                  f"conferir: saiu {sem_causa_saida or '—'}; entrou "
                  f"{sem_causa_entrada or '—'}")
        rep.detalhe["animais_entraram"] = sorted(entraram)
        rep.detalhe["animais_sairam"] = sorted(sairam)
    else:
        rep.headcount["delta_entradas"] = ent
        rep.headcount["delta_saidas"] = sai
    delta = rep.headcount.get("delta")
    if None in (ent, sai) or delta is None:
        return
    # Nascimento move o TOTAL sem ser entrada. O potro já era do plantel como embrião
    # confirmado (regra do Arthur, 18/09/2026): ele não chega de fora, só deixa de ser
    # embrião e passa a contar como animal. Por isso ele não entra em `entradas` — mas
    # a variação do total tem de ser lida com ele dentro, senão toda semana com
    # parição acusa movimentação fantasma. Em 17/09/2026 o aviso saiu com +2 de
    # headcount contra +0/-0 de movimento, e as duas contas estavam certas: os dois
    # potros das recep 453 e 440.
    nascidos = (rep.headcount.get("delta_nascimentos")
                or rep.saidas.get("entradas_nascimento") or 0)
    if ent + nascidos - sai != delta:
        extra = f" + {nascidos} nascimento(s)" if nascidos else ""
        print(f"  [Δ] headcount variou {delta:+d} mas o diff da população dá "
              f"+{ent}/-{sai}{extra} (líquido {ent + nascidos - sai:+d}) — conferir: "
              f"uma das duas fontes não registrou alguma movimentação")


def _chave_estavel(k: str) -> str:
    """A chave do confirmado (doadora|garanhão|receptora|data_ia) SEM a receptora.

    A receptora é o campo que o haras preenche depois: em 17/09/2026 o JAVA DA PAO
    GRANDE x XODÓ PORTEIRA AZUL, confirmado em 03/09 e já contado naquela semana,
    passou de '...|None|2026-08-10' para '...|7|2026-08-10' só porque digitaram o
    número da receptora — e reapareceu como confirmação nova duas semanas depois da
    real, levando 'Confirmados semana' a 1 numa semana sem confirmação nenhuma.
    Mesma doença do nome do cotista em PARICOES_EXTRA: identidade que muda quando a
    planilha é completada não serve de chave.

    Comparar sem a receptora NÃO perde o embrião gêmeo (mesma doadora × garanhão ×
    IA em duas receptoras): a comparação é por CONTAGEM, não por presença — ver
    _novos_confirmados."""
    partes = k.split("|")
    return "|".join(partes[:2] + partes[3:]) if len(partes) == 4 else k


def _novos_confirmados(cur: dict, prev_keys) -> list:
    """Confirmados de hoje que o snapshot anterior não tinha, por contagem de chave
    estável. Só o EXCESSO sobre a semana passada é confirmação nova — preencher um
    campo da linha antiga não cria excesso, lançar um embrião de verdade cria."""
    antes = Counter(_chave_estavel(k) for k in prev_keys)
    novos = []
    for k, e in cur.items():
        ke = _chave_estavel(k)
        if antes.get(ke):
            antes[ke] -= 1
        else:
            novos.append(e)
    return novos


CONFIRMADOS_EXTRA = BASE_DIR / "_cache" / "confirmados_extra.json"


def _chave_recep(nome) -> str:
    """Número da receptora como as duas planilhas o escrevem.

    A de receptoras batiza ('511 (ERA 207 SABE)', '07 ALAZA', '526 (CASTANHA
    ANDRADE)') e a ESTAÇÃO guarda só o número ('511', '7', '526'). Sem isto a mesma
    prenhez entra como duas — uma pela receptora, outra pela estação — e o acumulado
    conta em dobro assim que o haras lança o diagnóstico."""
    token = _norm(nome).split()[0] if _norm(nome) else ""
    return token.lstrip("0") or token


def _confirmados_por_receptora(rep: Report) -> list:
    """Confirmação que a PLANILHA DE RECEPTORAS mostra e a ESTAÇÃO ainda não tem.

    Mesmo padrão do fallback de saídas: o haras faz o diagnóstico, marca a receptora
    como PRENHA e só depois lança a série 15D/30D/45D/60D na estação de monta. Em
    18/09/2026 o quadro da sala listava 6 confirmados, as linhas existiam na ESTAÇÃO
    (recep 7, 526, 77, 362, 511, 517) e só a recep 7 tinha diagnóstico — as outras 5
    apareciam apenas como PRENHA na planilha de receptoras, e o card publicava 0.

    Receptora que passa a PRENHA no diff da semana conta como confirmação, marcada
    com a origem para a falta do lançamento continuar visível."""
    reg = {}
    if CONFIRMADOS_EXTRA.exists():
        try:
            reg = json.loads(CONFIRMADOS_EXTRA.read_text(encoding="utf-8"))
        except Exception:
            reg = {}

    ant = _arquivo_anterior(rep.semana_atual).get("receptoras") or []
    if ant:
        antes = {_norm(l.get("animal")): _norm(l.get("status")) for l in ant}
        for animal, info in sorted(_receptoras_info().items()):
            st_antes = antes.get(animal)
            if st_antes is None or _chave_recep(animal) in reg:
                continue
            if _norm(info.get("status")).startswith("PRENHA") and not st_antes.startswith("PRENHA"):
                reg[_chave_recep(animal)] = {"semana": rep.semana_atual,
                                             "safra": SAFRA_ATUAL}
        CONFIRMADOS_EXTRA.parent.mkdir(parents=True, exist_ok=True)
        CONFIRMADOS_EXTRA.write_text(json.dumps(reg, ensure_ascii=False, indent=2),
                                     encoding="utf-8")

    # O registro é CUMULATIVO em disco, pelo mesmo motivo de PARICOES_EXTRA: a prenhez
    # só é "nova" na semana em que aparece, e o acumulado da safra não pode cair na
    # semana seguinte por ela ter deixado de ser novidade.
    na_estacao = {_chave_recep(e.get("receptora")) for e in rep.confirmed}
    out, sem_linha = [], []
    for k, v in sorted(reg.items()):
        if v.get("safra") != SAFRA_ATUAL or k in na_estacao:
            continue
        linha = _EMBRIAO_POR_RECEP.get(k)
        if not linha:
            # Sem linha na ESTAÇÃO não há doadora nem garanhão para publicar. Vai com
            # a receptora e um aviso, em vez de entrar mudo no card.
            sem_linha.append(k)
            out.append({"doadora": None, "garanhao": None, "receptora": k,
                        "semana": v["semana"]})
            continue
        # MESMOS campos das outras confirmações — doadora × garanhão × receptora é o
        # que o haras publica e o que a tabela do dashboard já sabe renderizar. A
        # marca de origem fica no log e no registro em disco, não vira coluna nova.
        out.append(dict(linha, confirmado=True, semana=v["semana"]))
    desta = [c for c in out if c["semana"] == rep.semana_atual]
    if desta:
        print(f"  [confirmados] {len(desta)} confirmação(ões) vistas só pela planilha "
              f"de receptoras — a estação de monta não tem o 60D lançado:")
        for n in desta:
            print(f"    - {n.get('doadora') or '?'} x {n.get('garanhao') or '?'} "
                  f"(recep {n.get('receptora')})")
    if sem_linha:
        print(f"  [confirmados] {len(sem_linha)} receptora(s) prenhas sem linha na aba "
              f"ESTAÇÃO — sai sem doadora/garanhão: " + ", ".join(sem_linha))
    ja_na_estacao = [k for k in reg if k in na_estacao and reg[k].get("safra") == SAFRA_ATUAL]
    if ja_na_estacao:
        print(f"  [confirmados] {len(ja_na_estacao)} já lançada(s) na estação de monta, "
              f"contadas por lá: " + ", ".join(sorted(ja_na_estacao)))
    return sorted(out, key=lambda v: (v["semana"], str(v.get("receptora"))))


# Colunas que saem da TABELA de confirmados (só dela — em pendentes, por exemplo,
# `categoria` distingue potro de doadora e fica). A seção já diz que é embrião, da
# safra corrente e da semana do fechamento: repetir isso em três colunas empurrava
# doadora e garanhão, que é o que o haras publica, para o fim da linha.
CONFIRMADO_COLS_FORA = ("categoria", "safra", "semana")


def _confirmado_publicavel(e: dict) -> dict:
    return {k: v for k, v in e.items() if k not in CONFIRMADO_COLS_FORA}


def _compute_confirmados_diff(rep: Report):
    """Confirmados na semana = embriões que viraram +/-=OK vs o snapshot anterior
    (novos no conjunto de confirmados). Forward: precisa de 2 semanas capturadas."""
    hist = _load_hist()
    prev_keys = None
    for wid in sorted(hist):
        if wid < rep.semana_atual and hist[wid].get("confirmed_keys") is not None:
            prev_keys = hist[wid]["confirmed_keys"]
    cur = {e["key"]: e for e in rep.confirmed}
    if prev_keys is not None:
        candidatos = _novos_confirmados(cur, prev_keys)
        # Uma cobrição confirmada não pode ter IA no futuro — confirmação é IA+60d.
        # Achado em 28/08/2026: FACEIRA MAPEJO x IMPERIO SAPECADO só existe na cópia
        # do master na pasta da safra NOVA (a antiga nunca teve a linha), com IA
        # 26/09/2026 e parição JÁ LANÇADA em 15/08/2026 — nasceu antes de cobrir.
        # É erro de digitação na planilha do haras (ano da cobrição), não confirmação
        # nova; contar isso como "confirmado esta semana" é publicar lixo de dado.
        hoje = date.today()
        novos, suspeitos = [], []
        for e in candidatos:
            ia = date.fromisoformat(e["data_ia"]) if e.get("data_ia") else None
            if ia and ia > hoje:
                suspeitos.append(e)
            else:
                novos.append(e)
        if suspeitos:
            print(f"  [confirmados] {len(suspeitos)} confirmação(ões) com IA no futuro, "
                  f"fora da contagem (provável erro de digitação na fonte): " +
                  "; ".join(f"{e['doadora']} x {e['garanhao']} (IA {e['data_ia']})"
                            for e in suspeitos))
        # Confirmação que só a planilha de receptoras tem entra aqui, sem duplicar o
        # que a ESTAÇÃO já entregou (casa pela receptora).
        por_recep = _confirmados_por_receptora(rep)
        desta_semana = [c for c in por_recep if c["semana"] == rep.semana_atual]
        rep.producao["confirmados_semana"] = len(novos) + len(desta_semana)
        rep.detalhe["confirmados_semana"] = [_confirmado_publicavel(e)
                                             for e in novos + desta_semana]
        # ACUMULADO DA ESTAÇÃO conta confirmado, venha de onde vier (regra do Arthur,
        # 18/09/2026: "5 na semana, 5 no mês e 6 confirmados na estação"). O 6º é a
        # recep 7, que já está na aba ESTAÇÃO; as outras 5 só existem como PRENHA na
        # planilha de receptoras e somam aqui até o 60D ser lançado — quando for, elas
        # saem deste bloco e entram pela estação, sem trocar o total.
        rep.detalhe["confirmados_por_receptora"] = por_recep
        if por_recep:
            rep.producao["acumulado_estacao"] = (
                rep.producao.get("acumulado_estacao") or 0) + len(por_recep)
            rep.producao["acumulado_estacao_por_receptora"] = len(por_recep)
    else:
        # BOOTSTRAP: 1ª captura → semeia do relatório oficial
        dx = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("producao", {})
        rep.producao["confirmados_semana"] = dx.get("confirmados_semana")
        rep.detalhe["confirmados_semana"] = []

    # ACUMULADO NO MÊS = novos confirmados desde o último snapshot ANTES do mês de
    # referência (docx "--" = 0; acumulado parado no mês → 0). Não usa IA+60 (proxy ruim).
    #
    # Mês de referência = mês em que a JANELA COMEÇA, não o do fechamento. Regra do
    # Arthur em 04/09/2026: "como essa semana ainda pegou parte de agosto, estamos
    # considerando o final mês de agosto" — e a liberação daquela semana escreve
    # "Acumulado no mês (agosto): 01". Com o mês do fechamento (setembro) dava 0
    # contra 1 divulgado, na semana que atravessa a virada do mês.
    mes_ref = rep.semana_inicio[:7]                    # YYYY-MM da abertura da janela
    rep.producao["mes_referencia"] = mes_ref
    rep.producao["mes_referencia_rotulo"] = MESES_PT[int(mes_ref[5:7]) - 1]
    month_start = mes_ref + "-01"                      # YYYY-MM-01
    prev_month_keys = None
    for wid in sorted(hist):
        if wid < month_start and hist[wid].get("confirmed_keys") is not None:
            prev_month_keys = hist[wid]["confirmed_keys"]
    if prev_month_keys is not None:
        # Mesmo filtro de IA-no-futuro do "Confirmados semana" logo acima — esquecido
        # aqui até 28/08/2026: a FACEIRA MAPEJO (IA 26/09/2026, erro de digitação)
        # ficava fora de "Confirmados semana" mas ainda inflava "Acumulado no mês"
        # em +1, porque este bloco fazia o próprio diff sem reaproveitar `novos`.
        # Mesma chave estável do diff semanal: receptora preenchida depois não é
        # confirmação nova aqui também.
        hoje = date.today()
        def _ia_no_futuro(e):
            ia_iso = e.get("data_ia")
            return bool(ia_iso and date.fromisoformat(ia_iso) > hoje)
        # As confirmações vistas pela planilha de receptoras entram no mês pela SEMANA
        # em que foram registradas — elas não têm chave na ESTAÇÃO para o diff pegar.
        do_mes = [c for c in (rep.detalhe.get("confirmados_por_receptora") or [])
                  if c.get("semana", "") >= month_start]
        rep.producao["acumulado_mes"] = sum(
            1 for e in _novos_confirmados(cur, prev_month_keys)
            if not _ia_no_futuro(e)) + len(do_mes)
    else:
        dxp = (rep.docx_ref or {}).get(rep.semana_atual, {}).get("producao", {})
        rep.producao["acumulado_mes"] = dxp.get("acumulado_mes") or 0   # "--" = 0

    # PLACEHOLDER manual: confirmação que o haras já anunciou mas ainda não lançou
    # na ESTAÇÃO (falta detalhe, o Alexandre vai passar). Em 28/08/2026 é o 1º
    # confirmado da safra 26/27 — sem isto "Confirmados semana" e "Acumulado
    # estação (safra nova)" ficam 0 até o lançamento chegar, quando o relatório
    # oficial já publica 1. Só vale para a semana em que foi escrito — semana sem
    # entrada não herda nada, mesmo padrão do doadoras_ciclando.
    placeholder = _manual(rep.semana_atual).get("confirmado_placeholder")
    if placeholder:
        rep.producao["confirmados_semana"] = (rep.producao.get("confirmados_semana") or 0) + 1
        rep.producao["acumulado_estacao_proxima"] = (
            rep.producao.get("acumulado_estacao_proxima") or 0) + 1
        # "obs" fica só no manual (nota pra quem mexe no arquivo depois) — não vai
        # pro dashboard, que é visto pelo time. A linha ali é doadora/garanhão/safra,
        # igual às outras, sem comentário solto no meio do dado.
        rep.detalhe.setdefault("confirmados_semana", []).append({
            "doadora": placeholder.get("doadora"),
            "garanhao": placeholder.get("garanhao"),
            "safra": placeholder.get("safra"),
            "placeholder": True,
        })
        print(f"  [manual] placeholder de confirmação aplicado: "
              f"{placeholder.get('doadora')} x {placeholder.get('garanhao')} "
              f"— {placeholder.get('obs', 'aguardando lançamento na ESTAÇÃO')}")


def _snap_from_rep(rep: Report) -> dict:
    """Snapshot completo desta run (mesmo schema do _map_docx_to_snap)."""
    return {
        "source": "extractor",
        "acumulado_estacao": rep.producao.get("acumulado_estacao"),
        # Mesmo esquecimento do fontes_caminhos: computado em rep.producao, nunca
        # copiado pro snapshot. É a acumulada da SAFRA NOVA — conceito diferente de
        # 'acumulado no mês' (esse é confirmação nova no mês corrente, qualquer
        # safra; aquele é confirmação nova desde que a safra nova começou).
        "acumulado_estacao_proxima": rep.producao.get("acumulado_estacao_proxima"),
        # Mesmo esquecimento, quarta vez: sem isto o card do dashboard nunca aparecia
        # — o skip é 'sem safra_proxima_rotulo, esconde', e o rótulo nunca chegava
        # no snapshot pra além de ser calculado em rep.producao.
        "safra_atual_rotulo": rep.producao.get("safra_atual_rotulo"),
        "safra_proxima_rotulo": rep.producao.get("safra_proxima_rotulo"),
        "confirmados_semana": rep.producao.get("confirmados_semana"),
        "acumulado_mes": rep.producao.get("acumulado_mes"),
        # o mês do acumulado é o da ABERTURA da janela (ver mes_ref): na semana que
        # atravessa a virada, o card tem de dizer qual mês está somando
        "mes_referencia_rotulo": rep.producao.get("mes_referencia_rotulo"),
        "nascimentos": rep.producao.get("nascimentos"),
        "abortos_obitos": rep.producao.get("abortos_obitos"),
        "acumulado_estacao_split": rep.producao.get("acumulado_estacao_split"),
        "receptoras": rep.receptoras,
        "headcount": {k: v for k, v in rep.headcount.items() if k != "detalhe"},
        # a abertura animais/receptoras por local ia fora do snapshot, então quem quisesse
        # a contagem de um mês passado só tinha a aba CONTAGEM — que é retrato AO VIVO e
        # fazia o slide de junho do comitê exibir a contagem de agosto.
        "headcount_detalhe": rep.headcount.get("detalhe"),
        "terceiros": {k: v for k, v in rep.terceiros.items() if k != "painel"},
        "movimento": {"saidas": rep.saidas.get("saidas_semana"),
                      "entradas": rep.saidas.get("entradas_semana"),
                      "transferencias": rep.saidas.get("transferencias_semana")},
        "detalhe": {
            "confirmados": rep.detalhe.get("confirmados_semana"),
            "nascimentos": rep.detalhe.get("nascimentos_semana"),
            "abortos_obitos": rep.detalhe.get("abortos_obitos_semana"),
            "saidas": rep.detalhe.get("saidas_diff"),
            "entradas": rep.detalhe.get("entradas_diff"),
            # cancelamento que o cadastro do haras não acompanhou: fica de fora do
            # headcount e não aparecia em relatório nenhum (ver _cancelamentos_pendentes)
            "cancelamentos_pendentes": rep.detalhe.get("cancelamentos_pendentes"),
            "pendentes_saida": rep.detalhe.get("pendentes_saida"),
            # Faltava esta — mesmo esquecimento do fontes_caminhos e do
            # acumulado_estacao_proxima: existe em rep.detalhe, nunca chegava no
            # dict congelado. É a lista que a seção 4 (Terceiros na propriedade /
            # vendidos pendentes) mostra; sem isto ela renderiza vazia mesmo com o
            # KPI certo ao lado, porque o KPI lê rep.terceiros e a tabela lê o snapshot.
            "terceiros_propriedade": rep.detalhe.get("terceiros_propriedade"),
            "terceiros_vendidos": rep.detalhe.get("terceiros_vendidos"),
            "terceiros_vendidos_embrioes": rep.detalhe.get("terceiros_vendidos_embrioes"),
            "terceiros_sociedade": rep.detalhe.get("terceiros_sociedade"),
            "terceiros_sociedade_embrioes": rep.detalhe.get("terceiros_sociedade_embrioes"),
            "transferencias": rep.detalhe.get("transferencias_internas"),
        },
        "roster": rep.roster,
        "roster_fonte": ROSTER_FONTE,
        "receptoras_locais": rep.receptoras_locais,
        "populacao": rep.populacao,
        "confirmed_keys": [e["key"] for e in rep.confirmed],
        # Sem isto o snapshot esquece de onde cada número saiu — achado em
        # 28/08/2026 quando a auditoria publicou zero caminho de arquivo pra
        # semana inteira: rep.fontes_caminhos existia, _registra_caminhos rodava,
        # mas _snap_from_rep nunca copiava pro dict que de fato é congelado.
        "fontes_caminhos": rep.fontes_caminhos,
        "fontes_fora_de_lugar": rep.fontes_fora_de_lugar,
    }


def _persist_snapshot(rep: Report):
    """Congela o snapshot CALCULADO desta semana. O que vai pro dash é sempre calculado
    aqui — o docx nunca vira dado, só valida (ver rep.docx_ref)."""
    HIST_SNAPSHOTS.parent.mkdir(parents=True, exist_ok=True)
    hist = _load_hist()
    hist[rep.semana_atual] = _snap_from_rep(rep)
    HIST_SNAPSHOTS.write_text(json.dumps(hist, ensure_ascii=False, indent=2), encoding="utf-8")
    rep.snapshots = hist


def _load_docx_ref() -> dict:
    """Números dos relatórios oficiais (semanal_docx.json), keyed by data de referência.
    Usado SÓ como overlay de validação no dash — nunca como dado."""
    fp = BASES_DIR / "semanal_docx.json"
    if not fp.exists():
        return {}
    out = {}
    for w in json.loads(fp.read_text(encoding="utf-8")):
        out[w["ref"]] = w
    return out


class _SemAlvo(dict):
    """Sem liberação da semana, todo alvo é None e o _cmp não imprime comparação."""
    def __getitem__(self, k):
        return None
    def get(self, k, d=None):
        return None


def _cmp(label, got, target):
    flag = "" if target is None else ("  OK" if got == target else f"  (docx: {target})")
    print(f"    {label:38} {got}{flag}")


def print_report(rep: Report):
    alvo = _alvos(rep)
    DOCX_1707 = _SemAlvo() if not alvo else alvo
    print("=" * 66)
    print(f"ATUALIZAÇÃO SEMANAL — {rep.semana_inicio} a {rep.semana_fim}")
    print("=" * 66)
    print("Comparando com: " + ("liberação do haras de "
          + rep.semana_atual if alvo else "nada — o haras ainda não liberou esta semana"))
    print("Fontes:")
    for k, v in rep.fontes.items():
        print(f"    {k}: {v}")
    print("\n1) PRODUÇÃO")
    _cmp("Acumulado na estação", rep.producao["acumulado_estacao"], DOCX_1707["acumulado_estacao"])
    print(f"        split PG/sócio/vendido: {rep.producao['acumulado_estacao_split']}")
    _cmp("Embriões confirmados na semana", rep.producao["confirmados_semana"], DOCX_1707["confirmados_semana"])
    _cmp("Acumulado no mês", rep.producao["acumulado_mes"], None)
    _cmp("Nascimentos na semana", rep.producao["nascimentos"], DOCX_1707["nascimentos"])
    _cmp("Abortos / óbitos na semana", rep.producao["abortos_obitos"], DOCX_1707["abortos_obitos"])
    print("\n2) RECEPTORAS (rebanho ativo = FPG + ARRENDAMENTO)")
    _cmp("Total receptoras", rep.receptoras["total"], DOCX_1707["receptoras_total"])
    _cmp("Prenhas", rep.receptoras["prenhas"], DOCX_1707["receptoras_prenhas"])
    _cmp("Vazias", rep.receptoras["vazias"], DOCX_1707["receptoras_vazias"])
    _cmp("Índice eficiência (vazias/doadoras)", rep.receptoras["indice_eficiencia"], None)
    print("\n3) HEADCOUNT")
    _cmp("Total geral", rep.headcount["total"], DOCX_1707["headcount_total"])
    _cmp("Fazenda Pão Grande", rep.headcount["fazenda_pg"], DOCX_1707["headcount_fpg"])
    _cmp("Arrendamento", rep.headcount["arrendamento"], DOCX_1707["headcount_arr"])
    _cmp("Centro de Treinamento", rep.headcount["cte"], DOCX_1707["headcount_cte"])
    _cmp("Sócios", rep.headcount["socio"], DOCX_1707["headcount_soc"])
    _cmp("Δ vs semana anterior", rep.headcount.get("delta"), None)
    print("\n4) TERCEIROS / COMERCIAIS")
    _cmp("Vendidos pendentes de saída", rep.terceiros.get("vendidos_pendentes"), DOCX_1707["vendidos_pendentes"])
    _cmp("Em sociedade pendentes", rep.terceiros.get("sociedade_pendentes"), DOCX_1707["sociedade_pendentes"])
    _cmp("Terceiros na propriedade", rep.terceiros.get("terceiros_propriedade"), None)
    print(f"    painel comerciais: {rep.terceiros.get('painel')}")
    print("\n5) SAÍDAS")
    _cmp("Saídas na semana", rep.saidas["saidas_semana"], DOCX_1707["saidas_semana"])
    _cmp("Entradas na semana", rep.saidas["entradas_semana"], None)
    _cmp("Transferências internas", rep.saidas["transferencias_semana"], None)


def _parse_d(s: str) -> date:
    s = s.strip()
    if re.match(r"^\d{2}/\d{2}/\d{4}$", s):
        d, m, y = s.split("/"); return date(int(y), int(m), int(d))
    return date.fromisoformat(s)


MESES_PT = ("janeiro", "fevereiro", "março", "abril", "maio", "junho", "julho",
            "agosto", "setembro", "outubro", "novembro", "dezembro")


def _inicio_da_semana(fim: date) -> date:
    """Começo da janela = dia SEGUINTE ao último fechamento.

    Era `fim - 7` (sem argumento) e `fim - 14` (com um), e aí o dia do fechamento
    anterior caía nas DUAS semanas: rodando 04/09/2026 assim, a saída da INUSITADA
    DA PAO GRANDE (28/08, exatamente o fechamento anterior) foi contada de novo e a
    semana fechou com 9 saídas contra as 8 divulgadas. O orquestrador PGSemanal.py
    já fazia certo (`prev + 1`); quem rodava este script direto pegava a janela
    torta. O último fechamento sai do snapshot congelado ou do docx — o mais
    recente antes de `fim`."""
    refs = set()
    fp = BASES_DIR / "semanal_docx.json"
    if fp.exists():
        refs |= {w["ref"] for w in json.loads(fp.read_text(encoding="utf-8")) if w.get("ref")}
    if HIST_SNAPSHOTS.exists():
        refs |= {k for k in json.loads(HIST_SNAPSHOTS.read_text(encoding="utf-8")) if _is_iso(k)}
    ant = max((r for r in refs if r < fim.isoformat()), default=None)
    return date.fromisoformat(ant) + timedelta(days=1) if ant else fim - timedelta(days=6)


def main():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass
    args = sys.argv[1:]
    if len(args) >= 2:
        ini, fim = _parse_d(args[0]), _parse_d(args[1])
    else:
        fim = _parse_d(args[0]) if args else date.today()
        ini = _inicio_da_semana(fim)
    rep = build_report(ini, fim)
    print_report(rep)
    JSON_OUT.parent.mkdir(parents=True, exist_ok=True)
    JSON_OUT.write_text(json.dumps(asdict(rep), ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n-> {JSON_OUT.name} gravado ({JSON_OUT.stat().st_size} bytes)")


if __name__ == "__main__":
    main()
