"""Testa a aba Naturezas do Hub HPG com o supabase-js REAL e a rede interceptada.

A aba roda em iframe do hub e pega a sessão de `window.parent.HUB`. O teste monta isso de
verdade: uma página-pai (servida por HTTP local, porque iframes de file:// não dividem
origem) cria o cliente do supabase-js, põe em `window.HUB` e embute `naturezas.html`. O que
se confere é a requisição que o supabase-js produz (o UPDATE com a guarda de versão, o
download do bucket, a consulta a gastos_editores) e o que a tela faz com a resposta.

Roda sem credencial e sem tocar no Supabase de verdade. Cada teste abre a SUA página.

Uso: python assets/naturezas/test_painel.py   (da raiz do repo do hub)
"""

import copy
import functools
import http.server
import json
import pathlib
import threading
import unittest
import urllib.parse

RAIZ = pathlib.Path(__file__).resolve().parent.parent.parent

HARNESS = """<!doctype html><meta charset="utf-8">
<script src="/assets/vendor/supabase.min.js"></script>
<script>
const q = new URLSearchParams(location.search);
window.HUB = {
  sb: supabase.createClient('https://teste.supabase.co', 'chave-anon-de-teste'),
  email: 'fulano@luxor.com.br', role: q.get('role') || 'user', dashboards: ['gastos']
};
</script>
<iframe id="aba" src="/naturezas.html" style="width:100%;height:900px;border:0"></iframe>
"""


class _Handler(http.server.SimpleHTTPRequestHandler):
    def do_GET(self):
        if self.path.split("?")[0] == "/__harness.html":
            corpo = HARNESS.encode()
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.send_header("Content-Length", str(len(corpo)))
            self.end_headers()
            self.wfile.write(corpo)
            return
        super().do_GET()

    def log_message(self, *a):
        pass


SAN = ["Custos e Despesas", "Sanidade"]
INV = ["Investimentos", "Compra"]
STAMP = {"atualizado_em": "2026-08-01T10:00:00", "atualizado_por": "alguem@luxor.com.br"}

DOC = {
    "versao": 1,
    "naturezas": [
        {"ordem": 10, "cod": "10.1", "alt": "", "nome": "Vacina", "resp": "Diogo",
         "bimer": [101], "path": SAN, **STAMP},
        {"ordem": 20, "cod": "10.2", "alt": "", "nome": "Exame", "resp": "Diogo",
         "bimer": [102], "path": SAN, **STAMP},
        {"ordem": 30, "cod": "20.1", "alt": "", "nome": "Compra de egua", "resp": "Ana",
         "bimer": [201], "path": INV, **STAMP},
    ],
    "ignoradas": [{"bimer": 900, "nome": "APORTE DE SOCIOS", **STAMP}],
}

PENDENTES = {
    "gerado_em": "2026-10-13T08:30:00",
    "pendentes": [
        {"bimer": 1880, "nome": "MATERIAL DE REPRODUÇÃO", "qtd": 3, "valor": 1181.58, "meses": [9]},
        {"bimer": 1881, "nome": "OUTRA NATUREZA NOVA", "qtd": 1, "valor": 10.0, "meses": [10]},
    ],
}

V0 = "2026-09-01T00:00:00+00:00"


class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            import playwright  # noqa: F401
        except ImportError:
            raise unittest.SkipTest("playwright não instalado")
        from playwright.sync_api import sync_playwright
        cls._srv = http.server.ThreadingHTTPServer(
            ("127.0.0.1", 0), functools.partial(_Handler, directory=str(RAIZ)))
        threading.Thread(target=cls._srv.serve_forever, daemon=True).start()
        cls._pw = sync_playwright().start()
        cls._b = cls._pw.chromium.launch()

    @classmethod
    def tearDownClass(cls):
        cls._b.close()
        cls._pw.stop()
        cls._srv.shutdown()

    def abrir(self, doc=DOC, pendentes=PENDENTES, largura=1500, editor=True, role="user"):
        self.estado = copy.deepcopy(doc)
        self.versao = V0
        self.n_versao = 0
        self.gravacoes = []
        self.erros = []
        self.page = self._b.new_page(viewport={"width": largura, "height": 950})
        self.addCleanup(self.page.close)
        self.page.on("pageerror", lambda e: self.erros.append(str(e)))
        # um só tratador de diálogo (dois derrubam o Playwright); `prompt` responde o texto
        self.resposta_dialogo = None
        self.page.on("dialog", lambda d: d.accept(self.resposta_dialogo)
                     if d.type == "prompt" else d.accept())

        def rota(route, request):
            url, metodo = request.url, request.method
            if "/storage/v1/object/" in url:
                if pendentes is None:
                    return route.fulfill(status=404, body="")
                return route.fulfill(status=200, content_type="application/json",
                                     body=json.dumps(pendentes))
            if "/rest/v1/gastos_editores" in url:
                corpo = [{"email": "fulano@luxor.com.br"}] if editor else []
                return route.fulfill(status=200, content_type="application/json",
                                     body=json.dumps(corpo))
            if "/rest/v1/gastos_cadastro" in url and metodo in ("PATCH", "POST"):
                pedida = None
                if "updated_at=eq." in url:
                    pedida = urllib.parse.unquote(url.split("updated_at=eq.")[1].split("&")[0])
                # versão vencida: nenhuma linha casa e o PostgREST devolve lista vazia
                if pedida is not None and pedida != self.versao:
                    return route.fulfill(status=200, content_type="application/json", body="[]")
                corpo = json.loads(request.post_data)
                self.gravacoes.append(corpo)
                self.estado = corpo["data"]
                # quem decide a nova versão é o trigger do banco, não o cliente
                self.n_versao += 1
                self.versao = f"2026-10-14T12:00:{self.n_versao:02d}.123456+00:00"
                return route.fulfill(status=200, content_type="application/json",
                                     body=json.dumps([{"id": "naturezas",
                                                       "updated_at": self.versao}]))
            if "/rest/v1/gastos_cadastro" in url:
                return route.fulfill(status=200,
                                     content_type="application/vnd.pgrst.object+json",
                                     body=json.dumps({"data": self.estado,
                                                      "updated_at": self.versao}))
            return route.continue_()

        self.page.route("**/*.supabase.co/**", rota)
        self.page.goto(f"http://127.0.0.1:{self._srv.server_port}/__harness.html?role={role}")
        self.page.wait_for_selector("iframe")
        # a partir daqui `self.pg` é o QUADRO da aba: tem a mesma API de uma página
        self.pg = self.page.query_selector("iframe").content_frame()
        self.pg.wait_for_selector("#natSalvar, .erro-box", state="attached", timeout=5000)
        return self.pg

    def tearDown(self):
        self.assertEqual(self.erros, [], "erro de JavaScript na página")

    # ---- atalhos ----
    def linha(self, cod):
        """Seletor da linha pelo CÓDIGO: a ordem na tela não é a do cadastro."""
        uid = self.pg.evaluate(
            "cod => { const n = NAT_CADASTRO._nats().find(x => x.cod === cod);"
            "         return n ? String(n.uid) : null; }", cod)
        self.assertTrue(uid, f"não achei a natureza {cod}")
        return f'table.cadastro tr[data-uid="{uid}"]'

    def ordem_na_tela(self):
        return self.pg.eval_on_selector_all(
            'table.cadastro tr[data-uid] input[data-campo="cod"]', "els => els.map(e => e.value)")

    def salvar(self):
        self.pg.click("#natSalvar")
        self.pg.wait_for_function(
            "() => /gravado|não foi possível/.test(document.getElementById('natStatus').textContent)")
        return self.pg.inner_text("#natStatus")

    def gravado(self):
        self.assertTrue(self.gravacoes, "nada foi gravado")
        return self.gravacoes[-1]["data"]

    def preencher_form(self, nome=None, cod=None, grupo=None, resp=None, bimer=None):
        """Preenche o formulário "Nova natureza". `grupo` é o caminho com ' › '."""
        pg = self.pg
        if nome is not None:
            pg.fill("#natFNome", nome)
        if cod is not None:
            pg.fill("#natFCod", cod)
        if grupo is not None:
            pg.select_option("#natFGrupo", value=grupo)
        if resp is not None:
            pg.fill("#natFResp", resp)
        if bimer is not None:
            pg.fill("#natFBimer", bimer)

    def adicionar(self):
        self.pg.click("#natFAdd")
        self.pg.wait_for_selector("#natForm", state="detached")


class TestAbertura(Base):
    def test_abre_com_naturezas_agrupadas(self):
        pg = self.abrir()
        txt = pg.inner_text("body")
        self.assertIn("3 naturezas no Controle", txt)
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"])
        grupos = pg.eval_on_selector_all("tr.grupo td", "els => els.map(e => e.textContent)")
        self.assertEqual(len(grupos), 2)
        self.assertIn("Custos e Despesas › Sanidade", grupos[0])
        self.assertNotIn("undefined", txt)
        self.assertTrue(pg.is_disabled("#natSalvar"), "nada alterado, nada a gravar")

    def test_documento_vazio_abre_sem_erro(self):
        """O insert inicial grava '{}': a tela tem de abrir, e não quebrar."""
        pg = self.abrir(doc={}, pendentes=None)
        self.assertIn("0 naturezas no Controle", pg.inner_text("body"))
        self.assertEqual(self.ordem_na_tela(), [])

    def test_sem_pendentes_publicados_a_tela_funciona_e_diz_por_que(self):
        pg = self.abrir(pendentes=None)
        self.assertIn("ainda não foi publicada", pg.inner_text("body"))
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"])

    def test_pendentes_mostram_o_dinheiro(self):
        pg = self.abrir()
        txt = pg.inner_text("body")
        self.assertIn("MATERIAL DE REPRODUÇÃO", txt)
        self.assertIn("R$ 1.181,58", txt)
        self.assertEqual(pg.locator("tr[data-pend]").count(), 2)

    def test_sem_rolagem_horizontal_da_pagina_no_celular(self):
        pg = self.abrir(largura=390)
        self.assertLessEqual(pg.evaluate("document.documentElement.scrollWidth"), 390)


class TestPendentes(Base):
    def test_fora_do_escopo_vai_para_ignoradas_e_grava(self):
        pg = self.abrir()
        pg.click('tr[data-pend="1881"] [data-pa="fora"]')
        self.assertEqual(pg.locator("tr[data-pend]").count(), 1, "decidida, sai da lista na hora")
        self.assertFalse(pg.is_disabled("#natSalvar"))
        self.assertIn("gravado", self.salvar())
        ign = {g["bimer"]: g for g in self.gravado()["ignoradas"]}
        self.assertEqual(set(ign), {900, 1881})
        self.assertEqual(ign[1881]["atualizado_por"], "fulano@luxor.com.br")
        self.assertEqual(ign[900]["atualizado_por"], "alguem@luxor.com.br", "a antiga não é recarimbada")

    def test_incluir_no_controle_abre_o_formulario_ja_preenchido(self):
        pg = self.abrir()
        pg.click('tr[data-pend="1880"] [data-pa="incluir"]')
        self.assertTrue(pg.is_visible("#natForm"))
        self.assertEqual(pg.input_value("#natFNome"), "MATERIAL DE REPRODUÇÃO")
        self.assertEqual(pg.input_value("#natFBimer"), "1880")
        self.assertIn("Vinda do Bimer", pg.inner_text("#natForm"))
        # faltam código, grupo e responsável: a natureza NÃO entra incompleta
        self.assertTrue(pg.is_disabled("#natFAdd"))
        self.assertEqual(pg.inner_text("#natFFalta"), "Falta preencher: código, grupo, responsável.")
        self.assertEqual(pg.locator("tr[data-pend]").count(), 2, "só some da lista ao confirmar")
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"])

    def test_incluir_completo_entra_no_fim_do_grupo_e_grava(self):
        pg = self.abrir()
        pg.click('tr[data-pend="1880"] [data-pa="incluir"]')
        self.preencher_form(cod="10.4.2.03.21", grupo=" › ".join(SAN), resp="Diogo")
        self.assertFalse(pg.is_disabled("#natFAdd"))
        self.adicionar()
        self.assertEqual(pg.locator("tr[data-pend]").count(), 1, "decidida, sai da lista")
        self.assertEqual(pg.inner_text("#natProbs"), "")
        pg.click("#natSalvar")
        pg.wait_for_function("() => /gravado/.test(document.getElementById('natStatus').textContent)")
        d = self.gravado()
        novo = next(n for n in d["naturezas"] if n["cod"] == "10.4.2.03.21")
        self.assertEqual(novo["bimer"], [1880])
        self.assertEqual(novo["path"], SAN)
        self.assertEqual(novo["nome"], "MATERIAL DE REPRODUÇÃO")
        self.assertEqual(novo["atualizado_por"], "fulano@luxor.com.br")
        # fim do grupo Sanidade, ANTES de Investimentos: o grupo continua num bloco só
        ordem = [n["cod"] for n in sorted(d["naturezas"], key=lambda n: n["ordem"])]
        self.assertEqual(ordem, ["10.1", "10.2", "10.4.2.03.21", "20.1"])
        self.assertEqual(len({n["ordem"] for n in d["naturezas"]}), 4, "ordens distintas")
        # as que ninguém tocou mantêm o carimbo antigo
        velho = next(n for n in d["naturezas"] if n["cod"] == "20.1")
        self.assertEqual(velho["atualizado_por"], "alguem@luxor.com.br")
        self.assertEqual(velho["ordem"], 30, "quem não mudou de lugar não é renumerado")

    def test_reavaliar_ignorada_devolve_a_pendencia(self):
        """Tirar de 'fora do escopo' não apaga a decisão em silêncio: o código volta a ser
        pendente, se o Bimer ainda o traz."""
        pend = copy.deepcopy(PENDENTES)
        pend["pendentes"].append({"bimer": 900, "nome": "APORTE DE SOCIOS", "qtd": 2,
                                  "valor": 500.0, "meses": [8]})
        pg = self.abrir(pendentes=pend)
        self.assertEqual(pg.locator("tr[data-pend]").count(), 2, "900 já está ignorada")
        pg.click('summary:has-text("Fora do escopo")')
        pg.click('tr[data-ign="900"] [data-ia="reavaliar"]')
        self.assertEqual(pg.locator("tr[data-pend]").count(), 3)
        self.assertFalse(pg.is_disabled("#natSalvar"))


class TestEdicao(Base):
    def test_so_o_que_mudou_e_carimbado(self):
        pg = self.abrir()
        pg.fill(f'{self.linha("10.2")} input[data-campo="nome"]', "Exame de sangue")
        self.assertIn("gravado", self.salvar())
        por_cod = {n["cod"]: n for n in self.gravado()["naturezas"]}
        self.assertEqual(por_cod["10.2"]["atualizado_por"], "fulano@luxor.com.br")
        self.assertEqual(por_cod["10.1"]["atualizado_em"], "2026-08-01T10:00:00")
        self.assertEqual(por_cod["20.1"]["atualizado_por"], "alguem@luxor.com.br")

    def test_gravacao_preserva_chaves_desconhecidas_do_documento(self):
        doc = copy.deepcopy(DOC)
        doc["outra_coisa"] = {"x": 1}
        self.abrir(doc=doc)
        self.pg.fill(f'{self.linha("10.2")} input[data-campo="resp"]', "Ana")
        self.salvar()
        self.assertEqual(self.gravado()["outra_coisa"], {"x": 1})

    def test_mover_troca_so_dentro_do_grupo(self):
        pg = self.abrir()
        # 10.1 é o primeiro do grupo: não sobe; 10.2 é o último: não desce, e NÃO atravessa
        # para o grupo de baixo (racharia o grupo em dois blocos)
        self.assertTrue(pg.is_disabled(f'{self.linha("10.1")} [data-mv="-1"]'))
        self.assertTrue(pg.is_disabled(f'{self.linha("10.2")} [data-mv="1"]'))
        self.assertTrue(pg.is_disabled(f'{self.linha("20.1")} [data-mv="-1"]'))
        pg.click(f'{self.linha("10.2")} [data-mv="-1"]')
        self.assertEqual(self.ordem_na_tela(), ["10.2", "10.1", "20.1"])
        self.salvar()
        por_cod = {n["cod"]: n for n in self.gravado()["naturezas"]}
        self.assertEqual((por_cod["10.2"]["ordem"], por_cod["10.1"]["ordem"]), (10, 20))
        self.assertEqual(por_cod["20.1"]["atualizado_por"], "alguem@luxor.com.br",
                         "mover duas não carimba a terceira")

    def test_trocar_de_grupo_leva_para_o_fim_do_destino(self):
        pg = self.abrir()
        pg.select_option(f'{self.linha("10.1")} select[data-campo="grupo"]', value=" › ".join(INV))
        self.assertEqual(self.ordem_na_tela(), ["10.2", "20.1", "10.1"])
        grupos = pg.eval_on_selector_all("tr.grupo td", "els => els.map(e => e.textContent)")
        self.assertEqual(len(grupos), 2, "cada grupo continua em UM bloco")

    def test_novo_grupo_por_prompt(self):
        pg = self.abrir()
        self.resposta_dialogo = "Custos e Despesas > Pista"
        pg.select_option(f'{self.linha("10.1")} select[data-campo="grupo"]', value="__novo__")
        self.salvar()
        por_cod = {n["cod"]: n for n in self.gravado()["naturezas"]}
        self.assertEqual(por_cod["10.1"]["path"], ["Custos e Despesas", "Pista"])

    def test_remover_natureza(self):
        pg = self.abrir()
        pg.click(f'{self.linha("10.2")} .btn-lixo')
        self.assertEqual(self.ordem_na_tela(), ["10.1", "20.1"])
        self.salvar()
        self.assertEqual([n["cod"] for n in self.gravado()["naturezas"]], ["10.1", "20.1"])

    def test_filtro_desliga_mover(self):
        pg = self.abrir()
        pg.fill("#natBusca", "vacina")
        pg.wait_for_timeout(500)
        self.assertEqual(self.ordem_na_tela(), ["10.1"])
        self.assertEqual(pg.locator(".btn-mv:not([disabled])").count(), 0)

    def test_alterar_e_desfazer_nao_acende_o_gravar(self):
        pg = self.abrir()
        campo = f'{self.linha("10.1")} input[data-campo="nome"]'
        pg.fill(campo, "outra")
        self.assertFalse(pg.is_disabled("#natSalvar"))
        pg.fill(campo, "Vacina")
        self.assertTrue(pg.is_disabled("#natSalvar"))


class TestValidacao(Base):
    def test_bimer_nao_inteiro_bloqueia_e_aponta_o_campo(self):
        pg = self.abrir()
        campo = f'{self.linha("10.1")} input[data-campo="bimer"]'
        pg.fill(campo, "101, 13a9")
        self.assertTrue(pg.is_disabled("#natSalvar"))
        self.assertIn("13a9", pg.inner_text("#natProbs"))
        self.assertIn("ruim", pg.get_attribute(campo, "class"))
        pg.fill(campo, "101, 1363")
        self.assertFalse(pg.is_disabled("#natSalvar"))
        self.assertEqual(pg.inner_text("#natProbs"), "")
        self.salvar()
        self.assertEqual({n["cod"]: n for n in self.gravado()["naturezas"]}["10.1"]["bimer"],
                         [101, 1363])

    def test_codigo_do_bimer_repetido_bloqueia(self):
        pg = self.abrir()
        pg.fill(f'{self.linha("10.2")} input[data-campo="bimer"]', "101")
        self.assertTrue(pg.is_disabled("#natSalvar"))
        self.assertIn("código do Bimer — repetido", pg.inner_text("#natProbs"))

    def test_codigo_no_controle_e_nas_ignoradas_bloqueia(self):
        pg = self.abrir()
        pg.fill(f'{self.linha("10.2")} input[data-campo="bimer"]', "900")
        self.assertTrue(pg.is_disabled("#natSalvar"))
        self.assertIn("está no Controle E em ignoradas", pg.inner_text("#natProbs"))

    def test_classificacao_repetida_bloqueia(self):
        pg = self.abrir()
        pg.fill(f'{self.linha("10.2")} input[data-campo="cod"]', "10.1")
        self.assertTrue(pg.is_disabled("#natSalvar"))
        self.assertIn("classificação — repetido", pg.inner_text("#natProbs"))


class TestFormulario(Base):
    """"Nova natureza": os campos que a chefe pediu — nome, código e grupo — numa tela
    própria, e a natureza só entra no cadastro completa e sem conflito."""

    def test_botao_abre_o_formulario_com_os_campos_pedidos(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.assertTrue(pg.is_visible("#natForm"))
        rotulos = [t.lower() for t in pg.eval_on_selector_all(
            "#natForm .form-campo label", "els => els.map(e => e.textContent)")]
        for pedido in ("nome", "código", "grupo"):
            self.assertTrue(any(r.startswith(pedido) for r in rotulos), f"campo {pedido}")
        self.assertTrue(pg.is_disabled("#natFAdd"))
        self.assertEqual(pg.inner_text("#natFFalta"),
                         "Falta preencher: nome, código, grupo, responsável.")
        self.assertEqual(pg.evaluate("document.activeElement.id"), "natFNome", "foco no 1º campo")
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"], "nada entrou ainda")

    def test_nao_ha_mais_linha_em_branco_na_tabela(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.assertEqual(pg.locator("table.cadastro tr[data-uid]").count(), 3)

    def test_grupo_lista_os_existentes_e_a_opcao_de_criar(self):
        pg = self.abrir()
        pg.click("#natAdd")
        ops = pg.eval_on_selector_all("#natFGrupo option", "els => els.map(e => e.textContent)")
        self.assertIn("Custos e Despesas › Sanidade", ops)
        self.assertIn("Investimentos › Compra", ops)
        self.assertTrue(any("novo grupo" in o for o in ops))

    def test_adicionar_manual_entra_no_fim_do_grupo_e_grava(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="Exame de DNA", cod="10.9", grupo=" › ".join(SAN),
                            resp="Diogo", bimer="777")
        self.assertFalse(pg.is_disabled("#natFAdd"))
        self.adicionar()
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "10.9", "20.1"])
        self.assertIn("Falta gravar", pg.inner_text("#natStatus"))
        self.assertEqual(pg.locator("tr.nova").count(), 1, "a linha nova fica destacada")
        self.assertIn("gravado", self.salvar())
        novo = next(n for n in self.gravado()["naturezas"] if n["cod"] == "10.9")
        self.assertEqual((novo["nome"], novo["resp"], novo["bimer"], novo["path"]),
                         ("Exame de DNA", "Diogo", [777], SAN))

    def test_bimer_e_opcional(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="Sem Bimer", cod="10.8", grupo=" › ".join(INV), resp="Ana")
        self.assertFalse(pg.is_disabled("#natFAdd"))
        self.adicionar()
        self.salvar()
        novo = next(n for n in self.gravado()["naturezas"] if n["cod"] == "10.8")
        self.assertEqual(novo["bimer"], [])

    def test_novo_grupo_vai_para_o_fim(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.assertTrue(pg.is_hidden("#natFNovoGrupo"))
        pg.select_option("#natFGrupo", value="__novo__")
        self.assertTrue(pg.is_visible("#natFNovoGrupo"))
        self.assertEqual(pg.inner_text("#natFFalta"),
                         "Falta preencher: nome, código, grupo, responsável.")
        pg.fill("#natFNovoGrupo", "Custos e Despesas > Pista")
        self.preencher_form(nome="Ferrageamento", cod="10.7", resp="Ana")
        self.adicionar()
        self.salvar()
        novo = next(n for n in self.gravado()["naturezas"] if n["cod"] == "10.7")
        self.assertEqual(novo["path"], ["Custos e Despesas", "Pista"])
        self.assertEqual(novo["ordem"], 40, "grupo novo entra no fim de tudo")

    def test_codigo_repetido_bloqueia_com_a_mensagem(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="X", cod="10.1", grupo=" › ".join(SAN), resp="Ana")
        self.assertIn("já é de «Vacina»", pg.inner_text("#natForm"))
        self.assertTrue(pg.is_disabled("#natFAdd"))
        pg.fill("#natFCod", "10.11")
        self.assertFalse(pg.is_disabled("#natFAdd"), "corrigido, destrava")

    def test_codigo_do_bimer_de_outra_natureza_bloqueia(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="X", cod="10.6", grupo=" › ".join(SAN), resp="Ana", bimer="101")
        self.assertIn("O código 101 já pertence a «Vacina»", pg.inner_text("#natForm"))
        self.assertTrue(pg.is_disabled("#natFAdd"))

    def test_codigo_do_bimer_nas_ignoradas_bloqueia_e_aponta_o_caminho(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="X", cod="10.6", grupo=" › ".join(SAN), resp="Ana", bimer="900")
        txt = pg.inner_text("#natForm")
        self.assertIn("Fora do escopo", txt)
        self.assertTrue(pg.is_disabled("#natFAdd"))

    def test_codigo_do_bimer_que_nao_e_numero_bloqueia(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="X", cod="10.6", grupo=" › ".join(SAN), resp="Ana", bimer="13a9")
        self.assertIn("Não é número: 13a9", pg.inner_text("#natForm"))
        self.assertTrue(pg.is_disabled("#natFAdd"))

    def test_cancelar_nao_muda_nada(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="Descartada", cod="10.5")
        pg.click("#natFCancel")           # o confirm "Descartar?" é aceito pelo tratador
        self.assertEqual(pg.locator("#natForm").count(), 0)
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"])
        self.assertTrue(pg.is_disabled("#natSalvar"), "nada alterado")

    def test_esc_fecha(self):
        pg = self.abrir()
        pg.click("#natAdd")
        self.page.keyboard.press("Escape")
        self.assertEqual(pg.locator("#natForm").count(), 0)

    def test_clicar_fora_nao_fecha(self):
        """Janela com campos preenchidos: um clique perdido não pode jogar o trabalho fora."""
        pg = self.abrir()
        pg.click("#natAdd")
        self.preencher_form(nome="Preenchida")
        self.page.mouse.click(5, 5)
        self.assertTrue(pg.is_visible("#natForm"))
        self.assertEqual(pg.input_value("#natFNome"), "Preenchida")

    def test_responsavel_sugere_quem_ja_existe(self):
        pg = self.abrir()
        pg.click("#natAdd")
        sug = pg.eval_on_selector_all("#natFRespLista option", "els => els.map(e => e.value)")
        self.assertEqual(sorted(sug), ["Ana", "Diogo"])

    def test_ignorada_so_sai_da_lista_ao_confirmar(self):
        pg = self.abrir()
        pg.click('summary:has-text("Fora do escopo")')
        pg.click('tr[data-ign="900"] [data-ia="incluir"]')
        self.assertEqual(pg.input_value("#natFNome"), "APORTE DE SOCIOS")
        self.assertEqual(pg.input_value("#natFBimer"), "900", "o próprio código dela não conflita")
        pg.click("#natFCancel")
        self.assertEqual(pg.locator('tr[data-ign="900"]').count(), 1, "cancelar não tira da lista")
        self.assertTrue(pg.is_disabled("#natSalvar"))
        pg.click('tr[data-ign="900"] [data-ia="incluir"]')
        self.preencher_form(cod="10.4", grupo=" › ".join(INV), resp="Ana")
        self.adicionar()
        self.assertEqual(pg.locator('tr[data-ign="900"]').count(), 0)
        self.salvar()
        d = self.gravado()
        self.assertEqual(d["ignoradas"], [])
        self.assertIn(900, next(n for n in d["naturezas"] if n["cod"] == "10.4")["bimer"])


class TestGuardas(Base):
    """As duas guardas do validador (revisão do Arthur): documento que NÃO veio da tela
    tem de ser recusado pelo JS como o Python recusa, e não passar nem estourar."""

    def problemas(self, doc):
        return self.pg.evaluate("d => NAT_DADOS.problemas(d)", doc)

    def setUp(self):
        self.abrir(doc={}, pendentes=None)

    def nat(self, **extra):
        n = {"ordem": 10, "cod": "1", "alt": "", "nome": "n", "resp": "r", "bimer": [1],
             "path": ["A", "B"]}
        n.update(extra)
        return n

    def test_documento_valido_nao_acusa(self):
        self.assertEqual(self.problemas({"naturezas": [self.nat()], "ignoradas": []}), [])

    def test_forma_do_documento(self):
        casos = {
            "sem naturezas": {"ignoradas": []},
            "naturezas é objeto": {"naturezas": {}},
            "naturezas é texto": {"naturezas": "abc"},
            "documento nulo": None,
            "ignoradas é texto": {"naturezas": [], "ignoradas": "abc"},
            "ignoradas é objeto": {"naturezas": [], "ignoradas": {}},
            "ignoradas é nulo": {"naturezas": [], "ignoradas": None},
        }
        for nome, doc in casos.items():
            with self.subTest(nome):
                self.assertTrue(self.problemas(doc), f"{nome}: tinha de recusar")

    def test_tipo_de_cada_codigo_do_bimer(self):
        casos = {
            "bimer é texto": "abc",
            "bimer tem texto": ["999"],
            "bimer tem decimal": [1.5],
            "bimer tem booleano": [True],
            "bimer tem nulo": [None],
            "bimer é nulo": None,
        }
        for nome, bimer in casos.items():
            with self.subTest(nome):
                erros = self.problemas({"naturezas": [self.nat(bimer=bimer)], "ignoradas": []})
                self.assertTrue(any("lista de inteiros" in e["erro"] for e in erros),
                                f"{nome}: {erros}")

    def test_bimer_ausente_equivale_a_lista_vazia_como_no_python(self):
        n = self.nat()
        del n["bimer"]
        self.assertEqual(self.problemas({"naturezas": [n], "ignoradas": []}), [])

    def test_documento_malformado_no_banco_abre_com_mensagem_e_nao_com_erro_de_js(self):
        pg = self.abrir(doc={"naturezas": "abc"}, pendentes=None)
        txt = pg.inner_text("body")
        self.assertIn("formato inválido", txt)
        self.assertIn("naturezas", txt)
        self.assertEqual(pg.locator("table.cadastro").count(), 0, "não abre uma tela vazia")
        # (o tearDown confere que não houve erro de JavaScript)


class TestConcorrencia(Base):
    def test_outra_pessoa_gravou_nada_e_sobrescrito(self):
        pg = self.abrir()
        pg.fill(f'{self.linha("10.1")} input[data-campo="nome"]', "meu ajuste")
        # alguém grava depois que esta tela carregou
        self.versao = "2026-09-02T00:00:00+00:00"
        status = self.salvar()
        self.assertIn("outra pessoa gravou", status)
        self.assertEqual(self.gravacoes, [], "a guarda de versão impediu a gravação")
        self.assertFalse(pg.is_disabled("#natSalvar"), "a edição não se perde: dá para tentar de novo")

    def test_segunda_gravacao_na_mesma_tela_funciona(self):
        """A versão passa a ser a devolvida pelo banco; senão a 2ª gravação falharia sozinha."""
        pg = self.abrir()
        pg.fill(f'{self.linha("10.1")} input[data-campo="nome"]', "um")
        self.assertIn("gravado", self.salvar())
        pg.fill(f'{self.linha("10.1")} input[data-campo="nome"]', "dois")
        self.assertIn("gravado", self.salvar())
        self.assertEqual(len(self.gravacoes), 2)


class TestSomenteLeitura(Base):
    """Quem não é editor consulta; o banco recusaria a gravação de qualquer jeito, e a tela
    só tira o que não funcionaria."""

    def test_nao_editor_ve_o_cadastro_mas_nao_edita(self):
        pg = self.abrir(editor=False)
        self.assertIn("Somente consulta", pg.inner_text("body"))
        self.assertEqual(self.ordem_na_tela(), ["10.1", "10.2", "20.1"], "consulta é o uso")
        campos = pg.eval_on_selector_all(
            "table.cadastro input, table.cadastro select, .btn-mv, .btn-lixo",
            "els => els.map(e => e.disabled)")
        self.assertTrue(campos and all(campos), "nenhum campo editável")
        for sel in ("#natAdd", ".cpc-acoes", "[data-pa]"):
            self.assertTrue(pg.eval_on_selector_all(sel, "els => els.every(e => e.hidden || e.closest('[hidden]'))"),
                            f"{sel} escondido")

    def test_nao_editor_ainda_filtra_e_busca(self):
        pg = self.abrir(editor=False)
        pg.fill("#natBusca", "vacina")
        pg.wait_for_timeout(500)
        self.assertEqual(self.ordem_na_tela(), ["10.1"])
        self.assertEqual(self.gravacoes, [])

    def test_admin_edita_mesmo_fora_de_gastos_editores(self):
        pg = self.abrir(editor=False, role="admin")
        self.assertNotIn("Somente consulta", pg.inner_text("body"))
        pg.fill(f'{self.linha("10.1")} input[data-campo="nome"]', "ajuste do admin")
        self.assertIn("gravado", self.salvar())

    def test_editor_nao_ve_aviso_de_consulta(self):
        pg = self.abrir(editor=True)
        self.assertNotIn("Somente consulta", pg.inner_text("body"))


class TestAbaNoHub(unittest.TestCase):
    """A rota segue a permissão de 'gastos': não tem permissão própria, e por isso não
    precisa de linha nova em user_dashboard_access nem de valor novo no CHECK."""

    @classmethod
    def setUpClass(cls):
        try:
            import playwright  # noqa: F401
        except ImportError:
            raise unittest.SkipTest("playwright não instalado")
        from playwright.sync_api import sync_playwright
        cls._pw = sync_playwright().start()
        cls._b = cls._pw.chromium.launch()
        cls.pg = cls._b.new_page()
        cls.pg.goto((RAIZ / "index.html").as_uri())   # file://: o hub sobe offline, sem login
        cls.pg.wait_for_function("typeof allowed === 'function'")

    @classmethod
    def tearDownClass(cls):
        cls._b.close()
        cls._pw.stop()

    def rotas(self, dashboards):
        return self.pg.evaluate(
            "d => { window.HUB.dashboards = d; return allowed().map(r => r.id); }", dashboards)

    def test_quem_ve_gastos_ve_naturezas(self):
        self.assertIn("naturezas", self.rotas(["gastos"]))

    def test_quem_nao_ve_gastos_nao_ve_naturezas(self):
        self.assertNotIn("naturezas", self.rotas(["semanal", "comite", "plantel"]))

    def test_demais_abas_continuam_pela_propria_permissao(self):
        self.assertEqual(self.rotas(["semanal"]), ["", "semanal"])
        self.assertEqual(self.rotas(["gastos"]), ["", "gastos", "naturezas"])

    def test_a_aba_embute_a_pagina_do_cadastro(self):
        self.pg.evaluate("window.HUB.dashboards = ['gastos']; buildNav(); location.hash = '#/naturezas'")
        self.pg.wait_for_selector("iframe.embed")
        self.assertEqual(self.pg.get_attribute("iframe.embed", "src"), "naturezas.html")


class TestRegrasPuras(Base):
    """`hn_dados.js` sem rede: as contas de ordem, que são as que racham grupo se erradas."""

    def setUp(self):
        self.abrir(doc={}, pendentes=None)

    def ev(self, js):
        return self.pg.evaluate(js)

    def test_ordem_apos_cabe_no_vao(self):
        self.assertEqual(self.ev("NAT_DADOS.ordemApos([{ordem:10},{ordem:20}], 10)"), 15)
        self.assertEqual(self.ev("NAT_DADOS.ordemApos([{ordem:10},{ordem:11}], 10)"), None)
        self.assertEqual(self.ev("NAT_DADOS.ordemApos([{ordem:10}], 10)"), 20)

    def test_fim_do_grupo_renumera_quando_nao_cabe(self):
        js = """() => {
          const g = ['G'], h = ['H'];
          const nats = [{ordem:10,path:g},{ordem:11,path:h},{ordem:12,path:h}];
          const novo = {ordem:99, path:[]};
          nats.push(novo);
          const o = NAT_DADOS.ordemNoFimDoGrupo(nats, g, novo);
          novo.ordem = o; novo.path = g;
          return NAT_DADOS.porOrdem(nats).map(n => n.path[0]);
        }"""
        self.assertEqual(self.ev(js), ["G", "G", "H", "H"])

    def test_ler_bimer_nao_descarta_lixo(self):
        r = self.ev("NAT_DADOS.lerBimer('1339, 13a9;1363  x')")
        self.assertEqual(r["ints"], [1339, 1363])
        self.assertEqual(r["invalidos"], ["13a9", "x"])


if __name__ == "__main__":
    unittest.main(verbosity=2)
