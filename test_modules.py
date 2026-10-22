# -*- coding: utf-8 -*-
"""Pruebas de los modulos auxiliares y del servidor web, con datos ficticios y sin tocar Firefox ni el ERP.

    python -m unittest discover -s tests -v
"""
import os
import sys
import tempfile
import unittest

import pandas as pd

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
sys.path.insert(0, os.path.join(ROOT, "examples"))

import make_example_config as example      # noqa: E402
import DON_ARTICULO as core                # noqa: E402
import ROBOT                               # noqa: E402
import conjuntos                           # noqa: E402
import historico                           # noqa: E402
import row_choice                          # noqa: E402
import tariff_sql                          # noqa: E402


class HistoricoTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self._old_path = historico.DB_PATH
        historico.DB_PATH = os.path.join(self._tmp.name, "historico.db")

    def tearDown(self):
        historico.DB_PATH = self._old_path
        self._tmp.cleanup()

    def registrar(self, tecleado="ab1", erp="AB1", firma="DESC|1.000000"):
        historico.registrar(tecleado, erp, "PROVEEDOR_EJEMPLO", "1000", "tarifa.xlsx", "Hoja1", 5, "DESC", 1.0, firma)

    def test_same_article_is_recognised_by_its_signature(self):
        self.registrar()
        self.assertIsNotNone(historico.mismo_articulo("AB1", "PROVEEDOR_EJEMPLO", "DESC|1.000000"))
        self.assertIsNone(historico.mismo_articulo("AB1", "PROVEEDOR_EJEMPLO", "OTRA|2.000000"))

    def test_next_free_code_skips_the_ones_already_used(self):
        self.registrar(erp="AB1")
        self.assertEqual(historico.siguiente_codigo_libre("AB1"), "AB1.1")
        self.registrar(erp="AB1.1", firma="OTRA|2.000000")
        self.assertEqual(historico.siguiente_codigo_libre("AB1"), "AB1.2")

    def test_reset_archives_the_rows_before_deleting_them(self):
        self.registrar()
        filas = historico.reiniciar("AB1", esperadas=1)
        self.assertEqual(len(filas), 1)
        self.assertEqual(historico.coincidencias("AB1"), [])
        with historico._db() as con:
            self.assertEqual(con.execute("SELECT COUNT(*) FROM reinicios").fetchone()[0], 1)

    def test_reset_refuses_when_the_history_changed_meanwhile(self):
        self.registrar()
        with self.assertRaises(historico.HistorialCambiado):
            historico.reiniciar("AB1", esperadas=2)
        self.assertEqual(len(historico.coincidencias("AB1")), 1)       # no se borro nada


class RowChoiceTests(unittest.TestCase):
    @staticmethod
    def hit(rows):
        df = pd.DataFrame(rows)
        return {"df": df, "rows": list(df.index), "sheet": "Hoja1", "provider_name": "PROVEEDOR_EJEMPLO"}

    def test_identical_rows_do_not_ask(self):
        fila = {"Referencia": "R1", "Descripcion": "PIEZA", "Precio": 5.0}
        self.assertIsNone(row_choice.row_choice_payload(self.hit([fila, dict(fila)]), "R1"))

    def test_different_rows_ask_and_show_what_changes(self):
        base = {"Referencia": "R1", "Descripcion": "PIEZA", "Precio": 5.0}
        payload = row_choice.row_choice_payload(
            self.hit([dict(base, Color="BLANCO"), dict(base, Color="NEGRO", Precio=6.0)]), "R1")
        self.assertTrue(payload["needs_choice"])
        self.assertEqual(len(payload["candidates"]), 2)
        self.assertIn("Cambia: Color: NEGRO", payload["candidates"][1]["lines"][-1])

    def test_import_signature_collapses_rows_that_import_the_same(self):
        base = {"Referencia": "R1", "Descripcion": "PIEZA", "Precio": 5.0}
        hit = self.hit([dict(base, Nota="a"), dict(base, Nota="b")])
        self.assertIsNone(row_choice.row_choice_payload(hit, "R1", signature=lambda idx, row: "misma"))

    def test_too_many_rows_is_an_error_not_a_banner(self):
        rows = [{"Referencia": "R1", "Descripcion": "PIEZA %d" % i, "Precio": 1.0} for i in range(row_choice.MAX_OPCIONES + 5)]
        payload = row_choice.row_choice_payload(self.hit(rows), "R1")
        self.assertIn("error", payload)

    def test_subtitle_follows_what_is_already_imported(self):
        options = [{"title": "A", "disabled": True}, {"title": "B"}]
        self.assertIn("«B»", row_choice.subtitle_for(options, "x"))


class ConjuntosTests(unittest.TestCase):
    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        example.build(self._tmp.name)
        self.gestor = conjuntos.GestorAvisos(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_part_of_a_set_warns(self):
        aviso = self.gestor.aviso(example.PROVEEDOR_PERFIL, "a500")
        self.assertIsNotNone(aviso)
        self.assertEqual(aviso.conjunto_unico, "A499+A500")
        self.assertIn("A499", aviso.mensaje)

    def test_other_references_and_providers_never_warn(self):
        self.assertIsNone(self.gestor.aviso(example.PROVEEDOR_PERFIL, "A100"))
        self.assertIsNone(self.gestor.aviso(example.PROVEEDOR_CERAMICA, "A500"))

    def test_missing_table_means_no_warning(self):
        vacio = conjuntos.GestorAvisos(os.path.join(self._tmp.name, "no-existe"))
        self.assertIsNone(vacio.aviso(example.PROVEEDOR_PERFIL, "A500"))


class TariffSqlTests(unittest.TestCase):
    def test_sql_copy_gives_the_same_frames_as_the_excel_and_indexes_references(self):
        with tempfile.TemporaryDirectory() as tmp:
            database = example.build(tmp)
            path = os.path.join(database, example.PROVEEDOR_PERFIL, "TARIFA_EJEMPLO_2026.xlsx")
            read_excel = lambda: core._leer_excel_hojas(path)
            hojas = tariff_sql.hojas(path, read_excel, core._candidate_ref_columns, core._norm_cell)
            excel = read_excel()
            self.assertEqual(list(hojas), list(excel))
            for nombre in excel:
                pd.testing.assert_frame_equal(hojas[nombre], excel[nombre], check_exact=True)
            hits = tariff_sql.buscar(path, "A200")
            self.assertEqual(hits[0][0], "Sheet")
            self.assertEqual(hits[0][3], [1])


class WebFlowTests(unittest.TestCase):
    """Flujo completo de process_reference con el ERP y Firefox sustituidos por un doble."""

    class FakeAutomator:
        imported = 0

        def __init__(self, url, logger=None):
            pass

        def automate_import(self):
            type(self).imported += 1

    @classmethod
    def setUpClass(cls):
        os.environ.pop("DONARTICULO_RESET_HASH", None)
        os.environ.pop("DONARTICULO_RESET_SALT", None)
        import DON_ARTICULO_WEB as web
        cls.web = web
        cls._tmp = tempfile.TemporaryDirectory()
        cls.database = example.build(cls._tmp.name)
        cls._saved = (web.PROJECT_DIR, web.AYUDA_PATH, ROBOT.DEFAULT_DATABASE_PATH, ROBOT.DEFAULT_NETWORK_PATH,
                      ROBOT.FirefoxERPAutomator, historico.DB_PATH, web._AVISOS)
        web.PROJECT_DIR = cls._tmp.name
        web.AYUDA_PATH = os.path.join(cls._tmp.name, "ayuda.xlsx")
        web._AVISOS = conjuntos.GestorAvisos(cls._tmp.name)
        ROBOT.DEFAULT_DATABASE_PATH = cls.database
        ROBOT.DEFAULT_NETWORK_PATH = os.path.join(cls._tmp.name, "import")
        ROBOT.FirefoxERPAutomator = cls.FakeAutomator
        historico.DB_PATH = os.path.join(cls._tmp.name, "historico.db")
        web.load_config()

    @classmethod
    def tearDownClass(cls):
        (cls.web.PROJECT_DIR, cls.web.AYUDA_PATH, ROBOT.DEFAULT_DATABASE_PATH, ROBOT.DEFAULT_NETWORK_PATH,
         ROBOT.FirefoxERPAutomator, historico.DB_PATH, cls.web._AVISOS) = cls._saved
        core.load_overrides(os.path.join(cls._tmp.name, "no-existe.xlsx"))
        cls._tmp.cleanup()

    def test_unknown_reference_fails_cleanly(self):
        result = self.web.process_reference("NO-EXISTE")
        self.assertFalse(result["success"])
        self.assertIn("No se encontro", result["error"])

    def test_reference_in_two_providers_asks_which_one(self):
        result = self.web.process_reference("CER001")
        self.assertTrue(result["needs_choice"])
        self.assertEqual(result["kind"], "provider")
        self.assertEqual(len(result["candidates"]), 2)

    def test_repeated_row_asks_which_row(self):
        result = self.web.process_reference("CER004")
        self.assertTrue(result["needs_choice"])
        self.assertEqual(result["kind"], "row")

    def test_part_of_a_set_asks_before_importing(self):
        result = self.web.process_reference("A500")
        self.assertEqual(result["kind"], "conjunto")
        self.assertIn("A499+A500", result["conjuntos"])

    def test_import_writes_the_file_runs_the_robot_and_records_the_history(self):
        before = self.FakeAutomator.imported
        result = self.web.process_reference("A100")
        self.assertTrue(result["success"], result.get("error"))
        self.assertEqual(result["category"], "GRIFERÍA")
        self.assertEqual(self.FakeAutomator.imported, before + 1)
        self.assertTrue(os.path.exists(os.path.join(ROBOT.DEFAULT_NETWORK_PATH, "importart.xlsx")))
        self.assertEqual(len(historico.importaciones("A100")), 1)

    def test_zero_price_blocks_the_import(self):
        import openpyxl
        path = os.path.join(self.database, example.PROVEEDOR_CERAMICA, "TARIFA_CERAMICA_2026.xlsx")
        wb = openpyxl.load_workbook(path)
        ws = wb.active
        ws.append(["CER999", "SIN PRECIO", "SERIE UNO", "BLANCO", 0, "M2", 1.0, 1, 1, 1.0, ""])
        wb.save(path)
        before = self.FakeAutomator.imported
        result = self.web.process_reference("CER999")
        self.assertTrue(result.get("zero_price_blocked"))
        self.assertEqual(self.FakeAutomator.imported, before)           # el ERP no se toca

    def test_reset_tab_checks_the_password_on_the_server(self):
        client = self.web.app.test_client()
        self.assertEqual(client.post("/reiniciar/entrar", json={"password": "mal"}).status_code, 401)
        self.assertEqual(client.post("/reiniciar/entrar", json={"password": "y"}).status_code, 200)

    def test_history_list_is_built_with_the_dom_not_with_innerhtml(self):
        with open(os.path.join(ROOT, "web_templates", "index.html"), encoding="utf-8") as fh:
            html = fh.read()
        self.assertNotIn("${h.", html)                       # los datos de la tarifa no se interpolan en HTML
        self.assertNotIn("onclick=\"input.value", html)
        self.assertNotIn("fonts.googleapis.com", html)       # sin cargar recursos de terceros

    def test_static_folder_does_not_expose_the_sources(self):
        client = self.web.app.test_client()
        self.assertEqual(client.get("/static/ROBOT.py").status_code, 404)
        self.assertEqual(client.get("/health").status_code, 200)


class DesktopFlowTests(unittest.TestCase):
    """El flujo de la ventana de escritorio (DON_ARTICULO.py), sin abrir ninguna ventana."""

    class FakeRoot:
        def __init__(self):
            self.calls = []

        def after(self, _ms, func, *args):
            self.calls.append((func.__name__, args))

    def test_run_process_writes_the_import_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            database = example.build(tmp)
            saved = (ROBOT.DEFAULT_DATABASE_PATH, ROBOT.DEFAULT_NETWORK_PATH, ROBOT.FirefoxERPAutomator)
            ROBOT.DEFAULT_DATABASE_PATH = database
            ROBOT.DEFAULT_NETWORK_PATH = os.path.join(tmp, "import")
            ROBOT.FirefoxERPAutomator = WebFlowTests.FakeAutomator
            try:
                app = core.DonArticuloApp.__new__(core.DonArticuloApp)
                app.logs = []
                app._log = app.logs.append
                app.root = self.FakeRoot()
                app.robot = ROBOT
                app.provider_table = core.load_provider_table(os.path.join(tmp, "ayuda.xlsx"))
                app.provs_map = core.load_provs_map(os.path.join(tmp, "PROVS.xlsx"))
                app.exceltodo_rules = core.load_exceltodo_rules(os.path.join(tmp, "exceltodo.xlsx"))
                app._run_process("CER002")
                self.assertEqual(app.root.calls[0][0], "_on_success", "\n".join(app.logs))
                out = pd.read_excel(os.path.join(tmp, "import", "importart.xlsx"))
                self.assertEqual(str(out.loc[0, "Codigo"]), "CER002")
                base = 8.2 * ROBOT.CERAMIC_MULT + ROBOT.CERAMIC_ADD
                self.assertAlmostEqual(out.loc[0, "Precio Base Tarifa"], base, places=4)
                self.assertAlmostEqual(out.loc[0, "Tarifa 2"], base * (1 - example.PRO_DISCOUNT / 100), places=4)
            finally:
                ROBOT.DEFAULT_DATABASE_PATH, ROBOT.DEFAULT_NETWORK_PATH, ROBOT.FirefoxERPAutomator = saved


if __name__ == "__main__":
    unittest.main()
