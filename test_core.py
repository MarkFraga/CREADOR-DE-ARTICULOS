# -*- coding: utf-8 -*-
"""Pruebas del motor (busqueda, conversion al formato del ERP, clasificacion) con datos ficticios.

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
import ref_index                           # noqa: E402


class NormalizationTests(unittest.TestCase):
    def test_numeric_codes_read_as_float_match_the_typed_reference(self):
        self.assertEqual(core._norm_cell(123456789012.0), "123456789012")

    def test_text_is_upper_cased_and_whitespace_collapsed(self):
        self.assertEqual(core._norm_cell("  ab \n c "), "AB C")

    def test_nan_and_none_are_empty(self):
        self.assertEqual(core._norm_cell(float("nan")), "")
        self.assertEqual(core._norm_cell(None), "")

    def test_provider_names_are_normalized_without_spaces(self):
        self.assertEqual(core._normalize_provider(" proveedor ejemplo "), "PROVEEDOREJEMPLO")


class TextCleaningTests(unittest.TestCase):
    def test_typographic_quotes_and_dashes_are_normalized(self):
        text = "\u201cHola\u201d \u2018x\u2019 \u2013 y \u2014 z\u2026"
        self.assertEqual(ROBOT.clean_text_for_erp(text), "\"HOLA\" 'X' - Y - Z...")

    def test_line_breaks_become_spaces_and_text_is_upper_cased(self):
        self.assertEqual(ROBOT.clean_text_for_erp("  mezclador\r\nlavabo\tcromo "), "MEZCLADOR LAVABO CROMO")


class LocalDateTests(unittest.TestCase):
    def test_invalid_timezone_falls_back_to_the_machine_date(self):
        import datetime
        saved = ROBOT.LOCAL_TIMEZONE
        ROBOT.LOCAL_TIMEZONE = "No/Existe"
        try:
            self.assertEqual(ROBOT.get_local_date(), datetime.date.today())
        finally:
            ROBOT.LOCAL_TIMEZONE = saved

    def test_last_run_date_is_recorded_even_with_an_invalid_timezone(self):
        saved = ROBOT.LOCAL_TIMEZONE
        ROBOT.LOCAL_TIMEZONE = "No/Existe"
        try:
            with tempfile.TemporaryDirectory() as tmp:
                automator = ROBOT.FirefoxERPAutomator("http://erp.empresa-x.example/")
                automator.last_run_file = os.path.join(tmp, "last_run_date.txt")
                self.assertTrue(automator.is_first_run_today())
                automator.update_last_run_date()
                self.assertFalse(automator.is_first_run_today())
        finally:
            ROBOT.LOCAL_TIMEZONE = saved


class GeneratedEanTests(unittest.TestCase):
    def test_same_reference_always_gets_the_same_valid_ean(self):
        first = ROBOT.generate_ean13("REF-1")
        self.assertEqual(first, ROBOT.generate_ean13("REF-1"))
        self.assertNotEqual(first, ROBOT.generate_ean13("REF-2"))
        digits = [int(c) for c in first]
        self.assertEqual((sum(digits[0:12:2]) + 3 * sum(digits[1:12:2]) + digits[12]) % 10, 0)

    def test_it_does_not_touch_the_global_random_generator(self):
        import random
        random.seed(1234)
        expected = random.random()
        random.seed(1234)
        ROBOT.generate_ean13("REF-1")
        self.assertEqual(random.random(), expected)


class OnlineLookupsTests(unittest.TestCase):
    """Las consultas a terceros (buscador de imagenes, IA, buscador web) son opt-in: sin la variable no salen a la red."""

    def setUp(self):
        import urllib.request
        self._urlopen = urllib.request.urlopen
        urllib.request.urlopen = lambda *a, **k: self.fail("no debe haber trafico de red sin DONARTICULO_ONLINE_LOOKUPS=1")

    def tearDown(self):
        import urllib.request
        urllib.request.urlopen = self._urlopen

    def test_disabled_by_default(self):
        self.assertFalse(core.ONLINE_LOOKUPS)
        self.assertFalse(ROBOT.ONLINE_LOOKUPS)

    def test_image_search_and_ai_do_nothing_when_disabled(self):
        self.assertIsNone(core._fetch_google_image_bytes("mezclador"))
        self.assertIsNone(core._analyze_image_with_gemini(b"", "mezclador"))

    def test_web_verification_keeps_the_initial_category_when_disabled(self):
        classifier = ROBOT.ProductClassifier(enable_web_verification=True)
        self.assertEqual(classifier._verify_with_web_search("PIEZA SIN PALABRAS CLAVE", "VARIOS"), "VARIOS")


class DiscountParsingTests(unittest.TestCase):
    def test_two_discounts(self):
        self.assertEqual(core._parse_descuentos("37+13"), ("37", "13", None, None))

    def test_single_discount_goes_to_dto1(self):
        self.assertEqual(core._parse_descuentos(41), ("41", "", None, None))

    def test_decimal_pair_is_the_ceramic_formula_not_a_discount(self):
        self.assertEqual(core._parse_descuentos("1,8+4,5"), ("", "", 1.8, 4.5))

    def test_empty_values(self):
        self.assertEqual(core._parse_descuentos(None), ("", "", None, None))
        self.assertEqual(core._parse_descuentos("n/a"), ("", "", None, None))


class ExampleDatabaseTestCase(unittest.TestCase):
    """Crea la configuracion de ejemplo en una carpeta temporal."""

    @classmethod
    def setUpClass(cls):
        cls._tmp = tempfile.TemporaryDirectory()
        cls.root = cls._tmp.name
        cls.database = example.build(cls.root)

    @classmethod
    def tearDownClass(cls):
        cls._tmp.cleanup()


class FindReferenceTests(ExampleDatabaseTestCase):
    def test_exact_match_returns_provider_sheet_and_row(self):
        hit = core.find_reference_hit(self.database, "A100")
        self.assertEqual(hit["provider_name"], example.PROVEEDOR_PERFIL)
        self.assertEqual(hit["row_data"]["Cod. Articulo"], "A100")

    def test_lookup_is_case_insensitive(self):
        self.assertIsNotNone(core.find_reference_hit(self.database, "cer002"))

    def test_unknown_reference_is_not_found(self):
        self.assertIsNone(core.find_reference_hit(self.database, "NO-EXISTE"))

    def test_repeated_reference_reports_every_row(self):
        hit = core.find_reference_hit(self.database, "CER004")
        self.assertEqual(len(hit["rows"]), 2)

    def test_specific_row_can_be_forced(self):
        hit = core.find_reference_hit(self.database, "CER004", row_index=4)
        self.assertEqual(hit["row_data"]["Color"], "ANTRACITA")

    def test_reference_shared_by_two_providers_is_indexed_for_both(self):
        providers = ref_index.providers_for(self.database, "CER001")
        self.assertEqual(sorted(providers), sorted([example.PROVEEDOR_PERFIL, example.PROVEEDOR_CERAMICA]))


class ConfigLoadingTests(ExampleDatabaseTestCase):
    def test_provider_table(self):
        table = core.load_provider_table(os.path.join(self.root, "ayuda.xlsx"))
        cfg = table[example.PROVEEDOR_CERAMICA]
        self.assertEqual(cfg["provider_number"], "1001")
        self.assertEqual((cfg["ceramic_mult"], cfg["ceramic_add"]), (1.8, 4.5))

    def test_discounts_by_category_come_from_exceltodo(self):
        provs = core.load_provs_map(os.path.join(self.root, "PROVS.xlsx"))
        rules = core.load_exceltodo_rules(os.path.join(self.root, "exceltodo.xlsx"))
        pro, private, _reason = core.pick_discounts(example.PROVEEDOR_PERFIL, "GRIFERÍA", example.PRO_DISCOUNT,
                                                    example.PRIVATE_DISCOUNT, provs, rules)
        self.assertEqual((pro, private), (ROBOT.GRIFERIA_PRO_DISCOUNT, ROBOT.GRIFERIA_PRIVATE_DISCOUNT))

    def test_category_override_by_reference(self):
        core.load_overrides(os.path.join(self.root, "overrides.xlsx"))
        try:
            self.assertEqual(core.check_category_override(example.PROVEEDOR_PERFIL, "A300", ""), "ACCESORIOS")
            self.assertIsNone(core.check_category_override(example.PROVEEDOR_PERFIL, "A100", ""))
        finally:
            core.load_overrides(os.path.join(self.root, "no-existe.xlsx"))     # deja los overrides vacios


class ClassifierTests(unittest.TestCase):
    def setUp(self):
        self.classifier = ROBOT.ProductClassifier(enable_web_verification=False)

    def test_faucet(self):
        self.assertEqual(self.classifier.classify_product("MEZCLADOR LAVABO MONOMANDO CROMADO"), "GRIFERÍA")

    def test_adhesive(self):
        self.assertEqual(self.classifier.classify_product("ADHESIVO CEMENTO COLA FLEXIBLE C2"), "CEMENTO")

    def test_shower_screen(self):
        self.assertEqual(self.classifier.classify_product("MAMPARA DUCHA CORREDERA CRISTAL TEMPLADO"), "MAMPARAS")

    def test_empty_text_is_varios(self):
        self.assertEqual(self.classifier.classify_product(""), "VARIOS")

    def test_provider_category_column_wins(self):
        self.assertEqual(self.classifier.classify_product("PIEZA SIN PALABRAS CLAVE", provider_category="Sanitarios"),
                         "SANITARIO")


class ConvertRowTests(unittest.TestCase):
    USER_INPUTS = {"provider_number": "1001", "pro_discount": example.PRO_DISCOUNT, "private_discount": example.PRIVATE_DISCOUNT}

    def convert(self, row, provider):
        df = pd.DataFrame([row])
        return ROBOT.SingleRowConverter().convert_row(row, df, provider, dict(self.USER_INPUTS))

    def test_ceramic_uses_the_ceramic_formula(self):
        row = {"Referencia": "CER001", "Descripcion": "PAVIMENTO PORCELANICO 60X120", "Serie": "SERIE A",
               "Color": "BLANCO MATE", "Precio": 10.5, "Unidad": "M2", "EAN": ""}
        out, category, _info = self.convert(row, example.PROVEEDOR_CERAMICA)
        self.assertEqual(category, "CERÁMICA")
        self.assertAlmostEqual(out.loc[0, "Precio Base Tarifa"], 10.5 * ROBOT.CERAMIC_MULT + ROBOT.CERAMIC_ADD)
        self.assertEqual(out.loc[0, "Ud.Medida"], "M2")
        self.assertEqual(out.loc[0, "Descripcion"], "SERIE A PAVIMENTO PORCELANICO 60X120 BLANCO MATE REF: CER001")

    def test_generated_ean_is_a_valid_ean13_when_the_tariff_has_none(self):
        row = {"Referencia": "CER001", "Descripcion": "PAVIMENTO", "Precio": 1.0, "EAN": ""}
        out, _category, _info = self.convert(row, example.PROVEEDOR_CERAMICA)
        ean = str(out.loc[0, "Ean"])
        self.assertRegex(ean, r"^\d{13}$")
        digits = [int(c) for c in ean]
        self.assertEqual((sum(digits[0:12:2]) + 3 * sum(digits[1:12:2]) + digits[12]) % 10, 0)

    def test_provider_profile_overrides_columns_prices_and_discounts(self):
        row = {"Cod. Articulo": "A100", "Articulo": "MEZCLADOR LAVABO MONOMANDO CROMO (EXPOSICION)",
               "Precio": 120.0, "EAN": "8400000000017"}
        out, category, _info = self.convert(row, example.PROVEEDOR_PERFIL)
        self.assertEqual(category, "GRIFERÍA")
        self.assertEqual(out.loc[0, "Codigo"], "A100")
        self.assertEqual(out.loc[0, "Descripcion"], "MEZCLADOR LAVABO MONOMANDO CROMO REF: A100")   # sin el "(...)"
        self.assertEqual(out.loc[0, "Precio Compra"], 120.0)
        self.assertEqual(out.loc[0, "Precio Base Tarifa"], 120.0)                  # price_as_is: sin formula
        self.assertAlmostEqual(out.loc[0, "Tarifa 2"], 86.4)                       # -28 % (pro_discount del perfil)
        self.assertAlmostEqual(out.loc[0, "Tarifa 3"], 99.6)                       # -17 % (private_discount del perfil)
        self.assertEqual(out.loc[0, "%Dto1"], 38.0)
        self.assertEqual(out.loc[0, "Calculo Base"], "B")
        self.assertEqual(out.loc[0, "Fabricante"], 1000)

    def test_non_ceramic_is_always_sold_by_unit(self):
        row = {"Referencia": "X1", "Descripcion": "ESCOBILLERO DE PARED CROMADO", "Precio": 10.0, "Unidad": "KG"}
        out, _category, _info = self.convert(row, "PROVEEDOR_998")
        self.assertEqual(out.loc[0, "Ud.Medida"], "UD")


class CeramicLogisticsTests(unittest.TestCase):
    ROW = {"Referencia": "CER001", "Unidad": "M2", "M2/CAJA": 1.44, "Piezas/caja": 2, "Cajas/palet": 40, "Kg/caja": 30.0}

    def test_weight_is_per_sales_unit_not_per_box(self):
        lg = ROBOT._ceramic_logistics(self.ROW, "M2", example.PROVEEDOR_CERAMICA)
        self.assertAlmostEqual(lg["peso"], 30.0 / 1.44, places=5)
        self.assertEqual((lg["unid_minima"], lg["multiplo"]), (1.44, 1.44))
        self.assertEqual((lg["udx_envase"], lg["ud_palet"]), (2, 40))
        self.assertEqual(lg["peso_palet"], 1200.0)

    def test_providers_without_a_sheet_use_the_generic_logic(self):
        self.assertIsNone(ROBOT._ceramic_logistics(self.ROW, "M2", "PROVEEDOR_999"))

    def test_unit_translation(self):
        self.assertEqual(ROBOT._ceramic_unit({"Unidad": "PI"}, example.PROVEEDOR_CERAMICA), "PZ")

    def test_box_price_is_divided_by_pieces_per_box(self):
        row = {"Unidad": "CA", "Piezas/caja": 4}
        self.assertEqual(ROBOT._ceramic_price_divisor(row, example.PROVEEDOR_CERAMICA), 4)


if __name__ == "__main__":
    unittest.main()
