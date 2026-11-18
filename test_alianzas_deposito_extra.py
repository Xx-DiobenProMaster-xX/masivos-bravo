import unittest
from alianzas_deposito_extra import monto_de, dinero_colombiano, html_deposito_extra

class DepositoExtraTest(unittest.TestCase):
    def test_monto_entero(self):
        self.assertEqual(dinero_colombiano("1600000"), "$ 1.600.000")
    def test_separadores(self):
        self.assertEqual(dinero_colombiano("1.600.000"), "$ 1.600.000")
        self.assertEqual(dinero_colombiano("1600000,00"), "$ 1.600.000")
    def test_invalidos(self):
        self.assertIsNone(monto_de(""))
        self.assertIsNone(monto_de("-10"))
        self.assertIsNone(monto_de("NA"))
    def test_html_datos_protegidos(self):
        cuerpo = html_deposito_extra("Laura <test>", "Banco & Uno", "1600000",
            "https://example.com/logo.png","576017438029")
        self.assertIn("Laura &lt;test&gt;", cuerpo)
        self.assertIn("Banco &amp; Uno", cuerpo)
        self.assertIn("$ 1.600.000", cuerpo)
        self.assertIn("Aporte adicional estimado", cuerpo)
    def test_sin_valor_no_enviar(self):
        with self.assertRaises(ValueError):
            html_deposito_extra("Laura","Banco","0","","")

if __name__=="__main__":
    unittest.main()
