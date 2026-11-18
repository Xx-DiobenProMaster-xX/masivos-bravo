"""Regresión: DE usa un cliente gspread nativo y evita el proxy."""
import ast
import pathlib
import unittest

class DepositoExtraReadTest(unittest.TestCase):
    def test_de_no_usa_proxy_para_lectura(self):
        source=pathlib.Path("app .py").read_text(encoding="utf-8")
        tree=ast.parse(source)
        function=next(x for x in tree.body if isinstance(x,ast.FunctionDef)
                      and x.name=="leer_fuente_alianzas_deposito_extra")
        body=ast.get_source_segment(source,function)
        self.assertIn("cliente_con_retry._obj",body)
        self.assertIn("cliente_directo.open_by_key(",body)
        self.assertIn("libro.worksheet(ALIANZAS_DE_TAB)",body)
        self.assertIn("worksheet_de.get_all_values()",body)
        self.assertNotIn("fuente.values_get(",body)
        self.assertNotIn("fuente.worksheet(",body)
        self.assertIn('fila[9]',body)
        self.assertIn('fila[6]',body)

if __name__=="__main__":
    unittest.main()
