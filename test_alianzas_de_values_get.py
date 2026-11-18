"""Evita regresar a Worksheet.get() para leer DE."""
import ast
import pathlib
import unittest

class DepositoExtraReadTest(unittest.TestCase):
    def test_de_usa_values_get_y_no_worksheet_get(self):
        source=pathlib.Path("app .py").read_text(encoding="utf-8")
        tree=ast.parse(source)
        function=next(x for x in tree.body if isinstance(x,ast.FunctionDef)
                      and x.name=="leer_fuente_alianzas_deposito_extra")
        body=ast.get_source_segment(source,function)
        self.assertIn("values_get(",body)
        self.assertNotIn(".worksheet(ALIANZAS_DE_TAB).get(",body)
        self.assertIn('respuesta.get("values", [])',body)
        self.assertIn('fila[9]',body)
        self.assertIn('fila[6]',body)
if __name__=="__main__":
    unittest.main()
