"""Prueba que los datos ValueRange de gspread se mantengan como listas."""
import ast
import pathlib
import types
import unittest

class ProxyValueRangeTest(unittest.TestCase):
    def test_datos_de_sheet_no_se_envuelven(self):
        texto = pathlib.Path("app .py").read_text(encoding="utf-8")
        arbol = ast.parse(texto)
        funcion = next(
            n for n in arbol.body if isinstance(n, ast.FunctionDef)
            and n.name == "_wrap_gspread_result"
        )
        modulo = ast.Module(body=[funcion], type_ignores=[])
        ast.fix_missing_locations(modulo)
        class _GSpreadRetryProxy:
            def __init__(self, obj):
                self.obj = obj
        class ValueRange(list):
            pass
        ValueRange.__module__ = "gspread.worksheet"
        ctx = {"_GSpreadRetryProxy": _GSpreadRetryProxy}
        exec(compile(modulo, "<prueba>", "exec"), ctx)
        wrap = ctx["_wrap_gspread_result"]
        filas = ValueRange([["Referencia", "Correo"], ["123", "x@example.com"]])
        resultado = wrap(filas)
        self.assertIs(resultado, filas)
        self.assertEqual(len(resultado), 2)
        self.assertEqual(resultado[1][1], "x@example.com")
        class Worksheet: pass
        Worksheet.__module__ = "gspread.worksheet"
        self.assertIsInstance(wrap(Worksheet()), _GSpreadRetryProxy)

if __name__ == "__main__":
    unittest.main()
