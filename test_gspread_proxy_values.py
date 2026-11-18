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
        # Regresión: clase de gspread que NO hereda de list pero sí
        # proporciona el protocolo de secuencia.
        class ValueRangeSequence:
            def __init__(self, items):
                self.items = items
            def __len__(self):
                return len(self.items)
            def __getitem__(self, index):
                return self.items[index]
        ValueRangeSequence.__module__ = "gspread.worksheet"
        secuencia = ValueRangeSequence([["Referencia"], ["123"]])
        self.assertIs(wrap(secuencia), secuencia)
        self.assertEqual(len(wrap(secuencia)), 2)

        class Worksheet: pass
        Worksheet.__module__ = "gspread.worksheet"
        self.assertIsInstance(wrap(Worksheet()), _GSpreadRetryProxy)

if __name__ == "__main__":
    unittest.main()
