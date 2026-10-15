"""Pruebas sin Gmail ni Sheets: reglas reales de selección y calendario."""
import unittest
from datetime import date,timedelta
from mora1_reglas import puede_enviar,plantilla_para_fecha
class PruebasMora1(unittest.TestCase):
    def setUp(self):
        self.hoy=date(2026,10,15)
        self.inicio=self.hoy-timedelta(days=1)
    def test_cliente_mora1_dia_par(self):
        self.assertTrue(puede_enviar("MORA 1","SI","MORA_1","",self.inicio,self.hoy,"T001"))
    def test_cliente_al_dia(self):
        self.assertFalse(puede_enviar("AL DÍA","SI","MORA_1","",self.inicio,self.hoy,"T001"))
    def test_excluido(self):
        self.assertFalse(puede_enviar("MORA 1","SI","MORA_1","SI",self.inicio,self.hoy,"T001"))
    def test_no_habilitado(self):
        self.assertFalse(puede_enviar("MORA 1","NO","MORA_1","",self.inicio,self.hoy,"T001"))
    def test_tipo_distinto(self):
        self.assertFalse(puede_enviar("MORA 1","SI","BIENVENIDA","",self.inicio,self.hoy,"T001"))
    def test_dia_sin_correo(self):
        self.assertIsNone(plantilla_para_fecha(self.hoy,self.hoy))
    def test_dias_especiales(self):
        for dia,code in {4:"MORA101",10:"MORA102",15:"MORA103",19:"MORA104",24:"MORA105"}.items():
            with self.subTest(dia=dia):
                self.assertEqual(plantilla_para_fecha(self.hoy-timedelta(days=dia-1),self.hoy),code)
    def test_fecha_futura(self):
        self.assertIsNone(plantilla_para_fecha(self.hoy+timedelta(days=1),self.hoy))
    def test_plantilla_incorrecta(self):
        self.assertFalse(puede_enviar("MORA 1","SI","MORA_1","",self.inicio,self.hoy,"MORA101"))
if __name__=="__main__":unittest.main()
