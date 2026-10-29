"""Validación de la identificación de movimientos con cheques.

Cubre los casos que hacen que los movimientos no cuadren:
  - abreviaturas: CHQ, CHEQ, CHKS, CK, CHECK
  - símbolos: N°, Nº, N.º, No., #, №
  - números con ceros de relleno ("000011666") o decimales ("826486.0")
  - descripción y referencia que dicen el mismo número de formas distintas

Ejecutar desde la raíz del proyecto:

    python -m unittest discover -s tests -v
"""

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app  # noqa: E402


class TestNormalizarNumeroCheque(unittest.TestCase):
    """normalizar_numero_cheque() unifica todas las formas de escribir el número."""

    def test_celda_numerica_con_decimal(self):
        self.assertEqual(app.normalizar_numero_cheque(826486.0), "826486")
        self.assertEqual(app.normalizar_numero_cheque("826486.0"), "826486")
        self.assertEqual(app.normalizar_numero_cheque("826486,00"), "826486")

    def test_ceros_de_relleno(self):
        self.assertEqual(app.normalizar_numero_cheque("000011666"), "11666")
        self.assertEqual(app.normalizar_numero_cheque("0077"), "77")
        self.assertEqual(app.normalizar_numero_cheque("0"), "0")

    def test_simbolos(self):
        self.assertEqual(app.normalizar_numero_cheque("N° 12345"), "12345")
        self.assertEqual(app.normalizar_numero_cheque("Nº12345"), "12345")
        self.assertEqual(app.normalizar_numero_cheque("N.º12345"), "12345")
        self.assertEqual(app.normalizar_numero_cheque("No. 1234"), "1234")
        self.assertEqual(app.normalizar_numero_cheque("Nro. 1234"), "1234")
        self.assertEqual(app.normalizar_numero_cheque("#1234"), "1234")

    def test_abreviaturas(self):
        self.assertEqual(app.normalizar_numero_cheque("CHQ#1234"), "1234")
        self.assertEqual(app.normalizar_numero_cheque("CHQ No.1234"), "1234")
        self.assertEqual(app.normalizar_numero_cheque("Cheque: 1234"), "1234")
        self.assertEqual(app.normalizar_numero_cheque("CK 1234"), "1234")

    def test_identificador_alfanumerico(self):
        self.assertEqual(app.normalizar_numero_cheque("1234-A"), "1234-A")
        self.assertEqual(app.normalizar_numero_cheque("4321/AB"), "4321/AB")

    def test_valores_invalidos(self):
        self.assertEqual(app.normalizar_numero_cheque(None), "")
        self.assertEqual(app.normalizar_numero_cheque(""), "")
        self.assertEqual(app.normalizar_numero_cheque("Pago"), "")
        self.assertEqual(app.normalizar_numero_cheque("12/03/2026"), "")
        self.assertEqual(app.normalizar_numero_cheque("01/09/2026"), "")


class TestDeteccionCheques(unittest.TestCase):
    """detectar_movimiento_cheque() reconoce abreviaturas y símbolos."""

    @staticmethod
    def detectar(descripcion, referencia="", codigo="", debito=-100.0, credito=0.0):
        return app.detectar_movimiento_cheque(
            referencia=referencia,
            codigo=codigo,
            descripcion=descripcion,
            debito=debito,
            credito=credito,
        )

    def test_abreviaturas_detectan_cheque(self):
        casos = [
            "CHQ#12345 PAGO PROVEEDOR",
            "CHQ No.12345",
            "CHQ12345",
            "CK12345",
            "CK 12345",
            "CHEQ 7766",
            "CHKS 9988",
            "CHKS9988",
            "CHECK 4321",
            "Pago por cheque 55",
        ]
        for descripcion in casos:
            with self.subTest(descripcion=descripcion):
                resultado = self.detectar(descripcion)
                self.assertTrue(resultado["es_cheque"], descripcion)

    def test_simbolos_en_el_numero(self):
        casos = {
            "CHQ N°12345": "12345",
            "CHQ Nº12345": "12345",
            "CHQ N.° 12345": "12345",
            "CHEQUE Nº 5544 REEMBOLSO": "5544",
            "N° CHEQUE 3344": "3344",
            "No. de Cheque 3344": "3344",
            "CHEQUE.No.123456": "123456",
            "CHQ N° 4321/AB": "4321/AB",
            "CHEQUE :     000011661": "11661",
            "Cheque No. 0077": "77",
        }
        for descripcion, esperado in casos.items():
            with self.subTest(descripcion=descripcion):
                resultado = self.detectar(descripcion)
                self.assertTrue(resultado["es_cheque"], descripcion)
                self.assertEqual(resultado["numero_cheque"], esperado, descripcion)

    def test_palabras_entre_la_palabra_cheque_y_el_numero(self):
        resultado = self.detectar("Cheque devuelto No. 9911")
        self.assertTrue(resultado["es_cheque"])
        self.assertEqual(resultado["numero_cheque"], "9911")
        self.assertEqual(resultado["tipo_cheque"], "DEVUELTO")

    def test_tipos_de_movimiento(self):
        ingreso = self.detectar("Deposito de cheque 8899", debito=0.0, credito=500.0)
        self.assertEqual(ingreso["tipo_cheque"], "INGRESO")
        self.assertEqual(ingreso["numero_cheque"], "8899")

        egreso = self.detectar("PAGO CHEQUE No. 11666", debito=-31630.98, credito=0.0)
        self.assertEqual(egreso["tipo_cheque"], "EGRESO")

        devuelto = self.detectar("RECHAZO CHEQUE 4455")
        self.assertEqual(devuelto["tipo_cheque"], "DEVUELTO")

    def test_devoluciones_con_abreviaturas(self):
        casos = [
            "DEV CHEQ 4455",
            "RECH CHEQ 4455",
            "PROT CHEQUE 4455",
            "Cheque devuelto No. 9911",
            "REVERSADO CHEQUE 1234",
        ]
        for descripcion in casos:
            with self.subTest(descripcion=descripcion):
                resultado = self.detectar(descripcion)
                self.assertTrue(resultado["es_cheque"], descripcion)
                self.assertEqual(resultado["tipo_cheque"], "DEVUELTO", descripcion)

    def test_no_son_cheques(self):
        casos = [
            "TRANSFERENCIA A PROVEEDOR",
            "DEPOSITO CLIENTE",
            "PAGO NÓMINA QUINCENAL",
            "COMISIÓN BANCARIA",
            # "GRAVABLOCK" contiene "CK": no debe detectarse como cheque.
            "T365 DE: INDUSTRIAS GRAVABLOCK",
        ]
        for descripcion in casos:
            with self.subTest(descripcion=descripcion):
                resultado = self.detectar(descripcion)
                self.assertFalse(resultado["es_cheque"], descripcion)

        # Tampoco unido a la referencia numérica.
        resultado = app.detectar_movimiento_cheque(
            referencia="912406873",
            descripcion="T365 DE: INDUSTRIAS GRAVABLOCK",
            debito=0.0,
            credito=178.54,
        )
        self.assertFalse(resultado["es_cheque"])

    def test_no_toma_el_numero_de_cuenta(self):
        # "CTA No. 1234567" es la cuenta, no el número del cheque.
        resultado = self.detectar("PAGO CHEQUE CTA No. 1234567")
        self.assertTrue(resultado["es_cheque"])
        self.assertEqual(resultado["numero_cheque"], "")


class TestCuadreDescripcionReferencia(unittest.TestCase):
    """La descripción y la referencia deben dar el mismo número de cheque."""

    def test_descripcion_y_referencia_cuadran(self):
        casos = [
            ("11666", "PAGO CHEQUE No. 000011666"),
            ("826486", "PAGO DE CHEQUE COBRADO POR CAMARA"),
            ("000011677", "CHEQUE :     000011677"),
            ("CHQ-1234", "PAGO CHQ 1234"),
            ("N° 4321/AB", "CHQ N° 4321/AB"),
        ]
        for referencia, descripcion in casos:
            with self.subTest(referencia=referencia, descripcion=descripcion):
                resultado = app.detectar_movimiento_cheque(
                    referencia=referencia,
                    descripcion=descripcion,
                    debito=-10.0,
                    credito=0.0,
                )
                self.assertTrue(resultado["es_cheque"])
                self.assertEqual(
                    resultado["numero_cheque"],
                    app.normalizar_numero_cheque(referencia),
                )

    def test_referencia_numerica_sin_texto_de_cheque(self):
        # El banco solo pone "PAGO DE CHEQUE COBRADO POR CAMARA" y el número
        # vive en la columna de referencia.
        resultado = app.detectar_movimiento_cheque(
            referencia="826486.0",
            descripcion="Pago de Cheques",
            debito=-315.15,
            credito=0.0,
        )
        self.assertTrue(resultado["es_cheque"])
        self.assertEqual(resultado["numero_cheque"], "826486")

    def test_no_usa_fechas_como_numero(self):
        resultado = app.detectar_movimiento_cheque(
            referencia="12/03/2026",
            descripcion="CHEQUE PROTESTADO",
            debito=-10.0,
            credito=0.0,
        )
        self.assertTrue(resultado["es_cheque"])
        self.assertEqual(resultado["numero_cheque"], "")


class TestExtraerTransaccion(unittest.TestCase):
    """La clasificación llega hasta la transacción sin tocar los importes."""

    columnas = {
        "fecha": 0,
        "referencia": 1,
        "descripcion": 2,
        "debito": 3,
        "credito": 4,
        "saldo": 5,
    }

    def extraer(self, fila):
        return app.extraer_transaccion(
            fila,
            self.columnas,
            "BANCO INDUSTRIAL",
            "1234",
            expected_month=9,
            expected_year=2026,
            previous_balance=100000.0,
            previous_date=None,
            saldo_inicial_cuenta=100000.0,
        )

    def test_numero_normalizado_en_transaccion(self):
        transaccion = self.extraer(
            ["01/09/2026", "11666", "PAGO CHEQUE No. 000011666", -31630.98, "", 100000]
        )
        self.assertIsNotNone(transaccion)
        self.assertTrue(transaccion["es_cheque"])
        self.assertEqual(transaccion["tipo_cheque"], "EGRESO")
        self.assertEqual(transaccion["numero_cheque"], "11666")
        # Los importes no cambian por la detección de cheques.
        self.assertEqual(transaccion["debito"], -31630.98)
        self.assertEqual(transaccion["credito"], 0.0)

    def test_abreviatura_con_simbolo(self):
        transaccion = self.extraer(
            ["02/09/2026", "", "CHQ N° 4321/AB PAGO", -500, "", 99500]
        )
        self.assertIsNotNone(transaccion)
        self.assertTrue(transaccion["es_cheque"])
        self.assertEqual(transaccion["numero_cheque"], "4321/AB")

    def test_movimiento_sin_cheque(self):
        transaccion = self.extraer(
            ["04/09/2026", "", "TRANSFERENCIA A PROVEEDOR", -100, "", 99400]
        )
        self.assertIsNotNone(transaccion)
        self.assertFalse(transaccion["es_cheque"])
        self.assertEqual(transaccion["numero_cheque"], "")


class TestProcesarHoja(unittest.TestCase):
    """Prueba de extremo a extremo sobre una hoja de estado de cuenta."""

    def procesar(self):
        from openpyxl import Workbook

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "MOVIMIENTOS"
        worksheet.append(
            ["Fecha", "Referencia", "Descripción", "Débito", "Crédito", "Saldo"]
        )
        worksheet.append(["01/09/2026", "11666", "PAGO CHEQUE No. 000011666", -31630.98, "", 100000])
        worksheet.append(["02/09/2026", "", "CHQ N° 4321/AB PAGO", -500, "", 99500])
        worksheet.append(["03/09/2026", "", "Deposito de cheque 8899", "", 200, 99700])
        worksheet.append(["04/09/2026", "", "CK12345 PAGO PROVEEDOR", -250, "", 99450])
        worksheet.append(["05/09/2026", "", "RECHAZO CHEQUE 4455", -75, "", 99375])
        worksheet.append(["06/09/2026", "", "TRANSFERENCIA A PROVEEDOR", -100, "", 99275])

        return app.procesar_hoja(
            worksheet,
            moneda_respaldo="USD",
            nombre_archivo="estado_cuenta_banco.xlsx",
        )

    def test_hoja_completa(self):
        transacciones = self.procesar()
        self.assertEqual(len(transacciones), 6)

        por_descripcion = {t["descripcion"]: t for t in transacciones}

        principal = por_descripcion["PAGO CHEQUE No. 000011666"]
        self.assertTrue(principal["es_cheque"])
        self.assertEqual(principal["numero_cheque"], "11666")
        self.assertEqual(principal["tipo_cheque"], "EGRESO")
        self.assertEqual(principal["debito"], -31630.98)

        simbolo = por_descripcion["CHQ N° 4321/AB PAGO"]
        self.assertTrue(simbolo["es_cheque"])
        self.assertEqual(simbolo["numero_cheque"], "4321/AB")

        ingreso = por_descripcion["Deposito de cheque 8899"]
        self.assertEqual(ingreso["tipo_cheque"], "INGRESO")
        self.assertEqual(ingreso["numero_cheque"], "8899")

        devuelto = por_descripcion["RECHAZO CHEQUE 4455"]
        self.assertEqual(devuelto["tipo_cheque"], "DEVUELTO")
        self.assertEqual(devuelto["numero_cheque"], "4455")

        normal = por_descripcion["TRANSFERENCIA A PROVEEDOR"]
        self.assertFalse(normal["es_cheque"])

        # El total de la hoja sigue cuadrando con los importes originales.
        total_debito = sum(t["debito"] for t in transacciones)
        self.assertAlmostEqual(total_debito, -31630.98 - 500 - 250 - 75 - 100, places=2)
        total_credito = sum(t["credito"] for t in transacciones)
        self.assertAlmostEqual(total_credito, 200.0, places=2)


class TestHojaMovimientosCheques(unittest.TestCase):
    """La hoja "MOVIMIENTOS CHEQUES" lista los cheques con su número cuadrado."""

    def test_hoja_cheques(self):
        from openpyxl import Workbook

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "MOVIMIENTOS"
        worksheet.append(
            ["Fecha", "Referencia", "Descripción", "Débito", "Crédito", "Saldo"]
        )
        worksheet.append(["01/09/2026", "11666", "PAGO CHEQUE No. 000011666", -31630.98, "", 100000])
        worksheet.append(["02/09/2026", "", "CHQ N° 4321/AB PAGO", -500, "", 99500])
        worksheet.append(["03/09/2026", "", "TRANSFERENCIA A PROVEEDOR", -100, "", 99400])

        transacciones = app.procesar_hoja(
            worksheet,
            moneda_respaldo="USD",
            nombre_archivo="estado_cuenta_banco.xlsx",
        )

        salida = Workbook()
        total = app.escribir_movimientos_cheques(salida, transacciones)

        self.assertEqual(total, 2)
        hoja = salida[app.SHEET_CHEQUES]
        filas = list(hoja.iter_rows(min_row=2, values_only=True))
        self.assertEqual(len(filas), 2)

        numeros = {fila[6] for fila in filas}
        self.assertEqual(numeros, {"11666", "4321/AB"})

        tipos = {fila[7] for fila in filas}
        self.assertEqual(tipos, {"EGRESO"})

        # Los importes siguen intactos.
        self.assertAlmostEqual(sum(fila[9] for fila in filas), 31630.98 + 500, places=2)
        self.assertAlmostEqual(sum(fila[10] for fila in filas), 0.0, places=2)


class TestColumnaTransaccionAgricola(unittest.TestCase):
    """Banco Agrícola titula la columna del concepto solo "Transaccion".

    Regresión: esa palabra se clasificaba como "fecha" (por el alias
    "fecha transaccion"), la columna del concepto se descartaba y el texto
    del movimiento se perdía. Un cheque como "CHEQUE CONFORMADO PAGADO EN
    COMP" dejaba de detectarse y su número nunca se extraía.
    """

    def test_header_transaccion_es_descripcion(self):
        self.assertEqual(app._header_field("Transaccion")[0], "descripcion")
        self.assertEqual(app._header_field("Transacciones")[0], "descripcion")

    def test_no_desplaza_la_fecha(self):
        # "Fecha Transaccion" sigue siendo la columna de fecha.
        self.assertEqual(app._header_field("Fecha Transaccion")[0], "fecha")

    def test_hoja_agricola_detecta_cheque(self):
        from openpyxl import Workbook

        workbook = Workbook()
        worksheet = workbook.active
        worksheet.title = "AGRICOLA"
        worksheet.append(
            [
                "Fecha Transaccion", "Fecha Aplicada", "Hora", "Transaccion",
                "Canal", "Referencia", "Cargo", "Abono", "Saldo",
            ]
        )
        worksheet.append(
            ["01/10/2026", "01/10/2026", "01:25:22 AM",
             "LIQUIDACION POR PAGOS CON NPE", "OPERACIONES VARIAS", "0", "", 592.3, 88326.24]
        )
        worksheet.append(
            ["12/10/2026", "12/10/2026", "11:37:51 AM",
             "CHEQUE CONFORMADO PAGADO EN COMP",
             "DEPTO. DE COMPENSACION SAN SALVADOR", "4951", 40000, "", 82200.68]
        )
        worksheet.append(
            ["12/10/2026", "12/10/2026", "02:36:17 PM",
             "CHEQUE CONFORMADO PAGADO EN COMP",
             "DEPTO. DE COMPENSACION SAN SALVADOR", "4940", 50000, "", 32200.68]
        )

        transacciones = app.procesar_hoja(
            worksheet,
            moneda_respaldo="USD",
            nombre_archivo="estado_agricola.xlsx",
        )

        # El concepto no se pierde en ninguna fila.
        self.assertEqual(len(transacciones), 3)
        for transaccion in transacciones:
            self.assertTrue(
                transaccion["descripcion"],
                "La columna Transaccion debe conservar el concepto",
            )

        cheques = [t for t in transacciones if t["es_cheque"]]
        self.assertEqual(len(cheques), 2)
        self.assertEqual({c["numero_cheque"] for c in cheques}, {"4951", "4940"})
        for cheque in cheques:
            self.assertEqual(cheque["tipo_cheque"], "EGRESO")
            # Los importes de la hoja quedan intactos.
            self.assertLess(cheque["debito"], 0)


if __name__ == "__main__":
    unittest.main(verbosity=2)
