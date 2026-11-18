"""
Pruebas de ajustar_pedido.py.

Usan archivos Excel sinteticos creados en una carpeta temporal, asi que no
dependen de los datos reales del negocio (que no estan en GitHub).

Ejecutar desde la raiz del repo:
    python -m unittest discover -s tests -v
"""

from __future__ import annotations

import contextlib
import io
import sys
import tempfile
import unittest
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
import ajustar_pedido as ap  # noqa: E402


HEADERS_PEDIDO = ["Código", "Producto", "Modelo", "Saldo Actual", "González Ruiz",
                  "Promedio", "Mínimos", "Máximos", "Exp Smoothing", "Estacional",
                  "Recomendación Claude"]


def _clasif_headers(anio: int = 2025) -> list[str]:
    return ["Codigo", "Producto", "Modelo", "Saldo Actual", f"Total Uds {anio}",
            "Ingreso Total", "Costo Total", "Margen Bruto", "Promedio Mensual",
            "Meses c/Venta", "ABC", "XYZ", "ABC-XYZ", "Interpretacion"]


def _fila_clasif(codigo, clase, total_uds, costo_total, prom_mensual):
    return [codigo, f"Producto {codigo}", "", 0, total_uds, costo_total * 2,
            costo_total, costo_total, prom_mensual, 12, clase[0], clase[1], clase, ""]


def _escribir_clasificacion(ruta: Path, filas, anio: int = 2025):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Orden Original"
    ws.append(_clasif_headers(anio))
    for f in filas:
        ws.append(f)
    wb.save(ruta)


def _escribir_pedido(ruta: Path, filas):
    """filas: lista de (codigo, producto, saldo, sugerido)."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Todos los Productos"
    ws.append(HEADERS_PEDIDO)
    for codigo, producto, saldo, sugerido in filas:
        ws.append([codigo, producto, "", saldo, 0, 0, 0, 0, 0, 0, sugerido])
    ws.append(["TOTAL", "", "", 0, 0, 0, 0, 0, 0, 0, 0])
    wb.save(ruta)


def _info(clase, total_uds=120, prom=10.0, anio=2025, costo_unitario=None):
    return {"abc_xyz": clase, "total_uds": total_uds, "promedio_mensual": prom,
            "anio": anio, "costo_unitario": costo_unitario}


COL = "Recomendación Claude"


class TestCalcularAjuste(unittest.TestCase):
    """Reglas de negocio documentadas en el README."""

    def _aj(self, clase, sugerido, saldo=0, **kw):
        item = {COL: sugerido, "Saldo Actual": saldo}
        return ap.calcular_ajuste(item, _info(clase, **kw), COL)

    def test_factores_por_clase(self):
        esperado = {"AX": 100, "AY": 100, "BX": 100, "BY": 75,
                    "CX": 30, "CY": 25, "CZ": 0}
        for clase, cant in esperado.items():
            with self.subTest(clase=clase):
                # promedio alto para que el tope de clases Z no intervenga
                self.assertEqual(self._aj(clase, 100, prom=1000)["cantidad_ajustada"], cant)

    def test_clases_z_con_promedio_alto_usan_factor(self):
        self.assertEqual(self._aj("AZ", 100, prom=1000)["cantidad_ajustada"], 75)
        self.assertEqual(self._aj("BZ", 100, prom=1000)["cantidad_ajustada"], 50)

    def test_tope_dos_meses_en_clases_z(self):
        r = self._aj("AZ", 100, prom=10)
        self.assertEqual(r["cantidad_ajustada"], 20)
        self.assertIn("2 meses", r["justificacion"])

    def test_cero_ventas_clase_a_revisar_manual(self):
        r = self._aj("AX", 50, total_uds=0)
        self.assertEqual(r["accion"], "REVISAR MANUAL")
        self.assertEqual(r["cantidad_ajustada"], 50)

    def test_cero_ventas_otras_clases_no_ordenar(self):
        r = self._aj("BX", 50, total_uds=0)
        self.assertEqual((r["accion"], r["cantidad_ajustada"]), ("NO ORDENAR", 0))
        self.assertIn("Sin ventas en 2025", r["justificacion"])

    def test_sobrestock_no_ordenar(self):
        r = self._aj("BX", 50, saldo=200, total_uds=120)
        self.assertEqual((r["accion"], r["cantidad_ajustada"]), ("NO ORDENAR", 0))

    def test_sobrestock_ax_exento(self):
        r = self._aj("AX", 50, saldo=200, total_uds=120)
        self.assertEqual((r["accion"], r["cantidad_ajustada"]), ("ORDENAR", 50))

    def test_sin_clasificacion_revisar_manual(self):
        r = ap.calcular_ajuste({COL: 40, "Saldo Actual": 0}, None, COL)
        self.assertEqual((r["accion"], r["cantidad_ajustada"], r["abc_xyz"]),
                         ("REVISAR MANUAL", 40, "N/D"))

    def test_ano_en_justificacion(self):
        r = self._aj("BX", 50, total_uds=0, anio=2026)
        self.assertIn("Sin ventas en 2026", r["justificacion"])


class TestCarga(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)

    def tearDown(self):
        self._tmp.cleanup()

    def test_detecta_ano_de_la_columna_total(self):
        for anio in (2025, 2026):
            with self.subTest(anio=anio):
                ruta = self.dir / f"c{anio}.xlsx"
                _escribir_clasificacion(ruta, [_fila_clasif("001", "AX", 120, 600, 10)], anio)
                c = ap.cargar_clasificacion(ruta)
                self.assertEqual(c["001"]["anio"], anio)
                self.assertEqual(c["001"]["total_uds"], 120)

    def test_varias_columnas_total_usa_la_mas_reciente(self):
        self.assertEqual(ap._detectar_columna_total(["Total Uds 2025", "Total Uds 2026"]),
                         (2026, "Total Uds 2026"))

    def test_falta_columna_total_da_error_claro(self):
        ruta = self.dir / "c.xlsx"
        wb = openpyxl.Workbook()
        wb.active.title = "Orden Original"
        wb.active.append([h for h in _clasif_headers() if not h.startswith("Total Uds")])
        wb.save(ruta)
        with self.assertRaisesRegex(ValueError, "Total Uds"):
            ap.cargar_clasificacion(ruta)

    def test_costo_unitario(self):
        ruta = self.dir / "c.xlsx"
        _escribir_clasificacion(ruta, [
            _fila_clasif("001", "AX", 120, 600, 10),   # 600 / 120 = 5.0
            _fila_clasif("002", "CZ", 0, 0, 0),        # sin ventas -> sin costo
        ])
        c = ap.cargar_clasificacion(ruta)
        self.assertAlmostEqual(c["001"]["costo_unitario"], 5.0)
        self.assertIsNone(c["002"]["costo_unitario"])

    def test_pedido_omite_total_eliminado_y_sin_producto(self):
        ruta = self.dir / "p.xlsx"
        _escribir_pedido(ruta, [
            ("001", "Producto A", 0, 10),
            ("ELIMINADO_29389_43823", "Producto borrado", 0, 5),
            ("003", "", 0, 5),
        ])
        with contextlib.redirect_stdout(io.StringIO()):
            _, datos, hoja = ap.cargar_pedido(ruta)
        self.assertEqual([d["Código"] for d in datos], ["001"])
        self.assertEqual(hoja, "Todos los Productos")


class TestEndToEnd(unittest.TestCase):

    def setUp(self):
        self._tmp = tempfile.TemporaryDirectory()
        self.dir = Path(self._tmp.name)
        self.clasif = self.dir / ap.CLASIFICACION_DEFAULT
        _escribir_clasificacion(self.clasif, [
            _fila_clasif("001", "AX", 120, 600, 10),
            _fila_clasif("002", "CZ", 12, 24, 1),
            _fila_clasif("003", "BY", 120, 1200, 10),
        ])
        self.pedido = self.dir / "PROVEEDOR.xlsx"
        _escribir_pedido(self.pedido, [
            ("001", "Producto A", 5, 20),
            ("002", "Producto B", 0, 10),
            ("003", "Producto C", 0, 8),
            ("999", "Producto nuevo", 0, 4),
        ])

    def tearDown(self):
        self._tmp.cleanup()

    def _main(self, *args):
        # -c explicito: si no, el programa prefiere la clasificacion real que
        # exista en la carpeta actual (la raiz del repo) sobre la de prueba.
        out = io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(out):
            rc = ap.main([str(a) for a in args] + ["-c", str(self.clasif)])
        return rc, out.getvalue()

    def test_genera_salida_con_ambas_hojas(self):
        rc, out = self._main(self.pedido)
        self.assertEqual(rc, 0, out)
        salida = self.dir / "PROVEEDOR_ajustado.xlsx"
        wb = openpyxl.load_workbook(salida)
        self.assertEqual(wb.sheetnames, ["Pedido + Clasificacion", "Recomendaciones"])

        h1 = [c.value for c in wb["Pedido + Clasificacion"][1]]
        self.assertEqual(h1, HEADERS_PEDIDO + ["ABC", "XYZ", "ABC-XYZ"])

        ws = wb["Recomendaciones"]
        h2 = [c.value for c in ws[1]]
        self.assertEqual(h2[:12], [
            "Codigo", "Producto", "ABC-XYZ", "Saldo Actual", "Vtas 2025 (uds)",
            "Prom Mensual", "Meses c/Venta", "Sugerido Original", "Cantidad Ajustada",
            "Diferencia", "Accion", "Justificacion"])
        self.assertEqual(h2[12:], ["Costo Unit. Prom", "Costo Sugerido",
                                   "Costo Ajustado", "Ahorro"])

        filas = {r[0]: r for r in ws.iter_rows(min_row=2, values_only=True)}
        # Orden: AX, BY, CZ, N/D
        self.assertEqual(list(filas), ["001", "003", "002", "999"])
        self.assertEqual(filas["001"][8], 20)            # AX 100%
        self.assertEqual(filas["003"][8], 6)             # BY 75% de 8
        self.assertEqual(filas["002"][8], 0)             # CZ
        self.assertEqual(filas["999"][10], "REVISAR MANUAL")
        # Dinero: BY costo 1200/120 = 10 -> sugerido 80, ajustado 60, ahorro 20
        self.assertEqual(filas["003"][12:], (10.0, 80.0, 60.0, 20.0))
        self.assertEqual(filas["999"][12:], (None, None, None, None))
        self.assertIn("Ahorro estimado", out)

    def test_salida_personalizada(self):
        destino = self.dir / "otro.xlsx"
        rc, out = self._main(self.pedido, "-o", destino)
        self.assertEqual(rc, 0, out)
        self.assertTrue(destino.exists())

    def test_modo_lote(self):
        otro = self.dir / "OTRO.xlsx"
        _escribir_pedido(otro, [("001", "Producto A", 0, 3)])
        rc, out = self._main(self.pedido, otro)
        self.assertEqual(rc, 0, out)
        self.assertTrue((self.dir / "PROVEEDOR_ajustado.xlsx").exists())
        self.assertTrue((self.dir / "OTRO_ajustado.xlsx").exists())
        self.assertIn("2 de 2", out)

    def test_lote_sigue_si_un_archivo_falta(self):
        rc, out = self._main(self.pedido, self.dir / "NO_EXISTE.xlsx")
        self.assertEqual(rc, 1)
        self.assertTrue((self.dir / "PROVEEDOR_ajustado.xlsx").exists())
        self.assertIn("1 de 2", out)

    def test_lote_no_acepta_salida_unica(self):
        rc, out = self._main(self.pedido, self.pedido, "-o", self.dir / "x.xlsx")
        self.assertEqual(rc, 1)
        self.assertIn("unico pedido", out)

    def test_formato_con_gonzalez_ruiz_per_month(self):
        """Formato nuevo (nov 2026): columna extra despues de 'Gonzalez Ruiz'."""
        nuevo = self.dir / "NUEVO.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Todos los Productos"
        headers = HEADERS_PEDIDO[:5] + ["González Ruiz per-month"] + HEADERS_PEDIDO[5:]
        ws.append(headers)
        ws.append(["003", "Producto C", "", 0, 9, 9, 0, 0, 0, 0, 0, 8])
        wb.save(nuevo)

        rc, out = self._main(nuevo)
        self.assertEqual(rc, 0, out)
        wb = openpyxl.load_workbook(self.dir / "NUEVO_ajustado.xlsx")
        h1 = [c.value for c in wb["Pedido + Clasificacion"][1]]
        self.assertEqual(h1, headers + ["ABC", "XYZ", "ABC-XYZ"])
        fila = next(wb["Recomendaciones"].iter_rows(min_row=2, values_only=True))
        self.assertEqual((fila[0], fila[7], fila[8]), ("003", 8, 6))  # BY 75% de 8

    def test_columna_inexistente(self):
        rc, out = self._main(self.pedido, "-q", "No Existe")
        self.assertEqual(rc, 1)
        self.assertIn("no existe", out)


if __name__ == "__main__":
    unittest.main()
