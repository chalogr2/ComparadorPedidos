#!/usr/bin/env python3
"""
ajustar_pedido.py
=================
Toma un pedido sugerido (ej. OHUHU.xlsx) y el archivo de clasificacion
ABC-XYZ del inventario, y produce un Excel con dos tablas:

  Hoja 1 - "Pedido + Clasificacion": el pedido original con dos columnas
           extras (ABC y XYZ) y una combinada (ABC-XYZ).

  Hoja 2 - "Recomendaciones": una accion sugerida por producto
           (ORDENAR / ORDENAR REDUCIDO / EVALUAR / NO ORDENAR / etc.)
           y la cantidad ajustada, basadas en la clasificacion y el
           historial de ventas.

Uso:
    python ajustar_pedido.py <pedido.xlsx>
    python ajustar_pedido.py OHUHU.xlsx -o salida.xlsx
    python ajustar_pedido.py OHUHU.xlsx -s "Filtrado (Pedido > 0)"
    python ajustar_pedido.py OHUHU.xlsx -q "Recomendacion Claude"
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


CLASIFICACION_DEFAULT = "Analisis-ABC-XYZ-Rentabilidad-Diseno-2025.xlsx"

# (accion, factor sobre la cantidad sugerida, descripcion)
ACCION_POR_CLASE: dict[str, tuple[str, float, str]] = {
    "AX": ("ORDENAR",          1.00, "Estrella - alta rentabilidad y demanda estable. Nunca debe faltar."),
    "AY": ("ORDENAR",          1.00, "Clave variable - rentable pero fluctua. Mantener stock de seguridad."),
    "AZ": ("ORDENAR CAUTELA",  0.75, "Alta esporadica - rentable pero irregular. Considerar pedido especial o estacional."),
    "BX": ("ORDENAR",          1.00, "Estable medio - reordenar regularmente en cantidades moderadas."),
    "BY": ("ORDENAR REDUCIDO", 0.75, "Medio variable - monitorear y ajustar pedidos periodicamente."),
    "BZ": ("EVALUAR",          0.50, "Medio irregular - evaluar si el costo de inventario lo justifica."),
    "CX": ("MINIMO",           0.30, "Bajo constante - stock minimo, pedidos pequenos frecuentes."),
    "CY": ("REDUCIR",          0.25, "Bajo variable - reducir inventario, pedir bajo demanda."),
    "CZ": ("NO ORDENAR",       0.00, "Muerto/Obsoleto - candidato a liquidacion o descontinuacion."),
}

# Colores de fondo por accion en la hoja de Recomendaciones (formato AARRGGBB sin alpha)
COLOR_POR_ACCION: dict[str, str] = {
    "ORDENAR":          "C6EFCE",
    "ORDENAR CAUTELA":  "FFEB9C",
    "ORDENAR REDUCIDO": "FFEB9C",
    "EVALUAR":          "FFD580",
    "MINIMO":           "FFD580",
    "REDUCIR":          "FFC7CE",
    "NO ORDENAR":       "FFC7CE",
    "REVISAR MANUAL":   "D9D9D9",
}

ORDEN_CLASES = ["AX", "AY", "AZ", "BX", "BY", "BZ", "CX", "CY", "CZ", "N/D"]


# --- carga ----------------------------------------------------------------

def cargar_clasificacion(ruta: Path) -> dict[str, dict]:
    """
    Lee la hoja 'Orden Original' del archivo de clasificacion y devuelve
    un diccionario indexado por codigo de producto.
    """
    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)
    try:
        return _leer_clasificacion(wb, ruta)
    finally:
        # Cerrar siempre: en modo read_only el archivo queda bloqueado en
        # Windows si se sale con error sin cerrarlo.
        wb.close()


def _leer_clasificacion(wb, ruta: Path) -> dict[str, dict]:
    if "Orden Original" not in wb.sheetnames:
        raise ValueError(
            f"El archivo {ruta.name} no tiene la hoja 'Orden Original'. "
            f"Hojas disponibles: {wb.sheetnames}"
        )
    ws = wb["Orden Original"]
    rows = ws.iter_rows(values_only=True)
    headers = list(next(rows))
    idx = {h: i for i, h in enumerate(headers)}

    # La columna de ventas anuales lleva el ano en el nombre ("Total Uds 2025",
    # "Total Uds 2026", ...). Se detecta sola; si hay varias se usa la mas reciente.
    anio, col_total = _detectar_columna_total(headers)
    if col_total is None:
        raise ValueError(
            "Faltan columnas en la clasificacion: ['Total Uds <ano>'] "
            f"(ej. 'Total Uds 2025'). Columnas disponibles: {headers}"
        )

    requeridos = ["Codigo", "Producto", "ABC", "XYZ", "ABC-XYZ",
                  "Promedio Mensual", "Meses c/Venta",
                  "Margen Bruto", "Interpretacion"]
    faltantes = [c for c in requeridos if c not in idx]
    if faltantes:
        raise ValueError(f"Faltan columnas en la clasificacion: {faltantes}")

    # "Costo Total" es opcional: si existe, permite estimar el costo unitario
    # y calcular el impacto en dinero de cada ajuste.
    i_costo = idx.get("Costo Total")

    clasif: dict[str, dict] = {}
    for row in rows:
        codigo = row[idx["Codigo"]]
        if codigo is None:
            continue
        total_uds = row[idx[col_total]] or 0
        costo_unitario = None
        if i_costo is not None and total_uds > 0:
            costo_total = row[i_costo] or 0
            if costo_total > 0:
                costo_unitario = costo_total / total_uds
        clasif[str(codigo).strip()] = {
            "producto":         row[idx["Producto"]],
            "abc":              row[idx["ABC"]] or "",
            "xyz":              row[idx["XYZ"]] or "",
            "abc_xyz":          row[idx["ABC-XYZ"]] or "",
            "total_uds":        total_uds,
            "promedio_mensual": row[idx["Promedio Mensual"]] or 0,
            "meses_venta":      row[idx["Meses c/Venta"]] or 0,
            "margen_bruto":     row[idx["Margen Bruto"]] or 0,
            "interpretacion":   row[idx["Interpretacion"]] or "",
            "costo_unitario":   costo_unitario,
            "anio":             anio,
        }
    return clasif


def _detectar_columna_total(headers) -> tuple[int | None, str | None]:
    """Busca columnas 'Total Uds <ano>' y devuelve (ano, nombre) de la mas reciente."""
    import re

    candidatas = []
    for h in headers:
        if h is None:
            continue
        m = re.fullmatch(r"\s*Total Uds\s+(\d{4})\s*", str(h))
        if m:
            candidatas.append((int(m.group(1)), h))
    if not candidatas:
        return None, None
    return max(candidatas)


def _es_fila_eliminada(codigo) -> bool:
    """Codigos tipo 'ELIMINADO_29389_43823' son productos dados de baja en el sistema."""
    return str(codigo).strip().upper().startswith("ELIMINADO")


def cargar_pedido(ruta: Path, hoja: str | None = None) -> tuple[list, list[dict], str]:
    """
    Lee el pedido sugerido. Devuelve (encabezados, filas como dicts, nombre_hoja).
    Se asume que la PRIMERA columna es el codigo de producto.
    """
    wb = openpyxl.load_workbook(ruta, data_only=True)
    if hoja is None:
        ws = wb.worksheets[0]
    else:
        if hoja not in wb.sheetnames:
            raise ValueError(
                f"La hoja '{hoja}' no existe en {ruta.name}. "
                f"Hojas disponibles: {wb.sheetnames}"
            )
        ws = wb[hoja]

    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        raise ValueError(f"La hoja '{ws.title}' esta vacia.")

    headers = list(rows[0])
    datos: list[dict] = []
    eliminados = 0
    for r in rows[1:]:
        codigo = r[0]
        if codigo is None or str(codigo).strip() == "":
            continue
        # Saltar filas resumen tipo "TOTAL" / "TOTALES" al pie del pedido.
        if str(codigo).strip().upper().startswith("TOTAL"):
            continue
        # Saltar productos dados de baja (codigo "ELIMINADO_...").
        if _es_fila_eliminada(codigo):
            eliminados += 1
            continue
        # Saltar filas sin nombre de producto (resumenes, separadores).
        producto = r[1] if len(r) > 1 else None
        if producto is None or str(producto).strip() == "":
            continue
        datos.append({h: v for h, v in zip(headers, r)})

    if eliminados:
        print(f"  Aviso: se omitieron {eliminados} fila(s) de productos ELIMINADO.")

    return headers, datos, ws.title


# --- logica de recomendacion ----------------------------------------------

def calcular_ajuste(item, clasif_info, columna_sugerido):
    """
    Determina la accion y la cantidad ajustada para un producto.
    Devuelve un dict con accion, cantidad_ajustada, justificacion, abc_xyz.
    """
    sugerido = item.get(columna_sugerido) or 0
    saldo_actual = item.get("Saldo Actual") or 0

    if clasif_info is None:
        return {
            "accion": "REVISAR MANUAL",
            "cantidad_ajustada": sugerido,
            "justificacion": (
                "Producto no encontrado en la clasificacion ABC-XYZ "
                "(probablemente nuevo o sin historial). Revisar manualmente."
            ),
            "abc_xyz": "N/D",
        }

    abc_xyz = clasif_info["abc_xyz"] or "N/D"
    accion, factor, descripcion = ACCION_POR_CLASE.get(
        abc_xyz, ("REVISAR MANUAL", 1.0, "Clasificacion desconocida")
    )

    cantidad = round(sugerido * factor)
    promedio_mensual = clasif_info["promedio_mensual"] or 0
    total_uds = clasif_info["total_uds"] or 0
    anio = clasif_info.get("anio") or 2025

    # --- Overrides basados en historial real ---

    # 1) Sin ventas en el ano: nunca ordenar (excepto si por algun motivo es A,
    #    porque podria ser reciente; ahi solo recomendamos revisar).
    if total_uds == 0:
        if abc_xyz.startswith("A"):
            return {
                "accion": "REVISAR MANUAL",
                "cantidad_ajustada": sugerido,
                "justificacion": (
                    f"Clase {abc_xyz} pero CERO ventas en {anio}. "
                    "Posible producto nuevo o de reciente ingreso, revisar."
                ),
                "abc_xyz": abc_xyz,
            }
        return {
            "accion": "NO ORDENAR",
            "cantidad_ajustada": 0,
            "justificacion": f"{descripcion} Sin ventas en {anio} - no reordenar.",
            "abc_xyz": abc_xyz,
        }

    # 2) Ya hay mas inventario del que se vendio en todo el ano (sobrestock),
    #    excepto AX que tiene demanda estable y se justifica un colchon.
    if saldo_actual >= total_uds and abc_xyz != "AX":
        return {
            "accion": "NO ORDENAR",
            "cantidad_ajustada": 0,
            "justificacion": (
                f"Sobrestock: saldo actual ({saldo_actual}) >= ventas anuales "
                f"{anio} ({total_uds:.0f}). Consumir inventario antes de reordenar."
            ),
            "abc_xyz": abc_xyz,
        }

    # 3) Para clases Z (esporadicas/irregulares), nunca pedir mas de
    #    ~2 meses de venta promedio: el pronostico estadistico tiende a
    #    inflar pedidos cuando los datos son volatiles.
    if abc_xyz.endswith("Z") and promedio_mensual > 0 and cantidad > promedio_mensual * 2:
        cantidad = max(round(promedio_mensual * 2), 0)
        descripcion = (
            f"{descripcion} Cantidad limitada a ~2 meses de venta promedio "
            f"(~{promedio_mensual:.1f}/mes) por volatilidad."
        )

    # 4) Defensa: nunca pedir cantidad negativa.
    cantidad = max(cantidad, 0)

    return {
        "accion": accion,
        "cantidad_ajustada": cantidad,
        "justificacion": descripcion,
        "abc_xyz": abc_xyz,
    }


# --- escritura del archivo de salida --------------------------------------

def _set_header(cell):
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill("solid", fgColor="1F4E78")
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _autosize(ws, max_width: int = 60):
    for col_idx in range(1, ws.max_column + 1):
        letter = get_column_letter(col_idx)
        largo = 0
        for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, values_only=True):
            v = row[0]
            if v is None:
                continue
            largo = max(largo, len(str(v)))
        ws.column_dimensions[letter].width = min(max(largo + 2, 10), max_width)


def construir_salida(headers_pedido, datos_pedido, clasificacion,
                     columna_sugerido, ruta_salida):
    wb = openpyxl.Workbook()

    # ---------------- Hoja 1: pedido + ABC/XYZ ----------------
    ws1 = wb.active
    ws1.title = "Pedido + Clasificacion"
    nuevos_headers = list(headers_pedido) + ["ABC", "XYZ", "ABC-XYZ"]
    ws1.append(nuevos_headers)

    codigo_col = headers_pedido[0]  # 1ra columna = codigo (formato fijo del pedido)
    for item in datos_pedido:
        codigo = str(item.get(codigo_col) or "").strip()
        clasif = clasificacion.get(codigo)
        fila = [item.get(h) for h in headers_pedido]
        if clasif:
            fila += [clasif["abc"], clasif["xyz"], clasif["abc_xyz"]]
        else:
            fila += ["N/D", "N/D", "N/D"]
        ws1.append(fila)

    for cell in ws1[1]:
        _set_header(cell)
    ws1.freeze_panes = "A2"

    # ---------------- Hoja 2: recomendaciones ----------------
    ws2 = wb.create_sheet("Recomendaciones")
    anio = next((c.get("anio") for c in clasificacion.values() if c.get("anio")), 2025)
    headers_rec = [
        "Codigo", "Producto", "ABC-XYZ", "Saldo Actual",
        f"Vtas {anio} (uds)", "Prom Mensual", "Meses c/Venta",
        "Sugerido Original", "Cantidad Ajustada", "Diferencia",
        "Accion", "Justificacion",
        # Columnas de impacto en dinero (al final para no mover las anteriores).
        "Costo Unit. Prom", "Costo Sugerido", "Costo Ajustado", "Ahorro",
    ]
    ws2.append(headers_rec)

    items_recomendados = []
    for item in datos_pedido:
        codigo = str(item.get(codigo_col) or "").strip()
        producto = item.get(headers_pedido[1]) if len(headers_pedido) > 1 else ""
        saldo = item.get("Saldo Actual") or 0
        sugerido = item.get(columna_sugerido) or 0
        clasif = clasificacion.get(codigo)
        rec = calcular_ajuste(item, clasif, columna_sugerido)
        costo_u = clasif.get("costo_unitario") if clasif else None
        if costo_u is not None:
            costo_sug = round((sugerido or 0) * costo_u, 2)
            costo_aj = round((rec["cantidad_ajustada"] or 0) * costo_u, 2)
            ahorro = round(costo_sug - costo_aj, 2)
        else:
            costo_sug = costo_aj = ahorro = None
        items_recomendados.append({
            # 4 decimales: hay productos de centavos (ej. hojas sueltas a $0.0338).
            "costo_unitario":   round(costo_u, 4) if costo_u is not None else None,
            "costo_sugerido":   costo_sug,
            "costo_ajustado":   costo_aj,
            "ahorro":           ahorro,
            "codigo":           codigo,
            "producto":         producto,
            "abc_xyz":          rec["abc_xyz"],
            "saldo":            saldo,
            "total_uds":        clasif["total_uds"] if clasif else 0,
            "promedio_mensual": clasif["promedio_mensual"] if clasif else 0,
            "meses_venta":      clasif["meses_venta"] if clasif else 0,
            "sugerido":         sugerido,
            "ajustada":         rec["cantidad_ajustada"],
            "diferencia":       rec["cantidad_ajustada"] - (sugerido or 0),
            "accion":           rec["accion"],
            "justificacion":    rec["justificacion"],
        })

    # Ordenar de mayor prioridad (AX) a menor (CZ); dentro de cada clase, mas pedido primero.
    items_recomendados.sort(
        key=lambda x: (
            ORDEN_CLASES.index(x["abc_xyz"]) if x["abc_xyz"] in ORDEN_CLASES else 99,
            -(x["ajustada"] or 0),
        )
    )

    for it in items_recomendados:
        ws2.append([
            it["codigo"], it["producto"], it["abc_xyz"], it["saldo"],
            it["total_uds"], round(it["promedio_mensual"], 1),
            it["meses_venta"],
            it["sugerido"], it["ajustada"], it["diferencia"],
            it["accion"], it["justificacion"],
            it["costo_unitario"], it["costo_sugerido"],
            it["costo_ajustado"], it["ahorro"],
        ])

    for cell in ws2[1]:
        _set_header(cell)
    ws2.freeze_panes = "A2"

    # Formato de moneda en las columnas de dinero.
    col_costo_unit = len(headers_rec) - 3
    for col in range(col_costo_unit, len(headers_rec) + 1):
        formato = '"$"#,##0.0000' if col == col_costo_unit else '"$"#,##0.00'
        for row_idx in range(2, len(items_recomendados) + 2):
            ws2.cell(row=row_idx, column=col).number_format = formato

    # Pintar fila segun accion
    for row_idx, it in enumerate(items_recomendados, start=2):
        color = COLOR_POR_ACCION.get(it["accion"], "FFFFFF")
        if color == "FFFFFF":
            continue
        fill = PatternFill("solid", fgColor=color)
        for col in range(1, len(headers_rec) + 1):
            ws2.cell(row=row_idx, column=col).fill = fill

    _autosize(ws1)
    _autosize(ws2)

    wb.save(ruta_salida)

    return items_recomendados


# --- main / cli -----------------------------------------------------------

def imprimir_resumen(items_recomendados):
    from collections import Counter

    contador = Counter(it["accion"] for it in items_recomendados)
    total_sug = sum((it["sugerido"] or 0) for it in items_recomendados)
    total_aj  = sum((it["ajustada"] or 0) for it in items_recomendados)

    # Filas con sugerido 0 que terminan en NO ORDENAR no son un "recorte" real.
    ya_en_cero = Counter(
        it["accion"] for it in items_recomendados if not (it["sugerido"] or 0)
    )
    con_pedido = sum(1 for it in items_recomendados if (it["sugerido"] or 0) > 0)

    print("\n=== RESUMEN ===")
    print(f"Productos en pedido: {len(items_recomendados)}  "
          f"(con cantidad sugerida > 0: {con_pedido})")
    print(f"Total unidades sugeridas (original): {total_sug:,}")
    print(f"Total unidades ajustadas:            {total_aj:,}")
    if total_sug > 0:
        delta = (total_aj - total_sug) / total_sug * 100
        print(f"Variacion:                           {delta:+.1f}%")

    con_costo = [it for it in items_recomendados if it.get("costo_unitario") is not None]
    if con_costo:
        costo_sug = sum(it["costo_sugerido"] for it in con_costo)
        costo_aj = sum(it["costo_ajustado"] for it in con_costo)
        print(f"Costo estimado sugerido (original):  ${costo_sug:,.2f}")
        print(f"Costo estimado ajustado:             ${costo_aj:,.2f}")
        print(f"Ahorro estimado:                     ${costo_sug - costo_aj:,.2f}")
        sin_costo = [it for it in items_recomendados
                     if it.get("costo_unitario") is None and (it["ajustada"] or 0) > 0]
        if sin_costo:
            uds = sum(it["ajustada"] for it in sin_costo)
            print(f"  (sin costo conocido: {len(sin_costo)} producto(s), "
                  f"{uds:,} uds ajustadas no incluidas en el monto)")

    print("\nDistribucion por accion:")
    for accion in ["ORDENAR", "ORDENAR CAUTELA", "ORDENAR REDUCIDO",
                   "EVALUAR", "MINIMO", "REDUCIR",
                   "NO ORDENAR", "REVISAR MANUAL"]:
        n = contador.get(accion, 0)
        if n:
            extra = ""
            if accion == "NO ORDENAR" and ya_en_cero.get(accion):
                extra = f"  ({ya_en_cero[accion]} ya tenian sugerido 0)"
            print(f"  {accion:18s} {n:4d}{extra}")


def main(argv=None):
    parser = argparse.ArgumentParser(
        description="Ajusta un pedido sugerido segun clasificacion ABC-XYZ.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python ajustar_pedido.py OHUHU.xlsx\n"
            "  python ajustar_pedido.py OHUHU.xlsx -o pedido_revisado.xlsx\n"
            "  python ajustar_pedido.py OHUHU.xlsx -s \"Filtrado (Pedido > 0)\"\n"
        ),
    )
    parser.add_argument(
        "pedido", nargs="+",
        help="Archivo(s) Excel con el pedido sugerido (ej. OHUHU.xlsx, o varios: ARISTO.xlsx KUM.xlsx).",
    )
    parser.add_argument(
        "-c", "--clasificacion",
        default=CLASIFICACION_DEFAULT,
        help=f"Archivo Excel con la clasificacion ABC-XYZ (default: {CLASIFICACION_DEFAULT}).",
    )
    parser.add_argument(
        "-o", "--salida", default=None,
        help="Archivo de salida (default: <pedido>_ajustado.xlsx en la misma carpeta).",
    )
    parser.add_argument(
        "-s", "--hoja", default=None,
        help="Hoja del pedido a procesar (default: la primera).",
    )
    parser.add_argument(
        "-q", "--columna-sugerido", default="Recomendacion Claude",
        help=(
            "Columna con la cantidad sugerida a ordenar (default: 'Recomendacion Claude'). "
            "Acepta el nombre con o sin acentos."
        ),
    )

    args = parser.parse_args(argv)

    pedidos = [Path(p) for p in args.pedido]
    if args.salida and len(pedidos) > 1:
        print("ERROR: -o/--salida solo se puede usar con un unico pedido.", file=sys.stderr)
        return 1

    faltan = [p for p in pedidos if not p.exists()]
    if faltan and len(pedidos) == 1:
        print(f"ERROR: no se encontro el archivo de pedido: {faltan[0]}", file=sys.stderr)
        return 1

    clasif_path = Path(args.clasificacion)
    if not clasif_path.is_absolute() and not clasif_path.exists():
        # Intentar relativo a la carpeta del pedido
        alt = pedidos[0].parent / clasif_path.name
        if alt.exists():
            clasif_path = alt
    if not clasif_path.exists():
        print(f"ERROR: no se encontro la clasificacion ABC-XYZ: {clasif_path}", file=sys.stderr)
        return 1

    print(f"Cargando clasificacion ABC-XYZ desde {clasif_path.name} ...")
    clasificacion = cargar_clasificacion(clasif_path)
    print(f"  -> {len(clasificacion):,} productos clasificados.")

    if len(pedidos) == 1:
        salida = Path(args.salida) if args.salida else None
        return procesar_pedido(pedidos[0], clasificacion, args.hoja,
                               args.columna_sugerido, salida)

    # Modo lote: procesar cada pedido; un error en uno no detiene a los demas.
    fallidos = []
    for p in pedidos:
        print(f"\n{'=' * 60}\n{p.name}\n{'=' * 60}")
        if not p.exists():
            print(f"ERROR: no se encontro el archivo de pedido: {p}", file=sys.stderr)
            fallidos.append(p.name)
            continue
        try:
            rc = procesar_pedido(p, clasificacion, args.hoja, args.columna_sugerido, None)
        except Exception as e:  # noqa: BLE001 - reportar y seguir con el siguiente
            print(f"ERROR procesando {p.name}: {e}", file=sys.stderr)
            rc = 1
        if rc != 0:
            fallidos.append(p.name)

    print(f"\n{'=' * 60}")
    print(f"Lote terminado: {len(pedidos) - len(fallidos)} de {len(pedidos)} pedido(s) OK.")
    if fallidos:
        print(f"Con error: {', '.join(fallidos)}", file=sys.stderr)
        return 1
    return 0


def procesar_pedido(pedido_path: Path, clasificacion, hoja, columna_sugerido,
                    salida: Path | None) -> int:
    """Procesa un pedido y genera su archivo ajustado. Devuelve 0 si todo OK."""
    if salida is None:
        salida = pedido_path.parent / f"{pedido_path.stem}_ajustado.xlsx"

    print(f"Cargando pedido sugerido desde {pedido_path.name} ...")
    headers, datos, hoja_usada = cargar_pedido(pedido_path, hoja)
    print(f"  -> {len(datos):,} productos en hoja '{hoja_usada}'.")

    # Resolver columna sugerida (acepta variaciones con/sin acentos).
    columna = _resolver_columna(headers, columna_sugerido)
    if columna is None:
        print(
            f"ERROR: la columna '{columna_sugerido}' no existe en la hoja.\n"
            f"Columnas disponibles: {headers}",
            file=sys.stderr,
        )
        return 1
    print(f"Cantidad sugerida tomada de la columna: '{columna}'.")

    print(f"Generando archivo: {salida.name} ...")
    items = construir_salida(headers, datos, clasificacion, columna, salida)

    sin_clasif = sum(1 for it in items if it["abc_xyz"] == "N/D")
    if sin_clasif:
        print(f"Aviso: {sin_clasif} producto(s) sin clasificacion ABC-XYZ (marcados como REVISAR MANUAL).")

    imprimir_resumen(items)
    print(f"\nArchivo generado: {salida}")
    return 0


def _resolver_columna(headers, nombre_buscado: str) -> str | None:
    """Empareja nombres de columna ignorando acentos y mayusculas."""
    import unicodedata

    def norm(s: str) -> str:
        s = unicodedata.normalize("NFKD", str(s))
        return "".join(c for c in s if not unicodedata.combining(c)).strip().lower()

    objetivo = norm(nombre_buscado)
    for h in headers:
        if h is not None and norm(h) == objetivo:
            return h
    return None


if __name__ == "__main__":
    sys.exit(main())
