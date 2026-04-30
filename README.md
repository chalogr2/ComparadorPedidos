# Comparador de Pedidos

CLI que ajusta un **pedido sugerido** (ej. `OHUHU.xlsx`) usando la
**clasificacion ABC-XYZ** del inventario y produce un Excel con dos tablas:

1. **Pedido + Clasificacion** — el pedido original con tres columnas extra
   (`ABC`, `XYZ`, `ABC-XYZ`).
2. **Recomendaciones** — accion sugerida y cantidad ajustada por producto,
   ordenadas de mayor a menor prioridad y coloreadas por accion.

## Requisitos

- Python 3.10+
- `openpyxl` (ya instalado)

## Uso basico

```bash
python ajustar_pedido.py OHUHU.xlsx
```

Genera `OHUHU_ajustado.xlsx` en la misma carpeta.

## Opciones

| Flag | Descripcion |
| ---- | ----------- |
| `-c, --clasificacion` | Archivo Excel con la clasificacion ABC-XYZ. Default: `Analisis-ABC-XYZ-Rentabilidad-Diseno-2025.xlsx`. |
| `-o, --salida` | Archivo de salida. Default: `<pedido>_ajustado.xlsx`. |
| `-s, --hoja` | Nombre de la hoja del pedido a procesar. Default: la primera. |
| `-q, --columna-sugerido` | Columna con la cantidad sugerida. Default: `Recomendacion Claude` (acepta con/sin acentos). |

## Ejemplos

```bash
# Pedido + clasificacion + salida especifica
python ajustar_pedido.py OHUHU.xlsx -o pedido_revisado.xlsx

# Procesar otra hoja del archivo OHUHU
python ajustar_pedido.py OHUHU.xlsx -s "Filtrado (Pedido > 0)"

# Usar otra columna como base del pedido
python ajustar_pedido.py OHUHU.xlsx -q "Estacional"
```

## Reglas de recomendacion

Sobre la cantidad sugerida del pedido se aplica un **factor por clase
ABC-XYZ**:

| Clase | Accion | Factor | Razon |
| ----- | ------ | ------ | ----- |
| AX | ORDENAR | 100% | Estrella, alta prioridad. |
| AY | ORDENAR | 100% | Clave variable, mantener stock de seguridad. |
| AZ | ORDENAR CAUTELA | 75% | Alta esporadica, fluctua. |
| BX | ORDENAR | 100% | Estable medio. |
| BY | ORDENAR REDUCIDO | 75% | Medio variable. |
| BZ | EVALUAR | 50% | Medio irregular. |
| CX | MINIMO | 30% | Bajo constante. |
| CY | REDUCIR | 25% | Bajo y erratico. |
| CZ | NO ORDENAR | 0% | Muerto/Obsoleto. |

Despues se aplican **overrides** que dependen del historial real de ventas:

1. **Cero ventas en 2025** → `NO ORDENAR` (excepto clases A, que se marcan
   `REVISAR MANUAL` por si son productos nuevos).
2. **Sobrestock** (`Saldo Actual >= Total Uds 2025`) → `NO ORDENAR`,
   excepto AX donde se justifica el colchon de inventario.
3. **Clases Z volatiles**: la cantidad final se topa a ~2 meses de venta
   promedio para evitar que el pronostico estadistico infle pedidos.
4. **Producto sin clasificacion** → `REVISAR MANUAL`, manteniendo la
   cantidad original sugerida.

## Notas

- El emparejamiento de productos se hace por la **primera columna** del
  pedido (codigo). El programa asume que ese formato no cambia.
- El programa salta filas resumen tipo `TOTAL` al pie del pedido.
