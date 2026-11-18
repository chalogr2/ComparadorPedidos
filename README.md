# Comparador de Pedidos

A command line tool I built and use in my own business, Diseño, an art
supplies store in El Salvador. Before I send a purchase order to a supplier,
I run it through this tool to check it against the ABC-XYZ classification of
our inventory, so we stop reordering products that don't sell and keep the
ones that do well stocked.

## The problem

Every supplier order starts as a "pedido sugerido": an Excel file with
one row per product and a few forecasts of how much to buy (moving
averages, exponential smoothing, seasonal, a manual estimate, and a final
recommended quantity).

Those forecasts only look at the sales history of each product on its
own. They don't know that a product is barely profitable, sells in random
bursts, or already has a year's worth of stock sitting in the warehouse. A
good chunk of our inventory value was tied up in exactly that kind of
product, so I wanted a second filter before any money goes out.

## What it does

It reads two files:

1. The suggested order for one supplier (e.g. `FABRIANO.xlsx`).
2. The ABC-XYZ analysis of the whole inventory, where **ABC** ranks products
   by gross margin contribution and **XYZ** by how stable their monthly
   demand is.

And produces an Excel file with two sheets:

- **Pedido + Clasificacion**: the original order with the ABC, XYZ and
  ABC-XYZ class added to every row.
- **Recomendaciones**: an action and an adjusted quantity for every product,
  sorted by priority and color coded, plus the estimated cost of the
  original vs. the adjusted order and how much is saved per line.

It also prints a short summary in the terminal: units and dollars before
and after, and how many products fall under each action.

## How the adjustment works

Each class gets a factor applied to the suggested quantity:

| Class | Action | Factor | Why |
| ----- | ------ | ------ | --- |
| AX | ORDENAR | 100% | Top sellers with steady demand. Never run out. |
| AY | ORDENAR | 100% | Profitable but variable. Keep safety stock. |
| AZ | ORDENAR CAUTELA | 75% | Profitable but sporadic. |
| BX | ORDENAR | 100% | Mid range, steady. |
| BY | ORDENAR REDUCIDO | 75% | Mid range, variable. |
| BZ | EVALUAR | 50% | Mid range, irregular. Is the stock worth it? |
| CX | MINIMO | 30% | Low value but constant. Small, frequent orders. |
| CY | REDUCIR | 25% | Low value and erratic. |
| CZ | NO ORDENAR | 0% | Dead stock. Candidate for clearance. |

On top of that, a few rules based on actual sales:

1. **No sales during the year**: don't order. Class A products are flagged
   for manual review instead, since they might be new.
2. **Overstock** (current stock is at least a full year of sales): don't
   order, except AX products, where the extra cushion is justified.
3. **Z classes are capped** at about two months of average sales, because
   forecasts tend to overshoot on irregular demand.
4. **Products not found in the classification** keep their original
   quantity and are flagged for manual review.

## Usage

Requires Python 3.10+ and openpyxl:

```bash
pip install -r requirements.txt
```

Put the suggested order and the classification file in the same folder and
run:

```bash
python ajustar_pedido.py FABRIANO.xlsx
```

This creates `FABRIANO_ajustado.xlsx` next to the original. You can also
process several suppliers in one go:

```bash
python ajustar_pedido.py ARISTO.xlsx CANSON.xlsx FABRIANO.xlsx
```

### Options

| Flag | What it does |
| ---- | ------------ |
| `-c, --clasificacion` | ABC-XYZ classification file. Defaults to `Analisis-ABC-XYZ-Rentabilidad-Diseno-2025.xlsx`. |
| `-o, --salida` | Output file name (only when processing a single order). |
| `-s, --hoja` | Sheet to read from the order. Defaults to the first one. |
| `-q, --columna-sugerido` | Column holding the suggested quantity. Defaults to `Recomendación Claude`. Accents and case don't matter. |

## Input format

- **Order file**: the product code goes in the first column and the product
  name in the second. The rest of the columns are matched by name
  (`Saldo Actual`, `Recomendación Claude`, ...), so new columns can be
  added without breaking anything. `TOTAL` rows and products marked as
  `ELIMINADO_...` are skipped.
- **Classification file**: needs an `Orden Original` sheet with the usual
  columns (`Codigo`, `ABC`, `XYZ`, `ABC-XYZ`, `Promedio Mensual`,
  `Total Uds <year>`, `Costo Total`, ...). The year is picked up from the
  column name, so a 2026 analysis works the same as the 2025 one.

Unit cost is estimated as total cost / units sold for the year. Products
without sales don't have a cost, so they're left out of the dollar totals
(the summary tells you how many).

The business data itself (orders, inventory, sales) is not included in this
repository.

## Tests

The tests build small fake Excel files, so they run without any real data:

```bash
python -m unittest discover -s tests -v
```

## License

Copyright (c) 2026 Fernando González Ruiz. All rights reserved. See
[LICENSE](LICENSE).
