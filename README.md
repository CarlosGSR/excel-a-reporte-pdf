# Excel a Reporte PDF — Automatización de ventas en Python

Script en Python que toma un Excel de ventas desordenado, lo limpia y genera un reporte en PDF + un Excel limpio, listo para entregar al cliente.

**Ideal para:** negocios pequeños que hoy hacen sus reportes a mano en Excel.

## El problema

El cliente tenía un Excel con:
- Precios como texto: `$1,200.50`, `1200,50`
- Fechas en 3 formatos distintos
- Productos escritos de varias formas: `café molido 500g`, `Café Molido 500g`, `café  molido 500g`
- Filas duplicadas y cantidades vacías

Hacer el reporte a mano le tomaba ~4 horas a la semana.

## La solución

Un solo comando:

```bash
pip install -r requirements.txt
python3 reporte.py
```

Genera:

| Archivo | Qué contiene |
|---|---|
| `salida/ventas_limpias.xlsx` | Ventas limpias + hoja Resumen + Top productos |
| `salida/reporte_ventas.pdf` | Total de ventas, ticket promedio, top 5 productos, ventas por mes y muestra de ventas |

Con el ejemplo incluido, el resultado es:

- **Filas originales:** 45
- **Filas eliminadas:** 5 (duplicados / sin datos)
- **Filas limpias:** 40
- **Total de ventas:** $7,967.50 MXN
- **Ticket promedio:** $199.19 MXN

## Qué limpia el script

- Normaliza nombres de columnas y de productos
- Convierte precios con `$` y comas a número
- Unifica fechas en varios formatos
- Cantidad vacía → 1
- Elimina duplicados exactos y filas sin producto, fecha o precio
- Calcula `Total = Cantidad × Precio Unitario`

## Personalizarlo para un cliente

```bash
python3 reporte.py --entrada mi_archivo.xlsx --pdf salida/mi_reporte.pdf --excel salida/mi_excel.xlsx
```

El Excel del cliente solo necesita estas columnas (el nombre exacto no importa tanto, el script las normaliza): `Fecha`, `Producto`, `Cantidad`, `Precio Unitario`.

## Archivos

- `reporte.py` — el script principal
- `generar_ejemplo.py` — regenera el Excel de ejemplo con errores intencionales
- `datos/ventas_desordenadas.xlsx` — ejemplo de entrada
- `salida/reporte_ventas.pdf` — ejemplo de salida
- `salida/ventas_limpias.xlsx` — ejemplo de salida

## Tecnologías

Python · pandas · openpyxl · reportlab

Hecho con asistencia de IA (Codex / Antigravity) y revisado y probado a mano.
