"""Genera datos/ventas_desordenadas.xlsx con errores reales:
duplicados, precios como texto con $, fechas en varios formatos,
productos con mayúsculas inconsistentes y cantidades vacías.
Úsalo para regenerar el ejemplo:  python3 generar_ejemplo.py
"""
import random
from pathlib import Path

import pandas as pd

random.seed(42)

productos = [
    ("café molido 500g", 145.00),
    ("TÉ VERDE 100g", 89.50),
    ("chocolate  tableta ", 62.00),
    ("Miel 250ml", 128.00),
    ("Galletas avena", 45.00),
    ("Café Molido 500g", 145.00),  # duplicado con otro formato
]

filas = []
for i in range(42):
    producto, precio = random.choice(productos)
    # Precio a veces como texto con $ y coma
    precio_txt = f"${precio:,.2f}" if random.random() < 0.4 else precio
    # Fecha en varios formatos
    dia = random.randint(1, 28)
    mes = random.choice([1, 2, 3])
    fecha = random.choice([
        f"2026-{mes:02d}-{dia:02d}",
        f"{dia:02d}/{mes:02d}/2026",
        f"2026/{mes:02d}/{dia:02d}",
    ])
    filas.append({
        " Fecha ": fecha,
        "Producto": producto,
        "Cantidad": random.choice([1, 2, 3, None, 2, 1]),
        "Precio Unitario": precio_txt,
    })

# Añadir duplicados exactos y una fila sin producto
filas.append(filas[3])
filas.append(filas[10])
filas.append({" Fecha ": "2026-02-10", "Producto": " ", "Cantidad": 1, "Precio Unitario": 50})

df = pd.DataFrame(filas)
salida = Path(__file__).parent / "datos" / "ventas_desordenadas.xlsx"
salida.parent.mkdir(parents=True, exist_ok=True)
df.to_excel(salida, index=False)
print(f"Ejemplo generado: {salida} ({len(df)} filas, con errores intencionales)")
