"""
Excel a Reporte PDF — Proyecto 1 de portafolio freelance
=========================================================
Lee un Excel de ventas desordenado, lo limpia y genera:
  1. Un Excel limpio (salida/ventas_limpias.xlsx)
  2. Un reporte en PDF (salida/reporte_ventas.pdf)

Uso:
    python3 reporte.py
    python3 reporte.py --entrada datos/ventas_desordenadas.xlsx
    python3 reporte.py --entrada mi_archivo.xlsx --pdf salida/mi_reporte.pdf

Requisitos: pandas, openpyxl, reportlab
    pip install -r requirements.txt
"""

import argparse
import re
from datetime import date
from pathlib import Path

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import cm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle


# ---------- Limpieza ----------

def limpiar_precio(valor):
    """'$1,200.50' / '1200,50' / 1200 -> 1200.5. Devuelve None si no se puede leer."""
    if pd.isna(valor):
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).strip().replace("$", "").replace(" ", "")
    # Si usa coma como decimal y no hay punto: 1200,50 -> 1200.50
    if "," in texto and "." not in texto:
        texto = texto.replace(",", ".")
    else:
        texto = texto.replace(",", "")
    try:
        return float(texto)
    except ValueError:
        return None


def limpiar_fecha(valor):
    """Acepta varios formatos: 2026-01-05, 05/01/2026, 2026/01/05, etc.
    Prueba primero los formatos ISO (año-mes-día) para no confundir
    mes y día, y luego día/mes/año.
    """
    if pd.isna(valor):
        return pd.NaT
    texto = str(valor).strip()
    for fmt in ("%Y-%m-%d", "%Y/%m/%d", "%d/%m/%Y", "%d-%m-%Y"):
        try:
            return pd.to_datetime(texto, format=fmt)
        except (ValueError, TypeError):
            continue
    return pd.to_datetime(texto, errors="coerce", dayfirst=True)


def limpiar_ventas(df: pd.DataFrame) -> pd.DataFrame:
    """Normaliza y limpia el DataFrame de ventas."""
    # Normalizar nombres de columnas: ' Precio Unitario ' -> 'precio_unitario'
    df = df.copy()
    antes_total = len(df)
    df.columns = [
        re.sub(r"[^a-z0-9]+", "_", str(c).strip().lower()).strip("_")
        for c in df.columns
    ]

    # Mapear nombres habituales a los esperados
    renombrar = {
        "fecha_venta": "fecha",
        "producto_nombre": "producto",
        "precio": "precio_unitario",
        "precio_venta": "precio_unitario",
        "cantidad_vendida": "cantidad",
    }
    df = df.rename(columns={k: v for k, v in renombrar.items() if k in df.columns})

    requeridas = {"fecha", "producto", "cantidad", "precio_unitario"}
    faltantes = requeridas - set(df.columns)
    if faltantes:
        raise ValueError(f"Faltan columnas en el Excel: {sorted(faltantes)}. "
                         f"Columnas encontradas: {list(df.columns)}")

    # Producto: quitar espacios extra y unificar mayúsculas
    df["producto"] = (
        df["producto"].astype(str).str.strip()
        .str.replace(r"\s+", " ", regex=True)
        .str.title()
        .str.replace(r"(\d+)G\b", r"\1g", regex=True)
        .str.replace(r"(\d+)Ml\b", r"\1ml", regex=True)
    )
    df = df[~df["producto"].str.lower().isin(["nan", "none", ""])]

    # Fecha y precio
    df["fecha"] = df["fecha"].apply(limpiar_fecha)
    df["precio_unitario"] = df["precio_unitario"].apply(limpiar_precio)

    # Cantidad: vacía -> 1, texto -> número
    df["cantidad"] = pd.to_numeric(df["cantidad"], errors="coerce").fillna(1).astype(int)

    # Eliminar filas sin datos clave y duplicados exactos
    df = df.dropna(subset=["fecha", "precio_unitario"])
    df = df[df["cantidad"] > 0]
    df = df[df["precio_unitario"] > 0]
    df = df.drop_duplicates()
    eliminadas = antes_total - len(df)

    # Total por fila
    df["total"] = (df["cantidad"] * df["precio_unitario"]).round(2)
    df = df.sort_values("fecha").reset_index(drop=True)

    df.attrs["filas_eliminadas"] = eliminadas
    df.attrs["filas_originales"] = antes_total
    return df


def resumen(df: pd.DataFrame) -> dict:
    """Calcula los números clave del reporte."""
    top = (
        df.groupby("producto")["total"].sum()
        .sort_values(ascending=False).head(5)
    )
    por_mes = (
        df.assign(mes=df["fecha"].dt.strftime("%Y-%m"))
        .groupby("mes")["total"].sum()
    )
    return {
        "total_ventas": float(df["total"].sum()),
        "num_ventas": int(len(df)),
        "ticket_promedio": float(df["total"].mean()) if len(df) else 0.0,
        "unidades": int(df["cantidad"].sum()),
        "top_productos": top,
        "por_mes": por_mes,
    }


# ---------- Salidas ----------

def guardar_excel_limpio(df: pd.DataFrame, res: dict, ruta: Path):
    """Guarda el Excel limpio con hojas 'Ventas limpias' y 'Resumen'."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    resumen_df = pd.DataFrame([
        {"Concepto": "Total de ventas (MXN)", "Valor": round(res["total_ventas"], 2)},
        {"Concepto": "Número de ventas", "Valor": res["num_ventas"]},
        {"Concepto": "Ticket promedio (MXN)", "Valor": round(res["ticket_promedio"], 2)},
        {"Concepto": "Unidades vendidas", "Valor": res["unidades"]},
        {"Concepto": "Filas originales", "Valor": df.attrs.get("filas_originales", "")},
        {"Concepto": "Filas eliminadas (duplicados / sin datos)", "Valor": df.attrs.get("filas_eliminadas", "")},
    ])
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Ventas limpias", index=False)
        resumen_df.to_excel(writer, sheet_name="Resumen", index=False)
        res["top_productos"].reset_index().rename(
            columns={"total": "Total (MXN)"}
        ).to_excel(writer, sheet_name="Top productos", index=False)


def guardar_pdf(df: pd.DataFrame, res: dict, ruta: Path):
    """Genera el reporte PDF con KPIs, top productos y ventas por mes."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    estilos = getSampleStyleSheet()
    elementos = [
        Paragraph("Reporte de Ventas", estilos["Title"]),
        Paragraph(
            f"Generado el {date.today().strftime('%d/%m/%Y')} &nbsp;|&nbsp; "
            f"Periodo: {df['fecha'].min().date()} a {df['fecha'].max().date()}",
            estilos["Normal"],
        ),
        Spacer(1, 0.6 * cm),
    ]

    # KPIs
    kpis = [
        ["Total de ventas", "Núm. de ventas", "Ticket promedio", "Unidades"],
        [
            f"${res['total_ventas']:,.2f} MXN",
            str(res["num_ventas"]),
            f"${res['ticket_promedio']:,.2f} MXN",
            str(res["unidades"]),
        ],
    ]
    tabla_kpi = Table(kpis, colWidths=[4.2 * cm, 3.5 * cm, 4.2 * cm, 3.0 * cm])
    tabla_kpi.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#dbeafe")),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    elementos += [tabla_kpi, Spacer(1, 0.6 * cm)]

    def tabla_con_titulo(titulo, encabezados, filas):
        datos = [encabezados] + filas
        t = Table(datos, colWidths=[7.5 * cm, 7.5 * cm])
        t.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        return [Paragraph(titulo, estilos["Heading2"]), t, Spacer(1, 0.4 * cm)]

    elementos += tabla_con_titulo(
        "Top 5 productos",
        ["Producto", "Total (MXN)"],
        [[p, f"${v:,.2f}"] for p, v in res["top_productos"].items()],
    )
    elementos += tabla_con_titulo(
        "Ventas por mes",
        ["Mes", "Total (MXN)"],
        [[m, f"${v:,.2f}"] for m, v in res["por_mes"].items()],
    )

    # Muestra de ventas
    muestra = df.head(15)[["fecha", "producto", "cantidad", "precio_unitario", "total"]].copy()
    muestra["fecha"] = muestra["fecha"].dt.strftime("%d/%m/%Y")
    filas = [[str(r[0]), r[1], str(r[2]), f"${r[3]:,.2f}", f"${r[4]:,.2f}"]
             for r in muestra.itertuples(index=False)]
    t = Table(
        [["Fecha", "Producto", "Cant.", "Precio", "Total"]] + filas,
        colWidths=[2.5 * cm, 5.5 * cm, 1.8 * cm, 2.6 * cm, 2.6 * cm],
    )
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f1f5f9")]),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]))
    elementos += [Paragraph("Primeras 15 ventas (muestra)", estilos["Heading2"]), t]

    SimpleDocTemplate(str(ruta), pagesize=A4,
                      topMargin=1.5 * cm, bottomMargin=1.5 * cm,
                      leftMargin=1.5 * cm, rightMargin=1.5 * cm).build(elementos)


# ---------- Main ----------

def main():
    parser = argparse.ArgumentParser(description="Limpia un Excel de ventas y genera reporte PDF + Excel limpio.")
    parser.add_argument("--entrada", default="datos/ventas_desordenadas.xlsx")
    parser.add_argument("--pdf", default="salida/reporte_ventas.pdf")
    parser.add_argument("--excel", default="salida/ventas_limpias.xlsx")
    args = parser.parse_args()

    base = Path(__file__).parent
    entrada = base / args.entrada if not Path(args.entrada).is_absolute() else Path(args.entrada)

    print(f"Leyendo: {entrada}")
    df_crudo = pd.read_excel(entrada)
    df = limpiar_ventas(df_crudo)
    res = resumen(df)

    guardar_excel_limpio(df, res, base / args.excel)
    guardar_pdf(df, res, base / args.pdf)

    print(f"Filas originales: {df.attrs['filas_originales']}  |  "
          f"Filas eliminadas: {df.attrs['filas_eliminadas']}  |  "
          f"Filas limpias: {len(df)}")
    print(f"Total de ventas: ${res['total_ventas']:,.2f} MXN  |  "
          f"Ventas: {res['num_ventas']}  |  "
          f"Ticket promedio: ${res['ticket_promedio']:,.2f} MXN")
    print(f"Excel limpio: {args.excel}")
    print(f"Reporte PDF: {args.pdf}")


if __name__ == "__main__":
    main()
