"""
Prueba Técnica - Procesamiento de Ventas
=========================================

Script que carga un archivo Excel de ventas, procesa los datos y genera
un archivo de salida con dos hojas de resumen.

Pasos:
    1. Carga de datos en un DataFrame de pandas.
    2. Manejo de valores faltantes en Total_Venta (Cantidad * Precio_Unitario).
    3. Conversión de la columna Fecha a datetime.
    4. Filtrado de ventas del año 2023.
    5. Extracción del mes en formato numérico.
    6. Cálculo de totales por vendedor y por mes.
    7. Generación de archivo Excel con dos hojas.

Autor: Santiago González Medina
Fecha: 2026-08-20
"""

from pathlib import Path
import pandas as pd


# ---------------------------------------------------------------------------
# Configuración de rutas
# ---------------------------------------------------------------------------
ARCHIVO_ENTRADA = Path("data/datos_ventas.xlsx")
ARCHIVO_SALIDA = Path("output/resumen_ventas.xlsx")


def cargar_datos(ruta: Path) -> pd.DataFrame:
    """Carga el archivo Excel de ventas en un DataFrame."""
    df = pd.read_excel(ruta)
    print(f"[OK] Archivo cargado: {ruta.name} | {df.shape[0]} filas, {df.shape[1]} columnas")
    return df


def imputar_total_venta(df: pd.DataFrame) -> pd.DataFrame:
    """
    Completa los valores faltantes de Total_Venta usando Cantidad * Precio_Unitario.
    No sobrescribe valores existentes.
    """
    faltantes_antes = df["Total_Venta"].isna().sum()
    df["Total_Venta"] = df["Total_Venta"].fillna(df["Cantidad"] * df["Precio_Unitario"])
    faltantes_despues = df["Total_Venta"].isna().sum()
    print(f"[OK] Total_Venta imputados: {faltantes_antes - faltantes_despues} "
          f"(faltantes antes: {faltantes_antes}, despues: {faltantes_despues})")
    return df


def convertir_fecha(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte la columna Fecha a tipo datetime."""
    df["Fecha"] = pd.to_datetime(df["Fecha"], errors="coerce")
    invalidas = df["Fecha"].isna().sum()
    if invalidas:
        print(f"[!] {invalidas} fechas no pudieron convertirse.")
    else:
        print("[OK] Columna Fecha convertida a datetime.")
    return df


def filtrar_anio(df: pd.DataFrame, anio: int) -> pd.DataFrame:
    """Filtra el DataFrame por el año indicado."""
    df_filtrado = df[df["Fecha"].dt.year == anio].copy()
    print(f"[OK] Ventas del {anio}: {len(df_filtrado)} registros.")
    return df_filtrado


def agregar_columna_mes(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega la columna Mes (numerico) extraida de Fecha."""
    df["Mes"] = df["Fecha"].dt.month
    print("[OK] Columna Mes agregada.")
    return df


def resumen_por_vendedor(df: pd.DataFrame) -> pd.DataFrame:
    """Total de ventas agregado por vendedor, ordenado de mayor a menor."""
    resumen = (
        df.groupby("Vendedor", as_index=False)["Total_Venta"]
          .sum()
          .sort_values("Total_Venta", ascending=False)
          .reset_index(drop=True)
    )
    resumen.rename(columns={"Total_Venta": "Total_Ventas"}, inplace=True)
    return resumen


def resumen_por_vendedor_mes(df: pd.DataFrame) -> pd.DataFrame:
    """Total de ventas por vendedor y por mes (formato largo)."""
    resumen = (
        df.groupby(["Vendedor", "Mes"], as_index=False)["Total_Venta"]
          .sum()
          .sort_values(["Vendedor", "Mes"])
          .reset_index(drop=True)
    )
    resumen.rename(columns={"Total_Venta": "Total_Ventas"}, inplace=True)
    return resumen


def resumen_mensual(df: pd.DataFrame) -> pd.DataFrame:
    """Total de ventas agregado por mes."""
    resumen = (
        df.groupby("Mes", as_index=False)["Total_Venta"]
          .sum()
          .sort_values("Mes")
          .reset_index(drop=True)
    )
    resumen.rename(columns={"Total_Venta": "Total_Ventas"}, inplace=True)
    return resumen


def exportar_excel(resumen_vendedor: pd.DataFrame,
                   resumen_mes: pd.DataFrame,
                   resumen_vendedor_mes: pd.DataFrame,
                   ruta: Path) -> None:
    """Escribe las hojas de resumen en un archivo Excel."""
    ruta.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(ruta, engine="openpyxl") as writer:
        resumen_vendedor.to_excel(writer, sheet_name="Resumen_Ventas", index=False)
        resumen_mes.to_excel(writer, sheet_name="Ventas_Mensuales", index=False)
        # Hoja adicional de valor: cruce vendedor x mes
        resumen_vendedor_mes.to_excel(writer, sheet_name="Vendedor_x_Mes", index=False)
    print(f"[OK] Archivo generado: {ruta}")


def main() -> None:
    print("=" * 60)
    print("PROCESAMIENTO DE VENTAS")
    print("=" * 60)

    df = cargar_datos(ARCHIVO_ENTRADA)
    df = imputar_total_venta(df)
    df = convertir_fecha(df)
    df = filtrar_anio(df, 2023)
    df = agregar_columna_mes(df)

    r_vendedor = resumen_por_vendedor(df)
    r_mes = resumen_mensual(df)
    r_vendedor_mes = resumen_por_vendedor_mes(df)

    print("\n--- Total por Vendedor ---")
    print(r_vendedor.to_string(index=False))
    print("\n--- Total por Mes ---")
    print(r_mes.to_string(index=False))

    exportar_excel(r_vendedor, r_mes, r_vendedor_mes, ARCHIVO_SALIDA)
    print("=" * 60)
    print("PROCESO COMPLETADO")
    print("=" * 60)


if __name__ == "__main__":
    main()
