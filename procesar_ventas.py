import pandas as pd
import os

ENTRADA = "data/datos_ventas.xlsx"
SALIDA = "output/resumen_ventas.xlsx"


def procesar():
    # 1. Cargar datos
    df = pd.read_excel(ENTRADA)
    print(f"Filas cargadas: {len(df)}")

    # 2. Imputar Total_Venta cuando esta vacio
    nan_antes = df["Total_Venta"].isna().sum()
    df["Total_Venta"] = df["Total_Venta"].fillna(df["Cantidad"] * df["Precio_Unitario"])
    print(f"Valores imputados en Total_Venta: {nan_antes}")

    # 3. Convertir Fecha a datetime
    df["Fecha"] = pd.to_datetime(df["Fecha"])

    # 4. Filtrar solo ventas de 2023
    df = df[df["Fecha"].dt.year == 2023].copy()
    print(f"Ventas de 2023: {len(df)}")

    # 5. Columna Mes
    df["Mes"] = df["Fecha"].dt.month

    # 6. Agrupaciones
    ventas_vendedor = (
        df.groupby("Vendedor", as_index=False)["Total_Venta"]
          .sum()
          .sort_values("Total_Venta", ascending=False)
    )

    ventas_mes = (
        df.groupby("Mes", as_index=False)["Total_Venta"]
          .sum()
    )

    ventas_vendedor_mes = (
        df.groupby(["Vendedor", "Mes"], as_index=False)["Total_Venta"]
          .sum()
    )

    # 7. Exportar
    os.makedirs("output", exist_ok=True)
    with pd.ExcelWriter(SALIDA, engine="openpyxl") as writer:
        ventas_vendedor.to_excel(writer, sheet_name="Resumen_Ventas", index=False)
        ventas_mes.to_excel(writer, sheet_name="Ventas_Mensuales", index=False)
        ventas_vendedor_mes.to_excel(writer, sheet_name="Vendedor_x_Mes", index=False)

    print(f"\nArchivo generado en: {SALIDA}")
    print("\nTotal por vendedor:")
    print(ventas_vendedor.to_string(index=False))


if __name__ == "__main__":
    procesar()
