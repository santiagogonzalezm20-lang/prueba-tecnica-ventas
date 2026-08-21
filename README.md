# Prueba tecnica - Procesamiento de ventas

Script en Python que lee un archivo Excel con datos de ventas, hace la limpieza y las agrupaciones pedidas, y genera un archivo de salida con los resumenes.

## Estructura

```
data/                  archivo de entrada
output/                archivo generado
procesar_ventas.py     script principal
formato_excel.py       aplica formato al Excel de salida
requirements.txt       dependencias
```

## Como ejecutar

```bash
pip install -r requirements.txt
python procesar_ventas.py
python formato_excel.py
```

## Que hace el script

1. Carga `data/datos_ventas.xlsx` con pandas.
2. Rellena los valores faltantes de `Total_Venta` con `Cantidad * Precio_Unitario`.
3. Convierte la columna `Fecha` a datetime.
4. Filtra los registros del anio 2023.
5. Agrega una columna `Mes` con el mes numerico.
6. Calcula el total de ventas por vendedor y por mes.
7. Exporta el resultado a `output/resumen_ventas.xlsx`.

## Hojas del archivo de salida

- `Resumen_Ventas`: total por vendedor.
- `Ventas_Mensuales`: total por mes.
- `Vendedor_x_Mes`: cruce vendedor x mes.
