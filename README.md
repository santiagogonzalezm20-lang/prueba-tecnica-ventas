# Prueba Técnica – Procesamiento de Ventas

Solución a la prueba técnica de desarrollador. Script en Python que carga un archivo Excel de ventas, procesa los datos con `pandas` y genera un archivo de resumen con múltiples hojas.

## 📁 Estructura del proyecto

```
.
├── data/
│   └── datos_ventas.xlsx        # Archivo de entrada
├── output/
│   └── resumen_ventas.xlsx      # Archivo generado
├── procesar_ventas.py           # Script principal
├── formato_excel.py             # Aplica formato profesional al Excel de salida
├── requirements.txt             # Dependencias
├── .gitignore
└── README.md
```

## ⚙️ Requisitos

- Python 3.10+
- Dependencias en `requirements.txt`

## 🚀 Ejecución

```bash
# 1. Crear entorno virtual (recomendado)
python -m venv .venv
source .venv/bin/activate        # Linux/Mac
.venv\Scripts\activate           # Windows

# 2. Instalar dependencias
pip install -r requirements.txt

# 3. Ejecutar
python procesar_ventas.py
python formato_excel.py          # Opcional: aplica formato al Excel de salida
```

## 🔎 Lógica del procesamiento

| Paso | Acción |
|------|--------|
| 1 | Carga de `datos_ventas.xlsx` en un DataFrame de pandas. |
| 2 | Imputación de `Total_Venta` faltantes usando `Cantidad * Precio_Unitario`. |
| 3 | Conversión de `Fecha` a `datetime`. |
| 4 | Filtrado de ventas del año 2023. |
| 5 | Creación de la columna `Mes` (numérico). |
| 6 | Cálculo de totales por vendedor y por mes. |
| 7 | Exportación a `output/resumen_ventas.xlsx`. |

## 📊 Salida

El archivo `resumen_ventas.xlsx` contiene tres hojas:

| Hoja | Contenido |
|------|-----------|
| `Resumen_Ventas` | Total de ventas por vendedor. |
| `Ventas_Mensuales` | Total de ventas por mes. |
| `Vendedor_x_Mes` | Cruce vendedor × mes (hoja adicional de valor). |

## 🧠 Decisiones técnicas

- **Imputación con `fillna(Cantidad * Precio_Unitario)`**: no sobrescribe valores existentes; solo rellena `NaN`.
- **Filtro explícito por año**: aunque el insumo actual está 100 % en 2023, se conserva el filtro para robustez ante insumos futuros.
- **Función por paso**: separar cada etapa en una función facilita test unitarios y reutilización.
- **Hoja adicional `Vendedor_x_Mes`**: no la pide el enunciado, pero es el cruce más útil para lectura ejecutiva. Se marca como opcional en el README.

## 👤 Autor

Santiago González Medina
