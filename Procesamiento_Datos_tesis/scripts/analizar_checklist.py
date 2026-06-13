"""
Simulación y análisis del checklist de preservación forense digital
Tesis: Análisis del nivel de implementación de prácticas de preservación forense digital.

Uso:
1. Coloca este script en la misma carpeta que el archivo:
   simulacion_checklist_tesis.xlsx
2. Ejecuta:
   python analizar_checklist.py

Salidas:
- resumen_checklist_python.xlsx
- datos_powerbi_checklist.csv
- grafico_niveles.png
- grafico_dimensiones.png
"""

import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path

ARCHIVO_ENTRADA = "checklist.xlsx"
HOJA_CHECKLIST = "01_Checklist"

def clasificar_cumplimiento(valor: float) -> str:
    if valor >= 0.80:
        return "Alto"
    if valor >= 0.50:
        return "Medio"
    return "Bajo"

def main():
    ruta = Path(ARCHIVO_ENTRADA)
    if not ruta.exists():
        raise FileNotFoundError(f"No se encontró el archivo: {ARCHIVO_ENTRADA}")

    df = pd.read_excel(ruta, sheet_name=HOJA_CHECKLIST)

    # Asegurar tipos
    df["Puntaje"] = pd.to_numeric(df["Puntaje"], errors="coerce").fillna(0)

    total_indicadores = len(df)
    puntaje_total = df["Puntaje"].sum()
    cumplimiento_general = df["Puntaje"].mean()
    nivel_general = clasificar_cumplimiento(cumplimiento_general)

    resumen_general = pd.DataFrame({
        "Indicador": [
            "Total de indicadores evaluados",
            "Puntaje total obtenido",
            "Cumplimiento general",
            "Nivel general de implementación"
        ],
        "Valor": [
            total_indicadores,
            puntaje_total,
            round(cumplimiento_general * 100, 2),
            nivel_general
        ]
    })

    resumen_nivel = (
        df.groupby("Nivel", as_index=False)
          .agg(Cantidad=("ID", "count"))
    )
    resumen_nivel["Porcentaje"] = (resumen_nivel["Cantidad"] / total_indicadores * 100).round(2)

    resumen_dimension = (
        df.groupby("Dimensión", as_index=False)
          .agg(
              Puntaje_promedio=("Puntaje", "mean"),
              Indicadores=("ID", "count")
          )
    )
    resumen_dimension["Cumplimiento_%"] = (resumen_dimension["Puntaje_promedio"] * 100).round(2)
    resumen_dimension["Nivel"] = resumen_dimension["Puntaje_promedio"].apply(clasificar_cumplimiento)

    resumen_resultado = (
        df.groupby("Resultado simulado", as_index=False)
          .agg(Cantidad=("ID", "count"))
    )
    resumen_resultado["Porcentaje"] = (resumen_resultado["Cantidad"] / total_indicadores * 100).round(2)

    # Exportar Excel de resultados
    with pd.ExcelWriter("resumen_checklist_python.xlsx", engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Base_Checklist", index=False)
        resumen_general.to_excel(writer, sheet_name="Resumen_General", index=False)
        resumen_nivel.to_excel(writer, sheet_name="Resumen_Nivel", index=False)
        resumen_dimension.to_excel(writer, sheet_name="Resumen_Dimension", index=False)
        resumen_resultado.to_excel(writer, sheet_name="Resumen_Resultado", index=False)

    # CSV plano para Power BI
    df.to_csv("datos_powerbi_checklist.csv", index=False, encoding="utf-8-sig")

    # Gráfico 1: distribución por nivel
    plt.figure(figsize=(8, 5))
    plt.bar(resumen_nivel["Nivel"], resumen_nivel["Cantidad"])
    plt.title("Distribución de indicadores por nivel")
    plt.xlabel("Nivel")
    plt.ylabel("Cantidad de indicadores")
    plt.tight_layout()
    plt.savefig("grafico_niveles.png", dpi=150)
    plt.close()

    # Gráfico 2: cumplimiento por dimensión
    plt.figure(figsize=(10, 6))
    resumen_dimension_ordenado = resumen_dimension.sort_values("Cumplimiento_%")
    plt.barh(resumen_dimension_ordenado["Dimensión"], resumen_dimension_ordenado["Cumplimiento_%"])
    plt.title("Cumplimiento por dimensión")
    plt.xlabel("Cumplimiento (%)")
    plt.ylabel("Dimensión")
    plt.tight_layout()
    plt.savefig("grafico_dimensiones.png", dpi=150)
    plt.close()

    print("Análisis completado.")
    print(f"Cumplimiento general: {cumplimiento_general:.2%}")
    print(f"Nivel general: {nivel_general}")
    print("Archivos generados:")
    print("- resumen_checklist_python.xlsx")
    print("- datos_powerbi_checklist.csv")
    print("- grafico_niveles.png")
    print("- grafico_dimensiones.png")

if __name__ == "__main__":
    main()
