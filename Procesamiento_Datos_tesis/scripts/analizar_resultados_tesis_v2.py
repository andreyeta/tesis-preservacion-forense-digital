# -*- coding: utf-8 -*-
"""
analizar_resultados_tesis_v2.py

Versión corregida:
- Detecta columnas que empiezan con P1, P2, P3... aunque tengan texto largo.
- Lee la hoja Matriz_Checklist si existe.
- Genera tabulación, indicadores, síntesis cualitativa y gráficas.

Uso:
    python analizar_resultados_tesis_v2.py --input matriz_resultados_simulados_tesis.xlsx --output resultados_analisis_tesis.xlsx
"""

import argparse
from pathlib import Path

import pandas as pd
import matplotlib.pyplot as plt


PREGUNTAS = {
    "P1": "Políticas o lineamientos de gestión y protección de información digital",
    "P2": "Procedimientos documentados para generación y entrega de información institucional",
    "P3": "Controles de acceso a sistemas institucionales",
    "P4": "Registros de auditoría o logs",
    "P5": "Documentación del origen de los datos",
    "P6": "Mecanismos de verificación de integridad",
    "P7": "Políticas documentadas de administración de seguridad de sistemas",
    "P8": "Procedimientos de control de calidad de datos",
    "P9": "Informes relacionados con calidad del dato",
}


def normalizar_texto(valor):
    if pd.isna(valor):
        return "Sin respuesta"
    return str(valor).strip()


def detectar_hoja_checklist(xls):
    posibles = [
        "Matriz_Checklist",
        "Checklist",
        "Matriz Checklist",
        "Lista de Verificación",
        "Lista_Verificacion",
    ]

    for hoja in posibles:
        if hoja in xls.sheet_names:
            return hoja

    # Respaldo: busca una hoja que tenga columnas tipo P1...
    for hoja in xls.sheet_names:
        temp = pd.read_excel(xls, sheet_name=hoja, nrows=2)
        columnas = [str(c).strip().upper() for c in temp.columns]
        if any(c.startswith("P1") for c in columnas):
            return hoja

    return xls.sheet_names[0]


def mapear_columnas_preguntas(df):
    """
    Devuelve un diccionario:
    {"P1": "P1 Políticas / lineamientos ...", "P2": "P2 Procedimientos ..."}
    """
    mapa = {}

    for codigo in PREGUNTAS.keys():
        for columna in df.columns:
            col_limpia = str(columna).strip().upper()
            if col_limpia == codigo or col_limpia.startswith(codigo + " "):
                mapa[codigo] = columna
                break

    return mapa


def leer_checklist(ruta_excel):
    xls = pd.ExcelFile(ruta_excel)
    hoja = detectar_hoja_checklist(xls)

    df = pd.read_excel(ruta_excel, sheet_name=hoja)
    df.columns = [str(c).strip() for c in df.columns]

    mapa_columnas = mapear_columnas_preguntas(df)

    if not mapa_columnas:
        raise ValueError(
            "No encontré columnas que empiecen con P1, P2, P3... "
            "Revisa que el Excel tenga la matriz del checklist."
        )

    # Crea una versión estándar con columnas P1, P2, etc.
    df_estandar = df.copy()
    for codigo, columna_original in mapa_columnas.items():
        df_estandar[codigo] = df_estandar[columna_original].apply(normalizar_texto)

    columnas_p = list(mapa_columnas.keys())

    return df_estandar, hoja, columnas_p, mapa_columnas


def calcular_frecuencias(df, columnas_p):
    tablas = {}
    total = len(df)

    for col in columnas_p:
        conteo = (
            df[col]
            .value_counts(dropna=False)
            .rename_axis("Respuesta")
            .reset_index(name="Frecuencia")
        )
        conteo["Porcentaje"] = (conteo["Frecuencia"] / total * 100).round(2)
        conteo.insert(0, "Código", col)
        conteo.insert(1, "Pregunta", PREGUNTAS.get(col, col))
        tablas[col] = conteo

    resumen_general = pd.concat(tablas.values(), ignore_index=True)
    return tablas, resumen_general


def calcular_indicadores(df, columnas_p):
    positivas = {
        "Sí",
        "Si",
        "Implementado",
        "Alto",
        "Cumple",
    }

    registros = []
    total = len(df)

    for col in columnas_p:
        cantidad_positiva = df[col].isin(positivas).sum()
        registros.append({
            "Código": col,
            "Pregunta": PREGUNTAS.get(col, col),
            "Respuestas favorables": cantidad_positiva,
            "Total": total,
            "Porcentaje favorable": round(cantidad_positiva / total * 100, 2),
        })

    return pd.DataFrame(registros)


def generar_graficas(tablas, carpeta_graficas):
    carpeta_graficas.mkdir(parents=True, exist_ok=True)
    rutas = []

    for codigo, tabla in tablas.items():
        titulo = PREGUNTAS.get(codigo, codigo)

        plt.figure(figsize=(10, 5))
        plt.bar(tabla["Respuesta"], tabla["Frecuencia"])
        plt.title(f"{codigo}. {titulo}")
        plt.xlabel("Respuesta")
        plt.ylabel("Frecuencia")
        plt.xticks(rotation=25, ha="right")
        plt.tight_layout()

        ruta = carpeta_graficas / f"{codigo}_grafica.png"
        plt.savefig(ruta, dpi=200)
        plt.close()

        rutas.append(str(ruta))

    return rutas


def analizar_entrevista(ruta_excel):
    xls = pd.ExcelFile(ruta_excel)

    posibles = [
        "Matriz_Entrevistas",
        "Entrevista",
        "Entrevistas",
        "Matriz Entrevista",
        "Matriz_Entrevista",
    ]

    hoja_entrevista = None
    for hoja in posibles:
        if hoja in xls.sheet_names:
            hoja_entrevista = hoja
            break

    if hoja_entrevista is None:
        return pd.DataFrame([{
            "Categoría": "Entrevista",
            "Menciones estimadas": 0,
            "Hallazgo": "No se encontró una hoja de entrevista en el archivo de entrada.",
        }])

    df = pd.read_excel(ruta_excel, sheet_name=hoja_entrevista)
    df.columns = [str(c).strip() for c in df.columns]

    texto_total = " ".join(df.astype(str).fillna("").agg(" ".join, axis=1).tolist()).lower()

    categorias = {
        "Controles de acceso": ["acceso", "permisos", "roles", "credenciales"],
        "Logs y auditoría": ["log", "logs", "bitácora", "bitácoras", "auditoría", "registro", "registros"],
        "Trazabilidad": ["trazabilidad", "fuente", "origen", "procedencia"],
        "Integridad": ["integridad", "validación", "verificación", "consistencia", "inconsistencia"],
        "Documentación": ["documentación", "documentado", "procedimiento", "procedimientos", "lineamiento"],
        "Calidad del dato": ["calidad", "dato", "datos"],
        "Capacitación": ["capacitación", "formación"],
    }

    registros = []
    for categoria, palabras in categorias.items():
        menciones = sum(texto_total.count(p.lower()) for p in palabras)
        if menciones > 0:
            hallazgo = f"Se identificaron menciones relacionadas con {categoria.lower()}."
        else:
            hallazgo = f"No se identificaron menciones relevantes sobre {categoria.lower()}."

        registros.append({
            "Categoría": categoria,
            "Menciones estimadas": menciones,
            "Hallazgo": hallazgo,
        })

    return pd.DataFrame(registros)


def generar_texto_analisis(indicadores):
    promedio = round(indicadores["Porcentaje favorable"].mean(), 2)
    mayor = indicadores.sort_values("Porcentaje favorable", ascending=False).iloc[0]
    menor = indicadores.sort_values("Porcentaje favorable", ascending=True).iloc[0]

    lineas = [
        "ANÁLISIS AUTOMÁTICO DE RESULTADOS",
        "",
        f"A partir de la matriz analizada, se obtuvo un promedio general de respuestas favorables de {promedio}%.",
        f"El aspecto con mayor nivel de cumplimiento fue: {mayor['Pregunta']}, con {mayor['Porcentaje favorable']}%.",
        f"El aspecto con menor nivel de cumplimiento fue: {menor['Pregunta']}, con {menor['Porcentaje favorable']}%.",
        "",
        "Los resultados reflejan que la institución cuenta con avances en controles técnicos y gestión de información digital.",
        "Sin embargo, también se identifican oportunidades de mejora en documentación formal, trazabilidad, verificación de integridad y estandarización de procedimientos.",
        "",
        "Desde el enfoque de preservación forense digital, estos hallazgos sugieren que no basta con contar con sistemas y accesos controlados.",
        "También es necesario fortalecer la evidencia documental, los registros de auditoría y los mecanismos que permitan comprobar el origen, integridad y trazabilidad de la información.",
    ]

    return pd.DataFrame({"Texto": lineas})


def ajustar_excel(writer):
    for ws in writer.book.worksheets:
        ws.freeze_panes = "A2"
        for col_cells in ws.columns:
            col_letter = col_cells[0].column_letter
            max_len = 0
            for cell in col_cells:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = min(max_len + 2, 60)


def exportar_resultados(ruta_salida, df_original, resumen_general, indicadores, sintesis_entrevista, texto_analisis):
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        df_original.to_excel(writer, sheet_name="Matriz_Checklist", index=False)
        resumen_general.to_excel(writer, sheet_name="Tabulacion", index=False)
        indicadores.to_excel(writer, sheet_name="Indicadores", index=False)
        sintesis_entrevista.to_excel(writer, sheet_name="Sintesis_Entrevista", index=False)
        texto_analisis.to_excel(writer, sheet_name="Analisis_Automatico", index=False)
        ajustar_excel(writer)


def main():
    parser = argparse.ArgumentParser(description="Analiza resultados de checklist y entrevista de tesis.")
    parser.add_argument("--input", required=True, help="Ruta del Excel de entrada.")
    parser.add_argument("--output", default="resultados_analisis_tesis.xlsx", help="Ruta del Excel de salida.")
    parser.add_argument("--graficas", default="graficas_resultados", help="Carpeta donde se guardarán las gráficas.")

    args = parser.parse_args()

    ruta_input = Path(args.input)
    ruta_output = Path(args.output)
    carpeta_graficas = Path(args.graficas)

    if not ruta_input.exists():
        raise FileNotFoundError(f"No existe el archivo de entrada: {ruta_input}")

    print("Leyendo matriz del checklist...")
    df_checklist, hoja, columnas_p, mapa_columnas = leer_checklist(ruta_input)

    print(f"Hoja utilizada: {hoja}")
    print("Columnas detectadas:")
    for codigo, columna in mapa_columnas.items():
        print(f" - {codigo}: {columna}")

    print("Calculando frecuencias y porcentajes...")
    tablas, resumen_general = calcular_frecuencias(df_checklist, columnas_p)

    print("Calculando indicadores...")
    indicadores = calcular_indicadores(df_checklist, columnas_p)

    print("Generando gráficas...")
    rutas_graficas = generar_graficas(tablas, carpeta_graficas)

    print("Analizando entrevista...")
    sintesis_entrevista = analizar_entrevista(ruta_input)

    print("Generando análisis automático...")
    texto_analisis = generar_texto_analisis(indicadores)

    print("Exportando resultados...")
    exportar_resultados(
        ruta_output,
        df_checklist,
        resumen_general,
        indicadores,
        sintesis_entrevista,
        texto_analisis,
    )

    print("\nProceso finalizado correctamente.")
    print(f"Excel generado: {ruta_output.resolve()}")
    print(f"Gráficas guardadas en: {carpeta_graficas.resolve()}")
    print("\nArchivos de gráficas:")
    for ruta in rutas_graficas:
        print(f" - {ruta}")


if __name__ == "__main__":
    main()
