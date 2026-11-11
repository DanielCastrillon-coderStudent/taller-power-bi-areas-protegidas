# -*- coding: utf-8 -*-
"""
TALLER DE VISUALIZACIÓN DE DATOS EN POWER BI
ETAPA 1 — Análisis exploratorio de la base de datos

Base: ÁREAS PROTEGIDAS CON ECOSISTEMAS DE PÁRAMO Y GESTIÓN DE LA RAPE-E
      REGIÓN CENTRAL (datos abiertos de Colombia)

Este script NO modifica ningún dato. Solo lee datos/base_original.csv
(copia intacta del CSV descargado) y genera un reporte de diagnóstico:
  - Reporte en consola
  - evidencias/diagnostico_exploratorio.txt

Uso:  python scripts/analisis_exploratorio.py
Requiere: pandas (pip install pandas)
"""
import sys
from pathlib import Path
from collections import Counter

import pandas as pd

# Salida por consola compatible con Windows (acentos en español)
sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- rutas
RAIZ = Path(__file__).resolve().parents[1]          # TALLER_POWER_BI/
CSV_ORIGINAL = RAIZ / "datos" / "base_original.csv"
REPORTE = RAIZ / "evidencias" / "diagnostico_exploratorio.txt"

# ---------------------------------------------------------------- columnas
COL = {
    "cod_depto": "CÓDIGO DE DEPARTAMENTO",
    "nom_depto": "NOMBRE DEL DEPARTAMENTO",
    "cod_muni": "CÓDIGO DE MUNICIPIO",
    "nom_muni": "NOMBRE DE MUNICIPIO",
    "paramo": "NOMBRE PARAMO",
    "area": "NOMBRE ÁREA PROTEGIDA",
    "sinap": "ÁREAS PROTEGIDAS DEL SISTEMA NACIONAL DE AREAS PROTEGIDAS",
    "categoria": "CATEGORÍA DE ÁREAS PROTEGIDAS (Decreto 2372/2010)",
    "admin": "ADMINISTRADOR",
    "fecha": "FECHA INICIO DE ADMINISTRACIÓN",
    "ha": "ÁREA - HECTÁREAS",
    "proyectos": "EJECUCIÓN DE PROYECTOS DE LA RAPE REGION CENTRAL",
}
FECHA_MES = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
             "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

L = []          # líneas del reporte
def w(cadena=""):
    L.append(str(cadena))

def seccion(titulo):
    w()
    w("=" * 72)
    w(titulo)
    w("=" * 72)

# ---------------------------------------------------------------- 1. carga
df = pd.read_csv(CSV_ORIGINAL, dtype=str, keep_default_na=False, encoding="utf-8")
n_filas, n_cols = df.shape

w("DIAGNÓSTICO EXPLORATORIO — BASE ORIGINAL")
w(f"Archivo analizado: {CSV_ORIGINAL.name}")
w(f"Generado por: scripts/analisis_exploratorio.py (no modifica datos)")

seccion("1. ESTRUCTURA GENERAL")
w(f"Filas de datos (sin encabezado): {n_filas}")
w(f"Columnas: {n_cols}")
w("Nombres exactos de las columnas:")
for i, c in enumerate(df.columns, 1):
    w(f"  {i:2d}. {c}")

seccion("2. VALORES VACÍOS POR COLUMNA")
for c in df.columns:
    vacios = int((df[c].str.strip() == "").sum())
    w(f"  {c}: {vacios}")

seccion("3. ESPACIOS SOBRANTES (inicio/fin) POR COLUMNA")
total_esp = 0
for c in df.columns:
    afectados = df[c] != df[c].str.strip()
    n = int(afectados.sum())
    if n:
        total_esp += n
        valores = sorted(set(df.loc[afectados, c]))
        w(f"  {c}: {n} registros; valores afectados: {valores}")
w(f"  TOTAL registros con espacios sobrantes: {total_esp}")

seccion("4. ESPACIOS DOBLES INTERNOS")
for c in df.columns:
    afectados = df[c].str.contains("  ", regex=False)
    n = int(afectados.sum())
    if n:
        w(f"  {c}: {n} registros; ejemplos: {sorted(set(df.loc[afectados, c]))[:3]}")

seccion("5. DUPLICADOS EXACTOS (los 12 campos idénticos)")
dup_mask = df.duplicated(keep=False)
dup_completos = df[dup_mask].copy()
dup_completos["FILA_CSV"] = dup_completos.index + 2   # +2: encabezado y base 1
w(f"  Registros duplicados exactos: {int(df.duplicated().sum())}")
w(f"  Filas implicadas (encabezado = fila 1):")
for _, fila in dup_completos.iterrows():
    w(f"    Fila {fila['FILA_CSV']}: {fila[COL['cod_depto']]}-{fila[COL['nom_depto']]}, "
      f"municipio {fila[COL['cod_muni']]} {fila[COL['nom_muni']].strip()}, "
      f"area '{fila[COL['area']]}', {fila[COL['ha']]} ha")

seccion("6. VALORES ÚNICOS DE VARIABLES CATEGÓRICAS")
for clave in ["cod_depto", "nom_depto", "cod_muni", "nom_muni", "paramo",
              "area", "sinap", "categoria", "admin", "proyectos"]:
    c = COL[clave]
    conteo = Counter(df[c])
    w(f"  {c}: {len(conteo)} valores únicos")
    if len(conteo) <= 12:
        for valor, n in conteo.most_common():
            w(f"     - '{valor}': {n}")
    else:
        for valor, n in conteo.most_common(10):
            w(f"     - '{valor}': {n}")
        w(f"     ... y {len(conteo) - 10} valores más")

seccion("7. CONSISTENCIA CÓDIGO-NOMBRE DE DEPARTAMENTO")
pares = df.groupby(COL["cod_depto"])[COL["nom_depto"]].agg(lambda s: sorted(set(s)))
for cod, nombres in pares.items():
    marca = "  <-- INCONSISTENTE" if len(nombres) > 1 else ""
    w(f"  Código {cod}: {nombres}{marca}")
w("  Nota: el código DANE 11 corresponde a Bogotá D.C. y el 25 a Cundinamarca")
w("  (División Político-Administrativa - DIVIPOLA del DANE). Los 9 registros con")
w("  código 11 tienen municipio 11001 (BOGOTA) pero nombre 'Cundinamarca'.")

seccion("8. CONSISTENCIA PREFIJO MUNICIPIO vs CÓDIGO DEPARTAMENTO")
mal = df[df[COL["cod_muni"]].str[:2] != df[COL["cod_depto"]]]
w(f"  Registros cuyo código de municipio NO inicia con el código de departamento: {len(mal)}")

seccion("9. CONSISTENCIA CÓDIGO-NOMBRE DE MUNICIPIO")
pares_m = df.groupby(COL["cod_muni"])[COL["nom_muni"]].agg(lambda s: sorted(set(s)))
conflictos = {cod: ns for cod, ns in pares_m.items() if len(ns) > 1}
w(f"  Códigos de municipio con más de un nombre: {len(conflictos)}")
for cod, ns in conflictos.items():
    w(f"    {cod}: {ns}")
w("  (Se debe únicamente a espacios sobrantes: 'FOMEQUE '/'FOMEQUE' y 'UMBITA '/'UMBITA')")
nombres_v = df.assign(**{COL["nom_muni"]: df[COL["nom_muni"]].str.strip()})
pares_inv = nombres_v.groupby(COL["nom_muni"])[COL["cod_muni"]].agg(lambda s: sorted(set(s)))
compartidos = {nom: cods for nom, cods in pares_inv.items() if len(cods) > 1}
w(f"  Nombres de municipio compartidos por más de un código: {len(compartidos)}")
for nom, cods in compartidos.items():
    w(f"    '{nom}': códigos {cods}")
    for cod in cods:
        sub = nombres_v[(nombres_v[COL["cod_muni"]] == cod)]
        w(f"      Código {cod}: {len(sub)} registros")
w("  Nota: según la DIVIPOLA del DANE, el código 15835 corresponde a Túrmeque y")
w("  el 15897 a Ubalá (ambos en Boyacá). La inconsistencia se documenta y NO se")
w("  corrige para no alterar el dato territorial de la fuente.")

seccion("10. FORMATO Y ESTADÍSTICAS DE HECTÁREAS")
def clasifica(v):
    v = v.strip()
    if v.isdigit():
        return "entero (ej. 322)"
    if v.replace(",", "").replace(".", "").isdigit():
        if "," in v:
            return "coma de miles y punto decimal (ej. 5,716.92)"
        return "decimal con punto, sin miles (ej. 966.34)"
    return "OTRO FORMATO"
formatos = Counter(df[COL["ha"]].map(clasifica))
for f_, n in formatos.items():
    w(f"  {f_}: {n} registros")
ha = pd.to_numeric(df[COL["ha"]].str.replace(",", "", regex=False))
w(f"  Mínimo: {ha.min()}  |  Máximo: {ha.max()}")
w(f"  Suma total: {ha.sum():.2f} ha")
w(f"  Media: {ha.mean():.2f}  |  Mediana: {ha.median():.2f}")
w(f"  Q1: {ha.quantile(0.25):.2f}  |  Q3: {ha.quantile(0.75):.2f}")
q1, q3 = ha.quantile(0.25), ha.quantile(0.75)
iqr = q3 - q1
lim_inf, lim_sup = q1 - 1.5 * iqr, q3 + 1.5 * iqr
w(f"  Registros fuera del rango intercuartílico [{lim_inf:.2f}, {lim_sup:.2f}]: "
  f"{int(((ha < lim_inf) | (ha > lim_sup)).sum())} (áreas grandes legítimas, no se eliminan)")
w(f"  Registros con 0 hectáreas: {int((ha == 0).sum())}")

seccion("11. FORMATO DE FECHAS")
def formato_fecha(v):
    partes = v.strip().split(" ")
    if (len(partes) == 5 and len(partes[0]) == 4 and partes[1] in FECHA_MES
            and len(partes[2]) == 2 and partes[4] in ("AM", "PM")):
        return "AAAA MMM DD hh:mm:ss AM/PM (mes en inglés)"
    return "OTRO FORMATO"
w("  Formatos detectados:")
for f_, n in Counter(df[COL["fecha"]].map(formato_fecha)).items():
    w(f"    {f_}: {n} registros")
def a_fecha(v):
    p = v.strip().split(" ")
    return (int(p[0]), FECHA_MES[p[1]], int(p[2]))
anios = Counter(a_fecha(v)[0] for v in df[COL["fecha"]])
w(f"  Rango de años: {min(anios)} a {max(anios)}")
w("  Años con más registros: " +
  ", ".join(f"{a} ({n})" for a, n in anios.most_common(5)))

seccion("12. CONTEOS PARA LAS TRES VISUALIZACIONES (base original, 468 registros)")
w("  Gráfica 1 — Registros por departamento (nombre):")
for valor, n in Counter(df[COL["nom_depto"]]).most_common():
    w(f"    {valor}: {n} ({100 * n / n_filas:.2f} %)")
w("  Gráfica 2 — Registros por tipo SINAP (pública/privada):")
for valor, n in Counter(df[COL["sinap"]]).most_common():
    w(f"    {valor}: {n} ({100 * n / n_filas:.2f} %)")
w("  Gráfica 3 — Registros por categoría de área protegida:")
for valor, n in Counter(df[COL["categoria"]]).most_common():
    w(f"    {valor}: {n} ({100 * n / n_filas:.2f} %)")

seccion("13. LIMITACIÓN METODOLÓGICA: REGISTROS vs ÁREAS ÚNICAS")
w(f"  Registros totales: {n_filas}")
w(f"  Áreas protegidas con nombre único: {df[COL['area']].nunique()}")
por_area = df.groupby(COL["area"])[COL["cod_muni"]].nunique()
w(f"  Áreas presentes en más de un municipio: {int((por_area > 1).sum())}")
w("  Ejemplos: " + ", ".join(f"{a} ({m} municipios)"
                              for a, m in por_area.sort_values(ascending=False).head(5).items()))

# ---------------------------------------------------------------- guardar
texto = "\n".join(L) + "\n"
REPORTE.parent.mkdir(parents=True, exist_ok=True)
REPORTE.write_text(texto, encoding="utf-8")
print(texto)
print(f"[OK] Reporte guardado en: {REPORTE}")
