# -*- coding: utf-8 -*-
"""
TALLER DE VISUALIZACIÓN DE DATOS EN POWER BI
ETAPA 2 — Limpieza y transformación de la base de datos

Base: ÁREAS PROTEGIDAS CON ECOSISTEMAS DE PÁRAMO Y GESTIÓN DE LA RAPE-E
      REGIÓN CENTRAL (datos abiertos de Colombia)

Transformaciones aplicadas (todas documentadas y verificadas):
  1. Recorte de espacios sobrantes al inicio/fin de los textos (Text.Trim).
  2. Normalización de espacios dobles internos (1 registro afectado).
  3. Eliminación del único registro completamente duplicado (fila 420).
  4. Conversión de 'ÁREA - HECTÁREAS' a número decimal (formato inglés:
     coma como separador de miles y punto decimal).
  5. Conversión de 'FECHA INICIO DE ADMINISTRACIÓN' a fecha (AAAA-MM-DD)
     desde texto en inglés 'AAAA MMM DD hh:mm:ss AM/PM'.
  6. Códigos territoriales conservados como TEXTO (sin pérdida de ceros).

Decisiones acordadas con el grupo (documentadas en el informe):
  - Los 9 registros con código de departamento 11 (Bogotá D.C. según
    DIVIPOLA) conservan el nombre 'Cundinamarca' del dato fuente.
  - Se mantiene la ortografía original del dato fuente (sin corrección
    de tildes en valores).

Entrada:  datos/base_original.csv   (NO se modifica)
Salidas:  datos/base_limpia.csv
          evidencias/verificacion_limpieza.txt

Uso:  python scripts/limpieza_datos.py
Requiere: pandas (pip install pandas)
"""
import csv
import sys
from pathlib import Path
from collections import Counter

import pandas as pd

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ---------------------------------------------------------------- rutas
RAIZ = Path(__file__).resolve().parents[1]
CSV_ORIGINAL = RAIZ / "datos" / "base_original.csv"
CSV_LIMPIO = RAIZ / "datos" / "base_limpia.csv"
REPORTE = RAIZ / "evidencias" / "verificacion_limpieza.txt"

FECHA_MES = {"Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4, "May": 5, "Jun": 6,
             "Jul": 7, "Aug": 8, "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12}

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
COLS_TEXTO = list(COL.values())

L = []
def w(cadena=""):
    L.append(str(cadena))

def verificar(condicion, mensaje_exito, mensaje_fallo):
    """Verificación post-transformación: corta el proceso si no se cumple."""
    if condicion:
        w(f"  [OK] {mensaje_exito}")
    else:
        w(f"  [FALLO] {mensaje_fallo}")
        raise SystemExit(f"Verificación fallida: {mensaje_fallo}")

w("REGISTRO DE LIMPIEZA Y VERIFICACIÓN — BASE LIMPIA")
w(f"Entrada: {CSV_ORIGINAL.name} (intacta)")

# ================================================================ CARGA
df = pd.read_csv(CSV_ORIGINAL, dtype=str, keep_default_na=False, encoding="utf-8")
n0 = len(df)
w()
w("=" * 72)
w(f"Estado inicial: {n0} registros x {len(df.columns)} columnas")
w("=" * 72)

# ================================================================ PASO 1
# Recorte de espacios al inicio/fin en todas las columnas de texto
w()
w("PASO 1 — Recorte de espacios sobrantes (equivalente a Text.Trim en Power Query)")
antes = sum((df[c] != df[c].str.strip()).sum() for c in COLS_TEXTO)
for c in COLS_TEXTO:
    df[c] = df[c].str.strip()
despues = sum((df[c] != df[c].str.strip()).sum() for c in COLS_TEXTO)
w(f"  Celdas corregidas: {antes} (55 en municipio, 5 en páramo)")
verificar(despues == 0,
          f"0 celdas con espacios sobrantes (antes: {antes}).",
          f"Quedan {despues} celdas con espacios sobrantes.")
municipios = df[COL["nom_muni"]].nunique()
verificar(municipios == 129,
          f"Nombres de municipio distintos: {municipios} (antes: 131). Se unificaron "
          "'FOMEQUE '/'FOMEQUE' y 'UMBITA '/'UMBITA'; quedan 129 nombres para 130 "
          "códigos porque TURMEQUE aparece con dos códigos (15835 y 15897) en el "
          "dato fuente, inconsistencia que se documenta y NO se corrige.",
          f"Nombres de municipio inesperados: {municipios}")

# ================================================================ PASO 2
# Espacios dobles internos -> espacio simple
w()
w("PASO 2 — Normalización de espacios internos dobles")
dobles = df[COL["area"]].str.contains("  ", regex=False).sum()
df[COL["area"]] = df[COL["area"]].str.replace(r"\s+", " ", regex=True)
for c in COLS_TEXTO:
    df[c] = df[c].str.replace(r"\s+", " ", regex=True)
restantes = sum(df[c].str.contains("  ", regex=False).sum() for c in COLS_TEXTO)
w(f"  Corregidos: {dobles} registro(s) ('Paramos  y Bosques Altoandinos de Genova')")
verificar(restantes == 0,
          "0 celdas con espacios dobles internos.",
          f"Quedan {restantes} celdas con espacios dobles.")

# ================================================================ PASO 3
# Eliminación de duplicados completos (los 12 campos idénticos)
w()
w("PASO 3 — Eliminación de registros completamente duplicados")
# Se conserva la PRIMERA aparición (fila 208) y se elimina la repetida (fila 420)
dup_idx = df.index[df.duplicated(keep="first")].tolist()
for i in dup_idx:
    fila = df.loc[i]
    w(f"  Fila eliminada ({i + 2} del CSV): {fila[COL['cod_depto']]}-{fila[COL['nom_depto']]}, "
      f"municipio {fila[COL['cod_muni']]} {fila[COL['nom_muni']]}, "
      f"páramo '{fila[COL['paramo']]}', área '{fila[COL['area']]}', {fila[COL['ha']]} ha")
df = df.drop_duplicates(keep="first").reset_index(drop=True)
n1 = len(df)
verificar(n1 == n0 - 1 and len(dup_idx) == 1,
          f"Registros: {n0} -> {n1} (se eliminó exactamente 1 duplicado completo).",
          f"Registros tras deduplicar: {n1} (se esperaba {n0 - 1}).")
combo = ((df[COL["cod_muni"]] == "15550") &
         (df[COL["area"]] == "Cuenca del Rio Cravo Sur")).sum()
verificar(combo == 1,
          "El área 'Cuenca del Rio Cravo Sur' en el municipio PISVA queda exactamente una vez.",
          f"El área eliminada no quedó única: aparece {combo} veces.")

# ================================================================ PASO 4
# Hectáreas: texto -> número decimal
w()
w("PASO 4 — Conversión de 'ÁREA - HECTÁREAS' a número decimal")
w("  Formato del dato fuente: inglés (coma = miles, punto = decimal).")
ha_num = pd.to_numeric(df[COL["ha"]].str.replace(",", "", regex=False))
df[COL["ha"]] = ha_num.map(lambda v: f"{v:.2f}".rstrip("0").rstrip("."))
sin_numero = ha_num.isna().sum()
verificar(sin_numero == 0,
          f"Los {len(df)} valores se convirtieron a número decimal sin errores.",
          f"{sin_numero} valores no numéricos.")
w(f"  Mínimo: {ha_num.min()} | Máximo: {ha_num.max()} | "
  f"Suma: {ha_num.sum():.2f} ha | Media: {ha_num.mean():.2f} | Mediana: {ha_num.median():.2f}")
verificar(abs(ha_num.sum() - 681672.32) < 0.005,
          "Suma de hectáreas verificada: 681672.32 (original 681672.38 menos el duplicado de 0.06).",
          f"Suma de hectáreas inesperada: {ha_num.sum():.2f}")

# ================================================================ PASO 5
# Fecha: texto en inglés -> fecha AAAA-MM-DD
w()
w("PASO 5 — Conversión de 'FECHA INICIO DE ADMINISTRACIÓN' a tipo fecha")
def texto_a_fecha_iso(v):
    partes = v.split(" ")
    anio, mes_txt, dia = int(partes[0]), partes[1], int(partes[2])
    return f"{anio:04d}-{FECHA_MES[mes_txt]:02d}-{dia:02d}"
try:
    df[COL["fecha"]] = df[COL["fecha"]].map(texto_a_fecha_iso)
except KeyError:
    verificar(False, "", "Encontrado un mes en inglés no reconocido.")
patron_ok = df[COL["fecha"]].str.fullmatch(r"\d{4}-\d{2}-\d{2}").all()
verificar(patron_ok,
          f"Las {len(df)} fechas quedaron en formato AAAA-MM-DD "
          f"(rango {df[COL['fecha']].min()} a {df[COL['fecha']].max()}).",
          "Quedaron fechas sin convertir.")

# ================================================================ PASO 6
# Códigos territoriales como texto
w()
w("PASO 6 — Códigos territoriales conservados como TEXTO")
long_depto = df[COL["cod_depto"]].str.fullmatch(r"\d{2}").all()
long_muni = df[COL["cod_muni"]].str.fullmatch(r"\d{5}").all()
verificar(long_depto and long_muni,
          f"Código de departamento (2 dígitos) y de municipio (5 dígitos) "
          f"conservados como texto, sin pérdida de ceros iniciales.",
          "Algún código territorial no tiene la longitud esperada.")

# ================================================================ VERIFICACIÓN FINAL
w()
w("=" * 72)
w(f"VERIFICACIÓN FINAL — CONTEOS DE LAS TRES GRÁFICAS ({len(df)} registros)")
w("=" * 72)
esperados_depto = {"Cundinamarca": 234, "Boyaca": 161, "Tolima": 40, "Meta": 32}
esperados_sinap = {"Areas Protegidas Públicas": 430, "Areas Protegidas Privadas": 37}
esperados_cat = {"Distritos Regionales de Manejo Integrado": 111,
                 "Reservas Forestales Protectoras Regionales": 110,
                 "Parque Nacional Natural": 81,
                 "Parques Naturales Regionales": 62,
                 "Reservas Forestales Protectoras Nacionales": 55,
                 "Reserva Natural de la Sociedad Civil": 37,
                 "Santuario de Fauna y Flora": 7,
                 "Distritos de Conservacion de Suelos": 4}

w()
w("Gráfica 1 — Registros por departamento:")
obtenido = Counter(df[COL["nom_depto"]])
for valor in ["Cundinamarca", "Boyaca", "Tolima", "Meta"]:
    w(f"  {valor}: {obtenido[valor]} ({100 * obtenido[valor] / len(df):.2f} %)")
verificar(obtenido == Counter(esperados_depto),
          "Conteos por departamento coinciden con lo acordado (234/161/40/32).",
          f"Conteos por departamento no coinciden: {dict(obtenido)}")

w()
w("Gráfica 2 — Registros por tipo SINAP:")
obtenido = Counter(df[COL["sinap"]])
for valor, n in obtenido.most_common():
    w(f"  {valor}: {n} ({100 * n / len(df):.2f} %)")
verificar(obtenido == Counter(esperados_sinap),
          "Conteos pública/privada coinciden con lo acordado (430/37).",
          f"Conteos SINAP no coinciden: {dict(obtenido)}")

w()
w("Gráfica 3 — Registros por categoría:")
obtenido = Counter(df[COL["categoria"]])
for valor, n in obtenido.most_common():
    w(f"  {valor}: {n} ({100 * n / len(df):.2f} %)")
verificar(obtenido == Counter(esperados_cat),
          "Conteos por categoría coinciden con lo acordado (111/110/81/62/55/37/7/4).",
          f"Conteos por categoría no coinciden: {dict(obtenido)}")

w()
w("Otras verificaciones de integridad:")
w(f"  Nombres de páramo distintos: {df[COL['paramo']].nunique()} (original: 15)")
w(f"  Áreas protegidas con nombre único: {df[COL['area']].nunique()} (original: 101)")
w(f"  Administradores distintos: {df[COL['admin']].nunique()} (original: 11)")
w(f"  Códigos de municipio distintos: {df[COL['cod_muni']].nunique()} (original: 130)")
w(f"  Registros con proyecto RAPE 'Ninguno': "
  f"{(df[COL['proyectos']] == 'Ninguno').sum()} (original: 6)")

# ================================================================ GUARDAR
CSV_LIMPIO.parent.mkdir(parents=True, exist_ok=True)
df.to_csv(CSV_LIMPIO, index=False, quoting=csv.QUOTE_ALL,
          encoding="utf-8-sig", lineterminator="\n")

texto = "\n".join(L) + "\n"
REPORTE.write_text(texto, encoding="utf-8")
print(texto)
print(f"[OK] Base limpia guardada en: {CSV_LIMPIO}")
print(f"[OK] Reporte de verificación guardado en: {REPORTE}")
