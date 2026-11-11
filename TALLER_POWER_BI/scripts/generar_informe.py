# -*- coding: utf-8 -*-
"""
TALLER DE VISUALIZACIÓN DE DATOS EN POWER BI
ETAPA 5 — Generación automática del informe Word (.docx)

Genera: informe/Informe_Taller_PowerBI.docx

Todos los números del documento se calculan directamente desde
datos/base_limpia.csv para garantizar que el texto y las tablas
coincidan exactamente con la base analizada en Power BI.
Las gráficas se insertan desde evidencias/*.png (capturas reales
de Power BI, no imágenes simuladas).

Uso:  python scripts/generar_informe.py
Requiere: pandas, python-docx
"""
import sys
from pathlib import Path

import pandas as pd
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

sys.stdout.reconfigure(encoding="utf-8", errors="replace")

# ================================================================ RUTAS
RAIZ = Path(__file__).resolve().parents[1]
CSV_LIMPIO = RAIZ / "datos" / "base_limpia.csv"
EVIDENCIAS = RAIZ / "evidencias"
SALIDA = RAIZ / "informe" / "Informe_Taller_PowerBI.docx"

# ================================================================ PORTADA
INSTITUCION = "TECNOLÓGICO DE ANTIOQUIA"
ASIGNATURA = "Big Data"
TITULO = "Taller de visualización de datos en Power BI"
SUBTITULO = ("Análisis de áreas protegidas con ecosistemas de páramo y "
             "gestión de la RAP-E Región Central")
INTEGRANTE_1 = "Daniel Castrillón Betancur"
INTEGRANTE_2 = "Juan Manuel Monsalve Agudelo"
DOCENTE = "Cristian David Pérez Torres"
CIUDAD = "Medellín"
FECHA = "11 de noviembre de 2026"
FUENTE_URL = ("https://www.datos.gov.co/Ambiente-y-Desarrollo-Sostenible/"
              "-REAS-PROTEGIDAS-CON-ECOSISTEMAS-DE-P-RAMO-Y-GESTI/"
              "r2yz-k5z7")

# ================================================================ DATOS
df = pd.read_csv(CSV_LIMPIO, dtype=str, keep_default_na=False,
                 encoding="utf-8-sig")
COL = list(df.columns)
DEP, SINAP, CAT, HA, AREA, PARAMO = COL[1], COL[6], COL[7], COL[10], COL[5], COL[4]
df[HA] = pd.to_numeric(df[HA])

N = len(df)
UNICAS = df[AREA].nunique()
SUMA_HA = df[HA].sum()
por_dep = df[DEP].value_counts()
por_sinap = df[SINAP].value_counts()
por_cat = df[CAT].value_counts()
ha_dep = df.groupby(DEP)[HA].sum()
ha_cat = df.groupby(CAT)[HA].sum()
pct = lambda n, d=N: 100 * n / d


def es(valor, decimales=2):
    """Formato numérico español: 681672.32 -> 681.672,32"""
    texto = f"{valor:,.{decimales}f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def pc(n):
    return f"{es(pct(n), 2)} %"


# ================================================================ ESTILOS
doc = Document()

normal = doc.styles["Normal"]
normal.font.name = "Calibri"
normal.font.size = Pt(11)
normal.paragraph_format.space_after = Pt(6)
normal.paragraph_format.line_spacing = 1.15

for nombre, tam, color in [("Heading 1", 15, RGBColor(0x1F, 0x3B, 0x57)),
                           ("Heading 2", 13, RGBColor(0x2E, 0x55, 0x76))]:
    estilo = doc.styles[nombre]
    estilo.font.name = "Calibri"
    estilo.font.size = Pt(tam)
    estilo.font.bold = True
    estilo.font.color.rgb = color
    estilo.paragraph_format.space_before = Pt(14)
    estilo.paragraph_format.space_after = Pt(6)

seccion = doc.sections[0]
seccion.page_width = Cm(21.0)
seccion.page_height = Cm(29.7)
seccion.top_margin = Cm(2.54)
seccion.bottom_margin = Cm(2.54)
seccion.left_margin = Cm(2.54)
seccion.right_margin = Cm(2.54)

# Metadatos del documento
doc.core_properties.title = TITULO
doc.core_properties.subject = SUBTITULO
doc.core_properties.author = f"{INTEGRANTE_1}; {INTEGRANTE_2}"
doc.core_properties.last_modified_by = f"{INTEGRANTE_1}; {INTEGRANTE_2}"
doc.core_properties.comments = ("Taller de visualización de datos en Power BI - "
                                "Tecnológico de Antioquia")
doc.core_properties.keywords = "Power BI; datos abiertos; páramo; RAP-E; Big Data"
doc.core_properties.category = "Informe académico"
doc.core_properties.language = "es-CO"

# Numeración de páginas en el pie
pie = seccion.footer.paragraphs[0]
pie.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = pie.add_run()
for tipo, texto in [("begin", None), ("instr", "PAGE"), ("separate", None),
                    ("texto", "1"), ("end", None)]:
    if tipo == "instr":
        el = OxmlElement("w:instrText")
        el.set(qn("xml:space"), "preserve")
        el.text = texto
    elif tipo == "texto":
        el = OxmlElement("w:t")
        el.text = texto
    else:
        el = OxmlElement("w:fldChar")
        el.set(qn("w:fldCharType"), tipo)
    run._r.append(el)
run.font.size = Pt(9)


def parrafo(texto, negrita=False, cursiva=False, centrado=False,
            tam=None, color=None, espacio_antes=None, espacio_despues=None):
    p = doc.add_paragraph()
    if centrado:
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    if espacio_antes is not None:
        p.paragraph_format.space_before = Pt(espacio_antes)
    if espacio_despues is not None:
        p.paragraph_format.space_after = Pt(espacio_despues)
    r = p.add_run(texto)
    r.bold = negrita
    r.italic = cursiva
    if tam:
        r.font.size = Pt(tam)
    if color:
        r.font.color.rgb = color
    return p


def vineta(texto):
    p = doc.add_paragraph(style="List Bullet")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.add_run(texto)
    return p


def item_numerado(numero, texto):
    """Ítem numerado manual: cada lista reinicia en 1 sin depender de Word."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(0.75)
    p.paragraph_format.first_line_indent = Cm(-0.75)
    p.paragraph_format.space_after = Pt(6)
    p.add_run(f"{numero}. ")
    p.add_run(texto)
    return p


def tabla(encabezados, filas, anchos=None):
    t = doc.add_table(rows=1, cols=len(encabezados))
    t.style = "Table Grid"
    for i, texto in enumerate(encabezados):
        celda = t.rows[0].cells[i]
        celda.text = ""
        r = celda.paragraphs[0].add_run(texto)
        r.bold = True
        r.font.size = Pt(10)
    for fila in filas:
        celdas = t.add_row().cells
        for i, valor in enumerate(fila):
            celdas[i].text = ""
            r = celdas[i].paragraphs[0].add_run(str(valor))
            r.font.size = Pt(10)
    if anchos:
        for fila in t.rows:
            for i, ancho in enumerate(anchos):
                fila.cells[i].width = Cm(ancho)
    return t


def imagen(nombre_archivo, ancho_cm, leyenda):
    ruta = EVIDENCIAS / nombre_archivo
    doc.add_picture(str(ruta), width=Cm(ancho_cm))
    doc.paragraphs[-1].alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after = Pt(12)
    r = p.add_run(leyenda)
    r.italic = True
    r.font.size = Pt(9)
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)


# ================================================================ PORTADA
for _ in range(4):
    doc.add_paragraph()
parrafo(INSTITUCION, negrita=True, centrado=True, tam=18,
        color=RGBColor(0x1F, 0x3B, 0x57))
parrafo(f"Asignatura: {ASIGNATURA}", centrado=True, tam=12)
for _ in range(3):
    doc.add_paragraph()
parrafo(TITULO, negrita=True, centrado=True, tam=24,
        color=RGBColor(0x1F, 0x3B, 0x57))
parrafo(SUBTITULO, cursiva=True, centrado=True, tam=13)
for _ in range(4):
    doc.add_paragraph()
parrafo("Integrantes:", negrita=True, centrado=True, tam=12)
parrafo(INTEGRANTE_1, centrado=True, tam=12)
parrafo(INTEGRANTE_2, centrado=True, tam=12)
doc.add_paragraph()
parrafo("Docente:", negrita=True, centrado=True, tam=12)
parrafo(DOCENTE, centrado=True, tam=12)
for _ in range(4):
    doc.add_paragraph()
parrafo(CIUDAD, centrado=True, tam=12)
parrafo(FECHA, centrado=True, tam=12)
doc.add_page_break()

# ================================================================ 1. INTRO
doc.add_heading("1. INTRODUCCIÓN", level=1)
parrafo(
    "El presente informe recoge el desarrollo del taller de visualización de "
    "datos en Power BI, cuyo propósito es aplicar un proceso completo de "
    "análisis de datos —búsqueda, descarga, exploración, limpieza, "
    "transformación, preparación y visualización— sobre un conjunto de datos "
    "abierto del Estado colombiano. El ejercicio busca fortalecer habilidades "
    "para identificar patrones, tendencias y comportamientos a partir de "
    "información real, y para comunicar los hallazgos mediante visualizaciones "
    "claras y sustentadas."
)
parrafo(
    "Los datos abiertos son un recurso fundamental para la transparencia, la "
    "investigación y la toma de decisiones basada en evidencia. El Portal "
    "Nacional de Datos Abiertos de Colombia (datos.gov.co) publica información "
    "producida por entidades públicas, que puede reutilizarse libremente. Sin "
    "embargo, estos datos suelen requerir un proceso de depuración previo, "
    "porque pueden contener duplicados, formatos heterogéneos, espacios "
    "sobrantes o inconsistencias territoriales; precisamente por eso la "
    "limpieza y la transformación constituyen una etapa central del análisis."
)
parrafo(
    "El conjunto de datos analizado corresponde a las áreas protegidas con "
    "ecosistemas de páramo y la gestión de la RAP-E Región Central. Los "
    "páramos son ecosistemas estratégicos de alta montaña: regulan el ciclo "
    "del agua, almacenan carbono y albergan una biodiversidad única, por lo "
    "que las áreas protegidas son el principal instrumento para su "
    "conservación. La Región Administrativa y de Planeación Especial (RAP-E) "
    "Región Central articula a Bogotá D.C. con los departamentos de "
    "Cundinamarca, Boyacá, Meta y Tolima, exactamente el ámbito territorial "
    "que cubre el conjunto de datos."
)
parrafo(
    "A partir de esta base se construyeron tres visualizaciones en Power BI: "
    "la distribución de los registros por departamento, la proporción de áreas "
    "protegidas públicas y privadas, y la distribución por categoría de "
    "protección ambiental. Para cada una se presentan dos conclusiones "
    "sustentadas en los resultados obtenidos, junto con una conclusión general "
    "del análisis."
)

# ================================================================ 2. BASE
doc.add_heading("2. DESCRIPCIÓN DE LA BASE DE DATOS", level=1)
parrafo(
    "El conjunto de datos se denomina “Áreas protegidas con ecosistemas de "
    "páramo y gestión de la RAP-E Región Central”. Fue publicado por la "
    "Región Administrativa y de Planeación Especial – RAP-E Región Central en "
    "el Portal Nacional de Datos Abiertos de Colombia (datos.gov.co), en la "
    "categoría “Ambiente y Desarrollo Sostenible”. El archivo descargado "
    "corresponde a un archivo CSV en codificación UTF-8, delimitado por comas "
    "y con todos sus campos entre comillas."
)
tabla(
    ["Característica", "Detalle"],
    [["Nombre del conjunto de datos",
      "Áreas protegidas con ecosistemas de páramo y gestión de la RAP-E Región Central"],
     ["Entidad responsable",
      "Región Administrativa y de Planeación Especial – RAP-E Región Central"],
     ["Portal de origen", "Portal Nacional de Datos Abiertos – datos.gov.co"],
     ["Categoría temática", "Ambiente y Desarrollo Sostenible"],
     ["Registros originales", f"{N + 1}"],
     ["Registros después de la limpieza", f"{N}"],
     ["Columnas", "12"],
     ["Cobertura territorial",
      "Bogotá D.C. y los departamentos de Cundinamarca, Boyacá, Meta y Tolima"],
     ["Áreas protegidas con nombre único", f"{UNICAS}"],
     ["Complejos de páramo identificados", f"{df[PARAMO].nunique()}"]],
    anchos=[5.8, 9.6],
)
parrafo("Tabla 1. Características generales del conjunto de datos.",
        cursiva=True, centrado=True, tam=9, color=RGBColor(0x55, 0x55, 0x55))

parrafo(
    "La base contiene 12 variables que describen, para cada registro, la "
    "ubicación territorial del área protegida, el complejo de páramo "
    "asociado, el nombre y la categoría del área, la entidad administradora, "
    "la fecha de inicio de la administración, su extensión en hectáreas y la "
    "ejecución de proyectos de la RAP-E. Cada fila representa la presencia de "
    "un área protegida en un municipio determinado; por ello, un mismo nombre "
    "de área puede aparecer en varios registros si está presente en varios "
    "municipios."
)
tabla(
    ["Variable", "Descripción", "Tipo de dato"],
    [["CÓDIGO DE DEPARTAMENTO", "Código DANE del departamento (2 dígitos)", "Texto"],
     ["NOMBRE DEL DEPARTAMENTO", "Nombre del departamento", "Texto"],
     ["CÓDIGO DE MUNICIPIO", "Código DANE del municipio (5 dígitos)", "Texto"],
     ["NOMBRE DE MUNICIPIO", "Nombre del municipio", "Texto"],
     ["NOMBRE PARAMO", "Complejo de páramo asociado al área", "Texto"],
     ["NOMBRE ÁREA PROTEGIDA", "Nombre del área protegida", "Texto"],
     ["ÁREAS PROTEGIDAS DEL SISTEMA NACIONAL DE AREAS PROTEGIDAS",
      "Clasificación del área como pública o privada (SINAP)", "Texto"],
     ["CATEGORÍA DE ÁREAS PROTEGIDAS (Decreto 2372/2010)",
      "Categoría de manejo del área protegida", "Texto"],
     ["ADMINISTRADOR", "Entidad que administra el área", "Texto"],
     ["FECHA INICIO DE ADMINISTRACIÓN", "Fecha de inicio de la administración", "Fecha"],
     ["ÁREA - HECTÁREAS", "Extensión registrada del área, en hectáreas", "Número decimal"],
     ["EJECUCIÓN DE PROYECTOS DE LA RAPE REGION CENTRAL",
      "Indica si el área está vinculada al proyecto de la RAP-E", "Texto"]],
    anchos=[5.0, 8.0, 2.4],
)
parrafo("Tabla 2. Variables del conjunto de datos y tipo de dato final.",
        cursiva=True, centrado=True, tam=9, color=RGBColor(0x55, 0x55, 0x55))

parrafo(
    f"Entre las características generales de la información se destaca que "
    f"los {N} registros corresponden a {UNICAS} áreas protegidas con nombre "
    f"único, distribuidas en {df[COL[2]].nunique()} códigos de municipio y "
    f"{df[COL[8]].nunique()} entidades administradoras. Las fechas de inicio "
    f"de administración abarcan desde 1968 hasta 2017. La extensión total "
    f"registrada asciende a {es(SUMA_HA)} hectáreas, aunque esta cifra debe "
    f"interpretarse con cautela, pues suma los valores registrados en cada "
    f"fila y no la superficie única de cada área protegida."
)

# ================================================================ 3. LIMPIEZA
doc.add_heading("3. EXPLORACIÓN, LIMPIEZA Y TRANSFORMACIÓN", level=1)
parrafo(
    "La exploración inicial se realizó con scripts de Python (biblioteca "
    "pandas) sobre una copia intacta del archivo original, con el fin de "
    "diagnosticar la calidad de los datos antes de transformarlos. Se "
    "revisaron la estructura del archivo, los valores vacíos, los registros "
    "duplicados, los espacios sobrantes, los formatos numéricos y de fecha, "
    "los valores únicos de las variables categóricas y la consistencia entre "
    "los códigos y los nombres de departamentos y municipios. Posteriormente, "
    "las transformaciones se aplicaron en Power BI mediante Power Query, y el "
    "resultado se verificó de forma independiente con los scripts de Python."
)
parrafo("Durante la exploración se identificaron los siguientes problemas de calidad:",
        espacio_despues=4)
vineta("Un registro completamente duplicado: la fila 420 repetía exactamente "
       "los 12 campos de la fila 208 (área “Cuenca del Rio Cravo Sur”, "
       "municipio de Pisva, Boyacá, 0,06 ha).")
vineta("Sesenta celdas con espacios sobrantes al inicio o al final: 55 en "
       "NOMBRE DE MUNICIPIO (por ejemplo “SAMACA ”, “CHOACHI ”, “FOMEQUE ”) y "
       "5 en NOMBRE PARAMO (“Chili Barragan ”). Además, un nombre de área "
       "protegida contenía un espacio doble interno (“Paramos  y Bosques "
       "Altoandinos de Genova”).")
vineta("La columna ÁREA - HECTÁREAS estaba almacenada como texto y combinaba "
       "tres formatos: enteros (322), decimales con punto (966.34) y valores "
       "con coma de miles y punto decimal (57,062.02), es decir, formato "
       "numérico inglés.")
vineta("Las 468 fechas estaban en texto con formato inglés y hora "
       "(“1986 Apr 07 12:00:00 AM”), con el mes escrito en inglés.")
vineta("Inconsistencias territoriales: 9 registros con código de "
       "departamento 11 (que según la DIVIPOLA del DANE corresponde a Bogotá "
       "D.C.) aparecen con el nombre “Cundinamarca”; y el nombre “TURMEQUE” "
       "aparece asociado a dos códigos de municipio (15835 y 15897).")
vineta("Siete registros con 0 hectáreas y 96 valores de hectáreas alejados "
       "del rango intercuartílico; estos últimos corresponden a áreas "
       "legítimas de gran extensión (por ejemplo, el Parque Nacional Natural "
       "El Cocuy con 66.741,42 ha), por lo que no se eliminaron.")

parrafo(
    "Con base en este diagnóstico se aplicaron las siguientes transformaciones "
    "en Power Query, todas reproducibles y verificadas:"
)
pasos = [
    "Recorte de espacios sobrantes al inicio y al final de todas las columnas "
    "de texto (equivalente a Text.Trim), que corrigió 60 celdas.",
    "Normalización de espacios dobles internos, que corrigió 1 registro.",
    "Eliminación de registros completamente duplicados: se conservó la primera "
    "aparición y se eliminó la fila 420, con lo cual la base pasó de 468 a "
    "467 registros.",
    "Conversión de ÁREA - HECTÁREAS a número decimal con configuración "
    "regional inglés (Estados Unidos), lo que interpretó correctamente la coma "
    "como separador de miles y el punto como separador decimal.",
    "Conversión de FECHA INICIO DE ADMINISTRACIÓN a tipo fecha, "
    "transformando el texto inglés al formato estándar año-mes-día.",
    "Conservación de los códigos territoriales como texto, para no perder "
    "ceros iniciales ni alterar su naturaleza de identificadores.",
]
for i, paso in enumerate(pasos, 1):
    item_numerado(i, paso)

parrafo(
    "Las inconsistencias territoriales identificadas se documentaron pero no "
    "se corrigieron, con el criterio de no inventar equivalencias ni alterar "
    "el dato publicado por la entidad fuente; de igual forma, se conservó la "
    "ortografía original de los valores. La verificación posterior confirmó "
    f"que la base limpia quedó con {N} registros, sin duplicados completos, "
    "sin espacios sobrantes, con las hectáreas como número decimal, las "
    "fechas como tipo fecha y los códigos como texto. Los scripts de Python "
    "reprodujeron exactamente los mismos resultados, lo que permitió validar "
    "la transformación realizada en Power BI.",
)
tabla(
    ["Aspecto", "Antes (base original)", "Después (base limpia)"],
    [["Registros", "468", f"{N}"],
     ["Duplicados completos", "1", "0"],
     ["Celdas con espacios sobrantes", "60", "0"],
     ["Espacios dobles internos", "1", "0"],
     ["ÁREA - HECTÁREAS", "Texto con formatos mixtos", "Número decimal"],
     ["FECHA INICIO DE ADMINISTRACIÓN", "Texto en inglés", "Fecha"],
     ["Códigos territoriales", "Texto", "Texto (conservados)"]],
    anchos=[5.6, 4.9, 4.9],
)
parrafo("Tabla 3. Resumen de las transformaciones aplicadas.",
        cursiva=True, centrado=True, tam=9, color=RGBColor(0x55, 0x55, 0x55))

# ================================================================ 4. VISUALIZACIONES
doc.add_heading("4. VISUALIZACIONES Y ANÁLISIS", level=1)
parrafo(
    "Las tres visualizaciones se construyeron en Power BI Desktop a partir de "
    "la consulta limpia “AreasProtegidas”. Todas presentan título, ejes "
    "identificados, etiquetas de datos y un formato homogéneo, y responden a "
    "una pregunta de análisis específica."
)
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
r = p.add_run("Nota metodológica: ")
r.bold = True
p.add_run(
    f"los conteos presentados corresponden a registros, no a áreas protegidas "
    f"únicas. La base contiene {N} registros pero solo {UNICAS} áreas con "
    f"nombre único, porque una misma área puede estar presente en varios "
    f"municipios (por ejemplo, el área Sumapaz aparece en 12 municipios y "
    f"Chingaza en 10). Por lo tanto, las gráficas describen la distribución "
    f"territorial de los registros y no la cantidad de áreas protegidas "
    f"distintas."
)

# ---------------------------------------------------------------- Gráfica 1
doc.add_heading("4.1 Distribución de áreas protegidas por departamento", level=2)
parrafo(
    "La primera visualización es un gráfico de barras agrupadas que responde "
    "a la pregunta: ¿qué departamento concentra la mayor cantidad de "
    "registros de áreas protegidas? El eje vertical muestra el nombre del "
    "departamento y el eje horizontal el conteo de registros."
)
imagen("grafica_departamentos.png", 14.5,
       "Gráfica 1. Distribución de áreas protegidas por departamento "
       "(elaboración propia en Power BI).")
parrafo(
    f"Los resultados muestran {por_dep.iloc[0]} registros en Cundinamarca "
    f"({pc(por_dep.iloc[0])}), {por_dep.iloc[1]} en Boyacá "
    f"({pc(por_dep.iloc[1])}), {por_dep.iloc[2]} en Tolima "
    f"({pc(por_dep.iloc[2])}) y {por_dep.iloc[3]} en Meta "
    f"({pc(por_dep.iloc[3])}), sobre un total de {N} registros."
)
parrafo("Conclusiones:", negrita=True, espacio_despues=2)
item_numerado(
    1,
    f"Cundinamarca concentra la mitad de los registros ({por_dep.iloc[0]} de "
    f"{N}, equivalentes al {pc(por_dep.iloc[0])}). Junto con Boyacá "
    f"({por_dep.iloc[1]} registros, {pc(por_dep.iloc[1])}), estos dos "
    f"departamentos reúnen el {es(pct(por_dep.iloc[0] + por_dep.iloc[1]))} % "
    f"del total, es decir, aproximadamente ocho de cada diez registros. Esto "
    f"evidencia una alta concentración de las áreas protegidas con "
    f"ecosistemas de páramo registradas en la Región Central."
)
item_numerado(
    2,
    f"El número de registros no equivale a la superficie protegida: Boyacá, "
    f"con {por_dep.iloc[1]} registros, suma {es(ha_dep['Boyaca'])} ha "
    f"({es(100 * ha_dep['Boyaca'] / SUMA_HA)} % de las hectáreas "
    f"registradas), mientras que Cundinamarca, con {por_dep.iloc[0]} "
    f"registros, suma {es(ha_dep['Cundinamarca'])} ha "
    f"({es(100 * ha_dep['Cundinamarca'] / SUMA_HA)} %). Así, Cundinamarca "
    f"lidera en cantidad de registros, pero Boyacá concentra la mayor "
    f"extensión; Tolima y Meta aportan el resto de los registros y de la "
    f"superficie."
)

# ---------------------------------------------------------------- Gráfica 2
doc.add_heading("4.2 Distribución de áreas protegidas públicas y privadas", level=2)
parrafo(
    "La segunda visualización es un gráfico de anillos que responde a la "
    "pregunta: ¿qué proporción de registros corresponde a áreas protegidas "
    "públicas y privadas? La leyenda muestra la clasificación del Sistema "
    "Nacional de Áreas Protegidas (SINAP) y los valores el conteo de "
    "registros."
)
imagen("grafica_publicas_privadas.png", 13.0,
       "Gráfica 2. Distribución de áreas protegidas públicas y privadas "
       "(elaboración propia en Power BI).")
parrafo(
    f"Los resultados indican {por_sinap.get('Areas Protegidas Públicas', 0)} "
    f"registros de áreas protegidas públicas "
    f"({pc(por_sinap.get('Areas Protegidas Públicas', 0))}) y "
    f"{por_sinap.get('Areas Protegidas Privadas', 0)} de áreas protegidas "
    f"privadas ({pc(por_sinap.get('Areas Protegidas Privadas', 0))})."
)
parrafo("Conclusiones:", negrita=True, espacio_despues=2)
item_numerado(
    1,
    f"El {pc(por_sinap.get('Areas Protegidas Públicas', 0))} de los registros "
    f"({por_sinap.get('Areas Protegidas Públicas', 0)} de {N}) corresponde a "
    f"áreas protegidas públicas, frente a solo el "
    f"{pc(por_sinap.get('Areas Protegidas Privadas', 0))} "
    f"({por_sinap.get('Areas Protegidas Privadas', 0)}) de áreas privadas. "
    f"Esto muestra que la protección de los ecosistemas de páramo en la "
    f"Región Central recae principalmente en el Estado y en las autoridades "
    f"ambientales, como Parques Nacionales Naturales, las corporaciones "
    f"autónomas regionales y el Ministerio de Ambiente y Desarrollo "
    f"Sostenible."
)
item_numerado(
    2,
    f"Las {por_sinap.get('Areas Protegidas Privadas', 0)} áreas privadas "
    f"corresponden en su totalidad a Reservas Naturales de la Sociedad Civil: "
    f"al cruzar las variables, la categoría “Reserva Natural de la Sociedad "
    f"Civil” tiene exactamente {por_cat.get('Reserva Natural de la Sociedad Civil', 0)} "
    f"registros, todos dentro de la clasificación privada. Esto indica que la "
    f"figura privada de conservación es única en su tipo dentro del conjunto "
    f"de datos y cumple un papel complementario, aunque todavía minoritario, "
    f"frente al esquema público de protección."
)

# ---------------------------------------------------------------- Gráfica 3
doc.add_heading("4.3 Distribución por categoría de protección ambiental", level=2)
parrafo(
    "La tercera visualización es un gráfico de columnas agrupadas que responde "
    "a la pregunta: ¿cuáles son las categorías de protección ambiental con "
    "mayor y menor representación? El eje horizontal presenta la categoría de "
    "área protegida y el eje vertical el conteo de registros."
)
imagen("grafica_categorias.png", 14.5,
       "Gráfica 3. Áreas protegidas según categoría ambiental "
       "(elaboración propia en Power BI).")
parrafo(
    f"Las categorías con más registros son Distritos Regionales de Manejo "
    f"Integrado ({por_cat.get('Distritos Regionales de Manejo Integrado', 0)}; "
    f"{pc(por_cat.get('Distritos Regionales de Manejo Integrado', 0))}) y "
    f"Reservas Forestales Protectoras Regionales "
    f"({por_cat.get('Reservas Forestales Protectoras Regionales', 0)}; "
    f"{pc(por_cat.get('Reservas Forestales Protectoras Regionales', 0))}); "
    f"las de menor representación son Distritos de Conservación de Suelos "
    f"({por_cat.get('Distritos de Conservacion de Suelos', 0)}; "
    f"{pc(por_cat.get('Distritos de Conservacion de Suelos', 0))}) y "
    f"Santuarios de Fauna y Flora "
    f"({por_cat.get('Santuario de Fauna y Flora', 0)}; "
    f"{pc(por_cat.get('Santuario de Fauna y Flora', 0))})."
)
parrafo("Conclusiones:", negrita=True, espacio_despues=2)
item_numerado(
    1,
    f"Las categorías de gestión regional concentran la mayor cantidad de "
    f"registros: los Distritos Regionales de Manejo Integrado "
    f"({por_cat.get('Distritos Regionales de Manejo Integrado', 0)}; "
    f"{pc(por_cat.get('Distritos Regionales de Manejo Integrado', 0))}) y las "
    f"Reservas Forestales Protectoras Regionales "
    f"({por_cat.get('Reservas Forestales Protectoras Regionales', 0)}; "
    f"{pc(por_cat.get('Reservas Forestales Protectoras Regionales', 0))}) "
    f"suman el "
    f"{es(pct(por_cat.get('Distritos Regionales de Manejo Integrado', 0) + por_cat.get('Reservas Forestales Protectoras Regionales', 0)))} % "
    f"del total. Esto refleja el peso de las corporaciones autónomas "
    f"regionales como administradoras de los páramos en el ámbito de la "
    f"RAP-E Región Central."
)
item_numerado(
    2,
    f"Las categorías con menor representación son los Distritos de "
    f"Conservación de Suelos "
    f"({por_cat.get('Distritos de Conservacion de Suelos', 0)}; "
    f"{pc(por_cat.get('Distritos de Conservacion de Suelos', 0))}) y los "
    f"Santuarios de Fauna y Flora "
    f"({por_cat.get('Santuario de Fauna y Flora', 0)}; "
    f"{pc(por_cat.get('Santuario de Fauna y Flora', 0))}), que juntas apenas "
    f"alcanzan el "
    f"{es(pct(por_cat.get('Distritos de Conservacion de Suelos', 0) + por_cat.get('Santuario de Fauna y Flora', 0)))} % "
    f"de los registros. En contraste, el Parque Nacional Natural, aunque "
    f"representa el "
    f"{pc(por_cat.get('Parque Nacional Natural', 0))} de los registros "
    f"({por_cat.get('Parque Nacional Natural', 0)}), concentra la mayor "
    f"extensión: {es(ha_cat['Parque Nacional Natural'])} ha, equivalentes al "
    f"{es(100 * ha_cat['Parque Nacional Natural'] / SUMA_HA)} % de las "
    f"hectáreas registradas en toda la base. Las categorías nacionales, por "
    f"tanto, tienen menos registros pero mayor tamaño promedio por área."
)

# ================================================================ 5. CONCLUSIÓN
doc.add_heading("5. CONCLUSIÓN GENERAL", level=1)
parrafo(
    f"El análisis de los {N} registros de áreas protegidas con ecosistemas de "
    f"páramo de la RAP-E Región Central permite identificar un patrón claro: "
    f"la protección de estos ecosistemas es predominantemente pública "
    f"({pc(por_sinap.get('Areas Protegidas Públicas', 0))} de los registros) y "
    f"está concentrada territorialmente en Cundinamarca y Boyacá, que reúnen "
    f"el {es(pct(por_dep.iloc[0] + por_dep.iloc[1]))} % de los registros. En "
    f"cuanto a las categorías de manejo, las figuras de gestión regional "
    f"(Distritos Regionales de Manejo Integrado y Reservas Forestales "
    f"Protectoras Regionales) dominan en número de registros, mientras que el "
    f"Parque Nacional Natural, con menos registros, concentra la mayor "
    f"extensión registrada ({es(100 * ha_cat['Parque Nacional Natural'] / SUMA_HA)} % "
    f"de las hectáreas). Esto sugiere un esquema complementario: una presencia "
    f"territorial amplia a cargo de las corporaciones autónomas regionales y "
    f"un resguardo de grandes extensiones a cargo de la categoría nacional."
)
parrafo(
    "El ejercicio también evidencia que los datos abiertos, aunque valiosos, "
    "requieren un proceso riguroso de limpieza y verificación: la base "
    "contenía un duplicado exacto, espacios sobrantes, hectáreas y fechas en "
    "formatos heterogéneos e inconsistencias territoriales entre códigos y "
    "nombres. Estas últimas se documentaron sin alterar el dato fuente, lo "
    "que permite que el análisis sea transparente y reproducible. La "
    "combinación de Power Query para transformar y de scripts de Python para "
    "verificar demostró ser una estrategia eficaz para garantizar la calidad "
    "del resultado."
)
parrafo(
    f"Finalmente, debe tenerse presente la limitación metodológica señalada: "
    f"los {N} registros representan {UNICAS} áreas protegidas únicas, porque "
    f"una misma área puede estar presente en varios municipios. Por ello, las "
    f"conclusiones describen la distribución territorial de los registros y "
    f"no la cantidad de áreas protegidas distintas; futuros análisis podrían "
    f"profundizar en la superficie única por área y en la evolución temporal "
    f"de la administración, variables que la base también permite explorar."
)

# ================================================================ 6. FUENTE
doc.add_heading("6. FUENTE DE INFORMACIÓN", level=1)
parrafo(
    "Conjunto de datos: “ÁREAS PROTEGIDAS CON ECOSISTEMAS DE PÁRAMO Y GESTIÓN "
    "DE LA RAPE REGIÓN CENTRAL”."
)
parrafo(
    "Entidad responsable: Región Administrativa y de Planeación Especial – "
    "RAP-E Región Central, Bogotá D.C."
)
parrafo("Portal de origen: Portal Nacional de Datos Abiertos de Colombia "
        "(datos.gov.co), categoría Ambiente y Desarrollo Sostenible.")
parrafo(f"Enlace: {FUENTE_URL}")
parrafo(f"Fecha de consulta: {FECHA}.")
parrafo(
    "Archivos complementarios del taller: base de datos original descargada "
    "(CSV), base de datos limpia (CSV), archivo de Power BI "
    "(Analisis_Areas_Protegidas.pbix), scripts de verificación en Python y "
    "código M de Power Query."
)

# ================================================================ GUARDAR
SALIDA.parent.mkdir(parents=True, exist_ok=True)
doc.save(SALIDA)

# ---------------------------------------------------------------- verificación
doc2 = Document(SALIDA)
n_parrafos = len(doc2.paragraphs)
n_imagenes = len(doc2.inline_shapes)
n_tablas = len(doc2.tables)
tam = SALIDA.stat().st_size
print(f"[OK] Informe generado: {SALIDA}")
print(f"     Párrafos: {n_parrafos} | Tablas: {n_tablas} | "
      f"Imágenes insertadas: {n_imagenes} | Tamaño: {tam:,} bytes")
assert n_imagenes == 3, "Deben insertarse exactamente 3 gráficas"
print("[OK] Verificación: el archivo se abre correctamente y contiene las 3 gráficas.")
