# Taller de visualización de datos en Power BI

Análisis de las **áreas protegidas con ecosistemas de páramo y gestión de la RAP-E Región Central** a partir de datos abiertos del Estado colombiano.

**Institución:** Tecnológico de Antioquia
**Asignatura:** Big Data
**Integrantes:** Daniel Castrillón Betancur y Juan Manuel Monsalve Agudelo
**Docente:** Cristian David Pérez Torres
**Medellín, 11 de noviembre de 2026**

## Descripción

El taller aplica un proceso completo de análisis de datos —búsqueda, descarga, exploración, limpieza, transformación, preparación y visualización— sobre un conjunto de datos abierto publicado en el Portal Nacional de Datos Abiertos de Colombia (datos.gov.co).

La base original contiene 468 registros y 12 variables; después de la limpieza queda con **467 registros**. Cada fila representa la presencia de un área protegida en un municipio, por lo que un mismo nombre de área puede aparecer varias veces (467 registros corresponden a 101 áreas protegidas únicas).

## Contenido del repositorio

```
.
├── TALLER_POWER_BI/
│   ├── datos/
│   │   ├── base_original.csv      # Copia intacta del archivo descargado
│   │   └── base_limpia.csv        # Base final: 467 registros
│   ├── scripts/
│   │   ├── analisis_exploratorio.py  # Diagnóstico de calidad de datos
│   │   ├── limpieza_datos.py         # Limpieza reproducible y verificada
│   │   └── generar_informe.py        # Genera el informe Word
│   ├── power_query/
│   │   └── transformaciones.m     # Código M usado en Power Query
│   ├── powerbi/
│   │   └── Analisis_Areas_Protegidas.pbix
│   ├── evidencias/
│   │   ├── grafica_departamentos.png
│   │   ├── grafica_publicas_privadas.png
│   │   ├── grafica_categorias.png
│   │   ├── diagnostico_exploratorio.txt
│   │   └── verificacion_limpieza.txt
│   └── informe/
│       └── Informe_Taller_PowerBI.docx
├── ÁREAS_PROTEGIDAS_..._20261111.csv   # Archivo original descargado
└── ejercicioBigData.jpeg               # Enunciado del taller
```

## Resultados principales

| Visualización | Hallazgo |
|---|---|
| Registros por departamento | Cundinamarca 234 (50,11 %), Boyacá 161 (34,48 %), Tolima 40 (8,57 %) y Meta 32 (6,85 %). Cundinamarca y Boyacá reúnen el 84,58 %. |
| Públicas vs. privadas | Públicas 430 (92,08 %) y privadas 37 (7,92 %). Las 37 privadas son Reservas Naturales de la Sociedad Civil. |
| Categorías de protección | DRMI 111 (23,77 %), RFPR 110 (23,55 %), PNN 81 (17,34 %), PNR 62 (13,28 %), RFPN 55 (11,78 %), RNSC 37 (7,92 %), SFF 7 (1,50 %) y DCS 4 (0,86 %). |

**Hallazgo clave:** el Parque Nacional Natural representa el 17,34 % de los registros pero concentra el **72,17 % de las hectáreas** registradas (491.976,85 ha de 681.672,32 ha).

## Limpieza aplicada

- Eliminación de 1 registro completamente duplicado (468 → 467).
- Recorte de espacios sobrantes en 60 celdas y normalización de 1 espacio doble interno.
- Conversión de `ÁREA - HECTÁREAS` de texto a número decimal (formato inglés).
- Conversión de `FECHA INICIO DE ADMINISTRACIÓN` de texto en inglés a tipo fecha.
- Conservación de los códigos territoriales como texto.
- Las inconsistencias territoriales y la ortografía original se documentaron, pero no se alteraron.

## Cómo reproducirlo

Requisitos: Python 3.12 con `pandas`, `python-docx` y `matplotlib`.

```bash
python TALLER_POWER_BI/scripts/analisis_exploratorio.py
python TALLER_POWER_BI/scripts/limpieza_datos.py
python TALLER_POWER_BI/scripts/generar_informe.py
```

El archivo `.pbix` se abre con Power BI Desktop.

## Fuente de los datos

**ÁREAS PROTEGIDAS CON ECOSISTEMAS DE PÁRAMO Y GESTIÓN DE LA RAPE REGIÓN CENTRAL**
Región Administrativa y de Planeación Especial – RAP-E Región Central, Bogotá D.C.
Portal Nacional de Datos Abiertos – datos.gov.co, categoría Ambiente y Desarrollo Sostenible.

<https://www.datos.gov.co/Ambiente-y-Desarrollo-Sostenible/-REAS-PROTEGIDAS-CON-ECOSISTEMAS-DE-P-RAMO-Y-GESTI/r2yz-k5z7>

Consulta: 11 de noviembre de 2026.
