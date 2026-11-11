// =====================================================================
// TALLER DE VISUALIZACIÓN DE DATOS EN POWER BI
// Código M (Power Query) — Limpieza y transformación de la base
// "ÁREAS PROTEGIDAS CON ECOSISTEMAS DE PÁRAMO Y GESTIÓN DE LA RAPE-E
//  REGIÓN CENTRAL" (datos abiertos de Colombia)
//
// CÓMO USARLO EN POWER BI DESKTOP:
//   1. Abrir Power BI Desktop.
//   2. Inicio > Obtener datos > Archivo de texto/CSV >
//      seleccionar TALLER_POWER_BI\datos\base_original.csv >
//      (en la vista previa NO dar clic en Cargar) dar clic en
//      "Transformar datos" para abrir el Editor de Power Query.
//   3. Inicio > Editor avanzado, borrar el contenido y pegar TODO
//      este código. Clic en Listo.
//   4. Verificar en "Pasos aplicados" que cada paso se ejecutó sin
//      errores y que la tabla queda con 467 filas.
//
// RESULTADO ESPERADO (verificado con scripts/limpieza_datos.py):
//   - 468 filas originales -> 467 filas (1 duplicado completo eliminado)
//   - 0 espacios sobrantes (antes 60 celdas) y 0 espacios dobles internos
//   - ÁREA - HECTÁREAS como número decimal (coma de miles eliminada)
//   - FECHA INICIO DE ADMINISTRACIÓN como tipo fecha (texto inglés -> fecha)
//   - Códigos territoriales como texto (2 y 5 dígitos)
//   - Conteos finales: Cundinamarca 234, Boyaca 161, Tolima 40, Meta 32;
//     Públicas 430 / Privadas 37; 8 categorías (111/110/81/62/55/37/7/4)
// =====================================================================
let
    // Ruta del CSV original. Ajustar si el proyecto cambia de ubicación.
    RutaBase = "C:\Users\danie\Downloads\Ejercicio Big Data\TALLER_POWER_BI\datos\base_original.csv",

    // -------------------------------------------------------------
    // PASO 0 — Lectura del CSV original
    // UTF-8 (65001), delimitado por comas, campos entre comillas
    // -------------------------------------------------------------
    Origen = Csv.Document(File.Contents(RutaBase), [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),

    // -------------------------------------------------------------
    // PASO 1 — Promover la primera fila como encabezados
    // -------------------------------------------------------------
    EncabezadosPromovidos = Table.PromoteHeaders(Origen, [PromoteAllScalars = true]),

    // -------------------------------------------------------------
    // PASO 2 — Tipos de datos: códigos y textos como TEXTO
    // (los códigos territoriales se conservan como texto para no
    //  perder ceros iniciales ni que Power BI los trate como números)
    // -------------------------------------------------------------
    TipoTexto = Table.TransformColumnTypes(EncabezadosPromovidos, {
        {"CÓDIGO DE DEPARTAMENTO", type text},
        {"NOMBRE DEL DEPARTAMENTO", type text},
        {"CÓDIGO DE MUNICIPIO", type text},
        {"NOMBRE DE MUNICIPIO", type text},
        {"NOMBRE PARAMO", type text},
        {"NOMBRE ÁREA PROTEGIDA", type text},
        {"ÁREAS PROTEGIDAS DEL SISTEMA NACIONAL DE AREAS PROTEGIDAS", type text},
        {"CATEGORÍA DE ÁREAS PROTEGIDAS (Decreto 2372/2010)", type text},
        {"ADMINISTRADOR", type text},
        {"EJECUCIÓN DE PROYECTOS DE LA RAPE REGION CENTRAL", type text}
    }),

    // -------------------------------------------------------------
    // PASO 3 — Limpieza de espacios en TODAS las columnas de texto:
    // recorta inicio/fin y colapsa espacios dobles internos.
    // Equivale a: Text.Trim  +  eliminar " " vacíos al dividir por " ".
    // Corrige 55 celdas en NOMBRE DE MUNICIPIO, 5 en NOMBRE PARAMO y
    // 1 espacio doble en NOMBRE ÁREA PROTEGIDA.
    // -------------------------------------------------------------
    ColumnasTexto = Table.ColumnNames(TipoTexto),
    LimpiaEspacios = (t as nullable text) as text =>
        if t = null then ""
        else Text.Combine(List.RemoveItems(Text.Split(t, " "), {""}), " "),
    EspaciosLimpiados = Table.TransformColumns(
        TipoTexto,
        List.Transform(ColumnasTexto, each {_, LimpiaEspacios, type text})
    ),

    // -------------------------------------------------------------
    // PASO 4 — Eliminar filas completamente duplicadas (12 campos).
    // Se conserva la primera aparición. Elimina exactamente 1 fila:
    // el área "Cuenca del Rio Cravo Sur" en PISVA (Boyacá), 0.06 ha.
    // -------------------------------------------------------------
    DuplicadosEliminados = Table.Distinct(EspaciosLimpiados),

    // -------------------------------------------------------------
    // PASO 5 — Hectáreas a NÚMERO DECIMAL con configuración regional
    // Inglés (Estados Unidos): en el dato fuente la coma separa miles
    // y el punto los decimales (ej. "5,716.92" -> 5716.92).
    // En la interfaz equivale a: clic derecho sobre la columna >
    // "Cambiar tipo con la configuración regional" > Número decimal >
    // Inglés (Estados Unidos).
    // -------------------------------------------------------------
    HectareasNumero = Table.TransformColumnTypes(
        DuplicadosEliminados,
        {{"ÁREA - HECTÁREAS", type number}},
        "en-US"
    ),

    // -------------------------------------------------------------
    // PASO 6 — Fecha de inicio de administración a tipo FECHA.
    // El dato viene como texto en inglés: "1986 Apr 07 12:00:00 AM".
    //
    // OPCIÓN B (la que usa este código, 100% confiable): convertir
    // mediante una función que separa año, mes y día.
    //
    // NOTA — OPCIÓN A (puede probarse primero desde la interfaz):
    //   clic derecho sobre la columna > "Cambiar tipo con la
    //   configuración regional" > Fecha > Inglés (Estados Unidos).
    //   Si con esa opción la columna muestra valores "Error", usar
    //   esta OPCIÓN B del código (equivalente en resultado).
    // El componente de hora es siempre 12:00:00 AM (medianoche),
    // por lo que se convierte directamente a tipo fecha (sin hora).
    // -------------------------------------------------------------
    MesesIngles = {"Jan", "Feb", "Mar", "Apr", "May", "Jun",
                   "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"},
    TextoAFecha = (t as text) as date =>
        let
            partes = Text.Split(t, " "),
            anio = Number.FromText(partes{0}),
            mes = List.PositionOf(MesesIngles, partes{1}) + 1,
            dia = Number.FromText(partes{2})
        in
            #date(anio, mes, dia),
    FechaConvertida = Table.TransformColumns(
        HectareasNumero,
        {{"FECHA INICIO DE ADMINISTRACIÓN", TextoAFecha, type date}}
    ),

    // -------------------------------------------------------------
    // PASO 7 — Verificación rápida integrada (no transforma nada).
    // Conteos por departamento tras la limpieza (467 registros):
    // -------------------------------------------------------------
    Verificacion = Table.Group(
        FechaConvertida,
        {"NOMBRE DEL DEPARTAMENTO"},
        {{"Registros", each Table.RowCount(_), Int64.Type}}
    ),

    Resultado = FechaConvertida
in
    Resultado
