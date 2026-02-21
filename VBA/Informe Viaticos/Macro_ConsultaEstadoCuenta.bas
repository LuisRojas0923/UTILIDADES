Option Explicit

'Attribute VB_Name = "ModuloConsultaEstadoCuenta"
' Macro para consultar estado de cuenta de viaticos desde PostgreSQL
' Base de datos: solid
' Servidor: 192.168.0.21:5432
' Usuario: postgres
' Autor: Sistema de Utilidades
' Fecha: 2025



Private Function ASCIIaTexto(valoresASCII As String) As String
    ' Convierte valores ASCII separados por comas a texto
    ' Ejemplo: "65,100,109" -> "Adm"
    
    Dim partes() As String
    Dim i As Long
    Dim resultado As String
    
    partes = Split(valoresASCII, ",")
    resultado = ""
    
    For i = LBound(partes) To UBound(partes)
        resultado = resultado & Chr(Val(Trim(partes(i))))
    Next i
    
    ASCIIaTexto = resultado
    
End Function

Sub ConsultarEstadoCuentaViaticos(Optional v_empleado As String = "", _
                                    Optional v_nombreempleado As String = "", _
                                    Optional v_desde As String = "", _
                                    Optional v_hasta As String = "")
    ' Conecta a PostgreSQL y consulta el estado de cuenta de viaticos
    ' Parametros opcionales:
    '   v_empleado: Cedula del empleado (ej: '16162075')
    '   v_nombreempleado: Nombre del empleado (ej: 'gladys' o 'gladys amparo')
    '   v_desde: Fecha desde (formato: 'YYYY-MM-DD', ej: '2025-01-01')
    '   v_hasta: Fecha hasta (formato: 'YYYY-MM-DD', ej: '2025-01-31')
    ' Los resultados se muestran en una hoja de Excel
    
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim ws As Worksheet
    Dim fila As Long
    Dim columna As Long
    Dim cadenaConexion As String
    Dim consultaSQL As String
    Dim servidor As String
    Dim puerto As String
    Dim baseDatos As String
    Dim usuario As String
    Dim contrasena As String
    Dim mensajeError As String
    Dim paramEmpleado As String
    Dim paramNombre As String
    Dim paramDesde As String
    Dim paramHasta As String
    Application.ScreenUpdating = False
    
    ' Parametros de conexion
    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    ' Contrasena en formato ASCII (generada con convertir_contrasena.py)
    contrasena = ASCIIaTexto("65,100,109,105,110,83,111,108,105,100,50,48,50,53")
    
    ' Preparar parametros para la consulta SQL
    If v_empleado = "" Then
        paramEmpleado = "NULL"
    Else
        paramEmpleado = "'" & Replace(v_empleado, "'", "''") & "'"
    End If
    
    If v_nombreempleado = "" Then
        paramNombre = "NULL"
    Else
        paramNombre = "'" & Replace(v_nombreempleado, "'", "''") & "'"
    End If
    
    If v_desde = "" Then
        paramDesde = "NULL"
    Else
        paramDesde = "'" & v_desde & "'"
    End If
    
    If v_hasta = "" Then
        paramHasta = "NULL"
    Else
        paramHasta = "'" & v_hasta & "'"
    End If
    
    ' Consulta SQL con parametros (dividida en muchas partes para evitar limite de continuaciones)
    Dim sqlWith As String
    Dim sqlSelect1 As String
    Dim sqlSelect2 As String
    Dim sqlFrom As String ' se construye a partir de sqlFromTx + sqlFromJoins
    Dim sqlWhere1 As String
    Dim sqlWhere2 As String
    Dim sqlOrder As String
    
    ' Parte 1: WITH params
    sqlWith = "WITH params AS (" & vbCrLf & _
              "  SELECT" & vbCrLf & _
              "    " & paramEmpleado & "::text AS v_empleado," & vbCrLf & _
              "    " & paramNombre & "::text AS v_nombreempleado," & vbCrLf & _
              "    " & paramDesde & "::date AS v_desde," & vbCrLf & _
              "    " & paramHasta & "::date AS v_hasta" & vbCrLf & _
              ")" & vbCrLf
    
    ' Parte 2: SELECT inicial
    sqlSelect1 = "SELECT " & vbCrLf & _
                 "    t.codigo AS ""CODIGO""," & vbCrLf & _
                 "    t.fechaaplicacion::timestamp AS ""FECHA APLICACION""," & vbCrLf & _
                 "    t.empleado AS ""CEDULA""," & vbCrLf & _
                 "    COALESCE(c.nombreempleado, l.nombreempleado) AS ""EMPLEADO""," & vbCrLf & _
                 "    t.numerodocumento AS ""RADICADO""," & vbCrLf & _
                 "    CASE WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor ELSE 0 END AS ""VALOR CONSIGNACION""," & vbCrLf & _
                 "    CASE WHEN t.tipodocumento = 'LEGALIZACION' THEN t.valor ELSE 0 END AS ""VALOR LEGALIZACION""," & vbCrLf
    
    ' Parte 3: SUM y SALDO
    sqlSelect2 = "    SUM(" & vbCrLf & _
                 "        CASE " & vbCrLf & _
                 "            WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor" & vbCrLf & _
                 "            WHEN t.tipodocumento = 'LEGALIZACION' THEN -t.valor" & vbCrLf & _
                 "            ELSE 0" & vbCrLf & _
                 "        END" & vbCrLf & _
                 "    ) OVER (" & vbCrLf & _
                 "        PARTITION BY t.empleado" & vbCrLf & _
                 "        ORDER BY t.fechaaplicacion::timestamp ASC, t.codigo ASC" & vbCrLf & _
                 "    ) AS ""SALDO""," & vbCrLf & _
                 "    COALESCE(c.observaciones, l.observaciones) AS ""OBSERVACIONES""" & vbCrLf
    
    ' Parte 4: FROM con subqueries para agrupar datos y evitar duplicados
    Dim sqlFromTx As String
    Dim sqlFromJoins As String
    
    ' Subquery: agrupar transacciones por documento (unifica lineas de detalle de legalizaciones)
    sqlFromTx = "FROM (" & vbCrLf & _
                "    SELECT MIN(codigo) AS codigo," & vbCrLf & _
                "        fechaaplicacion, empleado, numerodocumento, tipodocumento," & vbCrLf & _
                "        SUM(valor::numeric) AS valor" & vbCrLf & _
                "    FROM transaccionviaticos" & vbCrLf & _
                "    GROUP BY fechaaplicacion, empleado, numerodocumento, tipodocumento" & vbCrLf & _
                ") t" & vbCrLf & _
                "CROSS JOIN params p" & vbCrLf
    
    ' Subqueries: una sola fila por consignacion/legalizacion para el JOIN
    sqlFromJoins = "LEFT JOIN (" & vbCrLf & _
                   "    SELECT codigoconsignacion," & vbCrLf & _
                   "        MAX(nombreempleado) AS nombreempleado, MAX(observaciones) AS observaciones" & vbCrLf & _
                   "    FROM consignacion GROUP BY codigoconsignacion" & vbCrLf & _
                   ") c ON c.codigoconsignacion = t.numerodocumento" & vbCrLf & _
                   "LEFT JOIN (" & vbCrLf & _
                   "    SELECT codigolegalizacion," & vbCrLf & _
                   "        MAX(nombreempleado) AS nombreempleado, MAX(observaciones) AS observaciones" & vbCrLf & _
                   "    FROM legalizacion GROUP BY codigolegalizacion" & vbCrLf & _
                   ") l ON l.codigolegalizacion = t.numerodocumento" & vbCrLf
    
    sqlFrom = sqlFromTx & sqlFromJoins
    
    ' Parte 5: WHERE inicial
    sqlWhere1 = "WHERE 1=1" & vbCrLf & _
                "  AND (p.v_empleado IS NULL OR t.empleado = p.v_empleado)" & vbCrLf & _
                "  AND (" & vbCrLf & _
                "        p.v_nombreempleado IS NULL" & vbCrLf & _
                "        OR COALESCE(c.nombreempleado, l.nombreempleado) ILIKE ('%' || p.v_nombreempleado || '%')" & vbCrLf & _
                "      )" & vbCrLf
    
    ' Parte 6: WHERE fechas
    sqlWhere2 = "  AND t.fechaaplicacion::timestamp >= COALESCE(p.v_desde::timestamp, '1900-01-01'::timestamp)" & vbCrLf & _
                "  AND t.fechaaplicacion::timestamp < (COALESCE(p.v_hasta::timestamp, '9999-12-31'::timestamp) + INTERVAL '1 day')" & vbCrLf
    
    ' Parte 7: ORDER BY
    sqlOrder = "ORDER BY t.empleado, t.fechaaplicacion::timestamp ASC, t.codigo ASC"
    
    ' Construir consulta completa usando concatenacion simple
    consultaSQL = sqlWith & sqlSelect1 & sqlSelect2 & sqlFrom & sqlWhere1 & sqlWhere2 & sqlOrder
    
    ' Cadena de conexion ODBC para PostgreSQL
    cadenaConexion = "Driver={PostgreSQL Unicode(x64)};" & _
                     "Server=" & servidor & ";" & _
                     "Port=" & puerto & ";" & _
                     "Database=" & baseDatos & ";" & _
                     "Uid=" & usuario & ";" & _
                     "Pwd=" & contrasena & ";"
    
    On Error GoTo ErrorHandler
    
    ' Variables para medir tiempos (diagnostico de rendimiento)
    Dim tiempoInicio As Double
    Dim tiempoConexion As Double
    Dim tiempoConsulta As Double
    Dim tiempoExtraccion As Double
    Dim tiempoFormato As Double
    Dim tiempoTotal As Double
    
    tiempoInicio = Timer
    
    ' Crear objetos ADO
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Mostrar estado
    Application.StatusBar = "Conectando a PostgreSQL..."
    Debug.Print "=========================================="
    Debug.Print "DIAGNOSTICO DE RENDIMIENTO"
    Debug.Print "=========================================="
    Debug.Print "Intentando conectar a PostgreSQL..."
    Debug.Print "Servidor: " & servidor & ":" & puerto
    Debug.Print "Base de datos: " & baseDatos
    
    ' Configurar timeout de conexion (30 segundos)
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 300 ' 5 minutos para consultas complejas
    
    ' Abrir conexion
    conn.Open cadenaConexion
    tiempoConexion = Timer - tiempoInicio
    Debug.Print "Conexion establecida correctamente en " & Format(tiempoConexion, "0.00") & " segundos"
    
    ' Configurar Recordset para optimizar rendimiento
    rs.CursorType = 0 ' adOpenForwardOnly (solo lectura hacia adelante, mas rapido)
    rs.CursorLocation = 2 ' adUseServer (usa cursor del servidor, mas eficiente)
    rs.LockType = 1 ' adReadOnly (solo lectura, mas rapido)
    
    ' Ejecutar consulta
    Application.StatusBar = "Ejecutando consulta SQL..."
    Debug.Print "Ejecutando consulta de estado de cuenta..."
    Dim tiempoAntesConsulta As Double
    tiempoAntesConsulta = Timer
    rs.Open consultaSQL, conn
    tiempoConsulta = Timer - tiempoAntesConsulta
    Debug.Print "Consulta ejecutada en " & Format(tiempoConsulta, "0.00") & " segundos"
    
    ' Crear o limpiar hoja de resultados
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Estado_Cuenta")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Estado_Cuenta"
    Else
        ' Limpiar solo desde la fila 7 en adelante para preservar las filas 1-6
        ws.Range("A7:ZZ" & ws.Rows.Count).Clear
    End If
    On Error GoTo ErrorHandler
    'Crear la tabla a partir de la fila 7
    ' Guardar numero de columnas antes del loop
    Dim numColumnas As Integer
    numColumnas = rs.Fields.Count
    
    ' Guardar nombres de columnas en un array para uso posterior
    Dim nombresColumnas() As String
    ReDim nombresColumnas(0 To numColumnas - 1)
    For columna = 0 To numColumnas - 1
        nombresColumnas(columna) = rs.Fields(columna).Name
    Next columna
    
    ' Identificar indice de la columna de fecha
    Dim indiceColumnaFecha As Integer
    indiceColumnaFecha = -1
    For columna = 0 To numColumnas - 1
        If UCase(nombresColumnas(columna)) = "FECHA APLICACION" Then
            indiceColumnaFecha = columna
            Exit For
        End If
    Next columna
    
    ' Identificar indice de la columna de cedula
    Dim indiceColumnaCedula As Integer
    indiceColumnaCedula = -1
    For columna = 0 To numColumnas - 1
        If UCase(nombresColumnas(columna)) = "CEDULA" Then
            indiceColumnaCedula = columna
            Exit For
        End If
    Next columna
    
    ' Escribir encabezados en la fila 7
    fila = 7
    For columna = 0 To numColumnas - 1
        ws.Cells(fila, columna + 1).Value = nombresColumnas(columna)
        ws.Cells(fila, columna + 1).Font.Bold = True
        ws.Cells(fila, columna + 1).Interior.Color = RGB(0, 32, 96) '002060azul
        ws.Cells(fila, columna + 1).Font.Color = RGB(255, 255, 255)
    Next columna
    
    ' Escribir datos usando CopyFromRecordset (MUCHO MAS RAPIDO que celda por celda)
    Application.StatusBar = "Extrayendo datos..."
    Debug.Print "Extrayendo datos desde recordset..."
    Dim tiempoAntesExtraccion As Double
    tiempoAntesExtraccion = Timer
    
    Dim filaInicial As Long
    filaInicial = 8
    Dim rangoDestino As Range
    Set rangoDestino = ws.Range("A" & filaInicial)
    
    ' Usar CopyFromRecordset para mejor rendimiento (hasta 100 veces mas rapido)
    ' Esta es la optimizacion mas importante
    rangoDestino.CopyFromRecordset rs
    
    tiempoExtraccion = Timer - tiempoAntesExtraccion
    Debug.Print "Datos extraidos en " & Format(tiempoExtraccion, "0.00") & " segundos"
    
    ' Contar registros escritos verificando la ultima fila con datos
    Dim ultimaFilaConDatos As Long
    ultimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Dim totalRegistros As Long
    If ultimaFilaConDatos >= filaInicial Then
        totalRegistros = ultimaFilaConDatos - filaInicial + 1
    Else
        totalRegistros = 0
    End If
    
    ' Formatear la columna de fecha si existe
    If indiceColumnaFecha >= 0 Then
        Dim letraColumnaFecha As String
        ' Convertir indice 0-based a numero de columna 1-based
        Dim numColFecha As Integer
        numColFecha = indiceColumnaFecha + 1
        ' Usar metodo simple: obtener letra desde la direccion de la celda
        letraColumnaFecha = Split(ws.Cells(1, numColFecha).Address, "$")(1)
        ws.Range(letraColumnaFecha & filaInicial & ":" & letraColumnaFecha & ultimaFilaConDatos).NumberFormat = "dd/mm/yyyy"
        ws.Columns(letraColumnaFecha).HorizontalAlignment = xlCenter
    End If
    
    ' Formatear la columna de cedula como numero si existe
    If indiceColumnaCedula >= 0 And totalRegistros > 0 Then
        Dim letraColumnaCedula As String
        ' Convertir indice 0-based a numero de columna 1-based
        Dim numColCedula As Integer
        numColCedula = indiceColumnaCedula + 1
        ' Usar metodo simple: obtener letra desde la direccion de la celda
        letraColumnaCedula = Split(ws.Cells(1, numColCedula).Address, "$")(1)
        ' Formatear como numero sin decimales
        ws.Range(letraColumnaCedula & filaInicial & ":" & letraColumnaCedula & ultimaFilaConDatos).NumberFormat = "0"
        ' Convertir valores de texto a numero si es necesario
        Dim celda As Range
        For Each celda In ws.Range(letraColumnaCedula & filaInicial & ":" & letraColumnaCedula & ultimaFilaConDatos)
            If Not IsEmpty(celda.Value) And IsNumeric(celda.Value) Then
                celda.Value = CDbl(celda.Value)
            End If
        Next celda
    End If
    
    ' Formatear tabla
    Application.StatusBar = "Formateando resultados..."
    Debug.Print "Total de registros extraidos: " & totalRegistros
    Dim tiempoAntesFormato As Double
    tiempoAntesFormato = Timer
    
    
    
    ' Aplicar formato de tabla
    Dim rangoTabla As Range
    Dim ultimaColumna As String
    ' Convertir numero de columna a letra (A, B, C, ..., Z, AA, AB, etc.)
    If numColumnas <= 26 Then
        ultimaColumna = Chr(64 + numColumnas)
    Else
        ultimaColumna = Chr(64 + Int((numColumnas - 1) / 26)) & Chr(65 + ((numColumnas - 1) Mod 26))
    End If
    Set rangoTabla = ws.Range("A7:" & ultimaColumna & ultimaFilaConDatos)
    
    On Error Resume Next
    ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TablaEstadoCuenta"
    On Error GoTo ErrorHandler
    
    ' Formatear tabla de Excel
    With ws.ListObjects("TablaEstadoCuenta")
        .TableStyle = "TableStyleMedium9"
        .ShowAutoFilter = True
    End With
    
    ' Centrar encabezados
    ws.Range("A7:I7").HorizontalAlignment = xlCenter
    
    ' Formatear columnas numericas
    On Error Resume Next
    'formatear las columnas F:H como moneda con cero decimales y con signo de pesos y negativo en rojo
    ws.Columns("F:H").NumberFormat = "$#,##0;[Red]-$#,##0"
    ws.Columns("B").NumberFormat = "dd/mm/yyyy"
    'centrar el contenido de la columna B
    ws.Columns("B").HorizontalAlignment = xlCenter
    
    ' Ajustar la altura de los encabezados
    ws.Rows("7:7").RowHeight = 66
    ' Definir ancho de las columnas segun la columna (ancho fijo)
    ws.Columns("A:C").ColumnWidth = 15
    ws.Columns("D").ColumnWidth = 40
    ws.Columns("E:H").ColumnWidth = 15
    ws.Columns("I").ColumnWidth = 100

    ws.Range("A7:I7").WrapText = True
    ws.Range("A7:I7").VerticalAlignment = xlCenter
    ws.Range("F6").FormulaR1C1 = "=+SUBTOTAL(9,TablaEstadoCuenta[VALOR CONSIGNACION])"
    ws.Range("G6").FormulaR1C1 = "=+SUBTOTAL(9,TablaEstadoCuenta[VALOR LEGALIZACION])"

    tiempoFormato = Timer - tiempoAntesFormato
    tiempoTotal = Timer - tiempoInicio
    
    ' Mostrar diagnostico completo de tiempos
    Debug.Print "=========================================="
    Debug.Print "RESUMEN DE TIEMPOS:"
    Debug.Print "  Conexion:        " & Format(tiempoConexion, "0.00") & " segundos"
    Debug.Print "  Ejecucion SQL:    " & Format(tiempoConsulta, "0.00") & " segundos"
    Debug.Print "  Extraccion datos: " & Format(tiempoExtraccion, "0.00") & " segundos"
    Debug.Print "  Formato tabla:    " & Format(tiempoFormato, "0.00") & " segundos"
    Debug.Print "  TOTAL:            " & Format(tiempoTotal, "0.00") & " segundos"
    Debug.Print "=========================================="
    Debug.Print "Registros procesados: " & totalRegistros
    If totalRegistros > 0 Then
        Debug.Print "Tiempo por registro: " & Format(tiempoTotal / totalRegistros * 1000, "0.00") & " ms"
    End If
    Debug.Print "=========================================="

    On Error GoTo ErrorHandler
    
    ' Cerrar conexiones
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    ' Mensaje de exito con informacion de rendimiento
    Application.StatusBar = "Proceso completado: " & totalRegistros & " registros extraidos"
    Debug.Print "Proceso completado exitosamente"
    
    Dim mensajeExito As String
    mensajeExito = "Consulta completada exitosamente." & vbCrLf & vbCrLf & _
                   "Total de registros: " & totalRegistros & vbCrLf & _
                   "Tiempo total: " & Format(tiempoTotal, "0.00") & " segundos" & vbCrLf & vbCrLf & _
                   "Desglose de tiempos:" & vbCrLf & _
                   "  - Conexion: " & Format(tiempoConexion, "0.00") & " s" & vbCrLf & _
                   "  - Consulta SQL: " & Format(tiempoConsulta, "0.00") & " s" & vbCrLf & _
                   "  - Extraccion: " & Format(tiempoExtraccion, "0.00") & " s" & vbCrLf & _
                   "  - Formato: " & Format(tiempoFormato, "0.00") & " s" & vbCrLf & vbCrLf & _
                   "Los resultados se encuentran en la hoja 'Estado_Cuenta'" & vbCrLf & _
                   "Ver ventana Inmediato (Ctrl+G) para mas detalles."
    
    MsgBox mensajeExito, vbInformation, "Consulta Estado de Cuenta"
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    ' Manejo de errores
    mensajeError = "Error en la conexion o consulta:" & vbCrLf & vbCrLf & _
                   "Numero de error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description
    
    Debug.Print "ERROR: " & mensajeError
    Debug.Print "Cadena de conexion: " & cadenaConexion
    
    ' Cerrar conexiones si estan abiertas
    On Error Resume Next
    If Not rs Is Nothing Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    If Not conn Is Nothing Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0
    
    Application.StatusBar = False
    MsgBox mensajeError, vbCritical, "Error en Consulta Estado de Cuenta"
    Application.ScreenUpdating = True
End Sub

Sub ConsultarEstadoCuentaCompleto()
    ' Ejecuta la consulta sin filtros (todos los registros)
    ' Util para obtener un reporte completo
    
    Call ConsultarEstadoCuentaViaticos("", "", "", "")
    
End Sub

Sub ConsultarEstadoCuentaPorEmpleado()
    ' Ejecuta la consulta filtrando por cedula de empleado
    ' Solicita la cedula mediante un InputBox
    
    Dim cedula As String
    
    cedula = InputBox("Ingrese la cedula del empleado:", "Consulta Estado de Cuenta")
    
    If cedula = "" Then
        MsgBox "Operacion cancelada. No se ingreso cedula.", vbInformation, "Consulta Cancelada"
        Exit Sub
    End If
    
    Call ConsultarEstadoCuentaViaticos(cedula, "", "", "")
    
End Sub

Sub ConsultarEstadoCuentaPorFecha()
    ' Ejecuta la consulta filtrando por rango de fechas
    ' Solicita las fechas mediante InputBox
    
    Dim FechaDesde As String
    Dim FechaHasta As String
    
    FechaDesde = InputBox("Ingrese la fecha inicial (formato: YYYY-MM-DD):", "Consulta Estado de Cuenta", "2025-01-01")
    
    If FechaDesde = "" Then
        MsgBox "Operacion cancelada. No se ingreso fecha inicial.", vbInformation, "Consulta Cancelada"
        Exit Sub
    End If
    
    FechaHasta = InputBox("Ingrese la fecha final (formato: YYYY-MM-DD):", "Consulta Estado de Cuenta", "2025-01-31")
    
    If FechaHasta = "" Then
        MsgBox "Operacion cancelada. No se ingreso fecha final.", vbInformation, "Consulta Cancelada"
        Exit Sub
    End If
    
    Call ConsultarEstadoCuentaViaticos("", "", FechaDesde, FechaHasta)
    
End Sub






