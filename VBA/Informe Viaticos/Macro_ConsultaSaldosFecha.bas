'Attribute VB_Name = "ModuloConsultaSaldosFecha"
' Macro para consultar saldos de viaticos por empleado a una fecha especifica
' Base de datos: solid
' Servidor: 192.168.0.21:5432
' Usuario: postgres
' Autor: Sistema de Utilidades
' Fecha: 2025

Option Explicit

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

Sub ConsultarSaldosPorFecha(Optional v_fecha As String = "")
    ' Conecta a PostgreSQL y consulta los saldos de viaticos por empleado a una fecha especifica
    ' Parametro opcional:
    '   v_fecha: Fecha de corte (formato: 'YYYY-MM-DD', ej: '2025-01-31')
    '   Si no se proporciona, solicita la fecha mediante InputBox
    ' Los resultados se muestran en una hoja de Excel con una linea por empleado
    
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
    Dim paramFecha As String
    Dim fechaInput As String
    Application.ScreenUpdating = False
    
    ' Si no se proporciona fecha, solicitarla
    If v_fecha = "" Then
        fechaInput = InputBox("Ingrese la fecha de corte para calcular los saldos (formato: YYYY-MM-DD):", "Consulta Saldos por Fecha", Format(Date, "yyyy-mm-dd"))
        If fechaInput = "" Then
            MsgBox "Operacion cancelada. No se ingreso fecha.", vbInformation, "Consulta Cancelada"
            Exit Sub
        End If
        v_fecha = fechaInput
    End If
    
    ' Validar formato de fecha (formato basico)
    If Len(v_fecha) <> 10 Or Mid(v_fecha, 5, 1) <> "-" Or Mid(v_fecha, 8, 1) <> "-" Then
        MsgBox "Formato de fecha invalido. Use el formato YYYY-MM-DD (ej: 2025-01-31)", vbExclamation, "Error de Formato"
        Exit Sub
    End If
    
    ' Parametros de conexion
    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    ' Contrasena en formato ASCII (generada con convertir_contrasena.py)
    contrasena = ASCIIaTexto("65,100,109,105,110,83,111,108,105,100,50,48,50,53")
    
    ' Preparar parametro de fecha para la consulta SQL
    paramFecha = "'" & v_fecha & "'"
    
    ' Consulta SQL para obtener saldos por empleado a una fecha especifica
    ' Agrupa por empleado y calcula el saldo final hasta esa fecha
    Dim sqlWith As String
    Dim sqlSelect As String
    Dim sqlFrom As String
    Dim sqlWhere As String
    Dim sqlGroupBy As String
    Dim sqlOrder As String
    
    ' Parte 1: WITH params
    sqlWith = "WITH params AS (" & vbCrLf & _
              "  SELECT " & paramFecha & "::date AS v_fecha" & vbCrLf & _
              ")," & vbCrLf & _
              "transacciones_hasta_fecha AS (" & vbCrLf & _
              "  SELECT " & vbCrLf & _
              "    t.empleado," & vbCrLf & _
              "    t.fechaaplicacion::timestamp," & vbCrLf & _
              "    t.codigo," & vbCrLf & _
              "    CASE " & vbCrLf & _
              "      WHEN t.tipodocumento = 'CONSIGNACION' THEN t.valor::numeric" & vbCrLf & _
              "      WHEN t.tipodocumento = 'LEGALIZACION' THEN -(t.valor::numeric)" & vbCrLf & _
              "      ELSE 0" & vbCrLf & _
              "    END AS movimiento," & vbCrLf & _
              "    COALESCE(c.nombreempleado, l.nombreempleado) AS nombreempleado" & vbCrLf & _
              "  FROM transaccionviaticos t" & vbCrLf & _
              "  CROSS JOIN params p" & vbCrLf & _
              "  LEFT JOIN consignacion c" & vbCrLf & _
              "    ON c.codigoconsignacion = t.numerodocumento" & vbCrLf & _
              "  LEFT JOIN legalizacion l" & vbCrLf & _
              "    ON l.codigolegalizacion = t.numerodocumento" & vbCrLf & _
              "  WHERE t.fechaaplicacion::timestamp <= (p.v_fecha::timestamp + INTERVAL '1 day' - INTERVAL '1 second')" & vbCrLf & _
              ")" & vbCrLf
    
    ' Parte 2: SELECT con saldo calculado
    sqlSelect = "SELECT " & vbCrLf & _
                "  t.empleado AS ""CEDULA""," & vbCrLf & _
                "  MAX(t.nombreempleado) AS ""EMPLEADO""," & vbCrLf & _
                "  SUM(t.movimiento) AS ""SALDO""" & vbCrLf
    
    ' Parte 3: FROM
    sqlFrom = "FROM transacciones_hasta_fecha t" & vbCrLf
    
    ' Parte 4: GROUP BY
    sqlGroupBy = "GROUP BY t.empleado" & vbCrLf
    
    ' Parte 5: ORDER BY
    sqlOrder = "ORDER BY t.empleado ASC"
    
    ' Construir consulta completa
    consultaSQL = sqlWith & sqlSelect & sqlFrom & sqlGroupBy & sqlOrder
    
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
    Debug.Print "DIAGNOSTICO DE RENDIMIENTO - SALDOS POR FECHA"
    Debug.Print "=========================================="
    Debug.Print "Fecha de corte: " & v_fecha
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
    Debug.Print "Ejecutando consulta de saldos por fecha..."
    Dim tiempoAntesConsulta As Double
    tiempoAntesConsulta = Timer
    rs.Open consultaSQL, conn
    tiempoConsulta = Timer - tiempoAntesConsulta
    Debug.Print "Consulta ejecutada en " & Format(tiempoConsulta, "0.00") & " segundos"
    
    ' Crear o limpiar hoja de resultados
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Saldos_Por_Fecha")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Saldos_Por_Fecha"
    Else
        ' Limpiar solo desde la fila 7 en adelante para preservar las filas 1-6
        ws.Range("A7:ZZ" & ws.Rows.Count).Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Guardar numero de columnas
    Dim numColumnas As Integer
    numColumnas = rs.Fields.Count
    
    ' Guardar nombres de columnas en un array para uso posterior
    Dim nombresColumnas() As String
    ReDim nombresColumnas(0 To numColumnas - 1)
    For columna = 0 To numColumnas - 1
        nombresColumnas(columna) = rs.Fields(columna).Name
    Next columna
    
    ' Identificar indices de columnas importantes
    Dim indiceColumnaCedula As Integer
    Dim indiceColumnaSaldo As Integer
    indiceColumnaCedula = -1
    indiceColumnaSaldo = -1
    For columna = 0 To numColumnas - 1
        If UCase(nombresColumnas(columna)) = "CEDULA" Then
            indiceColumnaCedula = columna
        End If
        If UCase(nombresColumnas(columna)) = "SALDO" Then
            indiceColumnaSaldo = columna
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
    
    ' Formatear columnas
    If totalRegistros > 0 Then
        ' Formatear la columna de cedula como numero si existe
        If indiceColumnaCedula >= 0 Then
            Dim letraColumnaCedula As String
            Dim numColCedula As Integer
            numColCedula = indiceColumnaCedula + 1
            letraColumnaCedula = Split(ws.Cells(1, numColCedula).Address, "$")(1)
            ws.Range(letraColumnaCedula & filaInicial & ":" & letraColumnaCedula & ultimaFilaConDatos).NumberFormat = "0"
            ' Convertir valores de texto a numero si es necesario
            Dim celda As Range
            For Each celda In ws.Range(letraColumnaCedula & filaInicial & ":" & letraColumnaCedula & ultimaFilaConDatos)
                If Not IsEmpty(celda.Value) And IsNumeric(celda.Value) Then
                    celda.Value = CDbl(celda.Value)
                End If
            Next celda
        End If
        
        ' Formatear la columna de saldo como moneda si existe
        If indiceColumnaSaldo >= 0 Then
            Dim letraColumnaSaldo As String
            Dim numColSaldo As Integer
            numColSaldo = indiceColumnaSaldo + 1
            letraColumnaSaldo = Split(ws.Cells(1, numColSaldo).Address, "$")(1)
            ws.Range(letraColumnaSaldo & filaInicial & ":" & letraColumnaSaldo & ultimaFilaConDatos).NumberFormat = "$#,##0;[Red]-$#,##0"
        End If
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
    ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TablaSaldosPorFecha"
    On Error GoTo ErrorHandler
    
    ' Formatear tabla de Excel
    With ws.ListObjects("TablaSaldosPorFecha")
        .TableStyle = "TableStyleMedium9"
        .ShowAutoFilter = True
    End With
    
    ' Centrar encabezados
    ws.Range("A7:" & ultimaColumna & "7").HorizontalAlignment = xlCenter
    
    ' Ajustar la altura de los encabezados
    ws.Rows("7:7").RowHeight = 30
    ' Definir ancho de las columnas
    ws.Columns("A").ColumnWidth = 15 ' CEDULA
    ws.Columns("B").ColumnWidth = 40 ' EMPLEADO
    ws.Columns("C").ColumnWidth = 20 ' SALDO
    
    ws.Range("A7:" & ultimaColumna & "7").WrapText = True
    ws.Range("A7:" & ultimaColumna & "7").VerticalAlignment = xlCenter
    
    ' Agregar formula de total en la fila 6
    If indiceColumnaSaldo >= 0 Then
        Dim letraColumnaSaldoTotal As String
        Dim numColSaldoTotal As Integer
        numColSaldoTotal = indiceColumnaSaldo + 1
        letraColumnaSaldoTotal = Split(ws.Cells(1, numColSaldoTotal).Address, "$")(1)
        ws.Range(letraColumnaSaldoTotal & "6").Value = "TOTAL"
        ws.Range(letraColumnaSaldoTotal & "6").Font.Bold = True
        ws.Range(letraColumnaSaldoTotal & "6").HorizontalAlignment = xlRight
        ws.Range(letraColumnaSaldoTotal & "6").NumberFormat = "$#,##0;[Red]-$#,##0"
        ws.Range(letraColumnaSaldoTotal & "6").FormulaR1C1 = "=+SUBTOTAL(9,TablaSaldosPorFecha[SALDO])"
    End If
    
    ' Agregar informacion de fecha en la fila 1
    ws.Range("A1").Value = "SALDOS DE VIATICOS AL " & Format(CDate(v_fecha), "dd/mm/yyyy")
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 14
    
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
                   "Fecha de corte: " & Format(CDate(v_fecha), "dd/mm/yyyy") & vbCrLf & _
                   "Total de empleados: " & totalRegistros & vbCrLf & _
                   "Tiempo total: " & Format(tiempoTotal, "0.00") & " segundos" & vbCrLf & vbCrLf & _
                   "Desglose de tiempos:" & vbCrLf & _
                   "  - Conexion: " & Format(tiempoConexion, "0.00") & " s" & vbCrLf & _
                   "  - Consulta SQL: " & Format(tiempoConsulta, "0.00") & " s" & vbCrLf & _
                   "  - Extraccion: " & Format(tiempoExtraccion, "0.00") & " s" & vbCrLf & _
                   "  - Formato: " & Format(tiempoFormato, "0.00") & " s" & vbCrLf & vbCrLf & _
                   "Los resultados se encuentran en la hoja 'Saldos_Por_Fecha'" & vbCrLf & _
                   "Ver ventana Inmediato (Ctrl+G) para mas detalles."
    
    MsgBox mensajeExito, vbInformation, "Consulta Saldos por Fecha"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
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
    Application.ScreenUpdating = True
    MsgBox mensajeError, vbCritical, "Error en Consulta Saldos por Fecha"
End Sub

Sub ConsultarSaldosPorFechaInteractivo()
    ' Ejecuta la consulta solicitando la fecha mediante InputBox
    ' Util para obtener un reporte de saldos a una fecha especifica
    
    Call ConsultarSaldosPorFecha("")
    
End Sub

Sub ActualizarTablaPowerQuerySaldos()
    ' Actualiza la tabla de Power Query llamada "TablaSaldosPorFecha_1"
    ' Esta macro busca la tabla en todas las hojas del libro y la actualiza
    ' Compatible con Excel 2016+ y versiones anteriores
    
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim queryTable As QueryTable
    Dim tablaEncontrada As Boolean
    Dim nombreTabla As String
    Dim tiempoInicio As Double
    Dim tiempoTotal As Double
    Dim hojaEncontrada As String
    
    nombreTabla = "TablaSaldosPorFecha_1"
    tablaEncontrada = False
    tiempoInicio = Timer
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Buscando tabla de Power Query..."
    Debug.Print "=========================================="
    Debug.Print "ACTUALIZACION DE TABLA POWER QUERY"
    Debug.Print "=========================================="
    Debug.Print "Buscando tabla: " & nombreTabla
    
    On Error Resume Next
    
    ' Metodo 1: Buscar la tabla ListObject en todas las hojas
    For Each ws In ThisWorkbook.Worksheets
        Set tbl = Nothing
        Set tbl = ws.ListObjects(nombreTabla)
        If Not tbl Is Nothing Then
            tablaEncontrada = True
            hojaEncontrada = ws.Name
            Debug.Print "Tabla encontrada en la hoja: " & hojaEncontrada
            
            ' Intentar actualizar a traves de QueryTable (metodo mas comun)
            Set queryTable = Nothing
            Set queryTable = tbl.QueryTable
            If Not queryTable Is Nothing Then
                Application.StatusBar = "Actualizando tabla de Power Query..."
                Debug.Print "Actualizando tabla a traves de QueryTable..."
                queryTable.Refresh BackgroundQuery:=False
                Debug.Print "Tabla actualizada correctamente via QueryTable"
                Exit For
            Else
                ' Si no tiene QueryTable, puede ser una tabla de Power Query moderna
                ' Intentar actualizar usando el metodo de conexion
                Application.StatusBar = "Actualizando tabla de Power Query (metodo alternativo)..."
                Debug.Print "Intentando actualizar usando metodo alternativo..."
                ' Refrescar todas las conexiones del libro (puede ser mas lento)
                ThisWorkbook.RefreshAll
                Debug.Print "Tabla actualizada usando RefreshAll"
                Exit For
            End If
        End If
    Next ws
    
    ' Metodo 2: Si no se encontro la tabla, intentar actualizar todas las consultas
    If Not tablaEncontrada Then
        Debug.Print "Tabla no encontrada como ListObject, intentando RefreshAll..."
        Application.StatusBar = "Actualizando todas las consultas de Power Query..."
        ThisWorkbook.RefreshAll
        tablaEncontrada = True
        Debug.Print "Todas las consultas actualizadas (RefreshAll)"
    End If
    
    On Error GoTo ErrorHandler
    
    tiempoTotal = Timer - tiempoInicio
    
    If tablaEncontrada Then
        Application.StatusBar = "Tabla actualizada correctamente"
        Debug.Print "=========================================="
        Debug.Print "Tiempo de actualizacion: " & Format(tiempoTotal, "0.00") & " segundos"
        If hojaEncontrada <> "" Then
            Debug.Print "Hoja: " & hojaEncontrada
        End If
        Debug.Print "=========================================="
        
        Dim mensajeExito As String
        mensajeExito = "Tabla de Power Query actualizada correctamente." & vbCrLf & vbCrLf & _
                       "Tabla: " & nombreTabla & vbCrLf & _
                       "Tiempo: " & Format(tiempoTotal, "0.00") & " segundos"
        If hojaEncontrada <> "" Then
            mensajeExito = mensajeExito & vbCrLf & "Hoja: " & hojaEncontrada
        End If
        
        MsgBox mensajeExito, vbInformation, "Actualizacion Completada"
    Else
        MsgBox "No se encontro la tabla de Power Query: " & nombreTabla & vbCrLf & _
               "Verifique que la tabla exista en el libro.", _
               vbExclamation, "Tabla No Encontrada"
    End If
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub
    
ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    
    Dim mensajeError As String
    mensajeError = "Error al actualizar la tabla de Power Query:" & vbCrLf & vbCrLf & _
                   "Numero de error: " & Err.Number & vbCrLf & _
                   "Descripcion: " & Err.Description & vbCrLf & vbCrLf & _
                   "Verifique que:" & vbCrLf & _
                   "1. La tabla existe en el libro" & vbCrLf & _
                   "2. El nombre de la tabla es correcto: " & nombreTabla & vbCrLf & _
                   "3. La conexion a la base de datos esta disponible" & vbCrLf & _
                   "4. Tiene permisos para actualizar consultas"
    
    Debug.Print "ERROR: " & mensajeError
    MsgBox mensajeError, vbCritical, "Error en Actualizacion"
End Sub
