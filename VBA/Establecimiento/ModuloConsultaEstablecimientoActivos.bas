'Attribute VB_Name = "ModuloConsultaEstablecimientoActivos"
' Macro variante: consulta establecimiento solo con registros ACTIVOS (ESTADO = 'A').
' Base de datos: solid. Misma estructura que ModuloConsultaEstablecimiento.
' Autor: Ing Luis Enrique Rojas | Fecha: 2026

Option Explicit

Private Const ADOPEN_FORWARDONLY As Long = 0
Private Const ADLOCK_READONLY As Long = 1
Private Const ADCMDTEXT As Long = 1
Private Const ADCURSORSERVER As Long = 2

Private Function Password(cadenaDatos As String) As String
    Dim segmentos() As String
    Dim i As Long
    Dim textoFinal As String
    segmentos = Split(cadenaDatos, ",")
    textoFinal = ""
    For i = LBound(segmentos) To UBound(segmentos)
        textoFinal = textoFinal & Chr(Val(Trim(segmentos(i))))
    Next i
    Password = textoFinal
End Function

' Carga la consulta base desde ModuloConsultaEstablecimiento y aplica filtro WHERE ESTADO = 'A' (solo activos).
' Fuente: consulta_establecimiento.sql (o ConsultaSQLEmbebida del modulo principal).
Private Function CargarConsultaSQLActivos() As String
    Dim consultaSQL As String
    Dim posOrder As Long
    Dim whereActivos As String

    consultaSQL = ModuloConsultaEstablecimiento.CargarConsultaSQL()
    posOrder = InStr(1, consultaSQL, "ORDER BY", vbTextCompare)
    If posOrder > 0 Then
        whereActivos = vbCrLf & "WHERE (CASE WHEN B.estado = 'Activo' THEN 'A' WHEN B.estado = 'Inactivo' THEN 'RET' WHEN B.estado IS NOT NULL AND B.estado <> '' THEN B.estado::text WHEN C.estado = 'Activo' THEN 'A' WHEN C.estado = 'Terminado' THEN 'RET' ELSE 'RET' END) = 'A' " & vbCrLf
        consultaSQL = Left(consultaSQL, posOrder - 1) & whereActivos & Mid(consultaSQL, posOrder)
    End If
    CargarConsultaSQLActivos = consultaSQL
End Function

Sub ConsultarEstablecimientoActivos()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim cadenaConexion As String
    Dim consultaSQL As String
    Dim servidor As String
    Dim puerto As String
    Dim baseDatos As String
    Dim usuario As String
    Dim contrasena As String
    Dim mensajeError As String
    Dim totalRegistros As Long
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim tiempoEjecucion As Double
    Dim numColumnas As Integer
    Dim nombresColumnas() As String
    Dim headerRow() As Variant
    Dim columna As Long
    Dim filaInicial As Long
    Dim ultimaFilaConDatos As Long
    Dim rangoTabla As Range
    Dim ultimaColumna As String
    Dim tblExistente As Object
    Dim colFecha As Integer
    Dim nombreCol As String
    Dim esColumnaFecha As Boolean
    Dim colMon As Long
    Dim nombreColMon As String

    tiempoInicio = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    contrasena = Password("65,100,109,105,110,83,111,108,105,100,50,48,50,53")

    consultaSQL = CargarConsultaSQLActivos()

    cadenaConexion = "Driver={PostgreSQL Unicode(x64)};Server=" & servidor & ";Port=" & puerto & ";Database=" & baseDatos & ";Uid=" & usuario & ";Pwd=" & contrasena & ";"

    On Error GoTo ErrorHandler

    ' Hoja destino: ESTABLECIMIENTO_ACTIVOS
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ESTABLECIMIENTO_ACTIVOS")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ESTABLECIMIENTO_ACTIVOS"
    Else
        Set tblExistente = ws.ListObjects("TB_ESTABLECIMIENTO_ACTIVOS")
        If Not tblExistente Is Nothing Then
            On Error Resume Next
            tblExistente.AutoFilter.ShowAllData
            On Error GoTo ErrorHandler
            If Not tblExistente.DataBodyRange Is Nothing Then tblExistente.DataBodyRange.Delete
        Else
            ws.Range("A13:ZZ" & ws.Rows.Count).Clear
        End If
    End If
    On Error GoTo ErrorHandler

    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    Application.StatusBar = "Conectando (solo activos)..."
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 120
    conn.CursorLocation = ADCURSORSERVER
    conn.Open cadenaConexion

    Application.StatusBar = "Ejecutando consulta (solo activos)..."
    rs.CursorType = ADOPEN_FORWARDONLY
    rs.CursorLocation = ADCURSORSERVER
    rs.LockType = ADLOCK_READONLY
    rs.Open consultaSQL, conn, ADOPEN_FORWARDONLY, ADLOCK_READONLY, ADCMDTEXT

    numColumnas = rs.Fields.Count
    ReDim nombresColumnas(0 To numColumnas - 1)
    For columna = 0 To numColumnas - 1
        nombresColumnas(columna) = rs.Fields(columna).Name
    Next columna

    ReDim headerRow(1 To 1, 1 To numColumnas)
    For columna = 1 To numColumnas
        headerRow(1, columna) = nombresColumnas(columna - 1)
    Next columna
    With ws.Range(ws.Cells(12, 1), ws.Cells(12, numColumnas))
        .Value = headerRow
        .Font.Bold = True
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(12).RowHeight = 30

    filaInicial = 13
    ws.Range("A13").CopyFromRecordset rs

    ultimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFilaConDatos >= filaInicial Then
        totalRegistros = ultimaFilaConDatos - filaInicial + 1
    Else
        totalRegistros = 0
    End If

    If numColumnas <= 26 Then
        ultimaColumna = Chr(64 + numColumnas)
    Else
        ultimaColumna = Chr(64 + Int((numColumnas - 1) / 26)) & Chr(65 + ((numColumnas - 1) Mod 26))
    End If

    If totalRegistros > 0 Then
        Set rangoTabla = ws.Range("A12:" & ultimaColumna & (ultimaFilaConDatos))
        On Error Resume Next
        Set tblExistente = ws.ListObjects("TB_ESTABLECIMIENTO_ACTIVOS")
        If tblExistente Is Nothing Then
            ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TB_ESTABLECIMIENTO_ACTIVOS"
        Else
            tblExistente.Resize rangoTabla
        End If
        On Error GoTo ErrorHandler
        With ws.ListObjects("TB_ESTABLECIMIENTO_ACTIVOS")
            .TableStyle = "TableStyleMedium9"
            .ShowAutoFilter = True
        End With

        ws.Range("A13:" & ultimaColumna & ultimaFilaConDatos).NumberFormat = "General"
        For colFecha = 0 To numColumnas - 1
            nombreCol = nombresColumnas(colFecha)
            esColumnaFecha = (InStr(1, UCase(nombreCol), "FECHA") > 0 And InStr(1, UCase(nombreCol), "MES") = 0)
            If esColumnaFecha Then
                ws.Columns(colFecha + 1).NumberFormat = "dd/mm/yyyy"
                ws.Columns(colFecha + 1).HorizontalAlignment = xlCenter
            End If
        Next colFecha
        For colMon = 0 To numColumnas - 1
            nombreColMon = UCase(Trim(nombresColumnas(colMon)))
            If InStr(1, nombreColMon, "BASE VIATICOS") > 0 Or InStr(1, nombreColMon, "SALARIO AÑO") > 0 Or _
               InStr(1, nombreColMon, "AUX ALIMENTACION") > 0 Or InStr(1, nombreColMon, "AUX VIVIENDA") > 0 Or _
               InStr(1, nombreColMon, "RODAMIENTO") > 0 Or InStr(1, nombreColMon, "VALOR MAXIMO BONO") > 0 Or _
               InStr(1, nombreColMon, "CAPACIDAD ENDEUDAMIENTO") > 0 Or InStr(1, nombreColMon, "RIESGO ARL") > 0 Then
                ws.Range(ws.Cells(13, colMon + 1), ws.Cells(ultimaFilaConDatos, colMon + 1)).NumberFormat = "$#,##0;[Red]-$#,##0"
            End If
        Next colMon
    End If

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    tiempoFin = Timer
    tiempoEjecucion = tiempoFin - tiempoInicio
    Application.StatusBar = "Completado (activos): " & totalRegistros & " registros"
    ws.Range("F2").Value = "Tiempo: " & Round(tiempoEjecucion, 1) & " s"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Consulta completada (solo activos)." & vbCrLf & "Total de registros: " & totalRegistros, vbInformation, "Consulta Establecimiento - Activos"
    Exit Sub

ErrorHandler:
    mensajeError = "Error: " & Err.Number & vbCrLf & Err.Description
    Debug.Print mensajeError
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
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox mensajeError, vbCritical, "Consulta Establecimiento - Activos"
End Sub
