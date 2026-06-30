'Attribute VB_Name = "ModuloConteoInventario"
' ====================================================================
' MACROS PARA EXTRACCIÓN DE CONTEO FÍSICO DE INVENTARIO
' Base de datos: project_manager (Portal)
' Tabla: conteoinventario + asignacioninventario
' Driver: PostgreSQL Unicode(x64)
' ====================================================================

Option Explicit

' --- AMBIENTE ACTIVO ---
' Valores: "LOCAL", "DEV", "PROD"
Private mAmbiente As String

' --- UTILIDADES DE CIFRADO ---

Private Function HexToString(ByVal hexVal As String) As String
    Dim i As Long
    Dim res As String
    res = ""
    For i = 1 To Len(hexVal) Step 2
        res = res & Chr(Val("&H" & Mid(hexVal, i, 2)))
    Next i
    HexToString = res
End Function

' --- SWITCH AMBIENTE ---

Sub CambiarAmbiente()
    Dim opcion As String
    opcion = InputBox( _
        "Seleccione el ambiente de conexión:" & vbCrLf & vbCrLf & _
        "1 = LOCAL  (192.168.40.166:5432)" & vbCrLf & _
        "2 = DEV    (192.168.0.21:5435)" & vbCrLf & _
        "3 = PROD   (192.168.0.21:5433)" & vbCrLf & vbCrLf & _
        "Ambiente actual: " & IIf(mAmbiente = "", "NO DEFINIDO", mAmbiente), _
        "Seleccionar Ambiente", "2")
    
    Select Case Trim(opcion)
        Case "1": mAmbiente = "LOCAL"
        Case "2": mAmbiente = "DEV"
        Case "3": mAmbiente = "PROD"
        Case Else
            MsgBox "Opción no válida. Se mantiene: " & IIf(mAmbiente = "", "DEV (default)", mAmbiente), vbExclamation
            If mAmbiente = "" Then mAmbiente = "DEV"
            Exit Sub
    End Select
    
    MsgBox "Ambiente cambiado a: " & mAmbiente, vbInformation, "Ambiente Activo"
End Sub

Private Function ObtenerCadenaConexion() As String
    Dim srv As String, prt As String, db As String
    
    If mAmbiente = "" Then mAmbiente = "DEV"
    
    Select Case mAmbiente
        Case "LOCAL"
            srv = "192.168.40.166"
            prt = "5432"
            db = "project_manager"
        Case "DEV"
            srv = "192.168.0.21"
            prt = "5435"
            db = "project_manager_pruebas3"
        Case "PROD"
            srv = "192.168.0.21"
            prt = "5433"
            db = "project_manager"
    End Select
    
    ObtenerCadenaConexion = "Driver={PostgreSQL Unicode(x64)};" & _
        "Server=" & srv & ";" & _
        "Port=" & prt & ";" & _
        "Database=" & db & ";" & _
        "Uid=user;" & _
        "Pwd=" & HexToString("70617373776f72645f7365677572615f726566726964636f6c") & ";"
End Function

' --- SUBS DE CONTEO ---

Sub EjecutarConteo1()
    Call EjecutarConteo(1)
End Sub

Sub EjecutarConteo2()
    Call EjecutarConteo(2)
End Sub

Sub EjecutarConteo3()
    Call EjecutarConteo(3)
End Sub

' --- PROCESO PRINCIPAL GENÉRICO ---

Private Sub EjecutarConteo(ByVal ronda As Integer)

    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim totalRegs As Long
    Dim nombreHoja As String
    Dim nombreTabla As String
    Dim tbl As ListObject
    Dim tblRange As Range
    Dim filaInicio As Long
    Dim ultimaColumnaDatos As Long
    Dim ultimaColumnaTabla As Long
    Dim formulaConteoAud As String
    nombreHoja = "TOTAL BODEGAS DIGITACION (" & ronda & ")"
    
    Select Case ronda
        Case 1: nombreTabla = "TOTAL_BODEGAS_DIGITACION1"
        Case 2: nombreTabla = "TOTAL_BODEGAS_DIGITACION2"
        Case 3: nombreTabla = "TOTAL_BODEGAS_DIGITACION3"
    End Select
    filaInicio = 11
    Application.ScreenUpdating = False
    Application.Calculation = -4135 ' xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo ErrorHandler

    sql = ""
    sql = sql & "SELECT " & vbCrLf
    sql = sql & "    ROW_NUMBER() OVER (" & vbCrLf
    sql = sql & "        ORDER BY c.bodega, c.bloque, c.estante, c.nivel, c.codigo" & vbCrLf
    sql = sql & "    ) AS ""No.""," & vbCrLf
    sql = sql & "    a.id AS ""No. Planilla""," & vbCrLf
    sql = sql & "    '' AS ""Column3""," & vbCrLf
    sql = sql & "    c.b_siigo AS ""B. Siigo""," & vbCrLf
    sql = sql & "    c.bodega AS ""Bodega""," & vbCrLf
    sql = sql & "    c.bloque AS ""Bloque""," & vbCrLf
    sql = sql & "    c.estante AS ""Estante""," & vbCrLf
    sql = sql & "    c.nivel AS ""Nivel""," & vbCrLf
    sql = sql & "    '' AS ""Column9""," & vbCrLf
    sql = sql & "    c.codigo AS ""Codigo""," & vbCrLf
    sql = sql & "    c.descripcion AS ""Descripcion""," & vbCrLf
    sql = sql & "    c.unidad AS ""Und.""," & vbCrLf
    sql = sql & "    '' AS ""Column13""," & vbCrLf
    sql = sql & "    c.cant_c" & ronda & " AS ""Cant.""," & vbCrLf
    sql = sql & "    '' AS ""Column15""," & vbCrLf
    sql = sql & "    COALESCE(c.obs_c" & ronda & ", '') AS ""Observaciones:""," & vbCrLf
    sql = sql & "    COALESCE(a.numero_pareja::text, '') AS ""DIGITADOR"" " & vbCrLf
    sql = sql & "FROM conteoinventario c" & vbCrLf
    sql = sql & "LEFT JOIN LATERAL (" & vbCrLf
    sql = sql & "    SELECT ai.id, ai.numero_pareja" & vbCrLf
    sql = sql & "    FROM asignacioninventario ai" & vbCrLf
    sql = sql & "    WHERE ai.bodega  = c.bodega" & vbCrLf
    sql = sql & "      AND ai.bloque  = c.bloque" & vbCrLf
    sql = sql & "      AND ai.estante = c.estante" & vbCrLf
    sql = sql & "      AND ai.nivel   = c.nivel" & vbCrLf
    sql = sql & "    ORDER BY ai.id" & vbCrLf
    sql = sql & "    LIMIT 1" & vbCrLf
    sql = sql & ") a ON true" & vbCrLf
    sql = sql & "WHERE c.user_c" & ronda & " IS NOT NULL AND c.user_c" & ronda & " <> ''" & vbCrLf
    sql = sql & "ORDER BY c.bodega, c.bloque, c.estante, c.nivel, c.codigo;"
    
    Dim t0 As Double, tSQL As Double, tWrite As Double
    t0 = Timer
    
    Application.StatusBar = "Conectando a PostgreSQL (" & mAmbiente & ")..."
    Set conn = CreateObject("ADODB.Connection")
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 300
    conn.Open ObtenerCadenaConexion()
    
    Application.StatusBar = "Ejecutando consulta Conteo " & ronda & "..."
    Set rs = CreateObject("ADODB.Recordset")
    rs.CursorLocation = 3 ' adUseClient
    rs.Open sql, conn, 0, 1 ' adOpenForwardOnly, adLockReadOnly
    
    tSQL = Timer - t0

    Set ws = ThisWorkbook.Worksheets(nombreHoja)
    Set tbl = ws.ListObjects(nombreTabla)
    
    Dim colCount As Long
    colCount = tbl.Range.Columns.Count
    Dim filasAnteriores As Long
    
    Application.StatusBar = "Limpiando datos anteriores..."
    If Not tbl.DataBodyRange Is Nothing Then
        filasAnteriores = tbl.ListRows.Count
        ws.Range(ws.Cells(filaInicio + 1, 1), ws.Cells(filaInicio + filasAnteriores, 17)).ClearContents
    End If
    
    ws.Columns(10).NumberFormat = "@"
    
    Application.StatusBar = "Escribiendo datos en Excel..."
    If Not rs.EOF Then
        ws.Cells(filaInicio + 1, 1).CopyFromRecordset rs
        totalRegs = ws.Cells(ws.Rows.Count, 1).End(-4162).Row - filaInicio
    Else
        totalRegs = 0
    End If
    
    If totalRegs > 0 Then
        tbl.Resize ws.Range(ws.Cells(filaInicio, 1), ws.Cells(filaInicio + totalRegs, colCount))
    Else
        tbl.Resize ws.Range(ws.Cells(filaInicio, 1), ws.Cells(filaInicio + 1, colCount))
    End If
    ws.Columns(14).NumberFormat = "#,##0.0"
    
    tWrite = Timer - t0 - tSQL

    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    
    Application.ScreenUpdating = True
    Application.Calculation = -4105 ' xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox "Conteo " & ronda & " descargado con éxito." & vbCrLf & _
           "Ambiente: " & mAmbiente & vbCrLf & _
           "Registros importados: " & totalRegs & vbCrLf & _
           "Tiempo SQL: " & Format(tSQL, "0.0") & "s" & vbCrLf & _
           "Tiempo escritura: " & Format(tWrite, "0.0") & "s" & vbCrLf & _
           "Tiempo total: " & Format(Timer - t0, "0.0") & "s", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = -4105 ' xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Error en Conteo " & ronda & ": " & Err.Description & vbCrLf & _
           "Ambiente: " & mAmbiente, vbCritical
    
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing

End Sub
