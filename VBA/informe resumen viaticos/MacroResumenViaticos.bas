Attribute VB_Name = "ModuloResumenViaticos"
' ====================================================================
' Resumen viáticos -> hoja BD_Viaticos (tabla desde fila 7, datos desde 8)
' Filtro de áreas: hoja "AJUSTE DE RUTA", rango Filtro_areas $AH$4:$AI$9
' Base de datos: solid (PostgreSQL) - misma cadena que MacroInformeRubro
' ====================================================================

Option Explicit

Private Const HOJA_DATOS As String = "BD_Viaticos"
Private Const HOJA_FILTRO As String = "AJUSTE DE RUTA"
Private Const FILA_ENCABEZADO As Long = 7
Private Const PRIMERA_FILA_DATOS As Long = 8
Private Const ULTIMA_COL As Long = 20
Private Const COL_FILTRO_AREA As Long = 34
Private Const COL_FILTRO_ACTIVA As Long = 35
Private Const FILA_FILTRO_INI As Long = 4
Private Const FILA_FILTRO_FIN As Long = 9

Private Function HexToString(ByVal hexVal As String) As String
    Dim i As Long
    Dim res As String
    res = ""
    For i = 1 To Len(hexVal) Step 2
        res = res & Chr(Val("&H" & Mid(hexVal, i, 2)))
    Next i
    HexToString = res
End Function

Private Function SqlLiteralText(ByVal s As String) As String
    SqlLiteralText = Replace(Trim(s), "'", "''")
End Function

Private Function CadenaConexionSolid() As String
    CadenaConexionSolid = "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & _
        HexToString("41646D696E536F6C696432303235") & ";"
End Function

Private Function ObtenerTablaBDViaticos(ByVal ws As Worksheet) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set ObtenerTablaBDViaticos = ws.ListObjects("BD_Viaticos")
    If Not ObtenerTablaBDViaticos Is Nothing Then Exit Function
    Set ObtenerTablaBDViaticos = ws.ListObjects("Tabla_BD_Viaticos")
    If Not ObtenerTablaBDViaticos Is Nothing Then Exit Function
    Set ObtenerTablaBDViaticos = ws.ListObjects("tbl_BD_Viaticos")
    If Not ObtenerTablaBDViaticos Is Nothing Then Exit Function
    On Error GoTo 0
    For Each lo In ws.ListObjects
        If Not lo.HeaderRowRange Is Nothing Then
            If lo.HeaderRowRange.Row = FILA_ENCABEZADO And lo.Range.Column = 1 Then
                Set ObtenerTablaBDViaticos = lo
                Exit Function
            End If
        End If
    Next lo
End Function

Private Function ConstruirListaAreasActivas(ByVal wsFiltro As Worksheet, ByRef mensajeAdvertencia As String) As String
    Dim r As Long
    Dim areaVal As String
    Dim activaVal As Variant
    Dim partes As Collection
    Dim esCabecera As Boolean
    Dim lista As String

    Set partes = New Collection
    mensajeAdvertencia = ""

    For r = FILA_FILTRO_INI To FILA_FILTRO_FIN
        areaVal = Trim(CStr(wsFiltro.Cells(r, COL_FILTRO_AREA).Value2))
        activaVal = wsFiltro.Cells(r, COL_FILTRO_ACTIVA).Value2

        esCabecera = (UCase(areaVal) = "AREAS" Or UCase(areaVal) = "AREA")
        If esCabecera Then GoTo Siguiente

        If Len(areaVal) = 0 Then
            If EsActiva(activaVal) Then
                mensajeAdvertencia = mensajeAdvertencia & "Fila " & r & ": ACTIVA sin área (se ignora)." & vbCrLf
            End If
            GoTo Siguiente
        End If

        If Not EsActiva(activaVal) Then GoTo Siguiente

        On Error Resume Next
        partes.Add areaVal, UCase(areaVal)
        If Err.Number <> 0 Then Err.Clear
        On Error GoTo 0
Siguiente:
    Next r

    If partes.Count = 0 Then
        ConstruirListaAreasActivas = ""
        Exit Function
    End If

    Dim i As Long
    Dim s As String
    For i = 1 To partes.Count
        s = s & "'" & SqlLiteralText(UCase$(CStr(partes(i)))) & "',"
    Next i
    If Len(s) > 0 Then s = Left(s, Len(s) - 1)
    ConstruirListaAreasActivas = s
End Function

Private Function EsActiva(ByVal v As Variant) As Boolean
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        EsActiva = (CLng(v) = 1)
        Exit Function
    End If
    EsActiva = (Trim(CStr(v)) = "1" Or UCase(Trim(CStr(v))) = "TRUE" Or UCase(Trim(CStr(v))) = "SI")
End Function

Private Function SqlResumenViaticos(ByVal listaInSql As String) As String
    Dim s As String
    s = ""
    s = s & "SELECT " & vbCrLf
    s = s & "    b.fechaaplicacion_entrega::DATE AS ""FECHA ENTREGA REPORTE""," & vbCrLf
    s = s & "    b.nombre_upper AS ""NOMBRE""," & vbCrLf
    s = s & "    b.empleado::BIGINT AS ""DOCUMENTO DE IDENTIDAD""," & vbCrLf
    s = s & "    b.ot_cc AS ""OT-CC""," & vbCrLf
    s = s & "    b.fe_real AS ""FECHA REAL DEL GASTO""," & vbCrLf
    s = s & "    EXTRACT(YEAR FROM b.fe_real)::INTEGER AS ""Año""," & vbCrLf
    s = s & "    CASE EXTRACT(MONTH FROM b.fe_real)::INTEGER" & vbCrLf
    s = s & "        WHEN 1 THEN 'enero' WHEN 2 THEN 'febrero' WHEN 3 THEN 'marzo' WHEN 4 THEN 'abril'" & vbCrLf
    s = s & "        WHEN 5 THEN 'mayo' WHEN 6 THEN 'junio' WHEN 7 THEN 'julio' WHEN 8 THEN 'agosto'" & vbCrLf
    s = s & "        WHEN 9 THEN 'septiembre' WHEN 10 THEN 'octubre' WHEN 11 THEN 'noviembre' WHEN 12 THEN 'diciembre'" & vbCrLf
    s = s & "    END AS ""Mes""," & vbCrLf
    s = s & "    EXTRACT(WEEK FROM b.fe_real)::INTEGER AS ""SEMANA DEL AÑO""," & vbCrLf
    s = s & "    b.obra AS ""OBRA""," & vbCrLf
    s = s & "    b.ciudad AS ""CIUDAD""," & vbCrLf
    s = s & "    b.centrocosto AS ""CENTRO DE COSTO""," & vbCrLf
    s = s & "    b.subcentrocosto AS ""SUB CENTRO""," & vbCrLf
    s = s & "    b.categoria AS ""DESCRIPCION""," & vbCrLf
    s = s & "    b.v_conf AS ""VALOR TOTAL FACTURA""," & vbCrLf
    s = s & "    b.v_sin AS ""VALOR SIN FACTURA""," & vbCrLf
    s = s & "    (b.v_conf + b.v_sin) AS ""APROBADO""," & vbCrLf
    s = s & "    (b.v_conf + b.v_sin) AS ""SOLICITADO""," & vbCrLf
    s = s & "    b.codigolegalizacion AS ""RADICADO""," & vbCrLf
    s = s & "    SPLIT_PART(b.codigolegalizacion, '-', 1) AS ""AREA""," & vbCrLf
    s = s & "    0::BIGINT AS ""DIFERENCIA""" & vbCrLf
    s = s & "FROM (" & vbCrLf
    s = s & "    SELECT " & vbCrLf
    s = s & "        l.fechaaplicacion AS fechaaplicacion_entrega," & vbCrLf
    s = s & "        UPPER(l.nombreempleado) AS nombre_upper," & vbCrLf
    s = s & "        l.empleado," & vbCrLf
    s = s & "        l.codigolegalizacion," & vbCrLf
    s = s & "        CASE WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot ELSE 'C' || ln.centrocosto END AS ot_cc," & vbCrLf
    s = s & "        COALESCE(ln.fecharealgasto, l.fechaaplicacion)::DATE AS fe_real," & vbCrLf
    s = s & "        o.cliente AS obra," & vbCrLf
    s = s & "        o.ciudad," & vbCrLf
    s = s & "        ln.centrocosto," & vbCrLf
    s = s & "        ln.subcentrocosto," & vbCrLf
    s = s & "        ln.categoria," & vbCrLf
    s = s & "        COALESCE(ln.valorconfactura, 0)::BIGINT AS v_conf," & vbCrLf
    s = s & "        COALESCE(ln.valorsinfactura, 0)::BIGINT AS v_sin" & vbCrLf
    s = s & "    FROM legalizacion l" & vbCrLf
    s = s & "    INNER JOIN linealegalizacion ln ON l.codigo = ln.legalizacion" & vbCrLf
    s = s & "    LEFT JOIN (SELECT DISTINCT ON (numero) numero, cliente, ciudad FROM otviaticos) o ON ln.ot = o.numero" & vbCrLf
    s = s & "    WHERE UPPER(TRIM(SPLIT_PART(l.codigolegalizacion, '-', 1))) IN (" & listaInSql & ")" & vbCrLf
    s = s & ") b" & vbCrLf
    s = s & "ORDER BY b.fechaaplicacion_entrega DESC;"
    SqlResumenViaticos = s
End Function

Private Sub AplicarFormatoSalida(ByVal ws As Worksheet)
    ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("C:C").NumberFormat = "@"
    ws.Columns("E:E").NumberFormat = "dd/mm/yyyy"
    ws.Columns("N:Q").NumberFormat = "$#,##0"
    ws.Columns("T:T").NumberFormat = "$#,##0"
End Sub

Public Sub EjecutarResumenViaticos()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim wsFiltro As Worksheet
    Dim tbl As ListObject
    Dim sql As String
    Dim listaIn As String
    Dim adv As String
    Dim totalRegs As Long
    Dim t0 As Double
    Dim arr As Variant
    Dim ultimaFila As Long

    t0 = Timer
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    On Error GoTo ErrorHandler

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsFiltro = ThisWorkbook.Worksheets(HOJA_FILTRO)
    Set ws = ThisWorkbook.Worksheets(HOJA_DATOS)
    
    ' Limpiar filtros existentes para asegurar visibilidad de nuevos datos
    If ws.FilterMode Then ws.ShowAllData
    
    Set tbl = ObtenerTablaBDViaticos(ws)
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 1, , "No se encontró la tabla en '" & HOJA_DATOS & "'. Cree una tabla Excel con encabezados en la fila " & FILA_ENCABEZADO & " o asígnele el nombre BD_Viaticos."
    End If

    listaIn = ConstruirListaAreasActivas(wsFiltro, adv)
    If Len(listaIn) = 0 Then
        MsgBox "No hay áreas activas en '" & HOJA_FILTRO & "' (" & RangeAddressFiltro() & ")." & vbCrLf & vbCrLf & _
               "Marque 1 en ACTIVAS para al menos un área con nombre.", vbExclamation, "Resumen viáticos"
        GoTo Cleanup
    End If

    If Len(adv) > 0 Then
        MsgBox "Advertencias (área vacía con ACTIVA):" & vbCrLf & vbCrLf & adv, vbInformation, "Resumen viáticos"
    End If

    sql = SqlResumenViaticos(listaIn)

    Application.StatusBar = "Conectando a PostgreSQL..."
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 300
    conn.Open CadenaConexionSolid()

    Application.StatusBar = "Ejecutando consulta..."
    rs.CursorLocation = 3
    rs.Open sql, conn, 3, 1

    Application.StatusBar = "Volcando datos en memoria..."
    If rs.EOF Then
        totalRegs = 0
    Else
        arr = rs.GetRows()
        totalRegs = UBound(arr, 2) - LBound(arr, 2) + 1
    End If
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing

    Application.StatusBar = "Escribiendo en BD_Viaticos..."
    If Not tbl.DataBodyRange Is Nothing Then
        tbl.DataBodyRange.ClearContents
    End If

    If totalRegs > 0 Then
        EscribirGetRowsEnRango ws, arr, PRIMERA_FILA_DATOS, ULTIMA_COL
        ultimaFila = PRIMERA_FILA_DATOS + totalRegs - 1
    Else
        ultimaFila = PRIMERA_FILA_DATOS
    End If

    tbl.Resize ws.Range(ws.Cells(FILA_ENCABEZADO, 1), ws.Cells(ultimaFila, ULTIMA_COL))
    Call AplicarFormatoSalida(ws)

    Application.StatusBar = False
    MsgBox "Resumen viáticos actualizado." & vbCrLf & _
           "Registros: " & totalRegs & vbCrLf & _
           "Tiempo: " & Format(Timer - t0, "0.0") & " s", vbInformation

Cleanup:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description & " (" & Err.Number & ")", vbCritical, "Resumen viáticos"
    On Error Resume Next
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Set rs = Nothing
    Set conn = Nothing
End Sub

Private Function RangeAddressFiltro() As String
    RangeAddressFiltro = "$AH$" & FILA_FILTRO_INI & ":$AI$" & FILA_FILTRO_FIN
End Function

Private Sub EscribirGetRowsEnRango(ByVal ws As Worksheet, ByVal arr As Variant, ByVal filaIni As Long, ByVal numCols As Long)
    Dim nR As Long
    Dim nC As Long
    Dim r As Long
    Dim c As Long
    Dim outArr() As Variant
    Dim r0 As Long
    Dim c0 As Long

    r0 = LBound(arr, 1)
    c0 = LBound(arr, 2)
    nC = UBound(arr, 1) - r0 + 1
    nR = UBound(arr, 2) - c0 + 1

    If nC <> numCols Then
        Err.Raise vbObjectError + 2, , "Número de columnas inesperado en el resultado."
    End If

    ReDim outArr(1 To nR, 1 To nC)
    For r = 1 To nR
        For c = 1 To nC
            outArr(r, c) = arr(r0 + c - 1, c0 + r - 1)
        Next c
    Next r

    ws.Cells(filaIni, 1).Resize(nR, nC).Value2 = outArr
End Sub
