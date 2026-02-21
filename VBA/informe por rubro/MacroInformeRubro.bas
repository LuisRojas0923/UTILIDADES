Attribute VB_Name = "ModuloInformeRubro"
' Macro para generar el reporte DETALLADO POR RUBRO
' Hoja salida: "INFORME_RUBRO"
' Base de datos: solid
' Servidor: 192.168.0.21:5432

Option Explicit

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

Sub EjecutarInformeRubro()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim filaEncabezado As Integer
    Dim totalRegs As Long
    Dim col As Integer
    Dim tbl As ListObject
    
    ' Definimos fila de inicio
    filaEncabezado = 8
    
    Application.ScreenUpdating = False
    
    ' 1. CONFIGURAR HOJA
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_Viaticos")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "BD_Viaticos"
    End If
    On Error GoTo ErrorHandler
    
    ' 2. LIMPIAR ZONA DE DATOS
    ' Si existe tabla, limpiamos contenido para NO romper Slicers
    Dim existeTabla As Boolean
    existeTabla = False
    
    On Error Resume Next
    Set tbl = ws.ListObjects("Consolidado_Rubro")
    On Error GoTo ErrorHandler
    
    If Not tbl Is Nothing Then
        existeTabla = True
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
            ' Reducir a 1 sola fila para limpiar
            If tbl.ListRows.Count > 1 Then
                tbl.DataBodyRange.Offset(1, 0).Resize(tbl.ListRows.Count - 1, tbl.ListColumns.Count).Rows.Delete
            End If
        End If
    Else
        ' Si no existe, limpiar rango desde fila 8
        ws.Range("A" & filaEncabezado & ":Z" & ws.Rows.Count).Clear
    End If
    
    ' 3. PREPARAR CONSULTA SQL
    sql = "SELECT " & vbCrLf
    sql = sql & "    l.fechaaplicacion::DATE AS ""FECHA ENTREGA REPORTE""," & vbCrLf
    sql = sql & "    UPPER(l.nombreempleado) AS ""NOMBRE""," & vbCrLf
    sql = sql & "    l.empleado::BIGINT AS ""DOCUMENTO DE IDENTIDAD""," & vbCrLf
    sql = sql & "    CASE " & vbCrLf
    sql = sql & "        WHEN ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' THEN ln.ot " & vbCrLf
    sql = sql & "        ELSE 'C' || ln.centrocosto " & vbCrLf
    sql = sql & "    END AS ""OT-CC""," & vbCrLf
    sql = sql & "    COALESCE(ln.fecharealgasto, l.fechaaplicacion)::DATE AS ""FECHA REAL DEL GASTO""," & vbCrLf
    sql = sql & "    EXTRACT(YEAR FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS ""AÑO""," & vbCrLf
    sql = sql & "    TO_CHAR(COALESCE(ln.fecharealgasto, l.fechaaplicacion), 'TMMonth') AS ""MES""," & vbCrLf
    sql = sql & "    EXTRACT(WEEK FROM COALESCE(ln.fecharealgasto, l.fechaaplicacion))::INTEGER AS ""SEMANA DEL AÑO""," & vbCrLf
    sql = sql & "    o.cliente AS ""OBRA""," & vbCrLf
    sql = sql & "    o.ciudad AS ""CIUDAD""," & vbCrLf
    sql = sql & "    ln.centrocosto AS ""CENTRO DE COSTO""," & vbCrLf
    sql = sql & "    ln.subcentrocosto AS ""SUB CENTRO""," & vbCrLf
    sql = sql & "    ln.categoria AS ""DESCRIPCION""," & vbCrLf
    sql = sql & "    COALESCE(ln.valorconfactura, 0)::BIGINT AS ""VALOR TOTAL FACTURA""," & vbCrLf
    sql = sql & "    COALESCE(ln.valorsinfactura, 0)::BIGINT AS ""VALOR SIN FACTURA""," & vbCrLf
    sql = sql & "    (COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS ""APROBADO""," & vbCrLf
    sql = sql & "    (COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS ""SOLICITADO""," & vbCrLf
    sql = sql & "    l.codigolegalizacion AS ""RADICADO""," & vbCrLf
    sql = sql & "    SPLIT_PART(l.codigolegalizacion, '-', 1) AS ""AREA""," & vbCrLf
    sql = sql & "    0::BIGINT AS ""DIFERENCIA""" & vbCrLf
    sql = sql & "FROM " & vbCrLf
    sql = sql & "    legalizacion l" & vbCrLf
    sql = sql & "JOIN " & vbCrLf
    sql = sql & "    linealegalizacion ln ON l.codigo = ln.legalizacion" & vbCrLf
    sql = sql & "LEFT JOIN otviaticos o ON ln.ot = o.numero" & vbCrLf
    sql = sql & "ORDER BY " & vbCrLf
    sql = sql & "    l.fechaaplicacion DESC;"
          
    ' 4. CONEXIÓN A POSTGRESQL (Hex Password)
    Set conn = CreateObject("ADODB.Connection")
    Application.StatusBar = "Conectando a DB..."
    conn.Open "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & HexToString("41646D696E536F6C696432303235") & ";"
    
    ' 5. EJECUTAR
    Set rs = conn.Execute(sql)
    
    ' 6. VOLCAR ENC (Solo si no existe tabla) Y DATOS
    Application.StatusBar = "Escribiendo datos..."
    
    If Not existeTabla Then
        For col = 0 To rs.Fields.Count - 1
            ws.Cells(filaEncabezado, col + 1).Value = rs.Fields(col).Name
        Next col
    End If
    
    ' Datos
    ws.Cells(filaEncabezado + 1, 1).CopyFromRecordset rs
    
    ' 7. FORMATO Y REDIMENSIONAR TABLA
    totalRegs = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' xlUp
    
    If totalRegs >= filaEncabezado + 1 Then
        Dim rangoDatos As Range
        Set rangoDatos = ws.Range(ws.Cells(filaEncabezado, 1), ws.Cells(totalRegs, rs.Fields.Count))
        
        If existeTabla Then
            tbl.Resize rangoDatos
        Else
            Set tbl = ws.ListObjects.Add(1, rangoDatos, , 1) ' xlSrcRange, xlYes
            tbl.Name = "Consolidado_Rubro"
            tbl.TableStyle = "TableStyleMedium9" ' Azul claro estandar
        End If
        
        ' Formato Encabezados
        With tbl.HeaderRowRange
             .Interior.Color = RGB(0, 32, 96)
             .Font.Color = RGB(255, 255, 255)
             .Font.Bold = True
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlCenter
        End With
        
        ' Formato Moneda Columns (N, O, P, Q, T -> 14, 15, 16, 17, 20)
        ' VALOR TOTAL FACTURA, VALOR SIN FACTURA, APROBADO, SOLICITADO, DIFERENCIA
        ws.Columns("N:Q").NumberFormat = "$#,##0"
        ws.Columns("T:T").NumberFormat = "$#,##0"
        
        ' Formato Fecha (A, E)
        ws.Columns("A:A").NumberFormat = "dd/mm/yyyy"
        ws.Columns("E:E").NumberFormat = "dd/mm/yyyy"
        
        ' Ajustar anchos
        ws.Columns.AutoFit
    End If
    
    ' Limpiar
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Informe Por Rubro generado exitosamente.", vbInformation
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub
