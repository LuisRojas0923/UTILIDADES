'Attribute VB_Name = "ModuloConsignaciones"
' Macro para generar el reporte de CONSIGNACIONES
' Formato: Encabezados en Fila 12, Hoja "CONSIGNACIONES"
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

Sub EjecutarConsignaciones()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim filaEncabezado As Integer
    Dim totalRegs As Long
    Dim col As Integer
    Dim tbl As ListObject
    
    ' Definimos fila de inicio
    filaEncabezado = 12
    Application.ScreenUpdating = False
    ' 1. CONFIGURAR HOJA
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CONSIGNACIONES")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "CONSIGNACIONES"
    End If
    On Error GoTo ErrorHandler
    
    ' 2. LIMPIAR ZONA DE DATOS (Filas 12 hacia abajo)
    ' Si existe tabla, limpiamos contenido para NO romper Slicers
    Dim existeTabla As Boolean
    existeTabla = False
    
    On Error Resume Next
    Set tbl = ws.ListObjects("Consignaciones_Viaticos")
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
        ' Si no existe, limpiar rango manual
         ws.Range("A" & filaEncabezado & ":Z" & ws.Rows.Count).Clear
    End If
    
    ' 3. PREPARAR CONSULTA SQL
    sql = "SELECT " & vbCrLf & _
          "    EXTRACT(YEAR FROM t.fechaaplicacion)::INTEGER AS ""AÑO""," & vbCrLf & _
          "    TO_CHAR(t.fechaaplicacion, 'TMMonth') AS ""MES""," & vbCrLf & _
          "    EXTRACT(DAY FROM t.fechaaplicacion)::INTEGER AS ""DIA""," & vbCrLf & _
          "    t.fechaaplicacion::DATE AS ""FECHA""," & vbCrLf & _
          "    t.empleado::BIGINT AS ""CEDULA""," & vbCrLf & _
          "    UPPER(COALESCE(c.nombreempleado, '')) AS ""EMPLEADO""," & vbCrLf & _
          "    t.numerodocumento AS ""CONTRATO""," & vbCrLf & _
          "    t.valor::BIGINT AS ""CONSIGNACION""," & vbCrLf & _
          "    (t.valor * 0.004)::BIGINT AS ""IMP 4 X 1000""," & vbCrLf & _
          "    (t.valor + (t.valor * 0.004))::BIGINT AS ""TOTAL CONSIGNACION""," & vbCrLf & _
          "    NULL::TEXT AS ""VIATICO A PAGAR?""" & vbCrLf & _
          "FROM transaccionviaticos t" & vbCrLf & _
          "LEFT JOIN consignacion c ON t.numerodocumento = c.codigoconsignacion" & vbCrLf & _
          "WHERE t.tipodocumento = 'CONSIGNACION'" & vbCrLf & _
          "    AND UPPER(c.estado) LIKE '%CONTABILIZADO%'" & vbCrLf & _
          "ORDER BY t.fechaaplicacion DESC;"
          
    ' 4. CONEXIÓN A POSTGRESQL (Hex Password)
    Set conn = CreateObject("ADODB.Connection")
    Application.StatusBar = "Conectando a DB..."
    conn.Open "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & HexToString("41646D696E536F6C696432303235") & ";"
    
    ' 5. EJECUTAR
    Set rs = conn.Execute(sql)
    
    ' 6. VOLCAR ENC (Solo si no existe tabla, para asegurar) Y DATOS
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
            tbl.Name = "Consignaciones_Viaticos"
            tbl.TableStyle = "TableStyleMedium13"
        End If
        
        ' 1. Color Encabezado (002060 -> RGB(0, 32, 96)) y Centrado
        With tbl.HeaderRowRange
            .Interior.Color = RGB(0, 32, 96)
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        
        ' 2. Totales (En Fila 11, encima de encabezados)
        tbl.ShowTotals = False
        ws.Range("H11").Formula = "=SUBTOTAL(9,Consignaciones_Viaticos[CONSIGNACION])"
        ws.Range("I11").Formula = "=SUBTOTAL(9,Consignaciones_Viaticos[IMP 4 X 1000])"
        ws.Range("J11").Formula = "=SUBTOTAL(9,Consignaciones_Viaticos[TOTAL CONSIGNACION])"
        
        ' Formato Moneda para los subtotales superiores
        ws.Range("H11:J11").NumberFormat = "$#,##0"
        ws.Range("H11:J11").Font.Bold = True
        
        ' 3. Formato Moneda Columna
        ws.Columns("H:J").NumberFormat = "$#,##0" ' Columnas de valor
        ws.Columns("E:E").NumberFormat = "General" ' Columna Cedula
        
        ' 4. Formato Condicional (Check si es 1 en VIATICO A PAGAR?)
        Dim colCond As Range
        Set colCond = tbl.ListColumns("VIATICO A PAGAR?").DataBodyRange
        colCond.FormatConditions.Delete
        
        Dim iconSetCond As IconSetCondition
        Set iconSetCond = colCond.FormatConditions.AddIconSetCondition
        With iconSetCond
            .IconSet = ActiveWorkbook.IconSets(xl3Symbols)
            .ReverseOrder = False
            .ShowIconOnly = True
            ' Configurar icono verde (Check) para valor >= 1
            With .IconCriteria(3)
                .Type = xlConditionValueNumber
                .Value = 1
                .Operator = xlGreaterEqual
            End With
            ' Los otros iconos quedaran por defecto para valores menores
        End With
            
        ' Ajustar anchos
        ' ws.Columns.AutoFit <-- Removido por solicitud del usuario
    End If
    Application.ScreenUpdating = True
    
    ' Limpiar
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    MsgBox "Reporte de Consignaciones generado exitosamente.", vbInformation
    Exit Sub
    Application.ScreenUpdating = True
ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
    Application.ScreenUpdating = True
End Sub
