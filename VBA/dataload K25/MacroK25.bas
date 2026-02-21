'Attribute VB_Name = "ModuloK25"
' Macro para ejecutar el Dataload K25 desde PostgreSQL
' Formato solicitado: DOCUMENTO DE IDENTIDAD, OT-CC, CENTRO DE COSTO, SUB CENTRO, RADICADO, VALOR OT
' Base de datos: solid
' Servidor: 192.168.0.21:5432

Option Explicit

' --- UTILIDADES DE CIFRADO ---

Sub MacroActualizarDatos()
    Call EjecutarDataloadK25
    Call EjecutarDataloadK25Resumen
End Sub


Private Function HexToString(ByVal hexVal As String) As String
    Dim i As Long
    Dim res As String
    res = ""
    For i = 1 To Len(hexVal) Step 2
        res = res & Chr(Val("&H" & Mid(hexVal, i, 2)))
    Next i
    HexToString = res
End Function

' --- PROCESO PRINCIPAL ---

Sub EjecutarDataloadK25()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim col As Long
    Dim totalRegs As Long
    
    On Error GoTo ErrorHandler
    
    ' 1. SQL
    sql = "SELECT " & vbCrLf & _
          "    l.empleado::BIGINT AS ""DOCUMENTO DE IDENTIDAD""," & vbCrLf & _
          "    ln.ot AS ""OT-CC""," & vbCrLf & _
          "    ln.centrocosto AS ""CENTRO DE COSTO""," & vbCrLf & _
          "    ln.subcentrocosto AS ""SUB CENTRO""," & vbCrLf & _
          "    l.codigolegalizacion AS ""RADICADO""," & vbCrLf & _
          "    SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS ""VALOR OT""" & vbCrLf & _
          "FROM linealegalizacion ln " & vbCrLf & _
          "JOIN legalizacion l ON ln.legalizacion = l.codigo " & vbCrLf & _
          "WHERE ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' " & vbCrLf & _
          "AND UPPER(l.estado) = 'CONTABILIZADO' " & vbCrLf & _
          "GROUP BY l.empleado, ln.ot, ln.centrocosto, ln.subcentrocosto, l.codigolegalizacion " & vbCrLf & _
          "ORDER BY l.codigolegalizacion ASC, ln.ot ASC;"

    ' 2. CONEXIÓN
    Debug.Print "Conectando a PostgreSQL (192.168.0.21)..."
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & HexToString("41646D696E536F6C696432303235") & ";"
    Debug.Print "Conexión exitosa."
    
    ' 3. EJECUTAR
    Debug.Print "Ejecutando SQL K25..."
    Application.StatusBar = "Ejecutando SQL K25..."
    Set rs = conn.Execute(sql)
    Debug.Print "Consulta SQL ejecutada. Recordset obtenido."
    
    ' 4. PREPARAR HOJA Y TABLA
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_VIATICOS1")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "BD_VIATICOS1"
    End If
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("BD_VIATICOS1")
    
    If Not tbl Is Nothing Then
        Debug.Print "Tabla BD_VIATICOS1 encontrada. Convirtiendo a rango (Unlist) para evitar bloqueos..."
        ' Convertir tabla a rango normal para poder manipularla sin errores de ListObject
        tbl.Unlist
        ' Limpiar SOLO las columnas A a F (donde van los datos) para no tocar la tabla espejo
        ws.Columns("A:F").ClearContents
    Else
        Debug.Print "Preparando espacio para BD_VIATICOS1 (Cols A a F)..."
        ws.Columns("A:F").ClearContents
    End If
    On Error GoTo ErrorHandler
    
    ' 5. CARGAR ENCABEZADOS Y DATOS
    Application.StatusBar = "Cargando datos en Excel..."
    
    Dim yaExistiaTabla As Boolean
    yaExistiaTabla = Not tbl Is Nothing
    
    ' Encabezados (Siempre los reponemos porque ClearContents los borra)
    Debug.Print "Escribiendo encabezados..."
    For col = 0 To rs.Fields.Count - 1
        ws.Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col
    
    ' Datos desde fila 2 (debajo de encabezados)
    Debug.Print "Copiando datos del recordset a Excel (celda A2)..."
    ws.Cells(2, 1).CopyFromRecordset rs
    
    ' 6. FORMATO DE TABLA
    totalRegs = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' -4162 es xlUp
    
    If totalRegs >= 1 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(totalRegs, rs.Fields.Count))
        
        ' Crear la tabla de nuevo
        Debug.Print "Creando nueva definición de tabla BD_VIATICOS1."
        Set tbl = ws.ListObjects.Add(1, tblRange, , 1) ' 1 es xlSrcRange, 1 es xlYes
        tbl.Name = "BD_VIATICOS1"
        tbl.TableStyle = "TableStyleMedium2"
        
        ' Formato estético
        With tbl.HeaderRowRange
            .Font.Bold = True
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = -4108 ' xlCenter
        End With
        
        ws.Columns.AutoFit
        ' Formato moneda para la última columna (VALOR OT)
        ws.Columns(rs.Fields.Count).NumberFormat = "$#,##0"
    End If
    
    ' 7. LIMPIEZA
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    'MsgBox "Dataload K25 completado con éxito." & vbCrLf & _
    '       "Registros importados: " & (totalRegs - 1), vbInformation, "Éxito"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "ERROR CRÍTICO: " & Err.Number & " - " & Err.Description
    MsgBox "Error en Macro K25: " & Err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close

End Sub

Sub EjecutarDataloadK25Resumen()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim col As Long
    Dim totalRegs As Long
    Dim yaExistiaTabla As Boolean
    
    On Error GoTo ErrorHandler
    
    ' 1. SQL RESUMEN
    sql = "SELECT " & vbCrLf & _
          "    UPPER(l.nombreempleado) AS ""EMPLEADO""," & vbCrLf & _
          "    SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT AS ""APROBADO""," & vbCrLf & _
          "    ln.ot AS ""OT-CC""," & vbCrLf & _
          "    ln.centrocosto AS ""CC""," & vbCrLf & _
          "    ln.subcentrocosto AS ""SUB CENTRO""," & vbCrLf & _
          "    l.empleado::BIGINT AS ""CEDULA""," & vbCrLf & _
          "    l.codigolegalizacion AS ""RADICADO""," & vbCrLf & _
          "    (l.empleado || '-' || ln.ot || '-' || SUM(COALESCE(ln.valorconfactura, 0) + COALESCE(ln.valorsinfactura, 0))::BIGINT) AS ""LLAVE""" & vbCrLf & _
          "FROM linealegalizacion ln " & vbCrLf & _
          "JOIN legalizacion l ON ln.legalizacion = l.codigo " & vbCrLf & _
          "WHERE ln.ot IS NOT NULL AND TRIM(ln.ot) <> '' " & vbCrLf & _
          "AND UPPER(l.estado) = 'CONTABILIZADO' " & vbCrLf & _
          "GROUP BY l.nombreempleado, ln.ot, ln.centrocosto, ln.subcentrocosto, l.empleado, l.codigolegalizacion " & vbCrLf & _
          "ORDER BY l.nombreempleado ASC, l.codigolegalizacion ASC;"

    ' 2. CONEXIÓN
    Debug.Print "Conectando a PostgreSQL para Resumen..."
    Set conn = CreateObject("ADODB.Connection")
    conn.Open "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & HexToString("41646D696E536F6C696432303235") & ";"
    
    ' 3. EJECUTAR
    Debug.Print "Ejecutando SQL K25 Resumen..."
    Application.StatusBar = "Ejecutando SQL K25 Resumen..."
    Set rs = conn.Execute(sql)
    
    ' 4. PREPARAR HOJA Y TABLA
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_GENERAL2")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "BD_GENERAL2"
    End If
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("BD_GENERAL2")
    yaExistiaTabla = Not tbl Is Nothing
    
    If yaExistiaTabla Then
        Debug.Print "Tabla BD_GENERAL2 encontrada. Convirtiendo a rango (Unlist) para compatibilidad con espejos..."
        ' Desvincular tabla para evitar el error de "cambiar límites"
        tbl.Unlist
        ' Limpiar solo las columnas A a G
        ws.Columns("A:G").ClearContents
    Else
        Debug.Print "Preparando espacio para BD_GENERAL2 (Cols A a G)..."
        ws.Columns("A:G").ClearContents
    End If
    On Error GoTo ErrorHandler
    
    ' 5. CARGAR ENCABEZADOS Y DATOS
    Debug.Print "Reponiendo encabezados en fila 1..."
    For col = 0 To rs.Fields.Count - 1
        ws.Cells(1, col + 1).Value = rs.Fields(col).Name
    Next col
    
    Debug.Print "Cargando datos resumen..."
    ws.Cells(2, 1).CopyFromRecordset rs
    
    ' 6. FORMATO DE TABLA
    totalRegs = ws.Cells(ws.Rows.Count, 1).End(-4162).Row ' -4162 es xlUp
    
    If totalRegs >= 1 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(totalRegs, rs.Fields.Count))
        
        ' Crear la tabla de nuevo
        Debug.Print "Creando nueva definición de tabla BD_GENERAL2."
        Set tbl = ws.ListObjects.Add(1, tblRange, , 1) ' 1 es xlSrcRange, 1 es xlYes
        tbl.Name = "BD_GENERAL2"
        tbl.TableStyle = "TableStyleMedium2"
        
        With tbl.HeaderRowRange
            .Font.Bold = True
            .Interior.Color = RGB(0, 112, 192)
            .Font.Color = RGB(255, 255, 255)
            .HorizontalAlignment = -4108 ' xlCenter
        End With
        
        ws.Columns.AutoFit
        ' Formato moneda para la columna APROBADO (columna 2)
        ws.Columns(2).NumberFormat = "$#,##0"
    End If
    
    ' 7. LIMPIEZA
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    Debug.Print "Dataload K25 Resumen finalizado exitosamente."
    'MsgBox "Resumen K25 completado en BD_GENERAL2." & vbCrLf & _
    '       "Registros: " & (totalRegs - 1), vbInformation, "Éxito"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Debug.Print "Error en Resumen K25: " & Err.Description
    MsgBox "Error en Resumen K25: " & Err.Description, vbCritical, "Error"
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub
