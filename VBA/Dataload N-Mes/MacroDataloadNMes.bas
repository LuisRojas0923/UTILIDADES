'Attribute VB_Name = "ModuloDataloadNMes"
' Macro para ejecutar el Dataload N-Mes desde PostgreSQL
' Formato simplificado: RADICADO, DOCUMENTO DE IDENTIDAD, OT-CC, APROBADO, LLAVE
' Base de datos: solid
' Servidor: 192.168.0.21:5432

Option Explicit

' --- UTILIDADES DE CIFRADO ---

Private Function HexToString(ByVal hexVal As String) As String
    ' Convierte una cadena Hexadecimal a texto plano
    ' Ejemplo: "41646D" -> "Adm"
    Dim i As Long
    Dim res As String
    res = ""
    For i = 1 To Len(hexVal) Step 2
        res = res & Chr(Val("&H" & Mid(hexVal, i, 2)))
    Next i
    HexToString = res
End Function

' --- PROCESO PRINCIPAL ---

Sub EjecutarDataloadNMes()
    ' Proceso que solicita filtros, conecta a DB y carga los datos en Excel
    
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim sql As String
    Dim fechaDesde As String, fechaHasta As String, empleadoID As String
    Dim fila As Long, col As Long
    Dim totalRegs As Long
    
    ' 1. SOLICITAR FILTROS AL USUARIO - ELIMINADOS POR SOLICITUD DEL USUARIO
    
    On Error GoTo ErrorHandler
    
    ' 2. PREPARAR CONSULTA SQL
    ' Usamos el formato simplificado solicitado por el usuario, sin filtros
    sql = "SELECT " & vbCrLf & _
          "    ""RADICADO""," & vbCrLf & _
          "    ""DOCUMENTO DE IDENTIDAD""," & vbCrLf & _
          "    ""OT-CC""," & vbCrLf & _
          "    ""APROBADO""," & vbCrLf & _
          "    ""RADICADO"" || '-' || ROW_NUMBER() OVER(PARTITION BY ""RADICADO"" ORDER BY ""OT-CC"") AS ""LLAVE""" & vbCrLf & _
          "FROM (" & vbCrLf & _
          "    SELECT " & vbCrLf & _
          "        l.codigolegalizacion AS ""RADICADO""," & vbCrLf & _
          "        l.empleado AS ""DOCUMENTO DE IDENTIDAD""," & vbCrLf & _
          "        CASE " & vbCrLf & _
          "            WHEN TRIM(COALESCE(ln.ot, '')) = '' THEN 'C' || COALESCE(ln.centrocosto, '')" & vbCrLf & _
          "            ELSE ln.ot " & vbCrLf & _
          "        END AS ""OT-CC""," & vbCrLf & _
          "        SUM(COALESCE(ln.valorsinfactura, 0) + COALESCE(ln.valorconfactura, 0))::BIGINT AS ""APROBADO""," & vbCrLf & _
          "        l.fechaaplicacion" & vbCrLf & _
          "    FROM linealegalizacion ln" & vbCrLf & _
          "    JOIN legalizacion l ON ln.legalizacion = l.codigo" & vbCrLf & _
          "    GROUP BY l.codigolegalizacion, l.empleado, l.fechaaplicacion," & vbCrLf & _
          "             CASE WHEN TRIM(COALESCE(ln.ot, '')) = '' THEN 'C' || COALESCE(ln.centrocosto, '') ELSE ln.ot END" & vbCrLf & _
          ") AS grouped_results" & vbCrLf & _
          "ORDER BY fechaaplicacion, ""RADICADO"", ""LLAVE"";"

    ' 3. CONEXIÓN A POSTGRESQL
    Set conn = CreateObject("ADODB.Connection")
    ' Usamos la cadena hexadecimal para mayor seguridad
    conn.Open "Driver={PostgreSQL Unicode(x64)};Server=192.168.0.21;Port=5432;Database=solid;Uid=postgres;Pwd=" & HexToString("41646D696E536F6C696432303235") & ";"
    
    ' 4. EJECUTAR Y CARGAR DATOS
    Application.StatusBar = "Ejecutando SQL..."
    Set rs = conn.Execute(sql)
    
    ' Preparar Hoja y Tabla
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("BD_VIATICOS1")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "BD_VIATICOS1"
    End If
    
    Dim tbl As ListObject
    Set tbl = ws.ListObjects("BD_VIATICOS1")
    
    ' Si la tabla existe, borrar contenido de datos pero mantener el objeto
    If Not tbl Is Nothing Then
        If Not tbl.DataBodyRange Is Nothing Then
            tbl.DataBodyRange.ClearContents
            ' Borrar filas sobrantes si las hay
            If tbl.ListRows.Count > 1 Then
                tbl.DataBodyRange.Offset(1, 0).Resize(tbl.ListRows.Count - 1, tbl.ListColumns.Count).Rows.Delete
            End If
        End If
    Else
        ws.Cells.Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Encabezados (Solo si la tabla no existe o para asegurar nombres)
    For col = 0 To rs.Fields.Count - 1
        ws.Cells(1, col + 1).Value = rs.Fields(col).Name
        ws.Cells(1, col + 1).Font.Bold = True
        ws.Cells(1, col + 1).Interior.Color = RGB(0, 112, 192)
        ws.Cells(1, col + 1).Font.Color = RGB(255, 255, 255)
    Next col
    
    ' Datos
    Application.StatusBar = "Cargando datos en Excel..."
    ws.Cells(2, 1).CopyFromRecordset rs
    
    ' 5. FORMATO Y REDIMENSIONAMIENTO
    totalRegs = ws.Cells(ws.Rows.Count, 1).End(-4162).Row - 1 '-4162 es xlUp
    
    If totalRegs > 0 Then
        Dim tblRange As Range
        Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(totalRegs + 1, rs.Fields.Count))
        
        If tbl Is Nothing Then
            ' Crear tabla nueva si no existía
            Set tbl = ws.ListObjects.Add(1, tblRange, , 1) '1 es xlSrcRange, 1 es xlYes
            tbl.Name = "BD_VIATICOS1"
            tbl.TableStyle = "TableStyleMedium2"
        Else
            ' Redimensionar tabla existente
            tbl.Resize tblRange
        End If
        
        ' Ajustar columnas y formatos
        ws.Columns.AutoFit
        ws.Columns("D:D").NumberFormat = "$#,##0"
    End If
    
    ' 6. LIMPIEZA
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    Application.StatusBar = False
    MsgBox "Dataload completado con éxito." & vbCrLf & _
           "Registros importados: " & totalRegs, vbInformation, "Éxito"
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    MsgBox "Error: " & Err.Description, vbCritical, "Error en Macro"
    If Not rs Is Nothing Then If rs.State = 1 Then rs.Close
    If Not conn Is Nothing Then If conn.State = 1 Then conn.Close
End Sub
