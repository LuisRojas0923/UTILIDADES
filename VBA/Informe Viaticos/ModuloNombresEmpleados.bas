'Attribute VB_Name = "ModuloNombresEmpleados"
' Macro para consultar lista de nombres unicos de empleados desde PostgreSQL
' Base de datos: solid
' Servidor: 192.168.0.21:5432
' Usuario: postgres
' Autor: Sistema de Utilidades
' Fecha: 2025

Option Explicit

Function ASCIIaTexto(valoresASCII As String) As String
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

Sub ConsultarNombresEmpleadosUnicos()
    ' Conecta a PostgreSQL y consulta los nombres unicos de empleados
    ' Los resultados se muestran en una hoja de Excel desde la fila 7
    
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
    Dim numColumnas As Integer
    Dim rangoTabla As Range
    Dim ultimaColumna As String
    Dim totalRegistros As Long
    
    ' Parametros de conexion
    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    ' Contrasena en formato ASCII (generada con convertir_contrasena.py)
    contrasena = ASCIIaTexto("65,100,109,105,110,83,111,108,105,100,50,48,50,53")
    
    ' Consulta SQL para obtener nombres unicos de empleados
    consultaSQL = "SELECT DISTINCT " & vbCrLf & _
                  "    COALESCE(c.nombreempleado, l.nombreempleado) AS ""EMPLEADO""" & vbCrLf & _
                  "FROM transaccionviaticos t" & vbCrLf & _
                  "LEFT JOIN consignacion c" & vbCrLf & _
                  "  ON c.codigoconsignacion = t.numerodocumento" & vbCrLf & _
                  "LEFT JOIN legalizacion l" & vbCrLf & _
                  "  ON l.codigolegalizacion = t.numerodocumento" & vbCrLf & _
                  "WHERE COALESCE(c.nombreempleado, l.nombreempleado) IS NOT NULL" & vbCrLf & _
                  "  AND TRIM(COALESCE(c.nombreempleado, l.nombreempleado)) != ''" & vbCrLf & _
                  "  AND (c.nombreempleado IS NOT NULL OR l.nombreempleado IS NOT NULL)" & vbCrLf & _
                  "ORDER BY ""EMPLEADO"" ASC"
    
    ' Cadena de conexion ODBC para PostgreSQL
    cadenaConexion = "Driver={PostgreSQL ODBC Driver(UNICODE)};" & _
                     "Server=" & servidor & ";" & _
                     "Port=" & puerto & ";" & _
                     "Database=" & baseDatos & ";" & _
                     "Uid=" & usuario & ";" & _
                     "Pwd=" & contrasena & ";"
    
    On Error GoTo ErrorHandler
    
    ' Crear objetos ADO
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Mostrar estado
    Application.StatusBar = "Conectando a PostgreSQL..."
    Debug.Print "Intentando conectar a PostgreSQL..."
    Debug.Print "Servidor: " & servidor & ":" & puerto
    Debug.Print "Base de datos: " & baseDatos
    
    ' Abrir conexion
    conn.Open cadenaConexion
    Debug.Print "Conexion establecida correctamente"
    
    ' Ejecutar consulta
    Application.StatusBar = "Ejecutando consulta SQL..."
    Debug.Print "Ejecutando consulta de nombres unicos de empleados..."
    Set rs = conn.Execute(consultaSQL)
    
    ' Crear o limpiar hoja de resultados
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Empleados")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Empleados"
    Else
        ' Limpiar solo desde la fila 7 en adelante para preservar las filas 1-6
        ws.Range("A7:ZZ" & ws.Rows.Count).Clear
    End If
    On Error GoTo ErrorHandler
    
    ' Guardar numero de columnas antes del loop
    numColumnas = rs.Fields.Count
    
    ' Escribir encabezados en la fila 7
    fila = 7
    For columna = 0 To numColumnas - 1
        ws.Cells(fila, columna + 1).Value = rs.Fields(columna).Name
        ws.Cells(fila, columna + 1).Font.Bold = True
        ws.Cells(fila, columna + 1).Interior.Color = RGB(0, 32, 96) '002060azul
        ws.Cells(fila, columna + 1).Font.Color = RGB(255, 255, 255)
    Next columna
    
    ' Escribir datos desde la fila 8 en adelante
    Application.StatusBar = "Extrayendo datos..."
    fila = 8
    totalRegistros = 0
    
    Do While Not rs.EOF
        For columna = 0 To numColumnas - 1
            ws.Cells(fila, columna + 1).Value = rs.Fields(columna).Value
        Next columna
        rs.MoveNext
        fila = fila + 1
        totalRegistros = totalRegistros + 1
        
        ' Actualizar estado cada 100 registros
        If totalRegistros Mod 100 = 0 Then
            Application.StatusBar = "Extrayendo datos... " & totalRegistros & " registros procesados"
            DoEvents
        End If
    Loop
    
    ' Formatear tabla
    Application.StatusBar = "Formateando resultados..."
    Debug.Print "Total de registros extraidos: " & totalRegistros
    
    ' Convertir numero de columna a letra para el rango de tabla
    If numColumnas <= 26 Then
        ultimaColumna = Chr(64 + numColumnas)
    Else
        ultimaColumna = Chr(64 + Int((numColumnas - 1) / 26)) & Chr(65 + ((numColumnas - 1) Mod 26))
    End If
    
    ' Aplicar formato de tabla
    Set rangoTabla = ws.Range("A7:" & ultimaColumna & (fila - 1))
    
    On Error Resume Next
    ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TablaEmpleados"
    On Error GoTo ErrorHandler
    
    ' Formatear tabla de Excel
    With ws.ListObjects("TablaEmpleados")
        .TableStyle = "TableStyleMedium9"
        .ShowAutoFilter = True
    End With
    
    ' Centrar encabezados
    ws.Range("A7:" & ultimaColumna & "7").HorizontalAlignment = xlCenter
    
    ' Ajustar ancho de columnas
    ws.Columns("A:" & ultimaColumna).AutoFit
    
    ' Ajustar altura de encabezados
    ws.Rows("7:7").RowHeight = 30
    
    ' Aplicar formato de texto a los encabezados
    ws.Range("A7:" & ultimaColumna & "7").WrapText = True
    ws.Range("A7:" & ultimaColumna & "7").VerticalAlignment = xlCenter
    
    ' Cerrar conexiones
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    ' Mensaje de exito
    Application.StatusBar = "Proceso completado: " & totalRegistros & " nombres unicos extraidos"
    Debug.Print "Proceso completado exitosamente"
    MsgBox "Consulta completada exitosamente." & vbCrLf & _
           "Total de nombres unicos: " & totalRegistros & vbCrLf & _
           "Los resultados se encuentran en la hoja 'Empleados'", _
           vbInformation, "Consulta Nombres Empleados"
    
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
    'MsgBox mensajeError, vbCritical, "Error en Consulta Nombres Empleados"
    
End Sub

