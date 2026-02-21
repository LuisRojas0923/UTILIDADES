'Attribute VB_Name = "ModuloConsultaContratos"
' Macro para consultar datos de consignacion desde PostgreSQL
' Base de datos: solid
' Servidor: 192.168.0.21:5432
' Usuario: postgres
' Autor: Sistema de Utilidades
' Fecha: 2025

Option Explicit

Private Function Password(cadenaDatos As String) As String
    ' Procesa cadena de datos separados por comas
    
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

Sub ConsultarContratos()
    ' Conecta a PostgreSQL y consulta datos de consignacion
    ' Los resultados se actualizan en la hoja CONTRATOS desde la fila 2
    
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
    Dim totalRegistros As Long
    Dim ultimaFila As Long
    Dim rangoLimpiar As Range
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    ' Parametros de conexion
    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    contrasena = Password("65,100,109,105,110,83,111,108,105,100,50,48,50,53")
    
    ' Consulta SQL
    consultaSQL = "SELECT empleado AS CEDULA, " & vbCrLf & _
                  "       nombreempleado AS EMPLEADO, " & vbCrLf & _
                  "       codigoconsignacion AS CONTRATO, " & vbCrLf & _
                  "       valor AS CONSIGNACION " & vbCrLf & _
                  "FROM consignacion"
    
    ' Cadena de conexion ODBC para PostgreSQL
    cadenaConexion = "Driver={PostgreSQL Unicode(x64)};" & _
                     "Server=" & servidor & ";" & _
                     "Port=" & puerto & ";" & _
                     "Database=" & baseDatos & ";" & _
                     "Uid=" & usuario & ";" & _
                     "Pwd=" & contrasena & ";"
    
    On Error GoTo ErrorHandler
    
    ' Verificar que la hoja CONTRATOS existe
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("CONTRATOS")
    If ws Is Nothing Then
        Application.ScreenUpdating = True
        MsgBox "Error: La hoja 'CONTRATOS' no existe en el libro actual.", vbCritical, "Error"
        Exit Sub
    End If
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
    Debug.Print "Ejecutando consulta de contratos..."
    Set rs = conn.Execute(consultaSQL)
    
    ' Determinar la ultima fila con datos en las columnas A:D
    ' Buscar en cada columna para encontrar la ultima fila con datos
    ultimaFila = 1
    Dim ultimaFilaA As Long
    Dim ultimaFilaB As Long
    Dim ultimaFilaC As Long
    Dim ultimaFilaD As Long
    
    ultimaFilaA = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    ultimaFilaB = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    ultimaFilaC = ws.Cells(ws.Rows.Count, 3).End(xlUp).Row
    ultimaFilaD = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    
    ' Obtener la mayor fila con datos
    ultimaFila = Application.WorksheetFunction.Max(ultimaFilaA, ultimaFilaB, ultimaFilaC, ultimaFilaD)
    
    ' Si no hay datos, empezar desde la fila 2
    If ultimaFila < 2 Then
        ultimaFila = 2
    End If
    
    ' Borrar solo el contenido de las columnas A:D desde la fila 2
    Application.StatusBar = "Limpiando datos anteriores..."
    Debug.Print "Limpiando contenido de columnas A:D desde fila 2 hasta fila " & ultimaFila
    Set rangoLimpiar = ws.Range("A2:D" & ultimaFila)
    rangoLimpiar.ClearContents
    
    ' Escribir datos desde la fila 2 usando CopyFromRecordset (mucho mas rapido)
    Application.StatusBar = "Extrayendo datos..."
    Debug.Print "Copiando datos desde recordset..."
    
    ' Guardar la fila inicial para contar registros despues
    Dim filaInicial As Long
    filaInicial = 2
    
    ' Usar CopyFromRecordset para mejor rendimiento
    Dim rangoDestino As Range
    Set rangoDestino = ws.Range("A2")
    rangoDestino.CopyFromRecordset rs
    
    ' Contar registros escritos verificando la ultima fila con datos
    Dim ultimaFilaConDatos As Long
    ultimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFilaConDatos >= filaInicial Then
        totalRegistros = ultimaFilaConDatos - filaInicial + 1
    Else
        totalRegistros = 0
    End If
    
    ' Si hay menos datos nuevos que antes, limpiar las filas adicionales
    Dim ultimaFilaNueva As Long
    ultimaFilaNueva = ultimaFilaConDatos
    If ultimaFilaNueva < ultimaFila And ultimaFilaNueva >= 2 Then
        Application.StatusBar = "Limpiando filas adicionales..."
        Debug.Print "Limpiando filas adicionales desde " & (ultimaFilaNueva + 1) & " hasta " & ultimaFila
        Set rangoLimpiar = ws.Range("A" & (ultimaFilaNueva + 1) & ":D" & ultimaFila)
        rangoLimpiar.ClearContents
    End If
    
    ' Redimensionar la tabla CONTRATOS
    Application.StatusBar = "Redimensionando tabla..."
    Debug.Print "Redimensionando tabla CONTRATOS..."
    Dim tablaContratos As ListObject
    Dim nuevaUltimaFila As Long
    Dim rangoTabla As Range
    
    On Error Resume Next
    Set tablaContratos = ws.ListObjects("CONTRATOS")
    On Error GoTo ErrorHandler
    
    If Not tablaContratos Is Nothing Then
        ' Guardar el tamaño anterior de la tabla antes de redimensionar
        Dim ultimaFilaTablaAnterior As Long
        ultimaFilaTablaAnterior = tablaContratos.Range.Rows.Count + tablaContratos.Range.Row - 1
        
        ' Determinar la nueva ultima fila (asegurarse de que sea al menos la fila 2)
        nuevaUltimaFila = ultimaFilaConDatos
        If nuevaUltimaFila < 2 Then
            nuevaUltimaFila = 2
        End If
        
        ' Redimensionar la tabla desde A1 hasta I[nuevaUltimaFila]
        Set rangoTabla = ws.Range("A1:I" & nuevaUltimaFila)
        tablaContratos.Resize rangoTabla
        
        Debug.Print "Tabla redimensionada a: A1:I" & nuevaUltimaFila
        
        ' Despues de redimensionar, limpiar contenido de A:D que quede por debajo
        ' (si la tabla anterior tenia mas filas)
        If nuevaUltimaFila < ultimaFilaTablaAnterior Then
            Application.StatusBar = "Limpiando contenido sobrante..."
            Debug.Print "Limpiando contenido de A:D desde fila " & (nuevaUltimaFila + 1) & " hasta " & ultimaFilaTablaAnterior
            Set rangoLimpiar = ws.Range("A" & (nuevaUltimaFila + 1) & ":D" & ultimaFilaTablaAnterior)
            rangoLimpiar.ClearContents
        End If
    Else
        Debug.Print "Advertencia: No se encontro la tabla 'CONTRATOS'"
    End If
    
    ' Cerrar conexiones
    rs.Close
    conn.Close
    Set rs = Nothing
    Set conn = Nothing
    
    ' Mensaje de exito
    Application.StatusBar = "Proceso completado: " & totalRegistros & " registros actualizados"
    Debug.Print "Proceso completado exitosamente"
    Debug.Print "Total de registros: " & totalRegistros
    MsgBox "Consulta completada exitosamente." & vbCrLf & _
           "Total de registros: " & totalRegistros & vbCrLf & _
           "Los datos se han actualizado en la hoja 'CONTRATOS'", _
           vbInformation, "Consulta Contratos"
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
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
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox mensajeError, vbCritical, "Error en Consulta Contratos"
    
End Sub

