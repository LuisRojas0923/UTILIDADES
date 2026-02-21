'Attribute VB_Name = "TestConsultaEstablecimiento"
' Script de testing para validar la consulta de establecimiento
' Este archivo sera eliminado despues de la validacion

Option Explicit

Sub TestConsultaEstablecimiento()
    ' Ejecuta la consulta y valida los resultados basicos
    
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim numColumnas As Integer
    Dim resultado As String
    
    On Error GoTo ErrorHandler
    
    ' Ejecutar la consulta
    Call ConsultarEstablecimiento
    
    ' Validar que la hoja existe
    Set ws = ThisWorkbook.Worksheets("ESTABLECIMIENTO")
    If ws Is Nothing Then
        resultado = "ERROR: La hoja ESTABLECIMIENTO no existe"
        GoTo MostrarResultado
    End If
    
    ' Validar que hay datos
    ultimaFila = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFila < 2 Then
        resultado = "ERROR: No se encontraron datos en la hoja"
        GoTo MostrarResultado
    End If
    
    ' Validar numero de columnas esperadas (debe ser alrededor de 60+ columnas)
    numColumnas = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If numColumnas < 50 Then
        resultado = "ADVERTENCIA: Se esperaban mas columnas. Encontradas: " & numColumnas
        GoTo MostrarResultado
    End If
    
    ' Validar que existen columnas clave
    Dim encontrado As Boolean
    encontrado = False
    Dim col As Integer
    For col = 1 To numColumnas
        If UCase(ws.Cells(1, col).Value) = "CEDULA" Then
            encontrado = True
            Exit For
        End If
    Next col
    
    If Not encontrado Then
        resultado = "ERROR: No se encontro la columna CEDULA"
        GoTo MostrarResultado
    End If
    
    ' Validar que hay registros
    Dim numRegistros As Long
    numRegistros = ultimaFila - 1
    
    resultado = "VALIDACION EXITOSA:" & vbCrLf & _
                "Hoja: ESTABLECIMIENTO" & vbCrLf & _
                "Columnas: " & numColumnas & vbCrLf & _
                "Registros: " & numRegistros & vbCrLf & _
                "Ultima fila: " & ultimaFila
    
    Debug.Print resultado
    
MostrarResultado:
    MsgBox resultado, vbInformation, "Test Consulta Establecimiento"
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR en test: " & Err.Description, vbCritical, "Test Consulta Establecimiento"
    Debug.Print "ERROR en test: " & Err.Description
    
End Sub

