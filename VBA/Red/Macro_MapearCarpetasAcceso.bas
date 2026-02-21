'Attribute VB_Name = "ModuloMapearCarpetas"
' Macro para mapear carpetas accesibles en unidad de red
' Ruta: \\192.168.0.3
' Autor: Sistema de Utilidades
' Fecha: 2024

Option Explicit

Sub MapearCarpetasAcceso()
    ' Mapea las carpetas a las que se tiene acceso en la unidad de red
    ' Primero lista los recursos compartidos del servidor, luego mapea carpetas dentro de cada uno
    ' y muestra los resultados en una hoja de Excel
    
    Dim rutaRed As String
    Dim objFSO As Object
    Dim objCarpeta As Object
    Dim objSubCarpeta As Object
    Dim ws As Worksheet
    Dim fila As Long
    Dim nivel As Integer
    Dim carpetasAccesibles As Collection
    Dim carpeta As Variant
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim rangoTabla As Range
    Dim ultimaFila As Long
    Dim mensajeError As String
    Dim testDir As String
    Dim objShell As Object
    Dim objFolder As Object
    Dim recursosCompartidos As Collection
    Dim recurso As Variant
    Dim rutaRecurso As String
    
    ' Inicializar
    tiempoInicio = Timer
    rutaRed = "\\192.168.0.3"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set carpetasAccesibles = New Collection
    Set recursosCompartidos = New Collection
    nivel = 0
    
    ' Crear o limpiar hoja de resultados
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Carpetas_Acceso")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Carpetas_Acceso"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Configurar encabezados
    With ws
        .Cells(1, 1).Value = "Estructura"
        .Cells(1, 2).Value = "Nombre Carpeta"
        .Cells(1, 3).Value = "Ruta Completa"
        .Cells(1, 4).Value = "Nivel"
        .Cells(1, 5).Value = "Fecha Acceso"
    End With
    
    fila = 2
    
    ' Obtener lista de recursos compartidos del servidor
    Debug.Print "Obteniendo lista de recursos compartidos de: " & rutaRed
    Application.ScreenUpdating = False
    
    Call ObtenerRecursosCompartidos(rutaRed, recursosCompartidos)
    
    If recursosCompartidos.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "No se encontraron recursos compartidos en: " & rutaRed & vbCrLf & _
               "Verifique la conexion de red y los permisos.", vbCritical, "Error de Acceso"
        Exit Sub
    End If
    
    Debug.Print "Se encontraron " & recursosCompartidos.Count & " recursos compartidos"
    
    ' Procesar cada recurso compartido
    For Each recurso In recursosCompartidos
        rutaRecurso = rutaRed & "\" & recurso
        Debug.Print "Procesando recurso compartido: " & rutaRecurso
        
        ' Intentar acceder al recurso compartido
        On Error Resume Next
        Set objCarpeta = objFSO.GetFolder(rutaRecurso)
        
        If Err.Number = 0 Then
            ' Agregar recurso compartido como carpeta raiz (solo nivel 0)
            ws.Cells(fila, 1).Value = "├─ " & recurso
            ws.Cells(fila, 2).Value = recurso
            ws.Cells(fila, 3).Value = rutaRecurso
            ws.Cells(fila, 4).Value = nivel
            ws.Cells(fila, 5).Value = Now
            fila = fila + 1
            
            ' Solo nivel 0 - no explorar subcarpetas
        Else
            ' Recurso compartido no accesible - registrar pero continuar
            Debug.Print "Sin acceso a: " & rutaRecurso & " - " & Err.Description
            ws.Cells(fila, 1).Value = "├─ " & recurso & " (SIN ACCESO)"
            ws.Cells(fila, 2).Value = recurso
            ws.Cells(fila, 3).Value = rutaRecurso
            ws.Cells(fila, 4).Value = nivel
            ws.Cells(fila, 5).Value = Now
            ws.Range("A" & fila & ":E" & fila).Font.Color = RGB(255, 0, 0)
            fila = fila + 1
            Err.Clear
        End If
        On Error GoTo 0
        
        DoEvents
    Next recurso
    
    ' Crear tabla de Excel
    ultimaFila = fila - 1
    
    If ultimaFila > 1 Then
        Set rangoTabla = ws.Range("A1:E" & ultimaFila)
        
        ' Convertir a tabla de Excel
        On Error Resume Next
        ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TablaCarpetas"
        On Error GoTo 0
        
        ' Formatear tabla
        With ws.ListObjects("TablaCarpetas")
            .TableStyle = "TableStyleMedium9"
            .ShowAutoFilter = True
        End With
        
        ' Ajustar columnas
        ws.Columns("A").ColumnWidth = 50
        ws.Columns("B").ColumnWidth = 30
        ws.Columns("C").ColumnWidth = 60
        ws.Columns("D").ColumnWidth = 8
        ws.Columns("E").ColumnWidth = 18
    End If
    
    Application.ScreenUpdating = True
    
    tiempoFin = Timer
    Debug.Print "Mapeo completado en " & Format(tiempoFin - tiempoInicio, "0.00") & " segundos"
    
    MsgBox "Mapeo completado." & vbCrLf & _
           "Total de carpetas encontradas: " & (fila - 2) & vbCrLf & _
           "Tiempo transcurrido: " & Format(tiempoFin - tiempoInicio, "0.00") & " segundos" & vbCrLf & _
           "Resultados en hoja: " & ws.Name, vbInformation, "Mapeo Finalizado"
    
    ' Limpiar objetos
    Set objFSO = Nothing
    Set objCarpeta = Nothing
    Set ws = Nothing
End Sub

Private Sub ObtenerRecursosCompartidos(ByVal servidor As String, ByRef recursosCompartidos As Collection)
    ' Obtiene la lista de recursos compartidos de un servidor usando WMI
    
    Dim objWMIService As Object
    Dim colItems As Object
    Dim objItem As Object
    Dim nombreServidor As String
    Dim nombreRecurso As String
    
    ' Extraer nombre del servidor de la ruta UNC
    nombreServidor = Replace(Replace(servidor, "\\", ""), "\", "")
    
    On Error Resume Next
    
    ' Intentar conectar a WMI del servidor remoto
    Set objWMIService = GetObject("winmgmts:\\" & nombreServidor & "\root\cimv2")
    
    If Err.Number <> 0 Then
        Debug.Print "No se pudo conectar a WMI, intentando metodo alternativo..."
        Err.Clear
        
        ' Metodo alternativo: usar Shell para ejecutar net view
        Call ObtenerRecursosCompartidosNetView(servidor, recursosCompartidos)
        Exit Sub
    End If
    
    ' Consultar recursos compartidos
    Set colItems = objWMIService.ExecQuery("SELECT Name FROM Win32_Share WHERE Type = 0")
    
    For Each objItem In colItems
        nombreRecurso = Trim(objItem.Name)
        ' Excluir recursos compartidos del sistema como IPC$, ADMIN$, etc.
        If Right(nombreRecurso, 1) <> "$" And nombreRecurso <> "" Then
            recursosCompartidos.Add nombreRecurso
            Debug.Print "Recurso compartido encontrado: " & nombreRecurso
        End If
    Next
    
    ' Si no se encontraron recursos, intentar metodo alternativo
    If recursosCompartidos.Count = 0 Then
        Debug.Print "WMI no devolvio recursos, intentando metodo alternativo..."
        Call ObtenerRecursosCompartidosNetView(servidor, recursosCompartidos)
    End If
    
    On Error GoTo 0
End Sub

Private Sub ObtenerRecursosCompartidosNetView(ByVal servidor As String, ByRef recursosCompartidos As Collection)
    ' Metodo alternativo: usar net view para obtener recursos compartidos
    
    Dim objShell As Object
    Dim objExec As Object
    Dim strOutput As String
    Dim lineas() As String
    Dim i As Long
    Dim linea As String
    Dim nombreRecurso As String
    Dim inicioNombre As Long
    Dim finNombre As Long
    
    Set objShell = CreateObject("WScript.Shell")
    
    On Error Resume Next
    
    ' Ejecutar comando net view
    Set objExec = objShell.Exec("net view """ & servidor & """")
    
    ' Esperar a que termine (maximo 10 segundos)
    Dim tiempoEspera As Double
    tiempoEspera = Timer
    Do While objExec.Status = 0 And (Timer - tiempoEspera) < 10
        DoEvents
    Loop
    
    ' Leer salida
    strOutput = objExec.StdOut.ReadAll
    
    ' Parsear salida - buscar lineas con recursos compartidos
    lineas = Split(strOutput, vbCrLf)
    
    Dim enSeccionRecursos As Boolean
    enSeccionRecursos = False
    
    For i = 0 To UBound(lineas)
        linea = Trim(lineas(i))
        
        ' Detectar inicio de la seccion de recursos compartidos
        If InStr(LCase(linea), "nombre de recurso compartido") > 0 Or _
           InStr(LCase(linea), "share name") > 0 Then
            enSeccionRecursos = True
            ' Saltar la siguiente linea que es el separador "---"
            i = i + 1
            If i <= UBound(lineas) Then
                linea = Trim(lineas(i))
            End If
        End If
        
        ' Si estamos en la seccion de recursos, procesar lineas
        If enSeccionRecursos And Len(linea) > 0 Then
            ' Buscar lineas que contengan "Disco" o que tengan formato de recurso compartido
            If InStr(linea, "Disco") > 0 Or (InStr(linea, "---") = 0 And Left(linea, 1) <> "" And Len(linea) > 3) Then
                ' Extraer nombre del recurso (primera palabra antes de espacios multiples o "Disco")
                inicioNombre = 1
                
                ' Buscar fin del nombre (antes de "Disco" o espacios multiples)
                finNombre = InStr(linea, "Disco")
                If finNombre <= 0 Then
                    finNombre = InStr(linea, "  ")
                End If
                If finNombre <= 0 Then
                    finNombre = InStr(linea, vbTab)
                End If
                If finNombre <= 0 Then
                    finNombre = Len(linea) + 1
                End If
                
                If finNombre > inicioNombre Then
                    nombreRecurso = Trim(Left(linea, finNombre - 1))
                    
                    ' Validar que sea un nombre valido
                    If nombreRecurso <> "" And Right(nombreRecurso, 1) <> "$" And _
                       nombreRecurso <> "Nombre de recurso compartido" And _
                       InStr(nombreRecurso, "---") = 0 And _
                       InStr(LCase(nombreRecurso), "share name") = 0 And _
                       InStr(LCase(nombreRecurso), "comentario") = 0 Then
                        ' Verificar que no este ya en la coleccion
                        Dim yaExiste As Boolean
                        yaExiste = False
                        Dim recursoExistente As Variant
                        For Each recursoExistente In recursosCompartidos
                            If recursoExistente = nombreRecurso Then
                                yaExiste = True
                                Exit For
                            End If
                        Next
                        
                        If Not yaExiste Then
                            recursosCompartidos.Add nombreRecurso
                            Debug.Print "Recurso compartido encontrado (net view): " & nombreRecurso
                        End If
                    End If
                End If
            End If
        End If
    Next i
    
    On Error GoTo 0
    Set objShell = Nothing
End Sub

Private Sub ExplorarSubcarpetas(ByRef carpetaPadre As Object, ByRef ws As Worksheet, _
                                ByRef fila As Long, ByVal nivelActual As Integer, _
                                ByRef objFSO As Object)
    ' Funcion recursiva para explorar subcarpetas
    
    Dim objSubCarpeta As Object
    Dim errorOcurrido As Boolean
    Dim indentacion As String
    Dim i As Integer
    
    On Error Resume Next
    
    ' Intentar acceder a las subcarpetas
    For Each objSubCarpeta In carpetaPadre.SubFolders
        errorOcurrido = False
        
        ' Intentar acceder a la carpeta
        Err.Clear
        Set objSubCarpeta = objFSO.GetFolder(objSubCarpeta.Path)
        
        If Err.Number = 0 Then
            ' Carpeta accesible - agregar a la lista
            indentacion = ""
            
            ' Crear indentacion visual para mostrar jerarquia
            For i = 1 To nivelActual
                indentacion = indentacion & "    "
            Next i
            indentacion = indentacion & "├─ "
            
            ws.Cells(fila, 1).Value = indentacion & objSubCarpeta.Name
            ws.Cells(fila, 2).Value = objSubCarpeta.Name
            ws.Cells(fila, 3).Value = objSubCarpeta.Path
            ws.Cells(fila, 4).Value = nivelActual
            ws.Cells(fila, 5).Value = Now
            
            fila = fila + 1
            
            ' Actualizar barra de estado
            Application.StatusBar = "Explorando: " & objSubCarpeta.Path
            
            ' Continuar explorando recursivamente (limitar profundidad para evitar loops infinitos)
            If nivelActual < 10 Then
                Call ExplorarSubcarpetas(objSubCarpeta, ws, fila, nivelActual + 1, objFSO)
            End If
        Else
            ' Error al acceder - registrar pero continuar
            Debug.Print "Sin acceso a: " & objSubCarpeta.Path & " - Error: " & Err.Description
        End If
        
        ' Pequena pausa para no sobrecargar la red
        DoEvents
    Next objSubCarpeta
    
    On Error GoTo 0
End Sub

Sub MapearCarpetasAccesoNivel1()
    ' Version simplificada que solo mapea el primer nivel de carpetas
    ' Util para exploraciones rapidas
    
    Dim rutaRed As String
    Dim objFSO As Object
    Dim objCarpeta As Object
    Dim objSubCarpeta As Object
    Dim ws As Worksheet
    Dim fila As Long
    Dim rangoTabla As Range
    Dim ultimaFila As Long
    Dim mensajeError As String
    Dim testDir As String
    
    rutaRed = "\\192.168.0.3"
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    ' Intentar acceso directo a la ruta de red
    Debug.Print "Accediendo a: " & rutaRed
    
    ' Intentar acceso directo - no bloquear si falla la verificacion
    On Error Resume Next
    Set objCarpeta = objFSO.GetFolder(rutaRed)
    
    ' Si GetFolder falla, intentar con Dir() que es mas tolerante
    If Err.Number <> 0 Then
        mensajeError = Err.Description
        Debug.Print "GetFolder() fallo: " & mensajeError
        Err.Clear
        
        ' Intentar con Dir() para verificar si realmente hay acceso
        testDir = Dir(rutaRed & "\*", vbDirectory)
        
        ' Si Dir() funciona, hay acceso - reintentar GetFolder
        If testDir <> "" Then
            Debug.Print "Dir() confirma acceso disponible"
            Err.Clear
            Set objCarpeta = objFSO.GetFolder(rutaRed)
        End If
        
        ' Si GetFolder sigue fallando pero Dir() funciono, continuar de todas formas
        If Err.Number <> 0 And testDir <> "" Then
            Debug.Print "Advertencia: GetFolder falla pero hay acceso confirmado por Dir()"
            Debug.Print "Continuando - el acceso se verificara al listar carpetas"
            Err.Clear
        End If
    End If
    
    ' Solo salir si realmente no hay forma de acceder
    If objCarpeta Is Nothing And (testDir = "" Or testDir = vbNullString) Then
        On Error GoTo 0
        MsgBox "No se puede acceder a: " & rutaRed & vbCrLf & _
               "Error: " & mensajeError & vbCrLf & vbCrLf & _
               "Verifique la conexion de red y los permisos.", vbCritical, "Error"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Crear o limpiar hoja
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Carpetas_Nivel1")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "Carpetas_Nivel1"
    Else
        ws.Cells.Clear
    End If
    On Error GoTo 0
    
    ' Encabezados
    With ws
        .Cells(1, 1).Value = "Estructura"
        .Cells(1, 2).Value = "Nombre Carpeta"
        .Cells(1, 3).Value = "Ruta Completa"
    End With
    
    fila = 2
    
    ' Solo primer nivel
    For Each objSubCarpeta In objCarpeta.SubFolders
        On Error Resume Next
        ws.Cells(fila, 1).Value = "├─ " & objSubCarpeta.Name
        ws.Cells(fila, 2).Value = objSubCarpeta.Name
        ws.Cells(fila, 3).Value = objSubCarpeta.Path
        If Err.Number = 0 Then
            fila = fila + 1
        End If
        On Error GoTo 0
        DoEvents
    Next
    
    ' Crear tabla de Excel
    ultimaFila = fila - 1
    
    If ultimaFila > 1 Then
        Set rangoTabla = ws.Range("A1:C" & ultimaFila)
        
        ' Convertir a tabla de Excel
        On Error Resume Next
        ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TablaCarpetasNivel1"
        On Error GoTo 0
        
        ' Formatear tabla
        With ws.ListObjects("TablaCarpetasNivel1")
            .TableStyle = "TableStyleMedium9"
            .ShowAutoFilter = True
        End With
        
        ' Ajustar columnas
        ws.Columns("A").ColumnWidth = 50
        ws.Columns("B").ColumnWidth = 30
        ws.Columns("C").ColumnWidth = 60
    End If
    
    MsgBox "Mapeo de primer nivel completado." & vbCrLf & _
           "Carpetas encontradas: " & (fila - 2), vbInformation, "Completado"
    
    Set objFSO = Nothing
    Set objCarpeta = Nothing
    Set ws = Nothing
End Sub

