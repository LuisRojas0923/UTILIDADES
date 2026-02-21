Sub AplicarFiltro()
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim colViatico As ListColumn
    Dim celda As Range
    Dim filaTabla As Range
    Dim filaIndex As Long
    Dim contadorCeldas As Long
    Dim contadorProcesadas As Long

    Debug.Print "=== INICIO AplicarFiltro ==="
    
    ' Establecer hoja y tabla
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("CONSIGNACIONES")
    If ws Is Nothing Then
        Debug.Print "ERROR: No se encontro la hoja CONSIGNACIONES"
        MsgBox "No se encontro la hoja 'CONSIGNACIONES'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "Hoja encontrada: " & ws.Name
    
    Set tbl = ws.ListObjects("Consignaciones_Viaticos")
    If tbl Is Nothing Then
        Debug.Print "ERROR: No se encontro la tabla Consignaciones_Viaticos"
        MsgBox "No se encontro la tabla 'Consignaciones_Viaticos'.", vbExclamation
        Exit Sub
    End If
    Debug.Print "Tabla encontrada: " & tbl.Name
    Debug.Print "Rango de datos tabla: " & tbl.DataBodyRange.Address
    On Error GoTo 0

    ' Validar si existe la columna "VIATICO A PAGAR?"
    On Error Resume Next
    Set colViatico = tbl.ListColumns("VIATICO A PAGAR?")
    On Error GoTo 0

    If colViatico Is Nothing Then
        Debug.Print "ERROR: No se encontro la columna VIATICO A PAGAR?"
        MsgBox "No se encontro la columna 'VIATICO A PAGAR?'.", vbExclamation
        Exit Sub
    End If
    Dim letraColumna As String
    letraColumna = Split(ws.Cells(1, colViatico.Range.Column).Address, "$")(1)
    Debug.Print "Columna encontrada: " & colViatico.Name & " (Indice: " & colViatico.Index & ", Columna Excel: " & letraColumna & ")"
    Debug.Print "Los valores 1 se asignaran en la COLUMNA K (" & colViatico.DataBodyRange.Address & ")"

    ' Limpiar toda la columna antes de continuar
    If Not colViatico.DataBodyRange Is Nothing Then
        Debug.Print "Limpiando columna. Rango: " & colViatico.DataBodyRange.Address
        colViatico.DataBodyRange.ClearContents
    Else
        Debug.Print "ADVERTENCIA: DataBodyRange de la columna esta vacio"
    End If

    ' Verificar si hay selección
    If Selection Is Nothing Then
        Debug.Print "ERROR: No hay seleccion activa"
        MsgBox "Selecciona al menos una celda dentro de la tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Validar que la seleccion sea un Range
    Dim selRange As Range
    On Error Resume Next
    Set selRange = Selection
    On Error GoTo 0
    
    If selRange Is Nothing Then
        Debug.Print "ERROR: La seleccion no es un rango valido"
        MsgBox "Selecciona al menos una celda dentro de la tabla.", vbExclamation
        Exit Sub
    End If
    
    ' Obtener direccion de forma segura
    Dim direccionSeleccion As String
    On Error Resume Next
    direccionSeleccion = selRange.Address
    If Err.Number <> 0 Then
        direccionSeleccion = "No disponible (Error: " & Err.Description & ")"
        Err.Clear
    End If
    On Error GoTo 0
    
    Debug.Print "Seleccion encontrada: " & direccionSeleccion
    Debug.Print "Numero de celdas seleccionadas: " & selRange.Cells.Count

    ' Recorrer filas unicas de la seleccion (evitar procesar la misma fila multiples veces)
    Dim filasProcesadas As Object
    Set filasProcesadas = CreateObject("Scripting.Dictionary")
    Dim filaExcel As Long
    Dim filaUnica As Variant
    
    contadorCeldas = 0
    contadorProcesadas = 0
    
    ' Primero, identificar todas las filas unicas dentro de la tabla
    For Each celda In selRange.Cells
        contadorCeldas = contadorCeldas + 1
        Debug.Print "Celda " & contadorCeldas & ": " & celda.Address & " (Fila: " & celda.Row & ", Col: " & celda.Column & ")"
        
        ' Verificar que esté dentro de la tabla
        If Not Intersect(celda, tbl.DataBodyRange) Is Nothing Then
            filaExcel = celda.Row
            ' Agregar fila al diccionario si no existe (evita duplicados)
            If Not filasProcesadas.Exists(filaExcel) Then
                filasProcesadas.Add filaExcel, True
                Debug.Print "  -> Fila " & filaExcel & " agregada para procesamiento"
            End If
        Else
            Debug.Print "  -> ADVERTENCIA: Celda fuera de la tabla, se omite"
        End If
    Next celda
    
    Debug.Print "Total filas unicas identificadas: " & filasProcesadas.Count
    
    ' Ahora procesar cada fila unica
    For Each filaUnica In filasProcesadas.Keys
        filaExcel = filaUnica
        ' Verificar que la fila este dentro del rango de datos de la tabla
        If filaExcel >= tbl.DataBodyRange.Cells(1, 1).Row And filaExcel <= tbl.DataBodyRange.Cells(1, 1).Row + tbl.ListRows.Count - 1 Then
            If Not ws.Rows(filaExcel).Hidden Then
                Debug.Print "Procesando fila Excel " & filaExcel
                ' Calcular indice relativo de fila dentro de la tabla
                filaIndex = filaExcel - tbl.DataBodyRange.Cells(1, 1).Row + 1
                Debug.Print "  -> Indice de fila en tabla: " & filaIndex
                Dim celdaAsignar As Range
                Set celdaAsignar = tbl.DataBodyRange.Cells(filaIndex, colViatico.Index)
                Debug.Print "  -> Asignando valor 1 en COLUMNA K, celda: " & celdaAsignar.Address & " (Fila tabla: " & filaIndex & ", Columna indice: " & colViatico.Index & ")"
                
                ' Asignar valor
                celdaAsignar.Value = 1
                
                ' Verificar valor asignado
                Dim valorAsignado As Variant
                valorAsignado = celdaAsignar.Value
                Debug.Print "  -> Valor verificado en COLUMNA K, celda " & celdaAsignar.Address & ": " & valorAsignado & " (Tipo: " & TypeName(valorAsignado) & ")"
                contadorProcesadas = contadorProcesadas + 1
                Debug.Print "  -> VALOR ASIGNADO CORRECTAMENTE en fila tabla " & filaIndex
            Else
                Debug.Print "  -> ADVERTENCIA: Fila " & filaExcel & " oculta, se omite"
            End If
        Else
            Debug.Print "  -> ADVERTENCIA: Fila " & filaExcel & " fuera del rango de datos de la tabla"
        End If
    Next filaUnica
    
    Debug.Print "Total celdas procesadas: " & contadorProcesadas & " de " & contadorCeldas

    ' Eliminar filtro existente en la tabla
    If tbl.AutoFilter.FilterMode Then
        Debug.Print "Eliminando filtro existente"
        tbl.AutoFilter.ShowAllData
    Else
        Debug.Print "No habia filtro activo"
    End If

    ' Verificar valores antes de filtrar
    Debug.Print "=== VERIFICANDO VALORES 1 EN COLUMNA ANTES DE FILTRAR ==="
    Dim celdaVerificar As Range
    Dim filasVisibles As Long
    Dim indiceFilaVerificar As Long
    filasVisibles = 0
    
    Debug.Print "Recorriendo COLUMNA K (" & colViatico.Name & ") - Rango: " & colViatico.DataBodyRange.Address
    
    For Each celdaVerificar In colViatico.DataBodyRange
        indiceFilaVerificar = celdaVerificar.Row - tbl.DataBodyRange.Cells(1, 1).Row + 1
        Debug.Print "  Revisando COLUMNA K, celda " & celdaVerificar.Address & " (Fila Excel: " & celdaVerificar.Row & ", Fila tabla: " & indiceFilaVerificar & ") - Valor: " & celdaVerificar.Value & " (Tipo: " & TypeName(celdaVerificar.Value) & ")"
        
        If celdaVerificar.Value = 1 Then
            filasVisibles = filasVisibles + 1
            Debug.Print "  *** ENCONTRADO VALOR 1 en COLUMNA K: " & celdaVerificar.Address & " (Fila Excel: " & celdaVerificar.Row & ", Fila tabla: " & indiceFilaVerificar & ") ***"
        End If
    Next celdaVerificar
    
    Debug.Print "=== RESUMEN: Total filas con valor 1 encontradas: " & filasVisibles & " ==="
    If filasVisibles = 0 Then
        Debug.Print "ADVERTENCIA: No se encontraron valores 1. El filtro no mostrara ninguna fila."
    Else
        Debug.Print "El filtro mostrara " & filasVisibles & " fila(s) con valor 1."
    End If
    
    ' Aplicar autofiltro para mostrar solo las filas con 1 en la columna "VIATICO A PAGAR?" (COLUMNA K)
    Debug.Print "Aplicando filtro en COLUMNA K (indice " & colViatico.Index & ") con criterio =1"
    
    ' Asegurar que AutoFilter este habilitado
    If Not tbl.ShowAutoFilter Then
        Debug.Print "Habilitando AutoFilter en la tabla"
        tbl.ShowAutoFilter = True
    End If
    
    On Error Resume Next
    tbl.Range.AutoFilter Field:=colViatico.Index, Criteria1:="=1"
    If Err.Number <> 0 Then
        Debug.Print "ERROR al aplicar filtro: " & Err.Description & " (Codigo: " & Err.Number & ")"
        On Error GoTo 0
    Else
        Debug.Print "Filtro aplicado correctamente"
        On Error GoTo 0
        
        ' Verificar estado del filtro
        DoEvents ' Permitir que Excel procese el filtro
        Debug.Print "Estado AutoFilter.FilterMode: " & tbl.AutoFilter.FilterMode
        Debug.Print "Estado ShowAutoFilter: " & tbl.ShowAutoFilter
        
        ' Contar filas visibles despues del filtro
        Dim filaVisible As Range
        Dim contadorVisibles As Long
        contadorVisibles = 0
        For Each filaVisible In tbl.DataBodyRange.Rows
            If Not filaVisible.Hidden Then
                contadorVisibles = contadorVisibles + 1
            End If
        Next filaVisible
        Debug.Print "Filas visibles despues del filtro: " & contadorVisibles & " de " & tbl.ListRows.Count
    End If
    
    Debug.Print "=== FIN AplicarFiltro ==="

End Sub


Sub QuitarFiltroViaticos()
    Dim ws As Worksheet
    Dim tbl As ListObject

    ' Establecer hoja y tabla
    Set ws = ThisWorkbook.Sheets("CONSIGNACIONES")
    Set tbl = ws.ListObjects("Consignaciones_Viaticos")

    ' Verificar si hay filtros activos
    If tbl.AutoFilter.FilterMode Then
        tbl.AutoFilter.ShowAllData
        MsgBox "Filtro eliminado correctamente.", vbInformation
    Else
        MsgBox "No hay filtros activos en la tabla.", vbExclamation
    End If
End Sub

