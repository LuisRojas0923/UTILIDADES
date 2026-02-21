Private Sub btn_ActualizarTablasDV_Click()

    'Actualizar las tablas dinamicas de la hoja DETALLE VIATICOS
    'Ciclo: guardar-actualizar-guardar-actualizar
    
    Dim i As Integer
    Dim tiempoEspera As Date
    
    'Configurar tiempo de espera (10 segundos)
    tiempoEspera = Now + TimeValue("00:00:10")
    
    'Ciclo de 2 iteraciones: guardar-actualizar, guardar-actualizar
    For i = 1 To 2
        
        'Guardar el libro
        Application.StatusBar = "Guardando libro... (Iteracion " & i & " de 2)"
        Debug.Print "Guardando libro - Iteracion " & i & " de 2"
        ThisWorkbook.Save
        DoEvents
        
        'Esperar 10 segundos
        Application.StatusBar = "Esperando 10 segundos antes de actualizar... (Iteracion " & i & " de 2)"
        Debug.Print "Esperando 10 segundos - Iteracion " & i & " de 2"
        Application.Wait tiempoEspera
        tiempoEspera = Now + TimeValue("00:00:10")
        
        'Actualizar las tablas
        Application.StatusBar = "Actualizando tablas dinamicas... (Iteracion " & i & " de 2)"
        Debug.Print "Actualizando tablas dinamicas - Iteracion " & i & " de 2"
        Call Actualizar_Tablas.ActualizarTablas
        DoEvents
        
        'Esperar 10 segundos antes de la siguiente iteracion (excepto en la ultima)
        If i < 2 Then
            Application.StatusBar = "Esperando 10 segundos antes de la siguiente iteracion..."
            Debug.Print "Esperando 10 segundos antes de la siguiente iteracion"
            Application.Wait tiempoEspera
            tiempoEspera = Now + TimeValue("00:00:10")
        End If
        
    Next i
    
    'Restaurar barra de estado y mostrar mensaje final
    Application.StatusBar = "Proceso completado: 2 guardados y 2 actualizaciones realizadas"
    Debug.Print "Proceso completado: 2 guardados y 2 actualizaciones realizadas"
    
    'Restaurar barra de estado a su estado normal
    Application.StatusBar = False

End Sub