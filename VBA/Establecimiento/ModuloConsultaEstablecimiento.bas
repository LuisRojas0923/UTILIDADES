'Attribute VB_Name = "ModuloConsultaEstablecimiento"
' Macro para consultar datos de establecimiento desde PostgreSQL
' Base de datos: solid. Optimizado para maxima velocidad.
' Autor: Ing Luis Enrique Rojas | Fecha: 2026

Option Explicit

Private Const ADOPEN_FORWARDONLY As Long = 0
Private Const ADLOCK_READONLY As Long = 1
Private Const ADCMDTEXT As Long = 1
Private Const ADCURSORSERVER As Long = 2

' Fallback cuando CopyFromRecordset da Error 13 (incompatibilidad de tipos con el driver).
Private Sub VolcarRecordsetEnHoja(conn As Object, consultaSQL As String, ws As Worksheet, filaInicio As Long)
    Dim rs2 As Object
    Dim datos As Variant
    Dim r As Long, c As Long
    Dim nCols As Long, nRows As Long
    Dim arr() As Variant
    Dim v As Variant
    Set rs2 = CreateObject("ADODB.Recordset")
    rs2.CursorType = ADOPEN_FORWARDONLY
    rs2.CursorLocation = ADCURSORSERVER
    rs2.LockType = ADLOCK_READONLY
    rs2.Open consultaSQL, conn, ADOPEN_FORWARDONLY, ADLOCK_READONLY, ADCMDTEXT
    If rs2.EOF Then rs2.Close: Set rs2 = Nothing: Exit Sub
    datos = rs2.GetRows
    rs2.Close
    Set rs2 = Nothing
    If IsEmpty(datos) Then Exit Sub
    nCols = UBound(datos, 1) + 1
    nRows = UBound(datos, 2) + 1
    ReDim arr(1 To nRows, 1 To nCols)
    For r = 0 To nRows - 1
        For c = 0 To nCols - 1
            v = datos(c, r)
            If IsNull(v) Then v = ""
            arr(r + 1, c + 1) = v
        Next c
    Next r
    ws.Range(ws.Cells(filaInicio, 1), ws.Cells(filaInicio + nRows - 1, nCols)).Value = arr
End Sub

Private Function Password(cadenaDatos As String) As String
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

' Carga la consulta desde consulta_establecimiento.sql en la misma carpeta del libro; si falla usa consulta embebida.
' Public para que ModuloConsultaEstablecimientoActivos pueda reutilizarla.
Public Function CargarConsultaSQL() As String
    Dim ruta As String
    Dim numArchivo As Integer
    Dim contenido As String
    Dim linea As String
    Dim posPuntoComa As Long
    
    On Error Resume Next
    ruta = ThisWorkbook.Path
    If Len(ruta) > 0 Then
        ruta = ruta & "\consulta_establecimiento.sql"
        numArchivo = FreeFile
        Open ruta For Input As #numArchivo
        If Err.Number = 0 Then
            contenido = Input(LOF(numArchivo), numArchivo)
            Close #numArchivo
            posPuntoComa = InStr(1, contenido, ";")
            If posPuntoComa > 0 Then contenido = Trim(Left(contenido, posPuntoComa - 1))
            CargarConsultaSQL = contenido
            On Error GoTo 0
            Exit Function
        End If
    End If
    On Error GoTo 0
    ' Fallback: consulta embebida (alineada con consulta_establecimiento.sql)
    CargarConsultaSQL = ConsultaSQLEmbebida()
End Function

' Fallback cuando no se encuentra el archivo .sql. Debe coincidir con consulta_establecimiento.sql.
Public Function ConsultaSQLEmbebida() As String
    Dim a(1 To 92) As String
    a(1) = "WITH primera_fecha AS ("
    a(2) = "    SELECT TRIM(CAST(establecimiento AS TEXT)) AS establecimiento_trim, MIN(fechainicio) AS primera_fecha_ingreso FROM contrato GROUP BY TRIM(CAST(establecimiento AS TEXT))"
    a(3) = ")"
    a(4) = "SELECT DISTINCT"
    a(5) = "    CAST(NULLIF(TRIM(CAST(E.nrocedula AS TEXT)), '') AS NUMERIC) AS ""CEDULA"","
    a(6) = "    E.nombre::text AS ""NOMBRE"","
    a(7) = "    CASE WHEN B.estado = 'Activo' THEN 'A' WHEN B.estado = 'Inactivo' THEN 'RET' WHEN B.estado IS NOT NULL AND B.estado <> '' THEN B.estado::text WHEN C.estado = 'Activo' THEN 'A' WHEN C.estado = 'Terminado' THEN 'RET' ELSE 'RET' END::text AS ""ESTADO"","
    a(8) = "    C.empresa::text AS ""EMPRESA"","
    a(9) = "    C.area::text AS ""AREA"","
    a(10) = "    C.cargo::text AS ""CARGO ESTABLECIMIENTO"","
    a(11) = "    C.gerencia::text AS ""GERENCIA"","
    a(12) = "    C.jefe::text AS ""JEFE INMEDIATO"","
    a(13) = "    C.regional::text AS ""REGIONAL QUE REPORTA"","
    a(14) = "    C.centrocosto::text AS ""CENTRO DE COSTO"","
    a(15) = "    CAST(NULLIF(TRIM(CAST(C.cuenta AS TEXT)), '') AS NUMERIC) AS ""CUENTA"","
    a(16) = "    CAST(NULLIF(TRIM(CAST(C.tipocontable AS TEXT)), '') AS NUMERIC) AS ""TIPO"","
    a(17) = "    C.nomina::text AS ""NOMINA"","
    a(18) = "    CAST(NULLIF(TRIM(CAST(C.riesgoarl AS TEXT)), '') AS DECIMAL) AS ""RIESGO ARL"","
    a(19) = "    C.arl::text AS ""A.R.L"","
    a(20) = "    E.correocorporativo::text AS ""CORREO CORPORTIVO"","
    a(21) = "    CASE WHEN E.viaticante = true THEN 'SI' ELSE 'NO' END::text AS ""VIATICANTE"","
    a(22) = "    COALESCE(E.baseviaticos, 0)::numeric AS ""BASE VIATICOS"","
    a(23) = "    TRIM(CAST(C.numerocontrato AS TEXT))::text AS ""NUMERO CONTRATO"","
    a(24) = "    C.estado::text AS ""ESTADO CONTRATO"","
    a(25) = "    C.tipo::text AS ""TIPO DE CONTRATO"","
    a(26) = "    C.ciudadcontratacion::text AS ""CIUDAD CONTRATACION"","
    a(27) = "    PF.primera_fecha_ingreso::date AS ""1A FECHA INGRESO"","
    a(28) = "    C.fechainicio::date AS ""FECHA ULTIMO INGRESO"","
    a(29) = "    CASE WHEN C.fechainicioetapaproductiva = '1900-01-01'::date THEN NULL ELSE C.fechainicioetapaproductiva END::date AS ""FECHA INICIO ETAPA PRODUCTIVA"","
    a(30) = "    C.fechavencimiento::date AS ""FECHA VENCIMIENTO CONTRATO"","
    a(31) = "    CASE WHEN C.fecharetiro IN ('1900-01-01'::date, '1900-01-02'::date, '1990-01-01'::date) THEN NULL ELSE C.fecharetiro END::date AS ""FECHA RETIRO"","
    a(32) = "    CASE WHEN C.estado = 'Activo' AND UPPER(TRIM(COALESCE(C.empresa, '')::text)) = 'REFRIDCOL' THEN (C.fechavencimiento - INTERVAL '35 days')::date ELSE NULL END::date AS ""CARTA DE PRORROGA"","
    a(33) = "    C.observaciones::text AS ""OBSERVACIONES CONTRATO"","
    a(34) = "    TRIM(CAST(B.contrato AS TEXT))::text AS ""NUMERO BENEFICIO"","
    a(35) = "    B.estado::text AS ""ESTADO BENEFICIO"","
    a(36) = "    CAST(NULLIF(TRIM(CAST(B.salario AS TEXT)), '') AS NUMERIC) AS ""SALARIO AÑO 2025"","
    a(37) = "    B.auxiliolegaltransporte::text AS ""AUXILIO LEGAL TRANSPORTE"","
    a(38) = "    CAST(NULLIF(TRIM(CAST(B.auxilioalimentacionmensual AS TEXT)), '') AS NUMERIC) AS ""AUX ALIMENTACION MENSUAL"","
    a(39) = "    CAST(NULLIF(TRIM(CAST(B.auxilioalimentacionquincenal AS TEXT)), '') AS NUMERIC) AS ""AUX ALIMENTACION (QUINCENAL)"","
    a(40) = "    CAST(NULLIF(TRIM(CAST(B.auxiliovivienda AS TEXT)), '') AS NUMERIC) AS ""AUX VIVIENDA "","
    a(41) = "    CAST(NULLIF(TRIM(CAST(B.rodamiento AS TEXT)), '') AS NUMERIC) AS ""RODAMIENTO"","
    a(42) = "    CAST(NULLIF(TRIM(CAST(B.baserodamiento AS TEXT)), '') AS NUMERIC) AS ""BASE RODAMIENTO"","
    a(43) = "    CAST(NULLIF(TRIM(CAST(B.valormaximobono AS TEXT)), '') AS NUMERIC) AS ""VALOR MAXIMO BONO 40%"","
    a(44) = "    CAST(NULLIF(TRIM(CAST(B.capacidadendeudamiento AS TEXT)), '') AS NUMERIC) AS ""CAPACIDAD ENDEUDAMIENTO"","
    a(45) = "    CASE WHEN B.autorizacionrodamiento THEN 'SI' ELSE 'NO' END::text AS ""AUTORIZAN RODAMIENTO"","
    a(46) = "    CASE WHEN B.autorizacionhorasextras THEN 'SI' ELSE 'NO' END::text AS ""AUTORIZAN HORAS EXTRAS"","
    a(47) = "    B.observaciones::text AS ""OBSERVACIONES BENEFICIO"","
    a(48) = "    C.banco::text AS ""BANCO"","
    a(49) = "    CAST(NULLIF(TRIM(CAST(C.cuentanomina AS TEXT)), '') AS NUMERIC) AS ""CTA DE NOMINA"","
    a(50) = "    C.eps::text AS ""E.P.S"","
    a(51) = "    C.afp::text AS ""A.F.P"","
    a(52) = "    C.ccf::text AS ""C.C.F"","
    a(53) = "    C.afc::text AS ""A.F.C"","
    a(54) = "    E.examenmedico::text AS ""EXAMEN DE INGRESO"","
    a(55) = "    E.fechavencimientoexamenmedico::date AS ""FECHA VENCIMIENTO EXAMEN DE INGRESO"","
    a(56) = "    E.certificadoalturas::text AS ""CERTIFICADO TRABAJO EN ALTURAS"","
    a(57) = "    E.fechavencimientoalturas::date AS ""FECHA VENCIMIENTO TRABAJO EN ALTURAS"","
    a(58) = "    E.certificadoprimerosauxilios::text AS ""CERTIFICADO PRIMEROS AUXILIOS"","
    a(59) = "    E.fechavencimientoprimerosauxilios::date AS ""FECHA VENCIMIENTO PRIMEROS AUXILIOS"","
    a(60) = "    C.archivo::text AS ""ARCHIVO CONTRATO"","
    a(61) = "    E.fechanacimiento::date AS ""FECHA NACIMIENTO"","
    a(62) = "    TRIM(TO_CHAR(E.fechanacimiento, 'Mon'))::text AS ""MES CUMPLEAÑOS"","
    a(63) = "    E.sexo::text AS ""SEXO"","
    a(64) = "    EXTRACT(YEAR FROM AGE(E.fechanacimiento::date))::numeric AS ""EDAD"","
    a(65) = "    E.rh::text AS ""RH"","
    a(66) = "    CAST(NULLIF(TRIM(CAST(E.nrohijos AS TEXT)), '') AS NUMERIC) AS ""CUANTOS HIJOS"","
    a(67) = "    E.correopersonal::text AS ""CORREO PERSONAL"","
    a(68) = "    E.gradoalcanzado::text AS ""GRADO ALCANZADO"","
    a(69) = "    E.tituloobtenido::text AS ""TITULO OBTENIDO"","
    a(70) = "    E.tarjetaprofesional::text AS ""TARJETA PROFESIONAL"","
    a(71) = "    E.ciudadresidencia::text AS ""CUIDAD DE RESIDENCIA"","
    a(72) = "    E.direccionresidencia::text AS ""DIRECCION RESIDENCIA"","
    a(73) = "    E.barrio::text AS ""BARRIO"","
    a(74) = "    E.telefono::text AS ""TELEFONO"","
    a(75) = "    E.tallacamisa::text AS ""TALLA CAMISA"","
    a(76) = "    E.tallapantalon::text AS ""TALLA PANTALON "","
    a(77) = "    E.tallabotas::text AS ""TALLA BOTAS"","
    a(78) = "    E.contactoemergencia::text AS ""CONTACTO EMERGENCIA"","
    a(79) = "    E.parentesco::text AS ""PARENTESCO"","
    a(80) = "    E.telefonocontactoemergencia::text AS ""TELEFONO CONTACTO EMERGENCIA"","
    a(81) = "    E.archivo1::text AS ""ARCHIVO 1"","
    a(82) = "    E.archivo2::text AS ""ARCHIVO 2"","
    a(83) = "    E.archivo3::text AS ""ARCHIVO 3"","
    a(84) = "    E.archivo4::text AS ""ARCHIVO 4"","
    a(85) = "    E.codigo AS ""CODIGO"","
    a(86) = "    E.fecha AS ""FECHA DE REGISTRO"","
    a(87) = "    date_trunc('second', E.hora) AS ""HORA DE REGISTRO"""
    a(88) = "FROM establecimiento E"
    a(89) = "LEFT JOIN primera_fecha PF ON PF.establecimiento_trim = TRIM(CAST(E.nrocedula AS TEXT))"
    a(90) = "LEFT JOIN contrato C ON TRIM(CAST(C.establecimiento AS TEXT)) = TRIM(CAST(E.nrocedula AS TEXT))"
    a(91) = "LEFT JOIN beneficio B ON TRIM(CAST(B.contrato AS TEXT)) = TRIM(CAST(C.numerocontrato AS TEXT))"
    a(92) = "ORDER BY CASE WHEN B.estado = 'Activo' THEN 'A' WHEN B.estado = 'Inactivo' THEN 'RET' WHEN B.estado IS NOT NULL AND B.estado <> '' THEN B.estado::text WHEN C.estado = 'Activo' THEN 'A' WHEN C.estado = 'Terminado' THEN 'RET' ELSE 'RET' END ASC, E.nombre ASC, C.fechainicio DESC NULLS LAST"
    ConsultaSQLEmbebida = Join(a, vbCrLf)
End Function

Sub ConsultarEstablecimiento()
    Dim conn As Object
    Dim rs As Object
    Dim ws As Worksheet
    Dim cadenaConexion As String
    Dim consultaSQL As String
    Dim servidor As String
    Dim puerto As String
    Dim baseDatos As String
    Dim usuario As String
    Dim contrasena As String
    Dim mensajeError As String
    Dim totalRegistros As Long
    Dim tiempoInicio As Double
    Dim tiempoFin As Double
    Dim tiempoEjecucion As Double
    Dim numColumnas As Integer
    Dim nombresColumnas() As String
    Dim headerRow() As Variant
    Dim columna As Long
    Dim filaInicial As Long
    Dim ultimaFilaConDatos As Long
    Dim rangoTabla As Range
    Dim ultimaColumna As String
    Dim tblExistente As Object
    Dim colFecha As Integer
    Dim nombreCol As String
    Dim esColumnaFecha As Boolean
    Dim usoFallback As Boolean

    usoFallback = False
    tiempoInicio = Timer
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    servidor = "192.168.0.21"
    puerto = "5432"
    baseDatos = "solid"
    usuario = "postgres"
    contrasena = Password("65,100,109,105,110,83,111,108,105,100,50,48,50,53")

    consultaSQL = CargarConsultaSQL()

    cadenaConexion = "Driver={PostgreSQL Unicode(x64)};Server=" & servidor & ";Port=" & puerto & ";Database=" & baseDatos & ";Uid=" & usuario & ";Pwd=" & contrasena & ";"

    On Error GoTo ErrorHandler

    ' Hoja destino
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ESTABLECIMIENTO")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ESTABLECIMIENTO"
    Else
        Set tblExistente = ws.ListObjects("TB_ESTABLECIMIENTO")
        If Not tblExistente Is Nothing Then
            ' Quitar filtros si hay alguno aplicado (con manejo de error si no tiene filtros puestos)
            On Error Resume Next
            tblExistente.AutoFilter.ShowAllData
            On Error GoTo ErrorHandler
            If Not tblExistente.DataBodyRange Is Nothing Then tblExistente.DataBodyRange.Delete
        Else
            ws.Range("A13:ZZ" & ws.Rows.Count).Clear
        End If
    End If
    On Error GoTo ErrorHandler

    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")

    Application.StatusBar = "Conectando..."
    conn.ConnectionTimeout = 30
    conn.CommandTimeout = 120
    conn.CursorLocation = ADCURSORSERVER
    conn.Open cadenaConexion

    Application.StatusBar = "Ejecutando consulta..."
    rs.CursorType = ADOPEN_FORWARDONLY
    rs.CursorLocation = ADCURSORSERVER
    rs.LockType = ADLOCK_READONLY
    rs.Open consultaSQL, conn, ADOPEN_FORWARDONLY, ADLOCK_READONLY, ADCMDTEXT

    numColumnas = rs.Fields.Count
    If numColumnas <= 0 Then
        rs.Close
        conn.Close
        Set rs = Nothing
        Set conn = Nothing
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = False
        MsgBox "La consulta no devolvió columnas.", vbExclamation, "Consulta Establecimiento"
        Exit Sub
    End If
    ReDim nombresColumnas(0 To numColumnas - 1)
    For columna = 0 To numColumnas - 1
        nombresColumnas(columna) = CStr(rs.Fields(columna).Name)
    Next columna

    ' Encabezados en un solo rango (mas rapido)
    ReDim headerRow(1 To 1, 1 To numColumnas)
    For columna = 1 To numColumnas
        headerRow(1, columna) = nombresColumnas(columna - 1)
    Next columna
    With ws.Range(ws.Cells(12, 1), ws.Cells(12, numColumnas))
        .Value = headerRow
        .Font.Bold = True
        .Interior.Color = RGB(0, 32, 96)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
        .WrapText = True
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(12).RowHeight = 30

    filaInicial = 13
    On Error Resume Next
    ws.Range("A13").CopyFromRecordset rs
    If Err.Number = 13 Then
        Err.Clear
        rs.Close
        Set rs = Nothing
        usoFallback = True
        VolcarRecordsetEnHoja conn, consultaSQL, ws, 13
    End If
    On Error GoTo ErrorHandler

    ultimaFilaConDatos = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If ultimaFilaConDatos >= filaInicial Then
        totalRegistros = ultimaFilaConDatos - filaInicial + 1
    Else
        totalRegistros = 0
    End If

    ' Columna ultima (letra)
    If numColumnas <= 26 Then
        ultimaColumna = Chr(64 + numColumnas)
    Else
        ultimaColumna = Chr(64 + Int((numColumnas - 1) / 26)) & Chr(65 + ((numColumnas - 1) Mod 26))
    End If

    If totalRegistros > 0 Then
        Set rangoTabla = ws.Range("A12:" & ultimaColumna & (ultimaFilaConDatos))
        On Error Resume Next
        Set tblExistente = ws.ListObjects("TB_ESTABLECIMIENTO")
        If tblExistente Is Nothing Then
            ws.ListObjects.Add(xlSrcRange, rangoTabla, , xlYes).Name = "TB_ESTABLECIMIENTO"
        Else
            tblExistente.Resize rangoTabla
        End If
        On Error GoTo ErrorHandler
        With ws.ListObjects("TB_ESTABLECIMIENTO")
            .TableStyle = "TableStyleMedium9"
            .ShowAutoFilter = True
        End With

        ws.Range("A13:" & ultimaColumna & ultimaFilaConDatos).NumberFormat = "General"
        For colFecha = 0 To numColumnas - 1
            nombreCol = nombresColumnas(colFecha)
            esColumnaFecha = (InStr(1, UCase(nombreCol), "FECHA") > 0 And InStr(1, UCase(nombreCol), "MES") = 0)
            If esColumnaFecha Then
                ws.Columns(colFecha + 1).NumberFormat = "dd/mm/yyyy"
                ws.Columns(colFecha + 1).HorizontalAlignment = xlCenter
            End If
        Next colFecha
        ' Columnas de valores: formato moneda sin decimales
        Dim colMon As Long
        Dim nombreColMon As String
        For colMon = 0 To numColumnas - 1
            nombreColMon = UCase(Trim(nombresColumnas(colMon)))
            If InStr(1, nombreColMon, "BASE VIATICOS") > 0 Or InStr(1, nombreColMon, "SALARIO AÑO") > 0 Or _
               InStr(1, nombreColMon, "AUX ALIMENTACION") > 0 Or InStr(1, nombreColMon, "AUX VIVIENDA") > 0 Or _
               InStr(1, nombreColMon, "RODAMIENTO") > 0 Or InStr(1, nombreColMon, "VALOR MAXIMO BONO") > 0 Or _
               InStr(1, nombreColMon, "CAPACIDAD ENDEUDAMIENTO") > 0 Or InStr(1, nombreColMon, "RIESGO ARL") > 0 Then
                ws.Range(ws.Cells(13, colMon + 1), ws.Cells(ultimaFilaConDatos, colMon + 1)).NumberFormat = "$#,##0;[Red]-$#,##0"
            End If
        Next colMon
        ' ws.Range("A:" & ultimaColumna).Columns.AutoFit  ' Desactivado: no ajustar ancho de columnas automaticamente
    End If

    If Not usoFallback And Not rs Is Nothing Then rs.Close
    Set rs = Nothing
    conn.Close
    Set conn = Nothing

    tiempoFin = Timer
    tiempoEjecucion = tiempoFin - tiempoInicio
    Application.StatusBar = "Completado: " & totalRegistros & " registros"
    ws.Range("F2").Value = "Tiempo: " & Round(tiempoEjecucion, 1) & " s"

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox "Consulta completada exitosamente." & vbCrLf & "Total de registros: " & totalRegistros, vbInformation, "Consulta Establecimiento"
    Exit Sub

ErrorHandler:
    mensajeError = "Error: " & Err.Number & vbCrLf & Err.Description
    Debug.Print mensajeError
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
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = False
    MsgBox mensajeError, vbCritical, "Consulta Establecimiento"
End Sub
