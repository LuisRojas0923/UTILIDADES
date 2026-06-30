let
    Origen = Csv.Document(File.Contents(RUTA),[Delimiter=";", Columns=8, Encoding=1252, QuoteStyle=QuoteStyle.None]),
    #"Otras columnas quitadas" = Table.SelectColumns(#"Origen",{"Column3", "Column5", "Column6", "Column7", "Column8"}),
    #"Encabezados promovidos" = Table.PromoteHeaders(#"Otras columnas quitadas", [PromoteAllScalars=true]),
    #"Consultas combinadas1" = Table.NestedJoin(#"Encabezados promovidos", {"CICLO"}, CICLO, {"CICLO A CONSULTAR:"}, "Tabla8", JoinKind.LeftOuter),
    #"Se expandió Tabla8" = Table.ExpandTableColumn(#"Consultas combinadas1", "Tabla8", {"CICLO A CONSULTAR:"}, {"Tabla8.CICLO A CONSULTAR:"}),
    #"Columna condicional agregada" = Table.AddColumn(#"Se expandió Tabla8", "Personalizado", each if [CICLO] = [#"Tabla8.CICLO A CONSULTAR:"] then "Igual" else "Null"),
    #"Filas filtradas1" = Table.SelectRows(#"Columna condicional agregada", each ([Personalizado] = "Igual")),
    #"Columnas quitadas2" = Table.RemoveColumns(#"Filas filtradas1",{"CICLO", "Personalizado"}),
    #"Columnas con nombre cambiado6" = Table.RenameColumns(#"Columnas quitadas2",{{"Tabla8.CICLO A CONSULTAR:", "CICLO"}, {"IMPUESTOS ", "IVA"}}),
    #"Consultas combinadas" = Table.NestedJoin(#"Columnas con nombre cambiado6", {"DESCRIPCION"}, CRITERIOS, {"DESCRIPCIÓN"}, "CRITERIOS", JoinKind.LeftOuter),
    #"Se expandió CRITERIOS" = Table.ExpandTableColumn(#"Consultas combinadas", "CRITERIOS", {"GRUPO"}, {"CRITERIOS.GRUPO"}),
    #"Columnas con nombre cambiado4" = Table.RenameColumns(#"Se expandió CRITERIOS",{{"DESCRIPCION", "DESCRIPCION1"}}),
    #"Personalizada agregada" = Table.AddColumn(#"Columnas con nombre cambiado4", "DESCRIPCION", each if [CRITERIOS.GRUPO]= null then "Validar" else [CRITERIOS.GRUPO]),
    #"Columnas reordenadas6" = Table.ReorderColumns(#"Personalizada agregada",{"MIN", "DESCRIPCION", "VALOR", "IVA", "CICLO"}),
    #"Tipo cambiado3" = Table.TransformColumnTypes(#"Columnas reordenadas6",{{"VALOR", type number}, {"IVA", type number}}),
    #"Nuevas columnas iva" = Table.AddColumn(#"Tipo cambiado3", "IVA 19%", each if [DESCRIPCION1]= "Devolución de Cargo Fijo Mensual" then 
 [IVA]
else 
 
if [DESCRIPCION1]= "DevoluciÃ³n de Cargo Fijo Mensual GSM" or [DESCRIPCION1] = "DevoluciÃ³n de Servicio Datos"
or [DESCRIPCION1]="Mensajes de Texto a Movistar" or [DESCRIPCION1]="Mensajes de Texto a TIGO"
then
    if [IVA] < 0 then [IVA] else 0
else
    if [IVA] > 0 then [IVA] else 0),
    #"Nueva columna iva2" = Table.AddColumn(#"Nuevas columnas iva", "DESCUENTO IVA", each if [DESCRIPCION1]="Devolución de Cargo Fijo Mensual" 
then 0 else
if [DESCRIPCION1] = "DevoluciÃ³n de Cargo Fijo Mensual GSM" or [DESCRIPCION1] = "DevoluciÃ³n de Servicio Datos"
or [DESCRIPCION1]="Mensajes de Texto a Movistar" or [DESCRIPCION1]="Mensajes de Texto a TIGO"
then
    if [IVA] > 0 then [IVA] else 0
else
    if [IVA] < 0 then [IVA] else 0),
    #"Columnas quitadas1" = Table.RemoveColumns(#"Nueva columna iva2",{"DESCRIPCION1", "CRITERIOS.GRUPO", "IVA"}),
    #"Filas agrupadas" = Table.Group(#"Columnas quitadas1", {"MIN", "DESCRIPCION", "CICLO"}, {{"VALOR", each List.Sum([VALOR]), type nullable number}, {"IVA 19%", each List.Sum([#"IVA 19%"]), type number}, {"DESCUENTO IVA", each List.Sum([DESCUENTO IVA]), type number}}),
    #"Columna dinamizada" = Table.Pivot(#"Filas agrupadas", List.Distinct(#"Filas agrupadas"[DESCRIPCION]), "DESCRIPCION", "VALOR", List.Sum),
    #"Valor reemplazado" = Table.ReplaceValue(#"Columna dinamizada",null,0,Replacer.ReplaceValue,{"CARGO FIJO",  "ESPECIALES"}),
    #"Filas agrupadas1" = Table.Group(#"Valor reemplazado", {"MIN", "CICLO"}, {{"IVA 19%", each List.Sum([#"IVA 19%"]), type number}, {"CARGO FIJO", each List.Sum([CARGO FIJO]), type nullable number}, {"DESCUENTO", each 0, type nullable number}, {"ESPECIALES", each List.Sum([ESPECIALES]), type nullable number}, {"IMPOCONSUMO", each 0, type nullable number}, {"DESCUENTO IVA", each 0, type number}}),
    #"Combinación Planes y Convios" = Table.NestedJoin(#"Filas agrupadas1", {"MIN"}, #"Planes Asignados", {"LINEA"}, "Planes Asignados", JoinKind.LeftOuter),
    #"Se expandió Planes Asignados1" = Table.ExpandTableColumn(#"Combinación Planes y Convios", "Planes Asignados", {"LINEA", "DOCUMENTO DE COBRO", "EMPLEADO DE COBRO", "CENTRO DE COSTO", "NOMBRE DEL PLAN ACTUAL", "CONVENIO #1", "CONVENIO #2"}, {"Planes Asignados.LINEA", "Planes Asignados.DOCUMENTO DE COBRO", "Planes Asignados.EMPLEADO DE COBRO", "Planes Asignados.CENTRO DE COSTO", "Planes Asignados.NOMBRE DEL PLAN ACTUAL", "Planes Asignados.CONVENIO  #1", "Planes Asignados.CONVENIO #2"}),
    #"Columnas con nombre cambiado" = Table.RenameColumns(#"Se expandió Planes Asignados1",{{"Planes Asignados.EMPLEADO DE COBRO", "EMPLEADO"}, {"Planes Asignados.DOCUMENTO DE COBRO","CEDULA"}}),
    #"Columnas con nombre cambiado1" = Table.RenameColumns(#"Columnas con nombre cambiado",{{"Planes Asignados.NOMBRE DEL PLAN ACTUAL", "PLANES ASIGNADOS"}, {"Planes Asignados.CONVENIO  #1", "COBRO FIJO #1"}, {"Planes Asignados.CONVENIO #2", "COBRO SERV. ESPECIALES #2"},{"CARGO FIJO","CARGO FIJO 1"}}),
    #"Columnas reordenadas2" = Table.ReorderColumns(#"Columnas con nombre cambiado1",{"CICLO", "MIN", "EMPLEADO", "CEDULA", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "IVA 19%", "DESCUENTO IVA", "CARGO FIJO 1", "DESCUENTO", "ESPECIALES", "IMPOCONSUMO"}),
    #"Texto recortado" = Table.TransformColumns(#"Columnas reordenadas2",{{"CEDULA", Text.Trim, type text}}),
    #"Combinado de Tabla de Ajuste" = Table.NestedJoin(#"Texto recortado", {"PLANES ASIGNADOS"}, #"Tabla de ajuste", {"PLAN A AJUSTAR"}, "Tabla7", JoinKind.LeftOuter),
    #"Se expandió Tabla de Ajuste" = Table.ExpandTableColumn(#"Combinado de Tabla de Ajuste", "Tabla7", {"PLAN A AJUSTAR", "VALOR", "CANTIDAD"}, {"Tabla7.PLAN A AJUSTAR", "Tabla7.VALOR", "Tabla7.CANTIDAD"}),
    #"Codicional para sumar Ajuste" = Table.AddColumn(#"Se expandió Tabla de Ajuste", "CARGO FIJO", each if [PLANES ASIGNADOS]<>null and [PLANES ASIGNADOS]=[Tabla7.PLAN A AJUSTAR] then ([Tabla7.VALOR]/[Tabla7.CANTIDAD])+[CARGO FIJO 1] else [CARGO FIJO 1]),
    #"Tipo cambiado1" = Table.TransformColumnTypes(#"Codicional para sumar Ajuste",{{"CARGO FIJO", type number}}),
    #"DESCUENTO REFRIDCOL" = Table.AddColumn(#"Tipo cambiado1", "DESC.REFRIDCOL", each if [CEDULA]="805005717" then [CARGO FIJO]+[DESCUENTO]+[ESPECIALES]+[#"IVA 19%"]+[DESCUENTO IVA]+[IMPOCONSUMO] else 

(1-[#"COBRO FIJO #1"])*[CARGO FIJO]+(1-[#"COBRO SERV. ESPECIALES #2"])*[ESPECIALES]+[#"IVA 19%"]+[DESCUENTO IVA]+[IMPOCONSUMO]+[DESCUENTO])

// El 1 significa el 100%.
,
    #"Combinación Establecim Act" = Table.NestedJoin(#"DESCUENTO REFRIDCOL", {"CEDULA"}, #"ESTABLECIMIENTO ACT", {"CEDULA"}, "ESTABLECIMIENTO", JoinKind.LeftOuter),
    #"Se expandió ESTABLECIMIENTO" = Table.ExpandTableColumn(#"Combinación Establecim Act", "ESTABLECIMIENTO", {"CEDULA", "EMPRESA", "CENTRO DE COSTO", "AREA"}, {"ESTABLECIMIENTO.CEDULA", "ESTABLECIMIENTO.EMPRESA", "ESTABLECIMIENTO.CENTRO DE COSTO", "ESTABLECIMIENTO.AREA"}),
    #"Duplicados quitados" = Table.Distinct(#"Se expandió ESTABLECIMIENTO", {"MIN"}),
    #"Combina CONTRATISTAS" = Table.NestedJoin(#"Duplicados quitados", {"MIN"}, CONTRATISTAS, {"LINEA TELEFONICA"}, "CONTRATISTAS", JoinKind.LeftOuter),
    #"Se expandió CONTRATISTAS" = Table.ExpandTableColumn(#"Combina CONTRATISTAS", "CONTRATISTAS", {"CEDULA", "CO"}, {"CONTRATISTAS.CEDULA", "CONTRATISTAS.CO"}),
    #"Consultas ESTABLECIMIENTO RET" = Table.NestedJoin(#"Se expandió CONTRATISTAS", {"CEDULA"}, #"ESTABLECIMIENTO RET", {"CEDULA"}, "ESTABLECIMIENTO RET", JoinKind.LeftOuter),
    #"Se expandió ESTABLECIMIENTO RET1" = Table.ExpandTableColumn(#"Consultas ESTABLECIMIENTO RET", "ESTABLECIMIENTO RET", {"CEDULA", "EMPRESA", "CENTRO DE COSTO", "AREA"}, {"ESTABLECIMIENTO RET.CEDULA", "ESTABLECIMIENTO RET.EMPRESA", "ESTABLECIMIENTO RET.CENTRO DE COSTO", "ESTABLECIMIENTO RET.AREA"}),
    #"Tipo cambiado2" = Table.TransformColumnTypes(#"Se expandió ESTABLECIMIENTO RET1",{{"ESTABLECIMIENTO RET.CEDULA", type text}}),
    #"Columnas con nombre cambiado2" = Table.RenameColumns(#"Tipo cambiado2",{{"CEDULA", "CEDULA1"}}),
    #"Columna condicional agregada2" = Table.AddColumn(#"Columnas con nombre cambiado2", "CEDULA", each if [CEDULA1] = null then [ESTABLECIMIENTO.CEDULA] else if [CEDULA1] = null then [ESTABLECIMIENTO RET.CEDULA] else if [CONTRATISTAS.CEDULA] <> null then [CONTRATISTAS.CEDULA] else [CEDULA1]),
    #"CONDICIONAL EMPRESA" = Table.AddColumn(#"Columna condicional agregada2", "EMPRESA", each if [CEDULA] = "805005717" then "REFRIDCOL" else if [CEDULA] = "900902500" then "CRUZTOR" else if [CEDULA] = [CONTRATISTAS.CEDULA] then "CONTRATISTA" else if [CEDULA] = [ESTABLECIMIENTO.CEDULA] then [ESTABLECIMIENTO.EMPRESA] else if [CEDULA] = [ESTABLECIMIENTO RET.CEDULA] then "RETIRADO" else "VALIDAR"),
    #"CONDICION DE AREA" = Table.AddColumn(#"CONDICIONAL EMPRESA", "AREA", each if [ESTABLECIMIENTO.AREA]<>"" then [ESTABLECIMIENTO.AREA] else if [ESTABLECIMIENTO RET.AREA]<>"" then [ESTABLECIMIENTO RET.AREA] else "NO APLICA"),
    #"Columna condicional agregada1" = Table.AddColumn(#"CONDICION DE AREA", "C.O", each if [EMPLEADO] = "RETIRADO" then [ESTABLECIMIENTO RET.CENTRO DE COSTO] else if [CEDULA] = "805005717" then "9910-99" else if [CEDULA] = "900902500" then "9910-99" else if [CEDULA] = "16206046" then [CONTRATISTAS.CO] else if [CEDULA] = "16501759" then [CONTRATISTAS.CO] else if [CEDULA] = "31977462" then [CONTRATISTAS.CO] else if [CEDULA] = null then "9910-99" else [ESTABLECIMIENTO.CENTRO DE COSTO]),
    #"Columnas quitadas" = Table.RemoveColumns(#"Columna condicional agregada1",{"CICLO", "ESTABLECIMIENTO.EMPRESA", "CONTRATISTAS.CEDULA", "CARGO FIJO 1", "Tabla7.PLAN A AJUSTAR", "Tabla7.VALOR", "Tabla7.CANTIDAD", "ESTABLECIMIENTO.CEDULA", "ESTABLECIMIENTO.AREA", "ESTABLECIMIENTO RET.CEDULA", "ESTABLECIMIENTO RET.EMPRESA", "ESTABLECIMIENTO RET.AREA"}),
    #"Tipo cambiado" = Table.TransformColumnTypes(#"Columnas quitadas",{ {"DESC.REFRIDCOL", type number}}),
    #"Columna en blanco 1" = Table.AddColumn(#"Tipo cambiado", "-", each null),
    #"Columna en blanco 2" = Table.AddColumn(#"Columna en blanco 1", "--", each null),
    #"Columna en blanco 3" = Table.AddColumn(#"Columna en blanco 2", "---", each null),
    #"Columnas con nombre cambiado5" = Table.RenameColumns(#"Columna en blanco 3",{{"MIN", "LINEAS TELEFONICAS"}, {"IMPOCONSUMO", "BASE IMPOCONSUMO"}}),
    #"Columnas reordenadas7" = Table.ReorderColumns(#"Columnas con nombre cambiado5",{"LINEAS TELEFONICAS", "CEDULA", "EMPLEADO", "AREA", "EMPRESA", "C.O", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "---", "CARGO FIJO", "ESPECIALES", "BASE IMPOCONSUMO", "DESCUENTO", "IVA 19%", "DESCUENTO IVA", "-",  "--", "DESC.REFRIDCOL"}),
    #"Calculo Impoconsumo" = Table.AddColumn(#"Columnas reordenadas7", "IMPOCONSUMO", each 0),
    #"Calculo Nuevo IVA" = Table.AddColumn(#"Calculo Impoconsumo", "IVA  19%", each [#"IVA 19%"]),
    #"Columnas reordenadas3" = Table.ReorderColumns(#"Calculo Nuevo IVA",{"LINEAS TELEFONICAS", "CEDULA", "EMPLEADO", "AREA", "EMPRESA", "C.O", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "---", "CARGO FIJO", "ESPECIALES", "IMPOCONSUMO", "BASE IMPOCONSUMO", "DESCUENTO", "IVA  19%", "IVA 19%", "DESCUENTO IVA", "-",  "--", "DESC.REFRIDCOL"}),
    #"Columnas quitadas3" = Table.RemoveColumns(#"Columnas reordenadas3",{ "IVA 19%"}),
    #"Columnas reordenadas4" = Table.ReorderColumns(#"Columnas quitadas3",{"LINEAS TELEFONICAS", "CEDULA", "EMPLEADO", "AREA", "EMPRESA", "C.O", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "---", "CARGO FIJO", "ESPECIALES", "BASE IMPOCONSUMO", "IMPOCONSUMO", "DESCUENTO", "IVA  19%", "DESCUENTO IVA", "-", "--", "DESC.REFRIDCOL"}),
    #"DESCUENTO EMPLEADO" = Table.AddColumn(#"Columnas reordenadas4", "DESC.EMPLEADO", each if [CEDULA]="805005717" then 0 else ([CARGO FIJO]*[#"COBRO FIJO #1"])+([ESPECIALES]*[#"COBRO SERV. ESPECIALES #2"])),
    #"1 Quincena" = Table.AddColumn(#"DESCUENTO EMPLEADO", "1QUINCENA", each [DESC.EMPLEADO]/2),
    #"2 Quincena" = Table.AddColumn(#"1 Quincena", "2QUINCENA", each [DESC.EMPLEADO]/2),
    #"Columnas reordenadas5" = Table.ReorderColumns(#"2 Quincena",{"LINEAS TELEFONICAS", "CEDULA", "EMPLEADO", "AREA", "EMPRESA", "C.O", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "---", "CARGO FIJO", "ESPECIALES", "BASE IMPOCONSUMO", "IMPOCONSUMO", "DESCUENTO", "IVA  19%", "DESCUENTO IVA", "-", "1QUINCENA", "2QUINCENA", "DESC.EMPLEADO", "--", "DESC.REFRIDCOL"}),
    #"Duplicados quitados1" = Table.Distinct(#"Columnas reordenadas5", {"LINEAS TELEFONICAS"}),
    #"Otras columnas quitadas1" = Table.SelectColumns(#"Duplicados quitados1",{"LINEAS TELEFONICAS", "CEDULA", "EMPLEADO", "AREA", "EMPRESA", "C.O", "PLANES ASIGNADOS", "COBRO FIJO #1", "COBRO SERV. ESPECIALES #2", "---", "CARGO FIJO", "ESPECIALES", "BASE IMPOCONSUMO", "IMPOCONSUMO", "DESCUENTO", "IVA  19%", "DESCUENTO IVA", "-", "1QUINCENA", "2QUINCENA", "DESC.EMPLEADO", "--", "DESC.REFRIDCOL"})
in
    #"Otras columnas quitadas1"