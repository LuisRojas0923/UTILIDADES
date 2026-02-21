# Tabla comparativa: alias Establecimiento

| # | Lo que diste | Campo real (tabla.columna) | Nombre final |
|---|---------------------------|-----------------------------|--------------|
| 1 | Codigo | establecimiento.codigo | CODIGO |
| 2 | Fecha de Registro | establecimiento.fecha | FECHA DE REGISTRO |
| 3 | Hora de Registro | establecimiento.hora | HORA DE REGISTRO |
| 4 | Numero de Cedula | establecimiento.nrocedula | NUMERO DE CEDULA |
| 5 | Nombre | establecimiento.nombre | NOMBRE |
| 6 | Fecha de Nacimiento | establecimiento.fechanacimiento | FECHA DE NACIMIENTO |
| 7 | Edad | EXTRACT(AGE(fechanacimiento)) | EDAD |
| 8 | Primera Fecha Ingreso | establecimiento.primerafechaingreso | PRIMERA FECHA INGRESO |
| 9 | Cargo | establecimiento.cargo | CARGO |
| 10 | Empresa | establecimiento.empresa | EMPRESA |
| 11 | Estado | contrato.estado *(en V2 se usa C.estado como Estado principal)* | ESTADO |
| 12 | Centro de Costo | establecimiento.centrocosto | CENTRO DE COSTO |
| 13 | Cuenta | establecimiento.cuenta | CUENTA |
| 14 | Tipo | establecimiento.tipo | TIPO |
| 15 | Nomina | establecimiento.nomina | NOMINA |
| 16 | Regional | establecimiento.regional | REGIONAL |
| 17 | Area | establecimiento.area | AREA |
| 18 | Jefe | establecimiento.jefe | JEFE |
| 19 | Gerencia | establecimiento.gerencia | GERENCIA |
| 20 | Ciudad de Contratacion | establecimiento.ciudadcontratacion | CIUDAD DE CONTRATACION |
| 21 | Grado Alcanzado | establecimiento.gradoalcanzado | GRADO ALCANZADO |
| 22 | Titulo Obtenido | establecimiento.tituloobtenido | TITULO OBTENIDO |
| 23 | Tarjeta Profesional | establecimiento.tarjetaprofesional | TARJETA PROFESIONAL |
| 24 | RH | establecimiento.rh | RH |
| 25 | Cuenta de Nomina | establecimiento.cuentanomina | CUENTA DE NOMINA |
| 26 | Banco | establecimiento.banco | BANCO |
| 27 | Riesgo ARL | establecimiento.riesgoarl | RIESGO ARL |
| 28 | Sexo | establecimiento.sexo | SEXO |
| 29 | Numero de Hijos | establecimiento.nrohijos | NUMERO DE HIJOS |
| 30 | Correo Corporativo | establecimiento.correocorporativo | CORREO CORPORATIVO |
| 31 | Correo Personal | establecimiento.correopersonal | CORREO PERSONAL |
| 32 | ARL | establecimiento.arl | ARL |
| 33 | EPS | establecimiento.eps | EPS |
| 34 | AFP | establecimiento.afp | AFP |
| 35 | CCF | establecimiento.ccf | CCF |
| 36 | AFC | establecimiento.afc | AFC |
| 37 | Viaticante | establecimiento.viaticante | VIATICANTE |
| 38 | Base Viaticos | establecimiento.baseviaticos | BASE VIATICOS |
| 39 | Talla Camisa | establecimiento.tallacamisa | TALLA CAMISA |
| 40 | Talla Pantalon | establecimiento.tallapantalon | TALLA PANTALON |
| 41 | Talla Botas | establecimiento.tallabotas | TALLA BOTAS |
| 42 | Examen de Ingreso | establecimiento.examenmedico | EXAMEN DE INGRESO |
| 43 | Fecha Vencimiento Examen de Ingreso | establecimiento.fechavencimientoexamenmedico | FECHA VENCIMIENTO EXAMEN DE INGRESO |
| 44 | Certificado Trabajo en Alturas | establecimiento.certificadoalturas | CERTIFICADO TRABAJO EN ALTURAS |
| 45 | Fecha Vencimiento Trabajo en Alturas | establecimiento.fechavencimientoalturas | FECHA VENCIMIENTO TRABAJO EN ALTURAS |
| 46 | Certificado Primeros Auxilios | establecimiento.certificadoprimerosauxilios | CERTIFICADO PRIMEROS AUXILIOS |
| 47 | Fecha Vencimiento Primeros Auxilios | establecimiento.fechavencimientoprimerosauxilios | FECHA VENCIMIENTO PRIMEROS AUXILIOS |
| 48 | C.numeroContrato | contrato.numerocontrato | NUMERO CONTRATO |
| 49 | C.tipo | contrato.tipo | TIPO CONTRATO |
| 50 | C.fechaInicio | contrato.fechainicio | FECHA INICIO CONTRATO |
| 51 | C.fechaInicioEtapaProductiva | contrato.fechainicioetapaproductiva | FECHA INICIO ETAPA PRODUCTIVA |
| 52 | C.fechaVencimiento | contrato.fechavencimiento | FECHA VENCIMIENTO |
| 53 | C.fechaRetiro | contrato.fecharetiro | FECHA RETIRO |
| 54 | C.duracionMeses | contrato.duracionmeses | DURACION MESES |
| 55 | C.estado | contrato.estado | ESTADO CONTRATO |
| 56 | C.causaTerminacion | contrato.causaterminacion | CAUSA TERMINACION |
| 57 | C.pazYSalvo | contrato.pazysalvo | PAZ Y SALVO |
| 58 | B.fechaInicio | beneficio.fechainicio | FECHA INICIO BENEFICIO |
| 59 | B.fechaFin | beneficio.fechafin | FECHA FIN |
| 60 | B.salario | beneficio.salario | SALARIO |
| 61 | B.moneda | beneficio.moneda | MONEDA |
| 62 | B.valorMaximoBono | beneficio.valormaximobono | VALOR MAXIMO BONO |
| 63 | B.capacidadEndeudamiento | beneficio.capacidadendeudamiento | CAPACIDAD ENDEUDAMIENTO |
| 64 | B.autorizacionHorasExtras | beneficio.autorizacionhorasextras | AUTORIZACION HORAS EXTRAS |
| 65 | B.auxilioAlimentacionQuincenal | beneficio.auxilioalimentacionquincenal | AUXILIO ALIMENTACION QUINCENAL |
| 66 | B.auxilioAlimentacionMensual | beneficio.auxilioalimentacionmensual | AUXILIO ALIMENTACION MENSUAL |
| 67 | B.baseRodamiento | beneficio.baserodamiento | BASE RODAMIENTO |
| 68 | B.autorizacionRodamiento | beneficio.autorizacionrodamiento | AUTORIZACION RODAMIENTO |
| 69 | B.rodamiento | beneficio.rodamiento | RODAMIENTO |
| 70 | B.auxilioVivienda | beneficio.auxiliovivienda | AUXILIO VIVIENDA |
| 71 | B.periodicidadPago | beneficio.periodicidadpago | PERIODICIDAD PAGO |
| 72 | B.estado | beneficio.estado | ESTADO BENEFICIO |

---

**Tablas:** `establecimiento` (E), `contrato` (C), `beneficio` (B).  
Nombres de columnas en PostgreSQL suelen estar en minúsculas; en el SQL V2 se usan camelCase (ej. `nroCedula`). Aquí se listan en minúsculas como en el esquema típico de la BD.
