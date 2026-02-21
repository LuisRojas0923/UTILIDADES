-- Consulta SQL para Establecimiento, Contrato y Beneficio
-- Orden de columnas según especificación. Optimizada: CTE para primera fecha (evita subconsulta por fila).
-- Encabezados con espacio: AUX VIVIENDA , TALLA PANTALON

WITH primera_fecha AS (
    SELECT TRIM(CAST(establecimiento AS TEXT)) AS establecimiento_trim,
           MIN(fechainicio) AS primera_fecha_ingreso
    FROM contrato
    GROUP BY TRIM(CAST(establecimiento AS TEXT))    
)
SELECT DISTINCT
    CAST(NULLIF(TRIM(CAST(E.nrocedula AS TEXT)), '') AS NUMERIC) AS "CEDULA",
    E.nombre::text AS "NOMBRE",
    CASE
        WHEN B.estado = 'Activo' THEN 'A'
        WHEN B.estado = 'Inactivo' THEN 'RET'
        WHEN B.estado IS NOT NULL AND B.estado <> '' THEN B.estado::text
        WHEN C.estado = 'Activo' THEN 'A'
        WHEN C.estado = 'Terminado' THEN 'RET'
        ELSE 'RET'
    END::text AS "ESTADO",
    C.empresa::text AS "EMPRESA",
    C.area::text AS "AREA",
    C.cargo::text AS "CARGO ESTABLECIMIENTO",
    C.gerencia::text AS "GERENCIA",
    C.jefe::text AS "JEFE INMEDIATO",
    C.regional::text AS "REGIONAL QUE REPORTA",
    C.centrocosto::text AS "CENTRO DE COSTO",
    CAST(NULLIF(TRIM(CAST(C.cuenta AS TEXT)), '') AS NUMERIC) AS "CUENTA",
    CAST(NULLIF(TRIM(CAST(C.tipocontable AS TEXT)), '') AS NUMERIC) AS "TIPO",
    C.nomina::text AS "NOMINA",
    CAST(NULLIF(TRIM(CAST(C.riesgoarl AS TEXT)), '') AS DECIMAL) AS "RIESGO ARL",
    C.arl::text AS "A.R.L",
    E.correocorporativo::text AS "CORREO CORPORTIVO",
    CASE WHEN E.viaticante = true THEN 'SI' ELSE 'NO' END::text AS "VIATICANTE",
    COALESCE(E.baseviaticos, 0)::numeric AS "BASE VIATICOS",
    TRIM(CAST(C.numerocontrato AS TEXT))::text AS "NUMERO CONTRATO",
    C.estado::text AS "ESTADO CONTRATO",
    C.tipo::text AS "TIPO DE CONTRATO",
    C.ciudadcontratacion::text AS "CIUDAD CONTRATACION",
    PF.primera_fecha_ingreso::date AS "1A FECHA INGRESO",
    C.fechainicio::date AS "FECHA ULTIMO INGRESO",
    CASE
        WHEN C.fechainicioetapaproductiva = '1900-01-01'::date THEN NULL
        ELSE C.fechainicioetapaproductiva
    END::date AS "FECHA INICIO ETAPA PRODUCTIVA",
    C.fechavencimiento::date AS "FECHA VENCIMIENTO CONTRATO",
    CASE
        WHEN C.fecharetiro IN ('1900-01-01'::date, '1900-01-02'::date, '1990-01-01'::date) THEN NULL
        ELSE C.fecharetiro
    END::date AS "FECHA RETIRO",
    CASE
        WHEN C.estado = 'Activo' AND UPPER(TRIM(COALESCE(C.empresa, '')::text)) = 'REFRIDCOL'
        THEN (C.fechavencimiento - INTERVAL '35 days')::date
        ELSE NULL
    END::date AS "CARTA DE PRORROGA",
    C.observaciones::text AS "OBSERVACIONES CONTRATO",
    TRIM(CAST(B.contrato AS TEXT))::text AS "NUMERO BENEFICIO",
    B.estado::text AS "ESTADO BENEFICIO",
    CAST(NULLIF(TRIM(CAST(B.salario AS TEXT)), '') AS NUMERIC) AS "SALARIO AÑO 2025",
    B.auxiliolegaltransporte::text AS "AUXILIO LEGAL TRANSPORTE",
    CAST(NULLIF(TRIM(CAST(B.auxilioalimentacionmensual AS TEXT)), '') AS NUMERIC) AS "AUX ALIMENTACION MENSUAL",
    CAST(NULLIF(TRIM(CAST(B.auxilioalimentacionquincenal AS TEXT)), '') AS NUMERIC) AS "AUX ALIMENTACION (QUINCENAL)",
    CAST(NULLIF(TRIM(CAST(B.auxiliovivienda AS TEXT)), '') AS NUMERIC) AS "AUX VIVIENDA ",
    CAST(NULLIF(TRIM(CAST(B.rodamiento AS TEXT)), '') AS NUMERIC) AS "RODAMIENTO",
    CAST(NULLIF(TRIM(CAST(B.baserodamiento AS TEXT)), '') AS NUMERIC) AS "BASE RODAMIENTO",
    CAST(NULLIF(TRIM(CAST(B.valormaximobono AS TEXT)), '') AS NUMERIC) AS "VALOR MAXIMO BONO 40%",
    CAST(NULLIF(TRIM(CAST(B.capacidadendeudamiento AS TEXT)), '') AS NUMERIC) AS "CAPACIDAD ENDEUDAMIENTO",
    CASE WHEN B.autorizacionrodamiento THEN 'SI' ELSE 'NO' END::text AS "AUTORIZAN RODAMIENTO",
    CASE WHEN B.autorizacionhorasextras THEN 'SI' ELSE 'NO' END::text AS "AUTORIZAN HORAS EXTRAS",
    B.observaciones::text AS "OBSERVACIONES BENEFICIO",
    C.banco::text AS "BANCO",
    CAST(NULLIF(TRIM(CAST(C.cuentanomina AS TEXT)), '') AS NUMERIC) AS "CTA DE NOMINA",
    C.eps::text AS "E.P.S",
    C.afp::text AS "A.F.P",
    C.ccf::text AS "C.C.F",
    C.afc::text AS "A.F.C",
    E.examenmedico::text AS "EXAMEN DE INGRESO",
    E.fechavencimientoexamenmedico::date AS "FECHA VENCIMIENTO EXAMEN DE INGRESO",
    E.certificadoalturas::text AS "CERTIFICADO TRABAJO EN ALTURAS",
    E.fechavencimientoalturas::date AS "FECHA VENCIMIENTO TRABAJO EN ALTURAS",
    E.certificadoprimerosauxilios::text AS "CERTIFICADO PRIMEROS AUXILIOS",
    E.fechavencimientoprimerosauxilios::date AS "FECHA VENCIMIENTO PRIMEROS AUXILIOS",
    C.archivo::text AS "ARCHIVO CONTRATO",
    E.fechanacimiento::date AS "FECHA NACIMIENTO",
    TRIM(TO_CHAR(E.fechanacimiento, 'Mon'))::text AS "MES CUMPLEAÑOS",
    E.sexo::text AS "SEXO",
    EXTRACT(YEAR FROM AGE(E.fechanacimiento::date))::numeric AS "EDAD",
    E.rh::text AS "RH",
    CAST(NULLIF(TRIM(CAST(E.nrohijos AS TEXT)), '') AS NUMERIC) AS "CUANTOS HIJOS",
    E.correopersonal::text AS "CORREO PERSONAL",
    E.gradoalcanzado::text AS "GRADO ALCANZADO",
    E.tituloobtenido::text AS "TITULO OBTENIDO",
    E.tarjetaprofesional::text AS "TARJETA PROFESIONAL",
    E.ciudadresidencia::text AS "CUIDAD DE RESIDENCIA",
    E.direccionresidencia::text AS "DIRECCION RESIDENCIA",
    E.barrio::text AS "BARRIO",
    E.telefono::text AS "TELEFONO",
    E.tallacamisa::text AS "TALLA CAMISA",
    E.tallapantalon::text AS "TALLA PANTALON ",
    E.tallabotas::text AS "TALLA BOTAS",
    E.contactoemergencia::text AS "CONTACTO EMERGENCIA",
    E.parentesco::text AS "PARENTESCO",
    E.telefonocontactoemergencia::text AS "TELEFONO CONTACTO EMERGENCIA",
    E.archivo1::text AS "ARCHIVO 1",
    E.archivo2::text AS "ARCHIVO 2",
    E.archivo3::text AS "ARCHIVO 3",
    E.archivo4::text AS "ARCHIVO 4",
    E.codigo AS "CODIGO",
    E.fecha AS "FECHA DE REGISTRO",
    date_trunc('second', E.hora) AS "HORA DE REGISTRO"
FROM establecimiento E
LEFT JOIN primera_fecha PF ON PF.establecimiento_trim = TRIM(CAST(E.nrocedula AS TEXT))
LEFT JOIN contrato C ON TRIM(CAST(C.establecimiento AS TEXT)) = TRIM(CAST(E.nrocedula AS TEXT))
LEFT JOIN beneficio B ON TRIM(CAST(B.contrato AS TEXT)) = TRIM(CAST(C.numerocontrato AS TEXT))
ORDER BY
    CASE
        WHEN B.estado = 'Activo' THEN 'A'
        WHEN B.estado = 'Inactivo' THEN 'RET'
        WHEN B.estado IS NOT NULL AND B.estado <> '' THEN B.estado::text
        WHEN C.estado = 'Activo' THEN 'A'
        WHEN C.estado = 'Terminado' THEN 'RET'
        ELSE 'RET'
    END ASC,
    E.nombre ASC,
    C.fechainicio DESC NULLS LAST;
