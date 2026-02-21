# Chat IA ERP - Módulo de Consultas Inteligentes para SOLID

Módulo híbrido Java-Python que permite realizar consultas inteligentes a la base de datos del ERP SOLID usando GPT-4o-mini.

## Características

- Consultas en lenguaje natural
- Generación automática de SQL
- Análisis inteligente de resultados
- Integración nativa con SOLID (Java)
- Seguridad: solo consultas SELECT, whitelist de tablas

## Inicio Rápido

1. **Instalar dependencias Python**:
```powershell
cd python
.\instalar_entorno.ps1
```

2. **Configurar API Key**:
   - El archivo `.env` ya está creado con la API key
   - Si necesitas cambiarla, edita `python/.env`

3. **Compilar Java**:
```bash
cd java
javac ChatIARunner.java
```

4. **Probar**:
```bash
java ChatIARunner "¿Cuántas legalizaciones hay este mes?"
```

## Estructura

- `java/`: Orquestador Java (integración con SOLID)
- `python/`: Motor de IA (generación SQL y análisis)
- `docs/`: Documentación completa

## Documentación

Ver [GUIA_CHAT_IA.md](docs/GUIA_CHAT_IA.md) para documentación completa.

## Tablas Disponibles

- `legalizacion`: Legalizaciones de gastos
- `linealegalizacion`: Detalle de gastos por legalización
- `consignacion`: Consignaciones a empleados

## Seguridad

- Solo consultas SELECT permitidas
- Whitelist de tablas
- Timeout de 30 segundos
- Límite de 10,000 filas

## Licencia

Uso interno - Equipo de Mejoramiento Continuo

