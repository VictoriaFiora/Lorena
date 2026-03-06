# Proyecto Lorena 🌸

Repositorio para la automatización de la limpieza y auditoría de datos geográficos para el proyecto Maestro Localidades.

## Objetivo
Migrar la lógica de Google Apps Script a Python para ganar en rendimiento, versatilidad y tener un control de versiones robusto con Git.

## Arquitectura de Trabajo
El proceso se divide en dos fases principales:
1. **Completado de Localización**: Enriquecer la `Tabla Sucursales Localizacion` con datos precisos de Georef AR (CP, Departamento, Coordenadas).
2. **Generación del Maestro**: Consolidar la información validada en la tabla `Maestro Localidades - Corregido`.

## Uso
1. Instalar dependencias: `pip install -r requirements.txt`
2. Configurar credenciales de Google Sheets en `creds.json`.
3. Ejecutar el script: `python lorena_audit.py`

## Referencia Legacy
Los scripts originales de Google Apps Script se encuentran en la raíz para consulta lógica durante la migración.
