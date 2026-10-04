# Changelog

Todas las notas importantes sobre los cambios en este proyecto se documentarán en este archivo.

## [Unreleased]
### Corregido
- Priorizar los importes que coinciden con etiquetas específicas como `TOTAL` o `MONTO`; las comisiones encontradas por la búsqueda genérica de importes en Bs ya no desplazan el total.
- Declarar el modelo de idioma español de Tesseract en las dependencias del despliegue, acorde con la solicitud OCR `spa+eng`.

### Pendiente
- Agregar mejoras en los patrones de busqueda
- Scrapear tasa del dolar 
- Agregar bot que manda a tu whatsapp el archivo

## [2.0.0] - 2026-09-03
### Cambiado
- Nuevo front "Caja Chica Pro": hero, sidebar por secciones, 4 pestañas (Cargar / Revisar / Resumen / Generar), tabla editable, métricas y gráfico por proveedor.
- Plantilla Excel 100% propia generada por código (`excel_builder.py`): sin depender de `base/caja_chica_base.xlsx`, empresa configurable, fórmulas vivas Bs/$ con tasa editable, totales, saldo, filtros, firmas e impresión apaisada.
- OCR robusto en `ocr_utils.py` con múltiples patrones y alta manual; `requirements.txt` limpio y tema actualizado.

## [1.0.1] - 2025-03-18
### Añadido
- Colocar en mayusculas las primeras letra de una palabra y las demas en minusculas

## [1.0.0] - 2025-01-24
### Añadido
- Inicialización del proyecto con la primera versión completa.
- Documentación inicial como el readme y ejemplos.
- Documentacion para contribuir

## [0.1.0] - 2 Meses Antes
### Añadido
- Funcionalidades principales básicas.
- Configuración inicial del proyecto.
