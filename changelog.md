# Changelog

Todas las notas importantes sobre los cambios en este proyecto se documentarán en este archivo.

## [Unreleased]
### Añadido
- Interfaz español/inglés con detección del idioma del navegador, selector manual y modo oscuro.
- Consulta cacheada por 15 minutos de la tasa USD del BCV y su `Fecha Valor`, con actualización explícita, ajuste manual y aviso si no hay conexión.
- Etiquetas bilingües y procedencia de la tasa en el Excel; F8 sigue siendo numérica y conserva las fórmulas de conversión.

### Cambiado
- Cabecera compacta con idioma y modo oscuro; navegación de cuatro pasos con etiquetas breves.
- Selección entre comprobantes y gasto manual para mostrar un formulario a la vez; referencia y descripción se agrupan como detalles opcionales.
- Datos del reporte, empresa opcional y tasa con procedencia reunidos en Exportar, donde se prepara la descarga del Excel.
- Textos de revisión y acciones más breves; vaciar los gastos requiere una confirmación explícita en la sesión.

### Corregido
- Priorizar los importes que coinciden con etiquetas específicas como `TOTAL` o `MONTO`; las comisiones encontradas por la búsqueda genérica de importes en Bs ya no desplazan el total.
- Declarar el modelo de idioma español de Tesseract en las dependencias del despliegue, acorde con la solicitud OCR `spa+eng`.
- Aceptar coma o punto decimal en la tabla de revisión y en el formulario de gasto manual, y bloquear la exportación mientras haya importes inválidos.
- Conservar hasta 8 decimales en la tasa, distinguir un dato BCV anterior si falla la actualización y señalar la procedencia al exportar.
- Mantener la fecha de emisión como `DD/MM/YYYY` en los dos idiomas y reintentar OCR en inglés cuando falte el modelo español de Tesseract.
- Retirar del historial público la plantilla y las capturas internas de versiones anteriores; la demo conserva solo recursos sintéticos.

### Pendiente
- Agregar mejoras en los patrones de busqueda
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
