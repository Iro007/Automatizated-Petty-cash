# Automatizated Petty Cash

Aplicación Python/Streamlit para preparar una relación de gastos de caja chica. La interfaz está disponible en español e inglés, detecta el idioma del navegador y ofrece modo claro u oscuro. Incluye navegación por pasos, una guía visual de carga y un resumen con los gastos y el total real de la sesión. Los controles de idioma y apariencia están en la cabecera; las cuatro pestañas muestran cada paso del trabajo. Carga imágenes, revisa los campos extraídos mediante OCR o ingresa gastos manualmente y genera una planilla Excel propia por código.

## Flujo

1. En **Cargar**, elige comprobantes PNG, JPG, JPEG o WEBP para procesar con OCR, o selecciona gasto manual. Cada método muestra sus propios controles; referencia y descripción son opcionales.
2. En **Revisar**, corrige proveedor, fecha, referencia, importe y descripción. Las fechas deben existir y usar `dd/mm/aaaa`.
3. Consulta los totales y el gráfico por proveedor en **Resumen**.
4. En **Exportar**, completa responsable, cédula, monto otorgado y fecha de emisión. En esa pestaña también están la tasa y su procedencia, y los datos opcionales de empresa. Al generar el archivo, la descarga reemplaza el botón de generación; si editas los datos, debes generar una nueva versión. El archivo incluye conversiones, totales, saldo y espacios para firmas.
5. Revisa el archivo antes de utilizarlo. Cambiar la tasa numérica en **F8** modifica las fórmulas de conversión a USD.

La consulta lee el valor USD y la fecha `Fecha Valor` de la [página oficial del BCV](https://www.bcv.org.ve/estadisticas/tipo-cambio-de-referencia-smc), se cachea durante 15 minutos y se puede actualizar desde **Exportar**, en **Tasa y procedencia**. La fecha mostrada es la publicada por el BCV, incluso cuando cae en el siguiente día hábil. Si la consulta falla, la app avisa y permite ingresar una tasa manual; una tasa anterior no se presenta como cotización vigente. Si se modifica la tasa BCV, la app y el Excel la identifican como ajuste manual y conservan la referencia BCV y su fecha.

La planilla se construye con `openpyxl`, sin depender de una plantilla externa. F8 se mantiene numérica y editable; la fila 9 registra la procedencia, la tasa y su `Fecha Valor` cuando existen. Los metadatos describen la exportación: si cambias F8 directamente en Excel, actualiza también la fila 9. Los campos de texto introducidos por el usuario se guardan como texto literal en las celdas, preservando las fórmulas propias de cálculo. En español se usa `1.250,50` y en inglés `1,250.50`; la fecha se conserva como `DD/MM/YYYY` en ambos idiomas. Comprueba siempre los importes antes de exportar.

## Demo con datos ficticios

Estas imágenes proceden de una ejecución local. No contienen comprobantes reales, clientes ni datos personales reales.

### Comprobante sintético

![Comprobante ficticio sin validez fiscal](assets/demo/comprobante_ficticio.png)

### Revisión y corrección en la app

El comprobante ficticio contiene un total de 749,50 Bs y una comisión incluida de 25,00 Bs. La tabla permite revisar los campos extraídos y corregirlos antes de exportar; el OCR puede requerir supervisión según el formato y la calidad de cada imagen.

La captura siguiente muestra la interfaz actual con un gasto manual ficticio.

![Revisión de un gasto ficticio en la interfaz minimalista](assets/demo/revision_y_correccion.png)

### Excel generado

Con tres gastos de 1.250,50, 749,50 y 500 Bs, el total fue 2.500 Bs. A una tasa ficticia de 50 Bs/USD, el gasto fue 50 USD y el saldo 50 USD sobre un fondo de 100 USD.

![Vista renderizada del Excel generado, no captura de Microsoft Excel Desktop](assets/demo/vista_renderizada_excel.png)

## Ejecutar localmente en Windows

Desde la raíz del proyecto, usa un entorno virtual. Si `.venv` ya existe y funciona, reutilízalo.

```powershell
# Solo si aún no existe el entorno
python -m venv .venv

# Dependencias de la aplicación
.\.venv\Scripts\python.exe -m pip install -r requirements.txt

# Iniciar la app; abrir la dirección que muestre Streamlit
.\.venv\Scripts\python.exe -m streamlit run app.py
```

Estas son instrucciones de preparación; el entorno probado fue Python 3.14.6, Streamlit 1.63.0, pandas 3.0.5, openpyxl 3.1.5, Pillow 12.3.0 y pytesseract 0.3.13.

### Motor OCR

`pytesseract` es un adaptador Python: también necesita el ejecutable de **Tesseract** y sus modelos de idioma. Consulta la [documentación de instalación de Tesseract](https://tesseract-ocr.github.io/tessdoc/Installation.html) y, para Windows, los [instaladores de UB Mannheim](https://github.com/UB-Mannheim/tesseract/wiki).

La función OCR intenta `spa+eng` y, si falta el modelo español, reintenta con `eng`. Para obtener mejores resultados en comprobantes en español, instala ambos modelos y confirma su disponibilidad con `tesseract --list-langs`. `packages.txt` declara el motor y los datos de idioma español para Streamlit Community Cloud.

Si el ejecutable no está en PATH, añade su carpeta **solo a la sesión de PowerShell que ejecuta la app**. Este ejemplo corresponde a la instalación local utilizada; adapta la carpeta si instalaste Tesseract en otro lugar.

```powershell
$tesseractDir = Join-Path $env:LOCALAPPDATA 'Programs\Tesseract-OCR'
$env:PATH = "$tesseractDir;$env:PATH"
tesseract --version
tesseract --list-langs
.\.venv\Scripts\python.exe -m streamlit run app.py
```

En Linux, instala el motor y los modelos mediante el gestor de paquetes de tu distribución. `packages.txt` declara `tesseract-ocr`; los modelos de idioma deben estar disponibles además del motor.

Si el OCR falla o no encuentra datos, la app ofrece una fila para completar manualmente. No es necesario usar OCR para ingresar gastos manuales.

## Pruebas locales

Con las dependencias Python instaladas:

```powershell
.\.venv\Scripts\python.exe tests\test_demo.py
.\.venv\Scripts\python.exe tests\test_regressions.py
.\.venv\Scripts\python.exe tests\test_i18n_bcv.py
```

Las pruebas guardan resultados y ejemplos sintéticos en `evidence/demo_20261003/`, carpeta ignorada por Git. Incluyen prioridad del total etiquetado frente a una comisión, entrada de importes con coma o punto, validación de fechas, selección de idioma, formatos numéricos, lectura del HTML oficial del BCV, generación del Excel bilingüe, metadatos de procedencia y referencias de fórmula. Las pruebas de estructuras de 0, 1, 30 y 1.000 filas no establecen una capacidad máxima ni evalúan rendimiento.

Además se recorrió la demo en Chromium con imágenes sintéticas y se comprobó el recálculo de fórmulas en un motor independiente. Microsoft Excel Desktop no se ejecutó en esta verificación. La vista del Excel mostrada arriba es renderizada; el archivo generado solicita recálculo al abrirlo y no contiene resultados calculados en caché por `openpyxl`.

## Límites

- El OCR y las reglas de extracción pueden confundir un total con otro importe. Revisa siempre los campos; no se midieron precisión ni ahorro de tiempo.
- La tasa BCV se consulta en línea; una interrupción de la página o de la conexión requiere una tasa manual. Revisa su `Fecha Valor` y el importe aplicado antes de exportar.
- La fecha del gasto debe revisarse contra el comprobante. Una fecha válida puede ser incorrecta para el gasto.
- Las filas con importe cero muestran una advertencia, pero permiten exportar. La interfaz impone mínimos numéricos; el generador invocado directamente todavía acepta importes y tasas negativos.
- Guardar texto literal evita que esos campos de la hoja se interpreten como fórmulas. No constituye una evaluación de seguridad general.
- No se verificaron autenticación, permisos, uso multiusuario, concurrencia, escalabilidad ni tratamiento de datos no confiables en otros contextos.

## Archivos y estado de publicación

`app.py` contiene la interfaz, `ocr_utils.py` la extracción y `excel_builder.py` el generador de la hoja. Las carpetas de comprobantes, exportaciones, evidencia y el entorno virtual están ignoradas por Git. No uses comprobantes reales para las capturas de publicación.

El [repositorio del proyecto](https://github.com/Iro007/Automatizated-Petty-cash) contiene el código fuente y las tres imágenes sintéticas de esta demo. La plantilla y las capturas internas de versiones anteriores se retiraron del historial público; la demo actual usa únicamente datos ficticios.

La demo también está disponible en [Streamlit](https://petty-cash-automatizated.streamlit.app/). Para esta versión se comprobó que la interfaz publicada carga desde la rama `main`; usa únicamente datos ficticios al probarla.


