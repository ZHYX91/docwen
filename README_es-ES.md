[English](README.md) | [简体中文](README_zh-CN.md) | [繁體中文](README_zh-TW.md) | [Deutsch](README_de-DE.md) | [Français](README_fr-FR.md) | [Русский](README_ru-RU.md) | [Português](README_pt-BR.md) | [日本語](README_ja-JP.md) | [한국어](README_ko-KR.md) | [Español](README_es-ES.md) | [Tiếng Việt](README_vi-VN.md)

# DocWen

Herramienta de conversión de formatos de documentos y tablas: conversión bidireccional Word/Markdown/Excel. Se ejecuta completamente en local, garantizando seguridad y fiabilidad de los datos.

## 📖 Contexto del proyecto

Este software se diseñó originalmente para resolver problemas comunes en entornos de oficina:
- Los documentos llegan con formatos inconsistentes y hay que normalizarlos.
- Hay muchos tipos de archivo y cada uno tiene requisitos de formato distintos.
- Debe funcionar sin conexión (intranet/equipos antiguos).

**Filosofía de diseño**: herramienta ligera y “lista para usar”, con coste de aprendizaje muy bajo. No pretende sustituir a herramientas profesionales como LaTeX o Pandoc.

## ✨ Funciones principales

- **📄 Conversión de documentos** - Word ↔ Markdown, con conversión de fórmulas y mapeo de separadores (---/***/___) a saltos de página/sección/líneas horizontales. DOCX/DOC/WPS/RTF/ODT.
- **📊 Conversión de hojas de cálculo** - Excel ↔ Markdown. XLSX/XLS/ET/ODS/CSV. Incluye herramientas de resumen de tablas.
- **📑 PDF y archivos de maquetación** - PDF/XPS/OFD → Markdown o DOCX. Soporta unir/dividir PDF.
- **🖼️ Imágenes** - Conversión y compresión JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **🔍 Reconocimiento de texto OCR** - RapidOCR integrado para extraer texto de imágenes y PDF.
- **✏️ Revisión** - Reglas personalizables para símbolos, errores tipográficos y palabras sensibles.
- **📝 Plantillas** - Sistema flexible para formatos de documentos e informes.
- **💻 GUI + CLI** - Interfaz gráfica y línea de comandos.
- **🔒 Funcionamiento completamente local** - Se ejecuta sin conexión y cuenta con mecanismos integrados de aislamiento de red para garantizar la seguridad.
- **🔗 Ejecución de instancia única** - Gestiona automáticamente las instancias del programa y admite integración con el plugin complementario de Obsidian.

## 📸 Capturas de pantalla

| Lote | Markdown |
| --- | --- |
| ![Panel de lote](assets/screenshots/batch.png) | ![Panel de Markdown](assets/screenshots/markdown.png) |

| Documento | Hoja de cálculo |
| --- | --- |
| ![Panel de documento](assets/screenshots/document.png) | ![Panel de hoja de cálculo](assets/screenshots/spreadsheet.png) |

| Imagen | Archivos de maquetación |
| --- | --- |
| ![Panel de imagen](assets/screenshots/image.png) | ![Panel de maquetación](assets/screenshots/layout.png) |

Registro de cambios: ver [doc/CHANGELOG.md](doc/CHANGELOG.md)

## 🚀 Inicio rápido

### Iniciar el programa

En la versión empaquetada de Windows: haz doble clic en `DocWen.exe` para abrir la interfaz gráfica. Si instalas desde el código fuente / pip: ejecuta `docwen-gui`.

### Guía de inicio rápido

1.  **Prepara un archivo Markdown**:

    ```markdown
    ---
    title: Test Document
    ---
    
    ## Test Title
    
    This is the test body content.
    ```

2.  **Conversión con arrastrar y soltar**:
    - Inicia el programa.
    - Arrastra el archivo `.md` a la ventana.
    - Selecciona una plantilla.
    - Haz clic en "Convert to DOCX".

3.  **Obtén el resultado**:
    - Se generará un documento Word estandarizado en el mismo directorio.

**Consejo**: Puedes usar los archivos de ejemplo del directorio `samples/` para probar rápidamente las funciones.

## 📝 Convenciones de Markdown

### Mapeo de niveles de encabezado

Para facilitar el uso a compañeros sin conocimientos técnicos, los encabezados de Markdown se corresponden **uno a uno** con los encabezados de Word:
- El título y subtítulo del documento se colocan en los metadatos YAML.
- Markdown `# Heading 1` corresponde a Word "Heading 1".
- Markdown `## Heading 2` corresponde a Word "Heading 2".
- Y así sucesivamente, hasta 9 niveles.

**Consejo**: Si prefieres usar `#` como título del documento y empezar los encabezados del cuerpo con `##`, puedes ajustar el estilo "Heading 1" en la plantilla Word para que parezca un título (p. ej., centrado, negrita, tamaño mayor) y elegir en la configuración un esquema de numeración que omita el nivel 1.

### Saltos de línea y párrafos

**Regla básica**: Cada línea no vacía se trata como un párrafo independiente por defecto.

**Párrafos mixtos**: Cuando un subencabezado debe mezclarse con el texto del cuerpo en el mismo párrafo, deben cumplirse estas condiciones:
1.  El subencabezado termina con un signo de puntuación final (admite puntuación multilingüe).
2.  El texto del cuerpo está en la **línea inmediatamente siguiente**.
3.  La línea del cuerpo no puede ser un elemento especial de Markdown (encabezados, bloques de código, tablas, listas, citas, bloques de fórmula, separadores, etc.).

**Ejemplo**:
```markdown
## I. Work Requirements.
This meeting requires all units to earnestly implement...
```
Las dos líneas anteriores se fusionarán en un solo párrafo: "I. Work Requirements." mantiene el formato de subencabezado y "This meeting..." el formato del cuerpo.

**Nota**:
- No puede haber una línea vacía entre el subencabezado y el cuerpo; de lo contrario, se reconocerán como párrafos separados.
- Si el subencabezado no termina con puntuación final y no hay línea vacía antes del cuerpo, el cuerpo se fusionará con la línea del encabezado y se ajustará el formato.

### Conversión bidireccional de separadores

Soporta la conversión bidireccional entre separadores de Markdown y saltos de página/saltos de sección/líneas horizontales de Word:

-   **DOCX → MD**: Los saltos de página, saltos de sección y líneas horizontales de Word se convierten automáticamente en separadores de Markdown.
-   **MD → DOCX**: Markdown `---`, `***`, `___` se convierte automáticamente en elementos de Word correspondientes.
-   **Configurable**: Las relaciones de mapeo se pueden personalizar en la interfaz de configuración.

### Inserción de imágenes y tamaño

Soporta imágenes incrustadas estilo Obsidian/Wiki y Markdown estándar, con tamaño opcional (px):

```markdown
![[image.png]]
![[image.png|300]]
![[image.png\|300]]
![alt](image.png =300x200)
![alt](image.png =300x)
![alt|300](image.png)
```

- Sin tamaño: tamaño original, limitado por el ancho disponible (página/celda)
- Con tamaño: permite agrandar, pero sigue limitado por el ancho disponible
- Párrafo solo imagen: usa el estilo de párrafo “Image” (centrado, interlineado simple)

## 📖 Guía de uso detallada

### Word a Markdown

1.  Arrastra el archivo `.docx` a la ventana del programa.
2.  El programa analiza automáticamente la estructura del documento.
3.  Genera un archivo `.md` que incluye metadatos YAML.

**Formatos compatibles**:
-   `.docx` - Documento Word estándar.
-   `.doc` - Se convierte automáticamente a DOCX para procesar.
-   `.wps` - Documento WPS convertido automáticamente.

**Opciones de exportación**:

| Opción | Descripción |
| :--- | :--- |
| **Extraer imágenes** | Si se activa, las imágenes se extraen a la carpeta de salida y se insertan enlaces en el archivo MD. |
| **OCR de imágenes** | Si se activa, realiza OCR sobre imágenes y crea un archivo `.md` de imagen (con el texto reconocido). |
| **Limpiar números de subtítulos** | Si se activa, elimina números antes de subtítulos (p. ej., "一、", "（一）", "1.", etc.). |
| **Añadir números de subtítulos** | Si se activa, añade números automáticamente según los niveles de encabezado (configurable). |

### Markdown a Word

1.  Prepara un archivo `.md` con cabecera YAML.
2.  Arrástralo a la ventana y selecciona la plantilla Word correspondiente.
3.  El programa rellena la plantilla y genera el documento.

**Opciones de conversión**:

| Opción | Descripción |
| :--- | :--- |
| **Limpiar números de subtítulos** | Si se activa, elimina números antes de subtítulos. |
| **Añadir números de subtítulos** | Si se activa, añade números automáticamente según el nivel de encabezado. |

**Nota**: Si hay párrafos donde se mezclan subtítulos y cuerpo, deben mantenerse estrictamente los saltos de línea en el archivo MD (consulta "Saltos de línea y párrafos" arriba).

### Procesamiento automático de estilos de plantilla

El convertidor detecta y procesa automáticamente los estilos de la plantilla durante la conversión Markdown → DOCX:

#### Clasificación de estilos

**Estilo de párrafo (Paragraph Style)**: Se aplica a todo el párrafo.

| Estilo | Comportamiento de detección | Inyección si falta | Origen |
| :--- | :--- | :--- | :--- |
| Heading (1~9) | Detecta estilo de párrafo | Estilos de encabezado de la plantilla | Word incorporado |
| Code Block | Detecta estilo de párrafo | Fuente Consolas + fondo gris | Definido por el software |
| Quote (1~9) | Detecta estilo de párrafo | Fondo gris + borde izquierdo | Definido por el software |
| Formula Block | Detecta estilo de párrafo | Estilo específico de fórmula | Definido por el software |
| Separator (1~3) | Detecta estilo de párrafo | Estilo de párrafo con borde inferior | Definido por el software |

**Estilo de carácter (Character Style)**: Se aplica al texto seleccionado.

| Estilo | Comportamiento de detección | Inyección si falta | Origen |
| :--- | :--- | :--- | :--- |
| Inline Code | Detecta estilo de carácter | Fuente Consolas + sombreado gris | Definido por el software |
| Inline Formula | Detecta estilo de carácter | Estilo específico de fórmula | Definido por el software |

**Estilo de tabla (Table Style)**: Se aplica a toda la tabla.

| Estilo | Comportamiento de detección | Inyección si falta | Origen |
| :--- | :--- | :--- | :--- |
| Three-Line Table | Prioridad de configuración del usuario | Definición de estilo de tabla de tres líneas | Definido por el software |
| Grid Table | Prioridad de configuración del usuario | Definición de estilo de tabla con rejilla | Definido por el software |

**Definición de numeración (Numbering Definition)**: Se usa para formatos de lista.

| Tipo | Comportamiento de detección | Manejo si falta |
| :--- | :--- | :--- |
| List Numbering | Escanea definiciones existentes de listas en la plantilla | Usa preajustes decimal/bullet |

#### Internacionalización de nombres de estilo

-   **Estilos incorporados de Word** (heading 1~9):
    -   Los nombres de estilo usan nombres estándar en inglés (p. ej., `heading 1`).
    -   Word muestra nombres localizados según el idioma del sistema.
-   **Estilos definidos por el software** (Code Block, Quote, Formula, Separator, Table, etc.):
    -   Inyecta nombres según el idioma de la interfaz del software.

**Sugerencia**: Tras personalizar estilos en la plantilla, el convertidor usará tus estilos; si no existen, usará estilos predefinidos.

### Procesamiento de archivos de hoja de cálculo

1.  **Excel/CSV a Markdown**: Arrastra archivos `.xlsx` o `.csv` para convertirlos automáticamente a tablas Markdown.
2.  **Markdown a Excel**: Prepara un archivo MD y selecciona una plantilla Excel para convertir.

**Formatos compatibles**:
-   `.xlsx` - Documento Excel estándar.
-   `.xls` - Se convierte automáticamente a XLSX para procesar.
-   `.et` - Hoja de cálculo WPS convertida automáticamente.
-   `.csv` - Tabla de texto CSV.

### Función de revisión de texto

El programa ofrece cuatro reglas de revisión personalizables:

1.  **Comprobación de pares de puntuación** - Detecta si paréntesis y comillas emparejadas coinciden.
2.  **Revisión de símbolos** - Detecta uso mixto de puntuación china e inglesa.
3.  **Comprobación de errores tipográficos** - Comprueba errores comunes basándose en un diccionario personalizado.
4.  **Detección de palabras sensibles** - Detecta palabras sensibles basándose en un diccionario personalizado.

**Diccionarios personalizados**: Edita visualmente los diccionarios de errores tipográficos y palabras sensibles en "Configuración".

**Uso**:
1.  Arrastra el documento Word a revisar al programa.
2.  Marca las reglas necesarias.
3.  Haz clic en "Revisión de texto".
4.  Los resultados aparecen como comentarios en el documento.

## 🛠️ Sistema de plantillas

### Usar plantillas existentes

El programa incluye varias plantillas, incluidas versiones multilingües. Los archivos de plantilla están en el directorio `templates/`.

### Plantillas personalizadas

1.  Crea un archivo de plantilla con Word o WPS.
2.  Consulta plantillas existentes e inserta marcadores como `{{Title}}`, `{{DocumentNumber}}`, etc., donde sea necesario rellenar.
3.  En la plantilla, los estilos incorporados Heading 1 ~ Heading 5 deben modificarse manualmente.
4.  Guarda la plantilla en el directorio `templates/`.
5.  Reinicia el programa y la nueva plantilla se cargará automáticamente.

También puedes copiar una plantilla existente, modificarla y renombrarla.

### Uso de marcadores de posición

#### Marcadores de posición en plantillas Word

**Marcadores de campos YAML**: Usa `{{Field Name}}` en la plantilla; se reemplazará por el valor correspondiente del encabezado YAML del archivo Markdown durante la conversión.

| Marcador | Descripción |
| :--- | :--- |
| `{{Title}}` | Título del documento (prioridad abajo) |
| `{{Body}}` | Posición donde se inserta el cuerpo Markdown |
| Otros | Admite cualquier campo personalizado |

**Prioridad para obtener el título**:

| Prioridad | Origen | Descripción |
| :--- | :--- | :--- |
| 1 | YAML `Title` | Prioridad más alta |
| 2 | YAML `aliases` | Toma el primer elemento de la lista o el valor de cadena |
| 3 | Nombre de archivo | Nombre sin extensión `.md` |

**Soporte multilingüe**: Los marcadores de título y cuerpo admiten múltiples idiomas, por ejemplo, título `{{title}}`, `{{标题}}`, `{{Titel}}`, etc., cuerpo `{{body}}`, `{{正文}}`, `{{Inhalt}}`, etc.

#### Marcadores de posición en plantillas Excel

Las plantillas Excel admiten tres tipos de marcadores:

**1. Marcador de campo YAML** `{{Field Name}}`

Rellena un valor único del encabezado YAML:

```markdown
---
ReportName: 2024 Annual Sales Statistics
Unit: Sales Dept
---
```

`{{ReportName}}`, `{{Unit}}` se reemplazan por los valores correspondientes. El título sigue las mismas reglas de prioridad.

**2. Marcador de relleno por columna** `{{↓Field Name}}`

Extrae datos de la tabla Markdown y rellena **hacia abajo** fila a fila desde la posición del marcador:

```markdown
| ProductName | Quantity |
|:--- |:--- |
| Product A | 100 |
| Product B | 200 |
```

`{{↓ProductName}}` se reemplaza por "Product A" y la siguiente fila se rellena con "Product B".

**3. Marcador de relleno por fila** `{{→Field Name}}`

Extrae datos de la tabla Markdown y rellena **hacia la derecha** columna a columna desde la posición del marcador:

```markdown
| Month |
|:--- |
| Jan |
| Feb |
| Mar |
```

`{{→Month}}` se rellenará como "Jan", "Feb", "Mar" hacia la derecha.

**Manejo de celdas combinadas**: El programa omite automáticamente las celdas no iniciales de las combinadas.

**Fusión de datos de múltiples tablas**: Si hay varias tablas en Markdown con el mismo encabezado, los datos se fusionan y se rellenan secuencialmente.

## 🖥️ Uso de la interfaz gráfica

La mayoría de los usuarios utilizan el software a través de la interfaz gráfica. A continuación se muestra una guía detallada.

### Vista general de la interfaz

El programa utiliza un **diseño adaptativo de tres columnas**:

| Área | Descripción | Cuándo se muestra |
| :--- | :--- | :--- |
| **Columna central (área principal)** | Zona de arrastrar archivos, panel de operaciones, barra de estado | Siempre visible |
| **Columna derecha** | Selector de plantillas / panel de conversión | Se expande automáticamente tras seleccionar un archivo |
| **Columna izquierda** | Lista de archivos por lotes (agrupada por tipo) | Se muestra al cambiar al modo por lotes |

### Flujo básico de operación

1.  **Inicia el programa**: Doble clic en `DocWen.exe` (Windows empaquetado) o ejecuta `docwen-gui`.
2.  **Importa un archivo**:
    -   Método 1: Arrastra el archivo a la ventana.
    -   Método 2: Pulsa "Add" en la zona de arrastre para seleccionar archivos.
3.  **Selecciona plantilla** (si es necesario): El panel derecho se expande; selecciona una plantilla adecuada.
4.  **Configura opciones**: Marca las opciones necesarias en el panel de operaciones.
5.  **Ejecuta**: Pulsa el botón correspondiente (p. ej., "Export MD", "Convert to DOCX", etc.).
6.  **Ver resultado**: La barra de estado muestra el progreso; pulsa el icono 📍 para localizar el archivo de salida.

### Modo de archivo único vs modo por lotes

El programa admite dos modos, conmutables desde el botón de la zona de arrastre:

**Modo de archivo único** (predeterminado):
-   Procesa un archivo cada vez.
-   Interfaz simple, adecuada para uso diario.

**Modo por lotes**:
-   Importa varios archivos a la vez.
-   La columna izquierda muestra una lista agrupada por tipo.
-   Permite añadir/eliminar/ordenar en lote.
-   Al hacer clic en un archivo de la lista, cambia el objetivo de operación.

### Funciones del panel de operaciones

El panel ajusta automáticamente las operaciones disponibles según el tipo de archivo:

| Tipo de archivo | Operaciones disponibles |
| :--- | :--- |
| Documento Word | Exportar MD, Convertir a PDF, Revisión de texto, OCR |
| Markdown | Convertir a DOCX, Convertir a PDF |
| Hoja de cálculo Excel | Exportar MD, Convertir a PDF, Resumen de tablas |
| PDF | Exportar MD, Unir, Dividir, OCR |
| Imagen | Conversión de formato, Compresión, OCR |

### Pantalla de configuración

Pulsa el botón ⚙️ para abrir la configuración:

-   **General**: Tema, idioma, opacidad.
-   **Conversión**: Valores predeterminados de opciones.
-   **Salida**: Directorio de salida y reglas de nombre.
-   **Revisión**: Editar diccionarios.
-   **Estilo**: Configuración de estilos.
### Atajos

-   **Arrastrar archivo externo**: Importa arrastrándolo.
-   **Doble clic en resultado de la barra de estado**: Abre la carpeta de salida.
-   **Clic derecho en plantilla**: Abre la ubicación de la plantilla.
---

## 🔧 Uso en línea de comandos

Además de la GUI, el programa ofrece una CLI adecuada para automatización y procesamiento por lotes.

### Modos de ejecución

-   **Modo CLI**: Usa subcomandos (p. ej. `convert`, `validate`) para automatización y procesamiento por lotes.
### Ejemplos comunes

```bash
# Versión empaquetada (Windows)
DocWenCLI.exe convert document.docx --to md

# Exportar Word a Markdown (Extraer imágenes + OCR)
DocWenCLI.exe convert report.docx --to md --extract-img --ocr

# Markdown a Word (Especificar plantilla)
DocWenCLI.exe convert document.md --to docx --template "Template Name"

# Conversión por lotes (Saltar confirmación, continuar si hay error)
DocWenCLI.exe convert *.docx --to md --batch --yes --continue-on-error

# Revisión de documento
DocWenCLI.exe validate document.docx --check typo --check punct

# Unir/Dividir PDF
DocWenCLI.exe merge-pdfs *.pdf
DocWenCLI.exe split-pdf report.pdf --pages "1-3,5,7-10"

# Desde el código fuente / pip
docwen convert document.docx --to md
docwen convert report.docx --to md --extract-img --ocr
```

### Comandos y opciones principales

| Comando / Opción | Descripción |
| :--- | :--- |
| `convert <files...> --to <fmt>` | Convertir a formato destino (incluido `md`) |
| `validate <files...> --check ...` | Revisar documentos (`--check typo/punct/symbol/sensitive/all/none`) |
| `merge-pdfs <files...>` | Unir archivos PDF/OFD/XPS |
| `split-pdf <file> --pages ...` | Dividir un PDF por rangos de páginas |
| `merge-tables <files...> --mode row|col|cell` | Unir tablas |
| `merge-images-to-tiff <files...>` | Unir imágenes en TIFF |
| `md-numbering <files...>` | Procesar numeración de encabezados Markdown |
| `templates list [--for docx|xlsx]` | Listar plantillas disponibles |
| `optimizations list [--scope ...]` | Listar optimizaciones disponibles |
| `formats list [--for-source document|spreadsheet|layout|image|markdown]` | Listar formatos de destino disponibles |
| `inspect <file>` | Inspeccionar categoría/formato y acciones compatibles |
| `--template <name>` | Nombre de plantilla (usado con `convert`) |
| `--extract-img` / `--no-extract-img` / `--ocr` | Opciones para `convert --to md` |
| `--optimize-for <id>` | Activar optimización explícitamente (p. ej., `gongwen`, `invoice_cn`) |
| `--batch` / `--jobs` / `--continue-on-error` | Controles de procesamiento por lotes |
| `--json` | Salida del resultado en formato JSON |
| `--quiet` / `-q` | Modo silencioso, reducir salida |
| `--lang` | Cambiar idioma (afecta help/mensajes) |

## 🔌 Plugin de Obsidian

El proyecto incluye un plugin de Obsidian para funcionar junto con el convertidor:

### Funciones principales

-   **🚀 Inicio con un clic** - Icono lateral para iniciar rápidamente el convertidor.
-   **📂 Transferencia automática** - Pasa automáticamente la ruta del archivo abierto.
-   **🔄 Gestión de instancia única** - Si ya está en ejecución, envía el archivo sin reiniciar.
-   **💪 Recuperación ante fallos** - Detecta el estado del proceso y limpia residuos automáticamente.

### Principio de funcionamiento

El plugin interactúa con el convertidor mediante IPC basado en el sistema de archivos:

1.  **Primer clic** → Inicia el convertidor y pasa el archivo actual.
2.  **Clic de nuevo (con archivo)** → Sustituye el archivo (modo de archivo único).
3.  **Clic de nuevo (sin archivo)** → Activa la ventana del convertidor.

### Instalación

El plugin se publica en un repositorio separado. Consulta [docwen-obsidian](https://github.com/ZHYX91/docwen-obsidian) para instrucciones.

## ❓ Preguntas frecuentes

### ¿Qué hago si falla la conversión?

-   Comprueba si el archivo está en uso por otra aplicación.
-   Confirma que el formato sea correcto.
-   Revisa los logs en `logs/`.

### ¿No aparece la plantilla?

-   Confirma que las plantillas estén en `templates/`.
-   Comprueba si el archivo de plantilla está dañado.
-   Reinicia el programa para recargar plantillas.

### ¿La función de revisión no funciona?

-   Confirma que el documento sea `.docx`.
-   Comprueba que el documento contenga texto editable.
-   Confirma que las reglas de revisión estén habilitadas en la configuración.

### ¿El formato de salida no es el esperado?

-   El programa genera documentos según los estilos de la plantilla. Para ajustar la salida, modifica los estilos en el archivo de plantilla.
-   Las plantillas están en `templates/`.
-   Tras modificar estilos, todos los documentos convertidos con esa plantilla aplicarán los cambios.

### ¿Las celdas de fórmula están vacías después de la conversión de Excel a Markdown?

Esto es un comportamiento esperado. El programa lee los **valores en caché** de las celdas y no las fórmulas.

**Razón técnica**:
-   En Excel, las celdas con fórmula almacenan la fórmula y el último resultado calculado (valor en caché).
-   El programa usa `data_only=True` y solo lee valores en caché.
-   Si el archivo nunca se abrió en Excel o no se guardó tras recalcular, el valor en caché puede estar vacío.

**Solución**:
1.  Abre el archivo en Excel.
2.  Espera a que termine el cálculo.
3.  Guarda el archivo.
4.  Convierte de nuevo.

## 🔒 Características de seguridad

-   **Funcionamiento completamente local**: Todo el procesamiento se realiza en local, sin depender de red.
-   **Aislamiento de red**: Mecanismo incorporado para evitar fugas de datos.
-   **Sin subida de datos**: Los archivos del usuario no se suben a ningún servidor.
-   **Modo de seguridad estricto**: activado por defecto; la aplicación se cierra si fallan las comprobaciones de seguridad. Ver [doc/技术文档.md](doc/技术文档.md).

## 📜 Licencia

Este proyecto está licenciado bajo **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Este proyecto usa PyMuPDF (AGPL-3.0), por lo que el proyecto completo también se distribuye bajo AGPL-3.0.
-   Puedes usar, modificar y distribuir este software.
-   Si modificas este software y prestas servicios a través de una red, debes proporcionar el código fuente modificado a los usuarios.
-   Para más información, consulta [LICENSE](LICENSE).

### Contacto

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Email**: zhengyx91@hotmail.com

---

**Author**: ZhengYX
