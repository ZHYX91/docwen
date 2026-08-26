# DocWen

<p align="center">
  <img src="https://raw.githubusercontent.com/ZHYX91/docwen/main/assets/icon.svg" alt="DocWen logo" width="120">
</p>

[English](https://github.com/ZHYX91/docwen/blob/main/README.md) · [简体中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-CN.md) · [繁體中文](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.zh-TW.md) · [Deutsch](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.de-DE.md) · [Français](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.fr-FR.md) · [Español](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.es-ES.md) · [Português](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.pt-BR.md) · [Русский](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ru-RU.md) · [日本語](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ja-JP.md) · [한국어](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.ko-KR.md) · [Tiếng Việt](https://github.com/ZHYX91/docwen/blob/main/docs/user-guides/README.vi-VN.md)

Herramienta de conversión de formatos de documentos y tablas: conversión bidireccional Word/Markdown/Excel. Se ejecuta completamente en local, garantizando seguridad y fiabilidad de los datos.

## 📖 Contexto del proyecto

Este software se diseñó originalmente para resolver problemas comunes en entornos de oficina:
- Los documentos llegan con formatos inconsistentes y hay que normalizarlos.
- Hay muchos tipos de archivo y cada uno tiene requisitos de formato distintos.
- Debe funcionar sin conexión (intranet/equipos antiguos).

**Filosofía de diseño**: herramienta ligera y “lista para usar”, con coste de aprendizaje muy bajo. No pretende sustituir a herramientas profesionales como LaTeX o Pandoc.

## ✨ Funciones principales

- **📄 Conversión de documentos** - Word ↔ Markdown, con conversión de fórmulas, mapeo de separadores (---/***/___) a saltos de página/sección/líneas horizontales y restauración de marker explícitos `<` / `^` de tablas Markdown como combinaciones rectangulares en Word. DOCX/DOC/WPS/RTF/ODT.
- **📊 Conversión de hojas de cálculo** - Excel ↔ Markdown. XLSX/XLS/ET/ODS/CSV/TSV. Incluye estrategias configurables de exportación de celdas combinadas (`fill / empty / marker`) y herramientas de resumen de tablas. Las plantillas Markdown→XLSX vuelven a admitir campos YAML y placeholders verticales y horizontales de columna; la restauración completa de plantillas Excel, imágenes y combinaciones sigue siendo un objetivo de paridad.
- **📑 PDF y archivos de maquetación** - PDF/XPS/OFD → Markdown o DOCX. Soporta unir/dividir PDF.
- **🖼️ Imágenes** - Conversión y compresión JPEG/PNG/GIF/BMP/TIFF/WebP/HEIC.
- **📥 Importación de otros formatos** - Soporta la conversión unidireccional de HTML/MHTML/ENEX/EPUB/PPTX/PPT a Markdown.
- **🔍 Reconocimiento de texto OCR** - RapidOCR integrado para extraer texto de imágenes y PDF.
- **✏️ Revisión** - Revisa archivos Word (.docx) y Markdown (.md) con reglas personalizables para puntuación, símbolos, errores tipográficos y palabras sensibles. Las reglas pueden editarse en la interfaz de configuración.
- **📝 Plantillas** - Sistema flexible para formatos de documentos e informes.
- **💻 GUI + CLI** - Interfaz gráfica y línea de comandos.
- **🔒 Procesamiento local con protección de salida de dependencias** - La conversión no depende de servicios en línea. Mientras DocWen se ejecuta, su proceso Python bloquea DNS e IPv4/IPv6 para dependencias internas; las aplicaciones Office externas conservan la política de red del sistema.
- **🔗 Ejecución de instancia única** - Gestiona automáticamente las instancias del programa y admite integración con el plugin complementario de Obsidian.

## 📸 Capturas de pantalla

| Lote | Markdown |
| --- | --- |
| ![Panel de lote](../assets/screenshots/batch-light.png) | ![Ventana principal](../assets/screenshots/main-light.png) |

| Documento | Hoja de cálculo |
| --- | --- |
| ![Panel de documento](../assets/screenshots/conversion-document-light.png) | ![Panel de hoja de cálculo](../assets/screenshots/conversion-spreadsheet-light.png) |

| Imagen | Archivos de maquetación |
| --- | --- |
| ![Panel de imagen](../assets/screenshots/conversion-image-light.png) | ![Panel de maquetación](../assets/screenshots/conversion-layout-light.png) |

Registro de cambios: ver [CHANGELOG.md](../CHANGELOG.md)

## 🚀 Inicio rápido

### Instalación desde el código fuente

**Requisitos previos**: Python 3.12

**Límite objetivo de 0.9**: Este código fuente crea paquetes para Windows x64 y Ubuntu 24.04 x64.
Las demás distribuciones Linux y macOS siguen siendo rutas de código fuente/desarrollo y no quedan
cubiertas por el paquete de Ubuntu.

**Opción 1: Usando uv (recomendado)**

Instala [uv](https://docs.astral.sh/uv/getting-started/), luego:

```bash
git clone https://github.com/ZHYX91/docwen.git
cd docwen
uv sync --frozen --all-extras
```

El código fuente, las pruebas y las compilaciones de DocWen 0.9 solo admiten el archivo de bloqueo incluido con `uv 0.12.0`; `pip install -e` no es compatible.

### Iniciar el programa

En la versión empaquetada de Windows: haz doble clic en `DocWen.exe` para abrir la interfaz gráfica. Tras la instalación desde el código fuente:

```bash
docwen-gui  # Modo GUI
docwen      # Modo CLI
```

### Notas para macOS

**Limitación actual**: En macOS, `convert`, `validate`, `number`, `merge` y `split` no están disponibles.
Las notas siguientes solo documentan dependencias opcionales para experimentos de desarrollo.

**Soporte de LibreOffice (Opcional)**

Para convertir formatos heredados como `.doc` y `.xls`, instala LibreOffice:  
Descarga: https://www.libreoffice.org/download/

**Soporte de imágenes HEIC (Opcional)**

Para procesar imágenes HEIC/HEIF:

```bash
brew install libheif
pip install pillow-heif
```

### Requisitos previos del GUI en Linux

**Destino de paquete compatible**: DocWen 0.9 admite la GUI y la CLI empaquetadas en Ubuntu 24.04
x64. Estos requisitos no amplían ese compromiso a otras distribuciones o arquitecturas.

- Entorno de escritorio instalado (GNOME, KDE, XFCE, etc.)
- La GUI usa PySide6 (Qt6) y ya no depende de Python Tk. Si el arranque falla por bibliotecas del sistema ausentes, instala las dependencias de Qt indicadas por el error (normalmente relacionadas con OpenGL/X11).
- Para servidores headless, prioriza la entrada CLI `docwen` en lugar de la GUI; las compilaciones empaquetadas de Windows también incluyen `DocWenCLI.exe`.

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
6.  **Ver resultado**: La barra de estado muestra el progreso y el resultado; pulsa la acción «Abrir salida» a la derecha para abrir la ubicación de salida.

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
| Markdown | Convertir a DOCX, Convertir a PDF, Revisión de texto |
| Hoja de cálculo Excel | Exportar MD, Convertir a PDF, Resumen de tablas |
| PDF | Exportar MD, Unir, Dividir, OCR |
| Imagen | Conversión de formato, Compresión, OCR |
| HTML/EPUB/PPTX etc. | Exportar MD |

### Pantalla de configuración

Pulsa el botón «Configuración» en la cabecera del área de operaciones para abrir la configuración:

La configuración está organizada en pestañas: **General**, **Texto**, **Corrección**, **Documento**, **Hoja**, **Imagen**, **Maquetación**, **Enlaces**, **Formato**, **Salida**, **Exportar**, **Registro**, **Otros**.
### Atajos

-   **Arrastrar archivo externo**: Importa arrastrándolo.
-   **Abrir salida**: Pulsa la acción «Abrir salida» en la parte derecha de la barra de estado para abrir la ubicación de salida.
-   **Clic derecho en plantilla**: Abre la ubicación de la plantilla.
---

## 🔧 Uso de la CLI

Además de la interfaz gráfica, DocWen ofrece una interfaz de línea de comandos (CLI) para automatización, procesamiento por lotes e integraciones externas.

### Flujo recomendado para automatización

Para scripts, agentes o plugins, se recomienda este orden:

1. `inspect <file> [--json]`: detectar primero la categoría real del archivo, el formato y las acciones disponibles.
2. `schema convert`: leer el contrato legible por máquina y las reglas condicionales de `convert`.
3. `convert <file> --to <fmt> --output <path> --dry-run --json`: previsualizar detección, normalización y enrutamiento sin escribir archivos.
4. `convert <file> --to <fmt> --output <path> ...`: ejecutar la conversión real después de validar la previsualización.

### Ejemplos comunes

```bash
# Paquete de Windows
DocWenCLI.exe inspect document.docx --json

# Exportar el contrato de convert para scripts / agentes
DocWenCLI.exe schema convert

# Previsualizar cómo se ejecutará la conversión sin escribir resultados
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr --dry-run --json

# Exportar Word a Markdown (extraer imágenes + OCR)
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --ocr

# Markdown a Word (plantilla + modo de combinación de encabezado/cuerpo)
DocWenCLI.exe convert document.md --to docx --output document.docx --template template.docx.eb02ecb15c4730622ac6522f8399b5ab5dd8ee42d10d8aa0866f8616dbda45ef --heading-merge-mode punct_required

# Controlar el modo de imagen y la ubicación del texto OCR en Markdown
DocWenCLI.exe convert report.docx --to md --output report.md --extract-img --image-mode file --ocr --ocr-placement image_md

# Consultar capacidades en tiempo de ejecución y puertas de dependencia
DocWenCLI.exe doctor --json
DocWenCLI.exe resources list formats --json

# Revisión de documentos
DocWenCLI.exe validate document.docx --check typo --check punct
DocWenCLI.exe validate input.md --check typo --check punct

# Desde código fuente / uv
# inspect -> schema -> dry-run -> convert
# docwen inspect document.docx --json
# docwen schema convert
# docwen convert document.docx --to md --output document.md --dry-run --json
# docwen convert document.docx --to md --output document.md
```

### Comandos y opciones habituales

La tabla siguiente solo enumera los comandos mas habituales. Para la superficie completa de comandos, usa `docwen --help` (codigo fuente / uv) o `DocWenCLI --help` (version empaquetada).

| Comando / opción | Descripción |
| --- | --- |
| `convert <file> --to <fmt> --output <path>` | Punto de entrada unificado para conversiones. |
| `convert <file> --to <fmt> --output <path> --dry-run --json` | Previsualiza detección, normalización, enrutamiento y opciones efectivas sin ejecutar la conversión real. |
| `schema convert` | Exporta el contrato legible por máquina, los valores por defecto, las condiciones y las claves canónicas de `convert`. |
| `validate <file> --check ...` | Revisión documental (`typo/punct/symbol/sensitive/all/none`). Use `--json` para la envoltura de la CLI; `--report` es una ruta de archivo de informe opcional. |
| `inspect <file> [--json]` | Inspecciona categoría/formato del archivo, acciones recomendadas y advertencias por desajuste entre extensión y contenido. |
| `doctor --json` | Devuelve diagnósticos junto con resúmenes de capacidades en tiempo de ejecución y puertas de dependencia. |
| `resources list formats --json` | Lista formatos de destino por categoría de origen e incluye resúmenes de dependencias en tiempo de ejecución y limitaciones. |
| `resources list templates` | Lista las plantillas disponibles. |
| `resources list numbering-schemes` | Lista los esquemas de numeración disponibles. |
| `--template <id>` | ID canónico exacto devuelto por `resources list templates`; se rechazan nombres visibles, nombres de archivo y rutas. Los ID DOCX se aplican a `docx/doc/odt/rtf/wps/pdf` y los ID XLSX a `xlsx/xls/ods/csv`. |
| `--extract-img` / `--no-extract-img` / `--ocr` | Extracción de imágenes y OCR para `convert --to md`. |
| `--image-mode file|base64` | Controla cómo se emiten las imágenes durante la exportación a Markdown. |
| `--ocr-placement image_md|main_md` | Controla si el texto OCR se escribe en el Markdown auxiliar de la imagen o en el Markdown principal. |
| `--heading-merge-mode punct_required|always|never` | Controla la estrategia de combinación de encabezado + cuerpo para `convert --to docx`. |
| `--optimization <id>` | Activa explícitamente un perfil de optimización (vea `resources list optimizations`). |
| `batch convert|validate ... --jobs <n> [--continue-on-error]` | Controles de procesamiento por lotes. |
| `--json` / `--quiet` / `--timing` | Salida estructurada, menos logs y datos de tiempo para scripts o plugins. |

En el modo `punct_required`, la lista predeterminada exacta es `。：！？.:!?`. Puede editarse en la configuración de formato; un valor vacío desactiva la combinación en este modo. Las comas, los puntos y coma, las comas de enumeración, los guiones y los puntos suspensivos se excluyen de forma predeterminada.


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

**Párrafos mixtos**: Cuando un subencabezado debe mezclarse con el texto del cuerpo en el mismo párrafo (modo predeterminado: "Se requiere puntuación"), deben cumplirse estas condiciones:
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
- De forma predeterminada (modo "Se requiere puntuación"), si el subencabezado no termina con un signo de puntuación final, no se fusionará con la línea siguiente aunque no haya línea en blanco.
- Puedes cambiarlo en Configuración → Formato → "MarkDown a documento" → "Heading + body merge mode".

### Conversión bidireccional de separadores

Soporta la conversión bidireccional entre separadores de Markdown y saltos de página/saltos de sección/líneas horizontales de Word:

-   **DOCX → MD**: Los saltos de página, saltos de sección y líneas horizontales de Word se convierten automáticamente en separadores de Markdown.
-   **MD → DOCX**: Markdown `---`, `***`, `___` se convierte automáticamente en elementos de Word correspondientes.
-   **Configurable**: Las relaciones de mapeo se pueden personalizar en la interfaz de configuración.

### Listas de tareas

Soporta la conversión bidireccional de listas de tareas GFM:

```markdown
- [ ] Pendiente
- [x] Completado
```

-   **MD → DOCX**: Se renderiza como lista con viñetas con prefijo de texto `☐` / `☑`.
-   **DOCX → MD**: Convierte elementos de lista que comienzan con `☐` / `☑` / `☒` a `- [ ]` / `- [x]`.
-   **Nota sobre fuentes**: `☐`/`☑` pueden no mostrarse en algunas fuentes. Si es necesario, use fuentes como "Segoe UI Symbol" en su plantilla de Word.

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

### Gestión de enlaces

Soporta enlaces clicables en Markdown -> DOCX:

```markdown
[Docwen](https://example.com)
[[Target]]
[[Target|Open target]]
<https://example.com>
<user@example.com>
```

- Los enlaces Markdown y Wiki se convierten por defecto en hipervínculos de Word
- Los enlaces Wiki se resuelven como enlaces locales `file:///` cuando se encuentra el archivo de destino
- Los autolinks entre ángulos admiten `https://...` y correos `mailto:...`
- El autoenlace de URL sin delimitadores se evalúa por solicitud para Markdown -> DOCX, está desactivado por defecto y se activa con `[non_embed_links].auto_link_bare_url` en `configs/link.toml`
- Markdown -> XLSX no genera marcadores de hipervínculo para DOCX y conserva la sintaxis original del enlace

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
| **Optimización avanzada de campos** | Si se activa, extrae metadatos estructurados más completos; de lo contrario, utiliza el modo simplificado con solo título y subtítulo. |
| **Limpiar números de subtítulos** | Si se activa, elimina números antes de subtítulos (p. ej., "一、", "（一）", "1.", etc.). |
| **Añadir números de subtítulos** | Si se activa, añade números automáticamente según los niveles de encabezado (configurable). |

Nota: DOCX -> MD ahora restaura también la numeración multinivel vinculada en numbering.xml mediante estilos de párrafo (pStyle). Así, los prefijos de encabezado creados con listas multinivel de Word/WPS como "一、", "（一）", "1．", "（1）" y "①" se conservan tanto en el modo simplificado como en el modo avanzado de campos; el nivel del encabezado sigue detectándose correctamente cuando está activada la opción "Limpiar números de subtítulos".

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
2.  **Markdown a Excel**: Las tablas Markdown pueden exportarse a XLSX. Las plantillas admiten campos YAML, placeholders de columna e imagen y celdas combinadas o protegidas.

**Formatos compatibles**:
-   `.xlsx` - Documento Excel estándar.
-   `.xls` - Se convierte automáticamente a XLSX para procesar.
-   `.et` - Hoja de cálculo WPS convertida automáticamente.
-   `.csv` - Tabla de texto CSV.
-   `.tsv` - Tabla TSV separada por tabulaciones.


### Función de revisión de texto

El programa ofrece cuatro reglas de revisión personalizables:

1.  **Comprobación de pares de puntuación** - Detecta si paréntesis y comillas emparejadas coinciden.
2.  **Revisión de símbolos** - Detecta uso mixto de puntuación china e inglesa.
3.  **Comprobación de errores tipográficos** - Comprueba errores comunes basándose en un diccionario personalizado.
4.  **Detección de palabras sensibles** - Detecta palabras sensibles basándose en un diccionario personalizado.

**Diccionarios personalizados**: Edita visualmente los diccionarios de errores tipográficos y palabras sensibles en "Configuración".

**Uso**:
1.  Arrastra el documento Word o el archivo Markdown a revisar al programa.
2.  Marca las reglas necesarias.
3.  Haz clic en "Revisión de texto".
4.  Los resultados aparecen como comentarios en el documento. Para archivos Markdown, se genera un informe JSON.

Nota (informe JSON de revisión para Markdown):
- Motor: `text_rules` + adaptador Markdown `md_spell`
- Salida: la ruta actual de revisión en la CLI es `validate`; use `--json` para la envoltura de la CLI. `--report` es una ruta de archivo de informe opcional.

- Diferente de `--json` (JSON envoltorio de la CLI)

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

#### Marcadores de posición en plantillas Excel (objetivo de paridad legado)

Las plantillas XLSX admiten campos YAML, placeholders verticales `{{↓campo}}` y horizontales `{{→campo}}`, placeholders de imagen y celdas combinadas o protegidas.

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

**Manejo de celdas combinadas**:

- Markdown -> Excel conserva los rangos combinados originales de la plantilla.
- En las regiones de plantilla de columnas conocidas compuestas por marcadores continuos `{{↓NombreDelCampo}}`, el programa puede restaurar combinaciones rectangulares a partir de marcadores explícitos `<` / `^` en tablas Markdown.
- Solo participan en la detección de combinaciones las celdas cuyo contenido, tras recortar espacios, sea exactamente `<` o `^`; `\<` y `\^` permanecen como texto literal.
- Los rectángulos no válidos o los conflictos con rangos combinados existentes de la plantilla se degradan a texto normal con una advertencia, en lugar de sobrescribir a la fuerza la estructura de la plantilla.

**Fusión de datos de múltiples tablas**: Si hay varias tablas en Markdown con el mismo encabezado, los datos se fusionan y se rellenan secuencialmente.

## 🔌 Plugin de Obsidian

Hay un plugin complementario de Obsidian publicado por separado que funciona junto con el convertidor:

### Funciones principales

-   **🚀 Inicio con un clic** - Icono lateral para iniciar rápidamente el convertidor.
-   **📂 Transferencia automática** - Pasa automáticamente la ruta del archivo abierto.
-   **🔄 Gestión de instancia única** - Si ya está en ejecución, envía el archivo sin reiniciar.
-   **🔒 Control local acotado** - Usa solicitudes tipadas `status`, `open` y `activate` sin buscar procesos por nombre ni usar archivos de comandos o estado.

### Principio de funcionamiento

El transporte runtime/control de DocWen Core usa una canalización con nombre de Windows o un socket
AF_UNIX en Linux/macOS. Un bloqueo de archivo solo establece la propiedad de la instancia única; los
comandos de control no se transportan mediante archivos. Esto solo describe la capacidad del Core.
DocWen Assistant 2.0 sigue limitado al escritorio de Windows y no tiene aceptación combinada en Linux/macOS.

1.  **Primer clic** → Inicia el convertidor y pasa el archivo actual.
2.  **Clic de nuevo (con archivo)** → Sustituye el archivo (modo de archivo único).
3.  **Clic de nuevo (sin archivo)** → Activa la ventana del convertidor.

### Instalación

DocWen Assistant 2.0 usa DocWen Machine Protocol v1 y el único contrato Artifact Bundle v2. La versión del código
fuente no demuestra que esté publicada; instala solo una versión numérica que identifique explícitamente una versión
publicada y compatible de DocWen.

## 🔌 OpenClaw (Plugin + Skill)

OpenClaw 2.0 usa DocWen Machine Protocol v1 y el único contrato Artifact Bundle v2. La versión del código fuente no
demuestra que esté publicada; consulta la página de la versión numérica e instala solo después de que supere su
control de publicación inmutable.

## ❓ Preguntas frecuentes

### ¿Qué hago si falla la conversión?

-   Comprueba si el archivo está en uso por otra aplicación.
-   Confirma que el formato sea correcto.
-   Revisa en la configuración el campo "Ruta real actual del archivo de registro" o consulta los errores en el directorio de registros del usuario del sistema; si la verificación del paquete usa `DOCWEN_LOG_DIR`, revisa en su lugar el directorio sobrescrito.

### ¿No aparece la plantilla?

-   Confirma que las plantillas estén en `templates/`.
-   Comprueba si el archivo de plantilla está dañado.
-   Reinicia el programa para recargar plantillas.

### ¿La función de revisión no funciona?

-   Confirma que el documento sea `.docx` o `.md`.
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

## 🔒 Funcionalidades de seguridad

-   **Funcionamiento completamente local**: El procesamiento se realiza localmente por defecto y no depende de servicios en linea.
-   **Protección de salida de dependencias**: Las entradas GUI/CLI compatibles activan un guardián de auditoría CPython durante toda la vida del proceso Python principal. Bloquea toda resolución DNS/de nombres y las operaciones AF_INET/AF_INET6 `bind`, `connect`, `connect_ex`, `sendto` y `sendmsg`, conservando canalizaciones con nombre de Windows y sockets de dominio Unix.
-   **Límite explícito**: Los procesos iniciados por separado, incluidos Office/WPS/LibreOffice y el ayudante de Office, no se administran. Es una defensa ante conexiones accidentales de dependencias, no un sandbox del sistema operativo.
-   **Sin subida de datos**: Por defecto, los archivos del usuario no se suben de forma activa a servidores externos.
-   **Modo de seguridad estricto**: activado por defecto; la aplicacion se cierra si fallan las comprobaciones de seguridad centrales. Ver [Troubleshooting](../maintenance/troubleshooting.md).

## 📜 Licencia

Este proyecto está licenciado bajo **GNU Affero General Public License v3.0 (AGPL-3.0)**.

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

-   Este proyecto usa PyMuPDF (AGPL-3.0), por lo que el proyecto completo también se distribuye bajo AGPL-3.0.
- La GUI actual puede usar `PySide6-Fluent-Widgets` (QFluentWidgets) en rutas de host compatibles; esta dependencia sigue un modelo de doble licencia `GPLv3 / comercial`, mientras que este repositorio se sigue distribuyendo bajo AGPL.
-   Puedes usar, modificar y distribuir este software.
-   Si modificas este software y prestas servicios a través de una red, debes proporcionar el código fuente modificado a los usuarios.
-   Para más información, consulta [LICENSE](../../LICENSE).
- Para avisos de componentes de terceros, consulta [LICENSE_THIRD_PARTY.txt](../../LICENSE_THIRD_PARTY.txt); el resumen de distribucion esta en [NOTICE.txt](../../NOTICE.txt).

### Contacto

-   **GitHub**: https://github.com/ZHYX91/docwen
-   **Correo electrónico**: zhengyx91@hotmail.com

---

**Autor**: ZhengYX
