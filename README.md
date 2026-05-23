# 📄 Convertidor a Markdown

Aplicación de escritorio para convertir documentos a Markdown usando [MarkItDown](https://github.com/microsoft/markitdown).

Soporta PDF, Word, Excel, PowerPoint e imágenes. Los archivos en formato antiguo (`.doc`, `.xls`, `.ppt`) se convierten automáticamente usando LibreOffice.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue) ![License](https://img.shields.io/badge/license-MIT-green)

---

## Requisitos

- **Python 3.10 o superior** — [python.org](https://www.python.org/downloads/)
- **LibreOffice** (solo si vas a convertir `.doc`, `.xls` o `.ppt`) — [libreoffice.org](https://www.libreoffice.org/download/download-libreoffice/)

---

## Instalación

### 1. Clonar o descargar el repositorio

```bash
git clone https://github.com/tu-usuario/tu-repo.git
cd tu-repo
```

O simplemente descargá `convertidor.py` suelto.

### 2. Crear un entorno virtual (recomendado)

```bash
python -m venv .venv
```

Activarlo:

- **Windows:** `.venv\Scripts\activate`
- **Mac / Linux:** `source .venv/bin/activate`

### 3. Instalar dependencias

```bash
pip install markitdown[pdf,docx,xlsx,pptx]
```

> Si querés soporte completo (imágenes con OCR, audio, YouTube, etc.):
> ```bash
> pip install markitdown[all]
> ```
> Nota: en Python 3.14+ el extra `[all]` puede fallar por incompatibilidades. Usá la instalación por partes en ese caso.

---

## Ejecución

```bash
python convertidor.py
```

En Windows también podés hacer doble clic sobre `convertidor.py` si Python está asociado a archivos `.py`.

---

## Uso

1. Hacé clic en **Elegir archivo** o en la barra gris para seleccionar un documento
2. Presioná **Convertir**
3. El texto en Markdown aparecerá en el panel inferior
4. Usá **Guardar .md** para exportar el archivo, o **Copiar texto** para pegarlo donde quieras

### Formatos soportados

| Formato | Extensiones | Requiere |
|---|---|---|
| PDF | `.pdf` | — |
| Word moderno | `.docx` | — |
| Word antiguo | `.doc` | LibreOffice |
| Excel moderno | `.xlsx` | — |
| Excel antiguo | `.xls` | LibreOffice |
| PowerPoint moderno | `.pptx` | — |
| PowerPoint antiguo | `.ppt` | LibreOffice |
| Imágenes | `.png` `.jpg` `.jpeg` `.webp` `.bmp` | — |

---

## Solución de problemas

**"MarkItDown no está instalado"**
Ejecutá el paso 3 de instalación con el entorno virtual activo.

**"LibreOffice no encontrado"**
La app muestra un aviso amarillo al iniciar. Instalá LibreOffice desde [libreoffice.org](https://www.libreoffice.org) y reiniciá la app. Los archivos `.docx`, `.xlsx` y `.pptx` siguen funcionando sin él.

**Error al convertir un `.doc` / `.xls` / `.ppt`**
Verificá que LibreOffice esté instalado en la ruta estándar:
- Windows: `C:\Program Files\LibreOffice\`
- Linux/Mac: disponible como `libreoffice` o `soffice` en el PATH

**La app no abre con doble clic en Windows**
Abrí una terminal, activá el entorno virtual y ejecutá `python convertidor.py`.

---

## Licencia

MIT
