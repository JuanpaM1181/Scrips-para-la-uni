# Scripts para la Uni

Modulo reutilizable de Python para automatizar la creacion de documentos academicos (reportes, rubricasy presentaciones) en formato DOCX/PDF.

## Instalacion

```bash
pip install -r requirements.txt
```

Opcional: instalar LibreOffice para conversion automatica a PDF.

```bash
sudo pacman -S libreoffice-fresh
```

## Estructura

```
uni_tools/
  cover.py       -- Llenar portada desde plantilla DOCX
  formatting.py  -- Parrafos, encabezados, tablas, cuerpo con formato
  toc.py         -- Insertar tabla de contenidos automatica
  signatures.py  -- Extraer firmas desde Firmas.docx
  pdf.py         -- Convertir DOCX a PDF con LibreOffice
  rubrics.py     -- Llenar rubricasy firmas
  text_mining.py -- Utilidades de mineria de texto
  config.py      -- Configuracion centralizada (rutas, estudiantes, profesor)

examples/
  generar_reporte_ejemplo.py
  generar_rubricas_ejemplo.py
```

## Uso basico

```python
from docx import Document
from uni_tools import add_heading_custom, add_body, insert_toc, convert_to_pdf
from uni_tools.cover import copy_template, fill_cover
from uni_tools.config import STUDENTS

dst = copy_template("/tmp/reporte.docx")
doc = Document(dst)

fill_cover(doc, STUDENTS, "Mi Materia", "9o (Periodo)", "1", "Titulo", "Reporte", "Fecha")
insert_toc(doc)
add_heading_custom(doc, "Introduccion", 1)
add_body(doc, "Texto del reporte...")

doc.save(dst)
convert_to_pdf(docx_path)
```

## Rubricas

```python
from uni_tools.rubrics import fill_rubric
from uni_tools.config import STUDENTS

template = "/ruta/a/plantilla_rubrica.docx"
for s in STUDENTS:
    fill_rubric(template, f"/tmp/rubrica_{s['matricula']}.docx", s["name"], s["matricula"])
```

## Configuracion

Editar `uni_tools/config.py` para cambiar rutas de plantillas, estudiantes o profesor.

## Dependencias

- python-docx
- lxml
- numpy (para text_mining.py)
- LibreOffice (para PDF)
