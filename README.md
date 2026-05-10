# DOCX Template Preparer And Filler Tool

Esta es una herramienta para hacer que rellenar plantillas de Word sea más fácil y amigable.

## Características

- Conversión automática de placeholders:
  - [NOMBRE] → {{ nombre }}

- Conserva estilos del documento:
  - negrita
  - cursiva
  - formato del texto

- Soporta variables divididas entre múltiples runs de Word

- Generación automática de contratos DOCX

- Interfaz web local con Flask

- Exportación de documentos personalizados

## Flujo de trabajo

1. Crear una plantilla DOCX usando placeholders:
   [NOMBRE]
   [CÉDULA]

2. Ejecutar el preparador de plantilla

3. El sistema convierte automáticamente:
   {{ nombre }}
   {{ cedula }}

4. Ejecutar el generador de contratos

5. Completar el formulario y exportar el documento final

## Instalación

```bash
pip install -r requirements.txt
---

## Uso

### Preparar plantilla

```bash
python preparador_plantilla/src/preparador_plantilla.py
```

### Generar DOCX

```bash
python generador_docx/src/generador_docx.py
```

## Limitaciones conocidas

- Los textboxes de Word no son procesados
- Algunos elementos avanzados de OpenXML pueden no ser compatibles
- El sistema usa el estilo del primer run al reconstruir variables fragmentadas

## Tecnologías

- Python
- Flask
- python-docx
- docxtpl
- Jinja2
