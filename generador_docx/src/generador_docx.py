from tkinter import Tk, filedialog
from flask import Flask, render_template, request, send_file
from docxtpl import DocxTemplate

import os
import sys
import json
import webbrowser

app = Flask(__name__)

# =====================================================
# BASE
# =====================================================

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# =====================================================
# ARCHIVOS
# =====================================================

PLANTILLA = os.path.join(
    BASE_DIR,
    "plantilla_jinja.docx"
)

ARCHIVO_VARIABLES = os.path.join(
    BASE_DIR,
    "plantilla_variables.json"
)

# =====================================================
# CARGAR VARIABLES
# =====================================================

def cargar_variables():

    if not os.path.exists(ARCHIVO_VARIABLES):
        return {}

    with open(
        ARCHIVO_VARIABLES,
        "r",
        encoding="utf-8"
    ) as f:

        return json.load(f)

# =====================================================
# INDEX
# =====================================================

@app.route("/")
def index():

    variables = cargar_variables()

    return render_template(
        "index.html",
        variables=variables
    )

# =====================================================
# GENERAR DOCX
# =====================================================

@app.route("/generar", methods=["POST"])
def generar():

    datos = request.form.to_dict()

    doc = DocxTemplate(PLANTILLA)

    doc.render(datos)

    # Crear ventana oculta
    root = Tk()

    # Ocultar ventana principal
    root.withdraw()

    # Mantener diálogo al frente
    root.attributes('-topmost', True)

    # Actualizar ventana
    root.update()

    # Mostrar "Guardar como"
    ruta_salida = filedialog.asksaveasfilename(
        title="Guardar documento",
        defaultextension=".docx",
        filetypes=[("Documento Word", "*.docx")],
    )

    # Restaurar prioridad
    root.attributes('-topmost', False)

    # Destruir ventana tkinter
    root.destroy()

    # Si cancela, no hacer nada
    if not ruta_salida:
        return ('', 204)

    # Guardar documento
    doc.save(ruta_salida)

    # No hacer nada más
    return ('', 204)

# =====================================================
# CERRAR
# =====================================================

@app.route('/cerrar', methods=['POST'])
def cerrar():

    os._exit(0)

# =====================================================
# MAIN
# =====================================================

if __name__ == "__main__":

    webbrowser.open(
        "http://127.0.0.1:5000"
    )

    app.run(
        debug=False,
        use_reloader=False
    )