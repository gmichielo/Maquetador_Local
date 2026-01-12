from flask import Flask, render_template, request, send_file
import os
from cv_engine import parse_cv, generate_cv_from_template

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
TEMPLATES_FOLDER = "templates_docx"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

PLANTILLAS = {
    "1": "Plantilla1.docx",
    "2": "Plantilla2.docx",
    "3": "Plantilla3.docx",
    "4": "Plantilla4.docx"
}

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        pdf_file = request.files.get("cv_pdf")
        plantilla_id = request.form.get("plantilla")

        if not pdf_file or not plantilla_id:
            return render_template("index.html", success=False, error="Faltan datos")

        print("📄 CV recibido:", pdf_file.filename)
        print("📄 Plantilla:", plantilla_id)

        # Limpiar output
        for f in os.listdir(OUTPUT_FOLDER):
            os.remove(os.path.join(OUTPUT_FOLDER, f))

        pdf_path = os.path.join(UPLOAD_FOLDER, pdf_file.filename)
        pdf_file.save(pdf_path)

        plantilla_path = os.path.join(
            TEMPLATES_FOLDER,
            PLANTILLAS.get(plantilla_id)
        )

        print("🔍 Parseando CV...")
        cv_json = parse_cv(pdf_path)
        print("✅ CV parseado:")
        print(cv_json)

        print("📝 Generando CV final...")
        docx_path, pdf_out = generate_cv_from_template(
            plantilla_path,
            cv_json,
            OUTPUT_FOLDER
        )

        return render_template(
            "index.html",
            success=True,
            docx=os.path.basename(docx_path),
            pdf=os.path.basename(pdf_out) if pdf_out else None
        )

    return render_template("index.html", success=False)

@app.route("/download/<filename>")
def download(filename):
    return send_file(
        os.path.join(OUTPUT_FOLDER, filename),
        as_attachment=True
    )

if __name__ == "__main__":
    app.run(debug=True)
