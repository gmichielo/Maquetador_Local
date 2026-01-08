import pdfplumber
from docx import Document
from docx2pdf import convert
import platform
import re
import unicodedata
import shutil
import os
import time
import threading
import pythoncom

# =========================================================
# 1. UTILIDADES
# =========================================================

def normalize_text(text):
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{2,}', '\n', text)
    return text.strip()


def read_pdf(path):
    blocks = []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            txt = page.extract_text(layout=True)
            if txt:
                blocks.append(txt)
    return "\n".join(blocks)


# =========================================================
# 2. NORMALIZACIÓN ATS
# =========================================================

def rebuild_structure(text):
    headers = [
        "perfil profesional", "profile",
        "experiencia laboral", "work experience",
        "education", "educacion",
        "skills", "habilidades",
        "languages", "idiomas"
    ]

    for h in headers:
        text = re.sub(rf"\s*{h}\s*", f"\n{h.upper()}\n", text, flags=re.IGNORECASE)

    # Insertar salto de línea antes de cualquier • (u otros bullets) y mantener el símbolo
    text = re.sub(r"\s*([•\*\|▪●])\s*", r"\n\1 ", text)
    
    return normalize_text(text)

def split_lines(text):
    return [l.strip() for l in text.split("\n") if len(l.strip()) > 2]


# =========================================================
# 3. EXTRACCIÓN
# =========================================================

def extract_name(lines):
    """
    Extrae el nombre de un CV, intentando ser más flexible.
    """
    for line in lines[:10]:  # solo revisar primeras 10 líneas
        line_clean = line.strip()
        # ignorar líneas vacías o con números o emails
        if not line_clean:
            continue
        if re.search(r'\d', line_clean):
            continue
        if re.search(r'\S+@\S+', line_clean):  # ignorar emails
            continue
        if len(line_clean.split()) >= 2:  # debe tener al menos 2 palabras
            # eliminar títulos o palabras comunes como "perfil", "experiencia"
            if re.search(r'perfil|experiencia|skills|educacion', line_clean, re.IGNORECASE):
                continue
            return line_clean  # devuelve tal cual
    return "Nombre no detectado"


def extract_contact(text):
    email = re.search(r'\S+@\S+', text)

    phone = re.search(
        r'(\+\d{1,3}[\s\-]?)?(\(?\d{2,4}\)?[\s\-]?)?\d{3,4}[\s\-]?\d{3,4}',
        text
    )

    github = re.search(r'github\.com/\S+', text, re.IGNORECASE)
    linkedin = re.search(r'https?://(www\.)?linkedin\.com/\S+', text, re.IGNORECASE)

    telefono = phone.group(0).replace("\n", " ").strip() if phone else ""

    return {
        "email": email.group(0) if email else "",
        "telefono": telefono,
        "github": github.group(0) if github else "",
        "linkedin": linkedin.group(0) if linkedin else ""
    }


SECTIONS = {
    "perfil": ["perfil profesional", "profile"],
    "experiencia": ["experiencia laboral", "work experience","experiencia profesional"],
    "educacion": ["educacion", "education","certificaciones","formacion academica"],
    "skills": ["skills", "habilidades","experticia tecnica"],
    "idiomas": ["idiomas", "languages", "language"],
    "proyectos":["proyectos","proyectos destacados"]
}


def split_by_sections(lines):
    data = {k: [] for k in SECTIONS}
    current = None

    for line in lines:
        low = line.lower()
        for k, keys in SECTIONS.items():
            if any(h in low for h in keys):
                current = k
                break
        else:
            if current:
                data[current].append(line)

    return data


def extract_skills(lines):
    skills = []
    seen = set()

    for l in lines:
        for s in re.split(r"[,:]", l):
            s = s.strip()

            if len(s) <= 2:
                continue

            # evitar duplicados manteniendo orden
            key = s.lower()
            if key in seen:
                continue

            seen.add(key)
            skills.append(s)

    return skills


def extract_idiomas(lines):
    idiomas = {}
    for l in lines:
        matches = re.findall(r'([A-Za-zÁÉÍÓÚáéíóú]+)\s*[:\-]\s*(\w+)', l)
        for lang, lvl in matches:
            idiomas[lang] = lvl
    return idiomas


# =========================================================
# 4. PARSER PRINCIPAL
# =========================================================

def parse_cv(pdf_path):
    raw = read_pdf(pdf_path)
    structured = rebuild_structure(raw)
    lines = split_lines(structured)
    sections = split_by_sections(lines)

    return {
        "nombre": extract_name(lines) or "Nombre no detectado",
        "contacto": extract_contact(structured),
        "perfil": " ".join(sections["perfil"]),
        "skills": extract_skills(sections["skills"]),
        "experiencia": sections["experiencia"],
        "educacion": sections["educacion"],
        "idiomas": extract_idiomas(sections["idiomas"])
    }


# =========================================================
# 5. DOCX / PDF
# =========================================================

def cv_json_to_docx_data(cv):
    return {
        "NOMBRE": cv["nombre"],
        "EMAIL": cv["contacto"]["email"],
        "TELEFONO": cv["contacto"]["telefono"],
        "GITHUB": cv["contacto"]["github"],
        "LINKEDIN": cv["contacto"]["linkedin"],
        "PERFIL": cv["perfil"],
        "SKILLS": ", ".join(cv["skills"]),
        "EXPERIENCIA": "\n".join(cv["experiencia"]),
        "EDUCACION": "\n".join(cv["educacion"]),
        "IDIOMAS": "\n".join(f"{k}: {v}" for k, v in cv["idiomas"].items())
    }


def replace_placeholders(doc, data):
    for p in doc.paragraphs:
        for run in p.runs:
            for k, v in data.items():
                run.text = run.text.replace(f"{{{{{k}}}}}", v)

    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for run in p.runs:
                        for k, v in data.items():
                            run.text = run.text.replace(f"{{{{{k}}}}}", v)


def generate_cv_from_template(template_path, cv_json, output_dir="output"):
    """
    Genera un DOCX y un PDF desde la plantilla usando docx2pdf.
    Usa threading + pythoncom.CoInitialize() para que funcione en Flask.
    """
    os.makedirs(output_dir, exist_ok=True)

    # -----------------------------
    # Preparar nombres únicos
    # -----------------------------
    safe_name = cv_json["nombre"].replace(" ", "_") or "CV"
    timestamp = int(time.time())
    docx_out = os.path.abspath(os.path.join(output_dir, f"CV_{safe_name}_{timestamp}.docx"))
    pdf_out = os.path.abspath(os.path.join(output_dir, f"CV_{safe_name}_{timestamp}.pdf"))
    template_path = os.path.abspath(template_path)

    # -----------------------------
    # Limpiar archivos antiguos (opcional)
    # -----------------------------
    if os.path.exists(docx_out):
        os.remove(docx_out)
    if os.path.exists(pdf_out):
        os.remove(pdf_out)

    # -----------------------------
    # Copiar plantilla y reemplazar placeholders
    # -----------------------------
    shutil.copy(template_path, docx_out)
    doc = Document(docx_out)

    data = cv_json_to_docx_data(cv_json)

    # Reemplazo de placeholders
    for p in doc.paragraphs:
        for run in p.runs:
            for k, v in data.items():
                run.text = run.text.replace(f"{{{{{k}}}}}", v)
    for t in doc.tables:
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    for run in p.runs:
                        for k, v in data.items():
                            run.text = run.text.replace(f"{{{{{k}}}}}", v)

    doc.save(docx_out)

    # -----------------------------
    # Función para generar PDF en hilo separado
    # -----------------------------
    pdf_generated = False

    def convert_pdf_thread():
        nonlocal pdf_generated
        try:
            pythoncom.CoInitialize()  # Inicializar COM en este hilo
            convert(docx_out, pdf_out)
            if os.path.exists(pdf_out):
                pdf_generated = True
        except Exception as e:
            print("Error generando PDF en hilo:", e)

    if platform.system().lower() == "windows":
        thread = threading.Thread(target=convert_pdf_thread)
        thread.start()
        thread.join()  # Esperamos a que termine el PDF

        if not pdf_generated:
            print("PDF no se pudo generar después de intentar en hilo.")
            pdf_out = None
    else:
        # En otros sistemas no se usa docx2pdf
        pdf_out = None

    # -----------------------------
    # Devolver DOCX siempre, PDF si se generó
    # -----------------------------
    return docx_out, pdf_out if pdf_generated else None