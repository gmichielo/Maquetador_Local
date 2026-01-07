from PyPDF2 import PdfReader
from docx import Document
from docx2pdf import convert
import pythoncom
import re
import unicodedata
import json
import shutil
import os

PLANTILLA_UNO = r"C:\Users\Gabriel\Documents\CVs Gabriel\Plantilla1.docx"
PLANTILLA_DOS = r"C:\Users\Gabriel\Documents\CVs Gabriel\Plantilla2.docx"
PLANTILLA_TRES = r"C:\Users\Gabriel\Documents\CVs Gabriel\Plantilla3.docx"

PLANTILLAS = (PLANTILLA_UNO, PLANTILLA_DOS, PLANTILLA_TRES)

plantilla = ""
selected_plantilla = "0"
pdf_original = r"C:\Users\Gabriel\Documents\CVs Gabriel\CV_Gabriel_Michielon (FSa).pdf"

# 1️ UTILIDADES BASE

def normalize_text(text):
    text = unicodedata.normalize("NFKD", text)
    text = "".join(c for c in text if not unicodedata.combining(c))
    text = re.sub(r'[ \t]+', ' ', text)
    text = re.sub(r'\n{2,}', '\n', text)
    return text.strip()


def read_pdf(path):
    reader = PdfReader(path)
    text = ""
    for page in reader.pages:
        text += page.extract_text() or ""
    return text

# PROTEGER / RESTAURAR URLs

def protect_urls(text):
    urls = {}
    pattern = r'https?://\S+|www\.\S+|linkedin\.com/\S+|github\.com/\S+'

    def replacer(match):
        key = f"__URL_{len(urls)}__"
        urls[key] = match.group(0)
        return key

    protected_text = re.sub(pattern, replacer, text, flags=re.IGNORECASE)
    return protected_text, urls


def restore_urls(text, urls):
    for k, v in urls.items():
        text = text.replace(k, v)
    return text


# 2️ RECONSTRUIR ESTRUCTURA ATS

def rebuild_structure(text):
    # 1️⃣ Proteger URLs ANTES de tocar el texto
    text, urls = protect_urls(text)

    headers = [
        "PERFIL PROFESIONAL",
        "EXPERIENCIA LABORAL",
        "HABILIDADES TECNICAS",
        "EDUCACION",
        "IDIOMAS",
        "LENGUAJES",
    ]

    for h in headers:
        text = re.sub(rf"\s*{h}\s*", f"\n{h}\n", text, flags=re.IGNORECASE)

    text = re.sub(r"[•\*\|]", "\n", text)

    # Emails (solo salto antes)
    text = re.sub(r"(?<!\S)(\S+@\S+)", r"\n\1", text)

    # Teléfonos (sin romperlos)
    text = re.sub(
        r'(?<!\S)(\+\d{1,3}[\s\-]?\(?\d{2,3}\)?[\s\-]?\d{3}[\s\-]?\d{3})',
        r'\n\1',
        text
    )

    text = normalize_text(text)

    # 2️⃣ Restaurar URLs intactas
    text = restore_urls(text, urls)

    return text


def split_lines(text):
    return [l.strip() for l in text.split("\n") if len(l.strip()) > 2]


# 3️ EXTRACCIÓN DE DATOS

def extract_name(lines):
    for line in lines[:5]:
        if line.isupper() and len(line.split()) >= 2 and not any(c.isdigit() for c in line):
            return line.title()
    return ""

def normalize_urls(text):
    # Une saltos de línea dentro de URLs
    text = re.sub(
        r'(https?://[^\s]+)\s*\n\s*([^\s]+)',
        r'\1\2',
        text
    )
    return text

def extract_contact(text):
    email_match = re.search(r'\S+@\S+', text)

    phone_match = re.search(
        r'(\+\d{1,3}[\s\-]?)?(\(?\d{2,3}\)?[\s\-]?)?\d{3}[\s\-]?\d{3}',
        text
    )

    github_match = re.search(
        r'github\.com/[A-Za-z0-9\-_]+',
        text,
        re.IGNORECASE
    )

    linkedin_match = re.search(
        r'(https?:\/\/)?(www\.)?linkedin\.com\/(in|pub)\/[A-Za-z0-9\-_%]+\/?',
        text,
        re.IGNORECASE
    )

    telefono = phone_match.group(0) if phone_match else ""
    telefono = telefono.replace("\n", " ").strip()
    telefono = re.sub(r'\s+', ' ', telefono)

    linkedin = linkedin_match.group(0) if linkedin_match else ""
    linkedin = linkedin.replace("\n", "").replace(" ", "").strip()

    return {
        "email": email_match.group(0) if email_match else "",
        "telefono": telefono,
        "github": github_match.group(0) if github_match else "",
        "linkedin": linkedin
    }


SECTIONS = {
    "perfil": ["perfil profesional"],
    "experiencia": ["experiencia laboral"],
    "skills": ["habilidades", "lenguajes"],
    "educacion": ["educacion"],
    "idiomas": ["idiomas"]
}


def split_by_sections(lines):
    sections = {k: [] for k in SECTIONS}
    current = None

    for line in lines:
        lower = line.lower()
        found = False

        for key, aliases in SECTIONS.items():
            if any(a in lower for a in aliases):
                current = key
                found = True
                break

        if not found and current:
            sections[current].append(line)

    return sections


def extract_skills(lines):
    skills = set()
    for line in lines:
        for p in re.split(r"[,:]", line):
            p = p.strip()
            if len(p) > 2:
                skills.add(p)
    return sorted(skills)


def extract_idiomas(lines):
    idiomas = {}
    for line in lines:
        matches = re.findall(r'([A-Za-zÁÉÍÓÚÑáéíóúñ]+)\s*:\s*([A-Za-z0-9]+)', line)
        for lang, level in matches:
            idiomas[lang] = level
    return idiomas


# 4️ PARSER PRINCIPAL

def parse_cv(pdf_path):
    raw = read_pdf(pdf_path)
    structured = rebuild_structure(raw)
    lines = split_lines(structured)
    sections = split_by_sections(lines)

    return {
        "nombre": extract_name(lines),
        "contacto": extract_contact(structured),
        "perfil": " ".join(sections["perfil"]),
        "skills": extract_skills(sections["skills"]),
        "experiencia": sections["experiencia"],
        "educacion": sections["educacion"],
        "idiomas": extract_idiomas(sections["idiomas"])
    }


# 5️ WORD TEMPLATE ENGINE

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
    def replace_in_paragraph(paragraph):
        for run in paragraph.runs:
            for key, value in data.items():
                placeholder = f"{{{{{key}}}}}"
                if placeholder in run.text:
                    run.text = run.text.replace(placeholder, value)

    for p in doc.paragraphs:
        replace_in_paragraph(p)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_in_paragraph(p)


def generate_cv_from_template(template_path, cv_json, output_dir="output"):
    os.makedirs(output_dir, exist_ok=True)

    output_docx = os.path.join(output_dir, f"CV_FINAL_{cv_json['nombre']}.docx")
    output_pdf = os.path.join(output_dir, f"CV_FINAL_{cv_json['nombre']}.pdf")

    if os.path.exists(output_docx):
        os.remove(output_docx)
    if os.path.exists(output_pdf):
        os.remove(output_pdf)

    shutil.copy(template_path, output_docx)

    doc = Document(output_docx)
    replace_placeholders(doc, cv_json_to_docx_data(cv_json))
    doc.save(output_docx)

    pythoncom.CoInitialize()
    try:
        convert(output_docx, output_pdf)
    finally:
        pythoncom.CoUninitialize()

    return output_docx, output_pdf

