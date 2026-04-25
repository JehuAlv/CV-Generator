import os
import re
from cv_templates import TEMPLATES


def pre_split_sections(text):
    """
    Inserta salto de linea antes de palabras clave de seccion y campos
    para tolerar texto pegado en una sola linea.
    """
    keywords = [
        "DATOS PERSONALES", "RESUMEN PROFESIONAL", "LOGROS CLAVE",
        "EXPERIENCIA PROFESIONAL", "EDUCACION", "CURSOS Y CERTIFICACIONES",
        "HABILIDADES", "IDIOMAS",
        "PERSONAL INFORMATION", "PROFESSIONAL SUMMARY", "KEY ACHIEVEMENTS",
        "PROFESSIONAL EXPERIENCE", "EDUCATION", "CERTIFICATIONS & COURSES",
        "SKILLS", "LANGUAGES",
        "NOMBRE:", "CONTACTO:", "TITULO:", "TEMA:",
        "PUESTO:", "EMPRESA:", "FECHAS:",
        "CARRERA:", "ESCUELA:",
        "NAME:", "CONTACT:", "TITLE:",
        "POSITION:", "ROLE:", "COMPANY:", "DATES:",
        "DEGREE:", "SCHOOL:", "UNIVERSITY:",
    ]

    for kw in keywords:
        pattern = r'(?<!\n)(?=' + re.escape(kw) + r')'
        text = re.sub(pattern, '\n', text, flags=re.IGNORECASE)

    return text


def normalize_text(text):
    text = text.lstrip('\ufeff')
    replacements = {
        "\u2013": "-",
        "\u2014": "-",
        "\u2212": "-",
        "\u2010": "-",
        "\u2011": "-",
        "\u2012": "-",
        "\u2015": "-",
        "\u2022": "-",
        "\u25CF": "-",
        "\u25AA": "-",
        "\u25A0": "-",
        "\u2023": "-",
        "\u2043": "-",
        "\u00A0": " ",
        "\u202F": " ",
        "\u2009": " ",
        "\u2008": " ",
        "\u2007": " ",
        "\u2006": " ",
        "\u2005": " ",
        "\u2004": " ",
        "\u2003": " ",
        "\u2002": " ",
        "\u200B": "",
        "\u200C": "",
        "\u200D": "",
        "\uFEFF": "",
        "\t": " ",
        "\u201C": '"',
        "\u201D": '"',
        "\u2018": "'",
        "\u2019": "'",
        "\u201A": "'",
        "\u201E": '"',
    }

    for k, v in replacements.items():
        text = text.replace(k, v)

    text = text.replace("\r\n", "\n").replace("\r", "\n")

    lines = []
    for line in text.splitlines():
        line = line.strip()
        line = re.sub(r" {2,}", " ", line)
        line = re.sub(r"\s+:", ":", line)
        lines.append(line)

    cleaned = []
    prev_empty = False
    for line in lines:
        is_empty = (line == "")
        if is_empty and prev_empty:
            continue
        cleaned.append(line)
        prev_empty = is_empty

    return "\n".join(cleaned).strip()


def _starts(text, *prefixes):
    up = re.sub(r"\s+", " ", text.strip()).upper()
    up_clean = up.rstrip(": ")
    for p in prefixes:
        p_clean = p.upper().rstrip(": ")
        if up_clean.startswith(p_clean):
            return True
    return False


def parse_text(text):
    data = {
        "theme": "clasico", "nombre": "", "contacto": "", "titulo": "",
        "resumen": "", "logros": [], "experiencia": [], "educacion": [],
        "cursos": [], "habilidades": [], "idiomas": [],
    }

    text = pre_split_sections(text)
    text = normalize_text(text)
    lines = text.split("\n")
    section = None
    buffer = []
    current_job = None
    current_edu = None

    def flush_job():
        nonlocal current_job
        if current_job and current_job.get("puesto"):
            data["experiencia"].append(current_job)
        current_job = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("====") or line.startswith("----"):
            continue

        upper = re.sub(r":$", "", line.strip().upper())
        EN_TO_ES = {
            "PERSONAL INFORMATION": "DATOS PERSONALES",
            "PROFESSIONAL SUMMARY": "RESUMEN PROFESIONAL",
            "KEY ACHIEVEMENTS": "LOGROS CLAVE",
            "PROFESSIONAL EXPERIENCE": "EXPERIENCIA PROFESIONAL",
            "EDUCATION": "EDUCACION",
            "CERTIFICATIONS & COURSES": "CURSOS Y CERTIFICACIONES",
            "SKILLS": "HABILIDADES",
            "LANGUAGES": "IDIOMAS",
        }
        mapped = EN_TO_ES.get(upper, upper)
        if mapped in ("DATOS PERSONALES", "RESUMEN PROFESIONAL", "LOGROS CLAVE",
                      "EXPERIENCIA PROFESIONAL", "EDUCACION", "CURSOS Y CERTIFICACIONES",
                      "HABILIDADES", "IDIOMAS"):
            if section == "RESUMEN PROFESIONAL" and buffer:
                data["resumen"] = " ".join(buffer)
                buffer = []
            if section == "EXPERIENCIA PROFESIONAL":
                flush_job()
            section = mapped
            continue

        stripped = line.strip()
        if not stripped:
            continue

        if stripped.upper().startswith("TEMA:"):
            val = stripped.split(":", 1)[1].strip().lower()
            if val in TEMPLATES:
                data["theme"] = val
            continue

        if section == "DATOS PERSONALES":
            if _starts(stripped, "NOMBRE:", "NAME:"):
                data["nombre"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "CONTACTO:", "CONTACT:"):
                data["contacto"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "TITULO:", "TITLE:"):
                data["titulo"] = stripped.split(":", 1)[1].strip()

        elif section == "RESUMEN PROFESIONAL":
            buffer.append(stripped)

        elif section == "LOGROS CLAVE":
            m = re.match(r"^[-*]\s*(.+)", stripped)
            if m:
                data["logros"].append(m.group(1).strip())

        elif section == "EXPERIENCIA PROFESIONAL":
            if _starts(stripped, "PUESTO:", "POSITION:", "ROLE:"):
                flush_job()
                current_job = {"puesto": "", "empresa": "", "fechas": "", "bullets": []}
                current_job["puesto"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "EMPRESA:", "COMPANY:") and current_job:
                current_job["empresa"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "FECHAS:", "DATES:") and current_job:
                current_job["fechas"] = stripped.split(":", 1)[1].strip()
            elif current_job:
                m = re.match(r"^[-*]\s*(.+)", stripped)
                if m:
                    current_job["bullets"].append(m.group(1).strip())

        elif section == "EDUCACION":
            if _starts(stripped, "CARRERA:", "DEGREE:"):
                current_edu = {"titulo": stripped.split(":", 1)[1].strip(), "institucion": ""}
            elif _starts(stripped, "ESCUELA:", "SCHOOL:", "UNIVERSITY:") and current_edu:
                current_edu["institucion"] = stripped.split(":", 1)[1].strip()
                data["educacion"].append(current_edu)
                current_edu = None

        elif section == "CURSOS Y CERTIFICACIONES":
            m = re.match(r"^[-*]\s*(.+)", stripped)
            if m:
                data["cursos"].append(m.group(1).strip())

        elif section == "HABILIDADES":
            if ":" in stripped:
                parts = stripped.split(":", 1)
                data["habilidades"].append((parts[0].strip(), parts[1].strip()))

        elif section == "IDIOMAS":
            m = re.match(r"^[-*]\s*(.+)", stripped)
            if m:
                data["idiomas"].append(m.group(1).strip())

    if section == "RESUMEN PROFESIONAL" and buffer:
        data["resumen"] = " ".join(buffer)
    if section == "EXPERIENCIA PROFESIONAL":
        flush_job()
    return data


def parse_data(filepath):
    try:
        with open(filepath, "r", encoding="utf-8-sig") as f:
            text = f.read()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="latin-1") as f:
            text = f.read()
    return parse_text(text)
