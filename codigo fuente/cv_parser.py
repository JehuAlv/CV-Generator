import os
import re
import unicodedata
from cv_templates import TEMPLATES


def _strip_accents(text):
    nfkd = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nfkd if not unicodedata.combining(c))


def pre_split_sections(text):
    keywords = [
        "DATOS PERSONALES", "INFORMACION PERSONAL",
        "RESUMEN PROFESIONAL", "PERFIL PROFESIONAL", "RESUMEN EJECUTIVO",
        "OBJETIVO PROFESIONAL", "ACERCA DE MI", "SOBRE MI",
        "LOGROS CLAVE", "LOGROS PRINCIPALES", "LOGROS DESTACADOS", "LOGROS",
        "EXPERIENCIA PROFESIONAL", "EXPERIENCIA LABORAL", "EXPERIENCIA DE TRABAJO",
        "HISTORIAL LABORAL", "TRAYECTORIA PROFESIONAL", "EXPERIENCIA",
        "EDUCACION", "FORMACION ACADEMICA", "FORMACION", "ESTUDIOS",
        "CURSOS Y CERTIFICACIONES", "CERTIFICACIONES Y CURSOS",
        "CERTIFICACIONES", "CURSOS", "CAPACITACION",
        "HABILIDADES", "COMPETENCIAS", "APTITUDES", "SKILLS",
        "IDIOMAS", "LENGUAS",
        "PERSONAL INFORMATION", "PROFESSIONAL SUMMARY", "KEY ACHIEVEMENTS",
        "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE",
        "EDUCATION", "CERTIFICATIONS & COURSES", "CERTIFICATIONS",
        "LANGUAGES",
        "NOMBRE:", "CONTACTO:", "TITULO:", "TEMA:",
        "PUESTO:", "EMPRESA:", "FECHAS:", "CARGO:", "PERIODO:",
        "CARRERA:", "ESCUELA:", "INSTITUCION:", "UNIVERSIDAD:",
        "TELEFONO:", "EMAIL:", "CORREO:",
        "NAME:", "CONTACT:", "TITLE:",
        "POSITION:", "ROLE:", "COMPANY:", "DATES:",
        "DEGREE:", "SCHOOL:", "UNIVERSITY:",
    ]

    for kw in keywords:
        pattern = r'(?<!\n)(?=' + re.escape(kw) + r')'
        upper = _strip_accents(text).upper()
        new_text = []
        last = 0
        for m in re.finditer(pattern, upper):
            pos = m.start()
            if pos > 0 and text[pos - 1] != '\n':
                new_text.append(text[last:pos])
                new_text.append('\n')
                last = pos
        new_text.append(text[last:])
        text = "".join(new_text)

    return text


def normalize_text(text):
    text = text.lstrip('﻿')
    replacements = {
        "–": "-", "—": "-", "−": "-",
        "‐": "-", "‑": "-", "‒": "-", "―": "-",
        "•": "-", "●": "-", "▪": "-", "■": "-",
        "‣": "-", "⁃": "-",
        " ": " ", " ": " ", " ": " ", " ": " ",
        " ": " ", " ": " ", " ": " ", " ": " ",
        " ": " ", " ": " ",
        "​": "", "‌": "", "‍": "", "﻿": "",
        "\t": " ",
        "“": '"', "”": '"',
        "‘": "'", "’": "'", "‚": "'", "„": '"',
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
    up = _strip_accents(re.sub(r"\s+", " ", text.strip())).upper()
    up_clean = up.rstrip(": ")
    for p in prefixes:
        p_clean = _strip_accents(p).upper().rstrip(": ")
        if up_clean.startswith(p_clean):
            return True
    return False


def _strip_bullet(line):
    m = re.match(r"^(?:[-*•]|\d{1,2}[.)]\s*-?\s*)\s*(.+)", line)
    if m:
        return m.group(1).strip()
    return line.strip()


_SECTION_MAP = {}
for _canonical, _variants in {
    "DATOS PERSONALES": [
        "DATOS PERSONALES", "INFORMACION PERSONAL",
        "PERSONAL INFORMATION",
    ],
    "RESUMEN PROFESIONAL": [
        "RESUMEN PROFESIONAL", "PERFIL PROFESIONAL", "RESUMEN EJECUTIVO",
        "OBJETIVO PROFESIONAL", "ACERCA DE MI", "SOBRE MI",
        "PROFESSIONAL SUMMARY",
    ],
    "LOGROS CLAVE": [
        "LOGROS CLAVE", "LOGROS PRINCIPALES", "LOGROS DESTACADOS", "LOGROS",
        "KEY ACHIEVEMENTS",
    ],
    "EXPERIENCIA PROFESIONAL": [
        "EXPERIENCIA PROFESIONAL", "EXPERIENCIA LABORAL",
        "EXPERIENCIA DE TRABAJO", "HISTORIAL LABORAL",
        "TRAYECTORIA PROFESIONAL", "EXPERIENCIA",
        "PROFESSIONAL EXPERIENCE", "WORK EXPERIENCE",
    ],
    "EDUCACION": [
        "EDUCACION", "FORMACION ACADEMICA", "FORMACION", "ESTUDIOS",
        "EDUCATION",
    ],
    "CURSOS Y CERTIFICACIONES": [
        "CURSOS Y CERTIFICACIONES", "CERTIFICACIONES Y CURSOS",
        "CERTIFICACIONES", "CURSOS", "CAPACITACION",
        "CERTIFICATIONS & COURSES", "CERTIFICATIONS",
    ],
    "HABILIDADES": [
        "HABILIDADES", "COMPETENCIAS", "APTITUDES",
        "SKILLS",
    ],
    "IDIOMAS": [
        "IDIOMAS", "LENGUAS",
        "LANGUAGES",
    ],
}.items():
    for _v in _variants:
        _SECTION_MAP[_v] = _canonical


def _match_section(line):
    clean = _strip_accents(line.strip()).upper()
    clean = re.sub(r"[:\s]+$", "", clean)
    return _SECTION_MAP.get(clean)


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

    def flush_edu():
        nonlocal current_edu
        if current_edu and current_edu.get("titulo"):
            data["educacion"].append(current_edu)
        current_edu = None

    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("====") or line.startswith("----"):
            continue

        matched_section = _match_section(line)
        if matched_section:
            if section == "RESUMEN PROFESIONAL" and buffer:
                data["resumen"] = " ".join(buffer)
                buffer = []
            if section == "EXPERIENCIA PROFESIONAL":
                flush_job()
            if section == "EDUCACION":
                flush_edu()
            section = matched_section
            continue

        stripped = line.strip()
        if not stripped:
            continue

        if _starts(stripped, "TEMA:", "THEME:"):
            val = stripped.split(":", 1)[1].strip().lower()
            if val in TEMPLATES:
                data["theme"] = val
            continue

        if section == "DATOS PERSONALES":
            if _starts(stripped, "NOMBRE:", "NAME:"):
                data["nombre"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "CONTACTO:", "CONTACT:", "TELEFONO:", "EMAIL:", "CORREO:"):
                val = stripped.split(":", 1)[1].strip()
                if data["contacto"]:
                    data["contacto"] += " | " + val
                else:
                    data["contacto"] = val
            elif _starts(stripped, "TITULO:", "TÍTULO:", "TITLE:", "PUESTO:", "POSICION:"):
                data["titulo"] = stripped.split(":", 1)[1].strip()

        elif section == "RESUMEN PROFESIONAL":
            buffer.append(stripped)

        elif section == "LOGROS CLAVE":
            data["logros"].append(_strip_bullet(stripped))

        elif section == "EXPERIENCIA PROFESIONAL":
            if _starts(stripped, "PUESTO:", "POSITION:", "ROLE:", "CARGO:"):
                flush_job()
                current_job = {"puesto": "", "empresa": "", "fechas": "", "bullets": []}
                current_job["puesto"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "EMPRESA:", "COMPANY:", "ORGANIZACION:", "COMPAÑIA:", "COMPANIA:"):
                if not current_job:
                    current_job = {"puesto": "", "empresa": "", "fechas": "", "bullets": []}
                current_job["empresa"] = stripped.split(":", 1)[1].strip()
            elif _starts(stripped, "FECHAS:", "DATES:", "PERIODO:", "FECHA:"):
                if not current_job:
                    current_job = {"puesto": "", "empresa": "", "fechas": "", "bullets": []}
                current_job["fechas"] = stripped.split(":", 1)[1].strip()
            elif current_job:
                current_job["bullets"].append(_strip_bullet(stripped))

        elif section == "EDUCACION":
            if _starts(stripped, "CARRERA:", "DEGREE:", "TITULO:", "TÍTULO:", "GRADO:"):
                flush_edu()
                current_edu = {"titulo": stripped.split(":", 1)[1].strip(), "institucion": ""}
            elif _starts(stripped, "ESCUELA:", "SCHOOL:", "UNIVERSITY:", "UNIVERSIDAD:",
                         "INSTITUCION:", "INSTITUCIÓN:", "CENTRO:"):
                if not current_edu:
                    current_edu = {"titulo": "", "institucion": ""}
                current_edu["institucion"] = stripped.split(":", 1)[1].strip()
                if not current_edu["titulo"]:
                    flush_edu()

        elif section == "CURSOS Y CERTIFICACIONES":
            data["cursos"].append(_strip_bullet(stripped))

        elif section == "HABILIDADES":
            if ":" in stripped:
                parts = stripped.split(":", 1)
                val = parts[1].strip()
                if val and val.upper() != "N/A":
                    data["habilidades"].append((parts[0].strip(), val))

        elif section == "IDIOMAS":
            data["idiomas"].append(_strip_bullet(stripped))

    if section == "RESUMEN PROFESIONAL" and buffer:
        data["resumen"] = " ".join(buffer)
    if section == "EXPERIENCIA PROFESIONAL":
        flush_job()
    if section == "EDUCACION":
        flush_edu()
    return data


def parse_data(filepath):
    try:
        with open(filepath, "r", encoding="utf-8-sig") as f:
            text = f.read()
    except UnicodeDecodeError:
        with open(filepath, "r", encoding="latin-1") as f:
            text = f.read()
    return parse_text(text)
