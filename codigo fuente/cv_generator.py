"""
CV Generator - 10 Plantillas ATS Friendly, Auto-fit 1 pagina
Genera PDF (reportlab) + Word (python-docx) desde datos.txt

PLANTILLAS:
  clasico      - Barras de color llenas, nombre izquierda, linea separadora
  ejecutivo    - Nombre centrado, secciones con linea debajo, espaciado elegante
  moderno      - Borde izquierdo de color en secciones, fechas en negrita arriba
  elegante     - Nombre centrado, secciones MAYUSCULAS con linea fina extendida
  profesional  - Secciones con borde inferior grueso, compacto y directo
  creativo     - Punto de color como marcador de seccion, acento rojo, minimalista
  corporativo  - Header oscuro con nombre blanco, secciones con fondo suave
  fresco       - Secciones con etiqueta corta de color, limpio y juvenil
  sobrio       - Ultra minimalista, solo texto bold como seccion, sin adornos
  audaz        - Bloque cuadrado de color + linea fina, fuerte y directo
"""
import sys
import os
import re
import io
from cv_templates import TEMPLATES, LABELS_ES, LABELS_EN
from cv_parser import parse_data, parse_text
from cv_renderer import generate_pdf, generate_word

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_DIR = os.path.join(SCRIPT_DIR, "CVs generados")
DATA_FILE  = os.path.join(SCRIPT_DIR, "datos.txt")


# ============================================================
# MAIN
# ============================================================
if __name__ == "__main__":
    if not os.path.exists(DATA_FILE):
        print(f"ERROR: No se encontro {DATA_FILE}")
        print("Edita el archivo datos.txt con la informacion del candidato.")
        input("\nPresiona Enter para cerrar...")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("  CV GENERATOR - 10 Plantillas ATS (Auto-fit 1 pagina)")
    print("=" * 60)

    data = parse_data(DATA_FILE)

    if not data["nombre"] or data["nombre"] == "NOMBRE COMPLETO DEL CANDIDATO":
        print("\nERROR: Edita datos.txt con la informacion real del candidato.")
        print("Hay un ejemplo en: ejemplos\\datos_ejemplo.txt")
        input("\nPresiona Enter para cerrar...")
        sys.exit(1)

    os.makedirs(OUTPUT_DIR, exist_ok=True)

    safe_name = data["nombre"].replace(" ", "_")
    safe_name = re.sub(r'[^\w_]', '', safe_name)

    docx_path = os.path.join(OUTPUT_DIR, safe_name + "_CV.docx")
    pdf_path  = os.path.join(OUTPUT_DIR, safe_name + "_CV.pdf")

    print(f"\n  Candidato:  {data['nombre']}")
    print(f"  Plantilla:  {data['theme']}")
    print()

    try:
        final_scale = generate_pdf(data, pdf_path)
    except Exception as e:
        print(f"\n  Error generando PDF: {e}")
        print("  Intenta: pip install reportlab")
        final_scale = 1.0

    generate_word(data, docx_path, final_scale)

    print("\n  Archivos en: " + OUTPUT_DIR)
    print("=" * 60)
    input("\nPresiona Enter para cerrar...")
