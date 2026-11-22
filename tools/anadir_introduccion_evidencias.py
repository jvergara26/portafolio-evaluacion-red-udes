#!/usr/bin/env python3
"""Añadir exclusivamente una introducción al documento de evidencias UDES.

Preserva todas las capturas, tablas, enlaces, textos y páginas posteriores.
"""
from collections import Counter
from hashlib import sha256
from pathlib import Path
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "entregas/Documento_Evidencias_Portafolio_UDES_Corregido.docx"
DST = ROOT / "entregas/Documento_Evidencias_Portafolio_UDES_Final.docx"

INTRO = [
    "La evaluación de recursos educativos digitales exige considerar tanto su pertinencia pedagógica como las condiciones tecnológicas que permiten utilizarlos en el aula. En el curso Evaluación de Recursos Educativos Digitales, de la Especialización en Aplicación de TIC para la Enseñanza de la Universidad de Santander, se desarrolló un portafolio web para organizar, presentar y sustentar el estudio de modelos de evaluación y su aplicación a materiales destinados a la enseñanza de las Matemáticas.",
    "El portafolio está organizado en cinco páginas principales y dos subpáginas. Presenta los modelos de evaluación de recursos educativos digitales, la selección de instrumentos, la caracterización de dos recursos sobre media aritmética, el rediseño del modelo COdA mediante cuarenta indicadores contextualizados y la aplicación del instrumento COdA-6. La organización permite seguir la relación entre los referentes teóricos, los criterios de calidad y los resultados obtenidos para PhET e INTEF/GeoGebra.",
    "Este documento reúne el enlace de acceso al sitio y capturas de pantalla de sus páginas y subpáginas, acompañadas de descripciones que identifican los contenidos representados. Su finalidad es evidenciar la estructura y el desarrollo de la actividad solicitada en el Aula Virtual. Las conclusiones y referencias bibliográficas complementan el recorrido visual; la consulta íntegra de los recursos, matrices y anexos se realiza mediante los enlaces del portafolio."
]

def image_digests(doc):
    return sorted(
        sha256(rel.target_part.blob).hexdigest()
        for rel in doc.part.rels.values()
        if "image" in rel.reltype
    )

def hyperlink_targets(doc):
    return sorted(
        rel.target_ref for rel in doc.part.rels.values()
        if "hyperlink" in rel.reltype
    )

def main():
    if not SRC.is_file():
        raise FileNotFoundError(SRC)
    doc = Document(SRC)
    original = Document(SRC)
    target = next((p for p in doc.paragraphs if p.text.strip()=="Acceso y estructura del portafolio"),None)
    if target is None:
        raise RuntimeError("No se encontró el inicio de la sección de acceso.")
    if any(p.text.strip()=="Introducción" for p in doc.paragraphs):
        raise RuntimeError("El documento ya tiene una introducción; se evitó duplicarla.")

    # La portada termina con un salto de página; insertar entre ese salto y 'Acceso'.
    heading = doc.add_heading("Introducción",level=1)
    target._p.addprevious(heading._p)
    for txt in INTRO:
        p=doc.add_paragraph(txt)
        p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent=Cm(.7)
        p.paragraph_format.line_spacing=1.35
        p.paragraph_format.space_after=Pt(13)
        target._p.addprevious(p._p)
    sep=doc.add_page_break()
    target._p.addprevious(sep._p)

    doc.core_properties.title="Evidencias del portafolio web UDES — documento final con introducción"
    DST.parent.mkdir(parents=True,exist_ok=True)
    doc.save(DST)

    verified=Document(DST)
    texts=[p.text.strip() for p in verified.paragraphs]
    i_intro=texts.index("Introducción")
    i_access=texts.index("Acceso y estructura del portafolio")
    if i_intro>=i_access or texts[i_intro+1:i_access].count(INTRO[0])!=1:
        raise RuntimeError("Ubicación incorrecta de la introducción.")
    if len(verified.inline_shapes)!=len(original.inline_shapes) or len(verified.inline_shapes)!=26:
        raise RuntimeError("Se modificó el número de capturas originales.")
    if image_digests(verified)!=image_digests(original):
        raise RuntimeError("Se modificaron los bytes de una o más capturas.")
    if hyperlink_targets(verified)!=hyperlink_targets(original):
        raise RuntimeError("Se alteraron los enlaces del documento.")
    if len(verified.tables)!=len(original.tables):
        raise RuntimeError("Se alteraron tablas ajenas a la introducción.")
    for section in ["Página 1. Inicio","Página 2. Modelos de evaluación y RED",
                    "Subpágina 2.1. Seleccionando modelos","Subpágina 2.2. Recursos educativos digitales",
                    "Página 3. Rediseño","Página 4. Aplicación de modelos de evaluación",
                    "Página 5. Referencias","Conclusiones","Referencias bibliográficas"]:
        if not any(section in v for v in texts):
            raise RuntimeError("Se perdió una sección: "+section)
    # La secuencia original de párrafos se mantiene: se agregaron solo 5 párrafos nuevos.
    old_paragraphs=[p.text for p in original.paragraphs]
    new_without_intro=[p.text for p in verified.paragraphs]
    start=new_without_intro.index("Introducción")
    del new_without_intro[start:start+5]  # título, tres párrafos, salto de página
    if new_without_intro!=old_paragraphs:
        raise RuntimeError("Se detectó alteración de párrafos existentes.")
    print(f"DOCUMENTO_FINAL_OK={DST.name}")
    print(f"CAPTURAS_INTACTAS={len(verified.inline_shapes)}")
    print(f"PARRAFOS_INTRODUCCION={len(INTRO)}")
    print(f"ENLACES_INTACTOS={len(hyperlink_targets(verified))}")
    print(f"BYTES={DST.stat().st_size}")

if __name__=="__main__":
    main()
