#!/usr/bin/env python3
"""Genera el documento exigido para el Tablero de Anotaciones del Aula Virtual.
Capturas reales del sitio publicado como archivos HTML del repositorio, mediante
servidor HTTP local + navegador Chromium. No simula ni redibuja el contenido web.
"""
from __future__ import annotations
import os, re, io, threading, time
from pathlib import Path
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from datetime import date
from PIL import Image
from playwright.sync_api import sync_playwright
from docx import Document
from docx.shared import Inches, Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION_START
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/"entregas"/"Documento_Evidencias_Portafolio_UDES.docx"
ASSETS=ROOT/"entregas"/"_qa_capturas"
ASSETS.mkdir(parents=True,exist_ok=True)
PUBLIC="https://jvergara26.github.io/portafolio-evaluacion-red-udes/"
GREEN=RGBColor(21,57,46)
GOLD=RGBColor(163,113,48)
DARK=RGBColor(40,49,45)
MUTED=RGBColor(94,102,98)
SECTIONS=[
    ("Página 1. Inicio","index.html",
     "Presenta el título del portafolio, la identificación académica, la pregunta orientadora y el contexto de grado sexto.",
     [("","Portada del sitio"),("#enfoque","Contexto y ficha del proyecto")]),
    ("Página 2. Modelos de evaluación y RED","modelos.html",
     "Organiza seis modelos de evaluación de recursos educativos digitales y explica sus posibilidades de uso.",
     [("","Presentación general de los modelos"),("#lori","Modelos LORI, COdA y UNE"),("#reeves","Modelos complementarios y referentes")]),
    ("Subpágina 2.1. Seleccionando modelos","seleccionando-modelos.html",
     "Documenta la selección y comparación de LORI, COdA y UNE 71362:2017, con instrumentos de valoración y exposición audiovisual.",
     [("","Presentación de la selección"),("#exposicion-audiovisual","Exposición audiovisual"),("#eval-lori-resumen","Resultados e instrumentos de evaluación")]),
    ("Subpágina 2.2. Recursos educativos digitales","recursos-educativos.html",
     "Caracteriza los recursos «Media: Distribuye y Equilibra» (PhET) y «Explorando la media aritmética» (INTEF/GeoGebra).",
     [("","Presentación y propósito de los recursos"),("#red-phet","Recurso 1: PhET"),("#red-intef","Recurso 2: INTEF/GeoGebra")]),
    ("Página 3. Rediseño","rediseno.html",
     "Expone el modelo contextualizado COdA-6: introducción, problema, objetivos, fundamentación, metodología, escala y 40 indicadores.",
     [("","Presentación del rediseño"),("#objetivos","Objetivos"),("#metodologia","Metodología de COdA-6"),("#instrumento","Instrumento contextualizado")]),
    ("Página 4. Aplicación de modelos de evaluación","aplicacion-modelo.html",
     "Presenta la aplicación de COdA-6, las evidencias, las matrices y los resultados globales de ambos recursos.",
     [("","Presentación interactiva y anexos"),("#resultados","Resultados comparativos"),("#preguntas","Interpretación académica"),("#conclusiones","Conclusiones de la aplicación")]),
    ("Página 5. Referencias","referencias.html",
     "Consolida las fuentes bibliográficas y las direcciones de consulta utilizadas para sustentar los modelos y recursos evaluados.",
     [("","Presentación del apartado de referencias"),("#lista-referencias","Lista bibliográfica")]),
]

class QuietHandler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs):
        super().__init__(*args,directory=str(ROOT),**kwargs)
    def log_message(self, format, *args):
        pass

def capture_all():
    server=ThreadingHTTPServer(("127.0.0.1",0),QuietHandler)
    thread=threading.Thread(target=server.serve_forever,daemon=True)
    thread.start()
    base=f"http://127.0.0.1:{server.server_port}/"
    records=[]
    try:
        with sync_playwright() as p:
            browser=p.chromium.launch(headless=True,args=["--no-sandbox","--disable-dev-shm-usage"])
            ctx=browser.new_context(viewport={"width":1440,"height":900},device_scale_factor=1)
            page=ctx.new_page()
            page.set_default_navigation_timeout(60000)
            for idx,(heading,path,description,views) in enumerate(SECTIONS,1):
                url=base+path
                page.goto(url,wait_until="domcontentloaded")
                page.wait_for_timeout(550)
                actual_title=page.title()
                if not actual_title:
                    raise RuntimeError(f"No se cargó correctamente {path}")
                section_record={"heading":heading,"path":path,"description":description,"url":PUBLIC+path,"shots":[]}
                for j,(anchor,label) in enumerate(views,1):
                    if anchor:
                        target=page.locator(anchor)
                        if target.count()<1:
                            raise RuntimeError(f"Ancla no encontrada: {path}{anchor}")
                        target.first.evaluate("""e => {const y=e.getBoundingClientRect().top + window.scrollY - 165; window.scrollTo(0,Math.max(0,y));}""")
                    else:
                        page.evaluate("window.scrollTo(0,0)")
                    if path=="rediseno.html" and anchor=="#instrumento":
                        # Mostrar una muestra real de los criterios desplegables.
                        page.eval_on_selector_all("#instrumento details", "(arr)=>arr.slice(0,2).forEach(d=>d.open=true)")
                    page.wait_for_timeout(350)
                    location=ASSETS/f"{idx:02}_{j:02}.png"
                    page.screenshot(path=str(location),animations="disabled",full_page=False)
                    with Image.open(location) as im:
                        if im.width!=1440 or im.height!=900:
                            raise RuntimeError(f"Captura con dimensiones inesperadas: {location} {im.size}")
                    section_record["shots"].append((str(location),label))
                    print(f"CAPTURA {idx}/7 · {path} {anchor or '(inicio)'} · {location.name}",flush=True)
                records.append(section_record)
            browser.close()
    finally:
        server.shutdown()
    return records

def set_cell_fill(cell, fill):
    tcPr=cell._tc.get_or_add_tcPr()
    shd=OxmlElement("w:shd")
    shd.set(qn("w:fill"),fill)
    tcPr.append(shd)

def hyperlink(paragraph,text,url):
    h=OxmlElement("w:hyperlink")
    h.set(qn("r:id"),paragraph.part.relate_to(url,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",is_external=True))
    r=OxmlElement("w:r")
    pr=OxmlElement("w:rPr")
    color=OxmlElement("w:color");color.set(qn("w:val"),"1C654E")
    pr.append(color);r.append(pr)
    t=OxmlElement("w:t");t.text=text;r.append(t);h.append(r)
    paragraph._p.append(h)

def p(doc,text="",style=None,after=4):
    para=doc.add_paragraph(style=style)
    if text:
        para.add_run(text)
    para.paragraph_format.space_after=Pt(after)
    return para

def make_document(records):
    doc=Document()
    sec=doc.sections[0]
    sec.page_width=Cm(21.59)
    sec.page_height=Cm(27.94)
    sec.top_margin=Cm(1.9)
    sec.bottom_margin=Cm(1.7)
    sec.left_margin=Cm(2.15)
    sec.right_margin=Cm(2.15)
    sec.header_distance=Cm(0.8)
    sec.footer_distance=Cm(0.85)
    styles=doc.styles
    normal=styles["Normal"]
    normal.font.name="Cambria";normal.font.size=Pt(10.5);normal.font.color.rgb=DARK
    normal.paragraph_format.line_spacing=1.2
    for name,size,color in [("Title",22,GREEN),("Heading 1",16,GREEN),("Heading 2",12,GREEN)]:
        st=styles[name];st.font.name="Georgia";st.font.size=Pt(size);st.font.bold=(name!="Title");st.font.color.rgb=color
        st.paragraph_format.space_before=Pt(7);st.paragraph_format.space_after=Pt(7)
    header=sec.header.paragraphs[0]
    header.text="UDES  |  EVALUACIÓN DE RECURSOS EDUCATIVOS DIGITALES"
    header.style="Caption"
    header.runs[0].font.color.rgb=MUTED
    header.runs[0].font.size=Pt(8)
    footer=sec.footer.paragraphs[0]
    footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("PORTAFOLIO ACADÉMICO · 2026   |   ")
    field=OxmlElement("w:fldSimple");field.set(qn("w:instr"),"PAGE")
    footer._p.append(field)
    for r in footer.runs:
        r.font.size=Pt(8);r.font.color.rgb=MUTED

    # Portada
    title=p(doc,after=8);title.alignment=WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_before=Pt(65)
    rr=title.add_run("Universidad de Santander\nCentro Virtual CVUDES")
    rr.font.name="Georgia";rr.font.size=Pt(16);rr.bold=True;rr.font.color.rgb=GREEN
    p(doc,"Especialización en Aplicación de TIC para la Enseñanza",after=12).alignment=WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph("")
    title2=p(doc,after=14);title2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    run=title2.add_run("Evidencias de publicación\ndel portafolio web")
    run.bold=True;run.font.name="Georgia";run.font.size=Pt(24);run.font.color.rgb=GREEN
    sub=p(doc,"Evaluación de Recursos Educativos Digitales",after=24)
    sub.alignment=WD_ALIGN_PARAGRAPH.CENTER
    sub.runs[0].italic=True
    table=doc.add_table(rows=4,cols=2)
    table.autofit=False
    labels=[
       ("Estudiante","John Fernando Vergara Arroyo"),
       ("Docente","Hugo Ferdinand Burgos Bareño"),
       ("Programa","Especialización en Aplicación de TIC para la Enseñanza"),
       ("Lugar y fecha","Montería, Córdoba, Colombia · 21 de noviembre de 2026")
    ]
    for i,(lab,val) in enumerate(labels):
        for c in table.rows[i].cells:
            set_cell_fill(c,"F6F5F0" if i%2==0 else "FFFFFF")
            c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
        table.cell(i,0).text=lab
        table.cell(i,1).text=val
        for run in table.cell(i,0).paragraphs[0].runs:
            run.bold=True;run.font.color.rgb=GREEN
    doc.add_page_break()
    
    doc.add_heading("Enlace de acceso y estructura del portafolio",1)
    p(doc,"El portafolio digital reúne cinco páginas principales y dos subpáginas. Las capturas que se presentan corresponden a vistas reales de los archivos publicados, obtenidas mediante navegador Chromium en una copia local del sitio web, sin reconstruir ni alterar su contenido.")
    para=p(doc,"Enlace público:",after=5)
    hyperlink(para,PUBLIC,PUBLIC)
    p(doc,"La estructura del sitio sigue la organización exigida en la actividad:",after=7)
    for item in records:
        line=p(doc,after=4)
        bold=line.add_run(item["heading"]+": ");bold.bold=True;bold.font.color.rgb=GREEN
        hyperlink(line,item["path"],item["url"])
    p(doc,"Las capturas documentan la navegación, los títulos, la distribución de contenido y secciones representativas de cada página. Para consultar los desarrollos completos, se incluyen enlaces directos a los respectivos apartados.",after=3)
    
    for k,item in enumerate(records):
        doc.add_page_break()
        doc.add_heading(item["heading"],1)
        line=p(doc,after=5)
        line.add_run("Acceso: ").bold=True
        hyperlink(line,item["url"],item["url"])
        p(doc,item["description"],after=8)
        for j,(png,label) in enumerate(item["shots"],1):
            cap=p(doc,after=4)
            run=cap.add_run(f"Figura {sum(len(x['shots']) for x in records[:k])+j}. {label}.")
            run.bold=True;run.font.color.rgb=GREEN;run.font.size=Pt(9)
            pic=p(doc,after=8);pic.alignment=WD_ALIGN_PARAGRAPH.CENTER
            pic.add_run().add_picture(png,width=Cm(16.4))
        note=p(doc,"Nota. Capturas de pantalla del portafolio digital desarrollado para el curso Evaluación de Recursos Educativos Digitales.",after=1)
        note.runs[0].italic=True;note.runs[0].font.size=Pt(8)
    
    doc.add_page_break()
    doc.add_heading("Conclusiones",1)
    paragraphs=[
      "La publicación del portafolio permitió organizar en una estructura navegable los modelos de evaluación de recursos educativos digitales, la selección de instrumentos, la caracterización de los materiales y la aplicación de un modelo contextualizado para la enseñanza de la media aritmética en grado sexto.",
      "La organización en cinco páginas y dos subpáginas facilita localizar el fundamento conceptual, la evaluación de los dos recursos y las evidencias asociadas. La presentación interactiva y sus anexos amplían los registros del modelo COdA-6 y hacen posible consultar las valoraciones de forma estructurada.",
      "La aplicación de COdA-6 otorgó 5,00/5 a «Media: Distribuye y Equilibra» (PhET) y 4,80/5 a «Explorando la media aritmética» (INTEF/GeoGebra). Estos resultados describen la calidad potencial de los materiales según los indicadores evaluados; no representan una medición del impacto en los aprendizajes de los estudiantes.",
      "El portafolio ofrece un soporte verificable para presentar los productos académicos y para orientar futuras decisiones docentes sobre recursos digitales. Su valor reside en la articulación de criterios pedagógicos, evidencias documentales y condiciones de uso pertinentes para el contexto escolar."
    ]
    for t in paragraphs:
        para=p(doc,t,after=10)
        para.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    
    doc.add_heading("Referencias bibliográficas",1)
    references=[
     ("Fernández-Pampillón Cesteros, A. M., Domínguez Romero, E., & de Armas Ranero, I. (2012). COdA: Herramienta de evaluación de la calidad de objetos de aprendizaje. Guía para la producción y evaluación de materiales didácticos digitales (versión 1.1). Universidad Complutense de Madrid.","https://eprints.ucm.es/12533/1/COdAv1_1_07jul2012.pdf"),
     ("Asociación Española de Normalización. (2017). UNE 71362:2017. Calidad de los materiales educativos digitales: Anexo F. Perfiles de aplicación.","https://intef.es/wp-content/uploads/2020/01/AnexoF_UNE_713622017.pdf"),
     ("PhET Interactive Simulations. (s. f.). Media: Distribuye y Equilibra [Simulación interactiva]. Universidad de Colorado Boulder.","https://phet.colorado.edu/es/simulations/mean-share-and-balance"),
     ("Baeza Rojas, O. (s. f.). Explorando la media aritmética [Applet de GeoGebra]. GeoGebra.","https://www.geogebra.org/m/WghmKJhY"),
     ("Instituto Nacional de Tecnologías Educativas y de Formación del Profesorado. (s. f.). Explorando la media aritmética [Recurso educativo digital].","https://descargas.intef.es/recursos_educativos/geogebra/Primaria/P60066/index.html"),
     ("Vergara Arroyo, J. F. (2026). Portafolio académico de evaluación de recursos educativos digitales [Sitio web]. GitHub Pages.",PUBLIC),
    ]
    for t,url in references:
        para=p(doc,after=9)
        para.paragraph_format.left_indent=Cm(.6)
        para.paragraph_format.first_line_indent=Cm(-.6)
        para.add_run(t+" ")
        hyperlink(para,url,url)
    doc.core_properties.title="Evidencias del portafolio web – Evaluación de Recursos Educativos Digitales"
    doc.core_properties.author="John Fernando Vergara Arroyo"
    OUT.parent.mkdir(parents=True,exist_ok=True)
    doc.save(OUT)
    check=Document(OUT)
    screenshot_count=len(check.inline_shapes)
    expected=sum(len(a["shots"]) for a in records)
    if screenshot_count!=expected:
        raise RuntimeError(f"El Word contiene {screenshot_count} capturas, se esperaban {expected}.")
    if OUT.stat().st_size<150000:
        raise RuntimeError("El archivo es demasiado pequeño; podrían faltar las capturas.")
    print(f"VERIFICADO: {OUT.name}, {OUT.stat().st_size:,} bytes; {screenshot_count} capturas reales; 7 secciones y conclusiones/referencias.",flush=True)

if __name__=="__main__":
    captures=capture_all()
    make_document(captures)
