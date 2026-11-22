#!/usr/bin/env python3
"""Documento de evidencias UDES: capturas acotadas a bloques de contenido completos.

Ninguna captura es un recorte arbitrario de un viewport desplazado.
Se registra selector, etiqueta y tamaño para reproducibilidad.
Las imágenes se toman con Chromium del HTML/CSS del repositorio publicado.
"""
from __future__ import annotations

import io
import math
import os
import threading
from pathlib import Path
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from datetime import date

from PIL import Image
from playwright.sync_api import sync_playwright
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT=Path(__file__).resolve().parent.parent
OUT=ROOT/"entregas"/"Documento_Evidencias_Portafolio_UDES_Corregido.docx"
SCREEN=ROOT/"entregas"/"_capturas_curadas"
SCREEN.mkdir(parents=True,exist_ok=True)
BASE="https://jvergara26.github.io/portafolio-evaluacion-red-udes/"
PAGES=[
  {"id":"01", "title":"Página 1. Inicio", "path":"index.html",
   "summary":"Identidad académica del portafolio, propósito de la evaluación y contextualización institucional.",
   "shots":[
     ("encabezado","section.hero","Portada, autoría y título de la actividad"),
     ("elemento","#enfoque","Propósito de la evaluación y ficha de contexto")
   ]},
  {"id":"02","title":"Página 2. Modelos de evaluación y RED","path":"modelos.html",
   "summary":"Panorama de seis modelos de evaluación y contraste de sus dimensiones pedagógicas y técnicas.",
   "shots":[
     ("encabezado","section.models-intro","Presentación de los seis modelos"),
     ("elemento","#lori","Modelo LORI: criterios y alcance"),
     ("elemento","#coda","Modelo COdA: dimensión didáctica y tecnológica"),
     ("elemento","#une","Norma UNE 71362:2017"),
     ("elemento","#reeves","Modelo Reeves"),
     ("elemento","#galvis","Modelo Galvis"),
     ("elemento","#furps","Modelo FURPS")
   ]},
  {"id":"03","title":"Subpágina 2.1. Seleccionando modelos","path":"seleccionando-modelos.html",
   "summary":"Fundamentación de la selección de LORI, COdA y UNE 71362:2017; exposición audiovisual y herramientas.",
   "shots":[
     ("encabezado","section.select-hero","Presentación de la selección"),
     ("elemento",".selection-cards","Los tres modelos seleccionados y sus características"),
     ("elemento","#exposicion-audiovisual .selection-video-figure","Video académico y su acceso")
   ]},
  {"id":"04","title":"Subpágina 2.2. Recursos educativos digitales","path":"recursos-educativos.html",
   "summary":"Caracterización de los dos RED empleados para estudiar la media aritmética en sexto grado.",
   "shots":[
     ("encabezado","section.resources-hero","Propósito de selección de los RED"),
     ("elemento","#red-phet .resource-profile-grid","PhET: ficha técnica y visualización del recurso"),
     ("elemento","#red-intef .resource-profile-grid","INTEF/GeoGebra: ficha técnica y visualización")
   ]},
  {"id":"05","title":"Página 3. Rediseño","path":"rediseno.html",
   "summary":"Modelo COdA-6: objetivos, fundamentación, metodología, escala y cuarenta indicadores contextualizados.",
   "shots":[
     ("encabezado","section.c6-hero","Identidad y propósito de COdA-6"),
     ("elemento","#objetivos .c6-header","Objetivo general y objetivos específicos"),
     ("elemento","#metodologia .c6-method","Cinco fases del procedimiento aplicado"),
     ("elemento","#instrumento .c6-criterion:first-child","Ejemplo completo de los indicadores del instrumento")
   ]},
  {"id":"06","title":"Página 4. Aplicación de modelos de evaluación","path":"aplicacion-modelo.html",
   "summary":"Aplicación del modelo contextualizado, puntuaciones de los RED, interpretación y conclusiones.",
   "shots":[
     ("encabezado","section.ca-hero","Presentación de la aplicación"),
     ("elemento","#resultados .ca-trio","Puntuaciones globales y cobertura de COdA-6"),
     ("elemento","#resultados .ca-table","Tabla comparativa completa de diez criterios"),
     ("elemento","#conclusiones .ca-body","Conclusiones de la aplicación")
   ]},
  {"id":"07","title":"Página 5. Referencias","path":"referencias.html",
   "summary":"Fuentes consultadas y bibliografía utilizada en la evaluación de los modelos y RED.",
   "shots":[
     ("encabezado","section.reference-hero","Portada del apartado de referencias"),
     ("multielemento","#lista-referencias .apa-references > li","Primeras referencias bibliográficas (extracto)"),
     ("multielemento_final","#lista-referencias .apa-references > li","Últimas referencias bibliográficas (continuación)")
   ]}
]
class Handler(SimpleHTTPRequestHandler):
    def __init__(self,*args,**kwargs):super().__init__(*args,directory=str(ROOT),**kwargs)
    def log_message(self,*args,**kwargs):pass

def screenshot_exact(page,shotkind,selector,filename,pathname):
    if shotkind=="encabezado":
        target=page.locator(selector).first
        if target.count()<1:raise RuntimeError("No existe "+pathname+" "+selector)
        bbox=target.bounding_box()
        if not bbox:raise RuntimeError("Elemento oculto "+pathname+" "+selector)
        # Inicio exacto de la página hasta el final del bloque de portada;
        # no incluye el comienzo de la sección siguiente.
        bottom=int(bbox["y"]+bbox["height"]+1)
        page.evaluate("window.scrollTo(0,0)")
        png=page.screenshot(clip={"x":0,"y":0,"width":1366,"height":bottom},animations="disabled")
    elif shotkind in ("multielemento","multielemento_final"):
        items=page.locator(selector)
        count=items.count()
        if count<8:raise RuntimeError(f"Solo {count} referencias en {pathname}")
        if shotkind=="multielemento":indices=list(range(0,5))
        else:indices=list(range(max(5,count-5),count))
        boxes=[items.nth(k).bounding_box() for k in indices]
        print("REFERENCIAS_DEBUG",pathname,"indices",indices,"total",count,"boxes",boxes,flush=True)
        if any(b is None for b in boxes):raise RuntimeError("Referencia no visible")
        left=min(b["x"] for b in boxes);right=max(b["x"]+b["width"] for b in boxes)
        top=min(b["y"] for b in boxes);bottom=max(b["y"]+b["height"] for b in boxes)
        # La altura se calcula a partir del borde inferior del último <li>,
        # no desde la ventana de 900 px.
        clip={"x":max(0,left-6),"y":top,"width":min(1366-left+6,right-left+12),"height":bottom-top}
        png=page.screenshot(clip=clip,animations="disabled")
    else:
        target=page.locator(selector).first
        if target.count()<1:raise RuntimeError("Selector inexistente "+pathname+" "+selector)
        if pathname=="rediseno.html" and selector.startswith("#instrumento"):
            target.evaluate("(el)=>el.open=true")
        target.scroll_into_view_if_needed()
        page.wait_for_timeout(190)
        png=target.screenshot(animations="disabled")
    im=Image.open(io.BytesIO(png)).convert("RGB")
    if im.width<280 or im.height<90:raise RuntimeError(f"Captura demasiado pequeña: {pathname} {selector} {im.size}")
    # Evitar escalar una captura demasiado alta en una sola página y volver ilegible el texto.
    # Estas figuras tendrán ancho de 16.3 cm y altura máxima de 15.8 cm en Word.
    if im.height/im.width > 1.6:
        raise RuntimeError(f"Figura desproporcionadamente alta: {pathname}, {selector}, {im.size}")
    im.save(filename,format="JPEG",quality=91,subsampling=0,optimize=True)
    print(f"CAPTURA_OK|{pathname}|{selector}|{im.width}x{im.height}|{filename.name}",flush=True)

def capture():
    server=ThreadingHTTPServer(("127.0.0.1",0),Handler)
    thread=threading.Thread(target=server.serve_forever,daemon=True);thread.start()
    url=f"http://127.0.0.1:{server.server_port}/"
    records=[]
    try:
      with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=True,args=["--no-sandbox","--disable-dev-shm-usage"])
        context=browser.new_context(viewport={"width":1366,"height":900},device_scale_factor=1)
        page=context.new_page()
        for section in PAGES:
            page.goto(url+section["path"],wait_until="domcontentloaded",timeout=60000)
            page.wait_for_timeout(480)
            if not page.title():raise RuntimeError("La página no tiene título: "+section["path"])
            section["captures"]=[]
            for i,(kind,sel,description) in enumerate(section["shots"],1):
                fname=SCREEN/f'{section["id"]}_{i:02}.jpg'
                screenshot_exact(page,kind,sel,fname,section["path"])
                section["captures"].append((fname,description,sel))
        browser.close()
    finally:
      server.shutdown()
    return PAGES

def add_link(p,url,text=None):
    h=OxmlElement('w:hyperlink')
    h.set(qn('r:id'),p.part.relate_to(url,"http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",is_external=True))
    r=OxmlElement("w:r");rp=OxmlElement("w:rPr")
    color=OxmlElement("w:color");color.set(qn("w:val"),"1D644E");rp.append(color);r.append(rp)
    t=OxmlElement("w:t");t.text=text or url;r.append(t);h.append(r);p._p.append(h)

def protect_with_next(p):
    p.paragraph_format.keep_with_next=True

def create_doc(data):
    doc=Document();sec=doc.sections[0];sec.page_width=Cm(21.59);sec.page_height=Cm(27.94)
    sec.left_margin=sec.right_margin=Cm(2.1);sec.top_margin=Cm(2);sec.bottom_margin=Cm(1.7)
    sec.header_distance=Cm(.7);sec.footer_distance=Cm(.8)
    st=doc.styles['Normal'];st.font.name="Cambria";st.font.size=Pt(10.5);st.font.color.rgb=RGBColor(32,51,43)
    st.paragraph_format.space_after=Pt(6)
    for name,sz in [("Heading 1",16),("Heading 2",12)]:
        h=doc.styles[name];h.font.name="Georgia";h.font.size=Pt(sz);h.font.color.rgb=RGBColor(21,57,46)
        h.paragraph_format.space_before=Pt(9);h.paragraph_format.space_after=Pt(9)
    header=sec.header.paragraphs[0];header.text="UDES  ·  EVALUACIÓN DE RECURSOS EDUCATIVOS DIGITALES";header.runs[0].font.size=Pt(8)
    footer=sec.footer.paragraphs[0];footer.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    footer.add_run("PORTAFOLIO ACADÉMICO / 2026  ·  ")
    fld=OxmlElement("w:fldSimple");fld.set(qn("w:instr"),"PAGE");footer._p.append(fld)
    for run in footer.runs:run.font.size=Pt(8)
    # Portada
    pa=doc.add_paragraph();pa.paragraph_format.space_before=Pt(66);pa.alignment=1
    run=pa.add_run("Universidad de Santander\nCentro Virtual CVUDES")
    run.bold=True;run.font.name="Georgia";run.font.size=Pt(17);run.font.color.rgb=RGBColor(21,57,46)
    pa=doc.add_paragraph("Especialización en Aplicación de TIC para la Enseñanza");pa.alignment=1
    doc.add_paragraph("")
    pa=doc.add_paragraph();pa.alignment=1;pa.paragraph_format.space_before=Pt(20)
    run=pa.add_run("Evidencias del portafolio web\nEvaluación de Recursos Educativos Digitales")
    run.bold=True;run.font.name="Georgia";run.font.size=Pt(23);run.font.color.rgb=RGBColor(21,57,46)
    doc.add_paragraph("")
    for label,value in [
        ("Estudiante","John Fernando Vergara Arroyo"),
        ("Docente","Hugo Ferdinand Burgos Bareño"),
        ("Institución","Institución Educativa Isabel La Católica, Montería"),
        ("Fecha","21 de noviembre de 2026")
    ]:
        pa=doc.add_paragraph();rr=pa.add_run(label+": ");rr.bold=True;pa.add_run(value)
    doc.add_page_break()
    doc.add_heading("Acceso y estructura del portafolio",1)
    pa=doc.add_paragraph("Portafolio web publicado: ");add_link(pa,BASE)
    pa=doc.add_paragraph("El sitio contiene cinco páginas principales y dos subpáginas. Las figuras son capturas del navegador sobre los archivos HTML y CSS publicados, tomadas sobre elementos completos de la interfaz. Se priorizaron títulos, fichas de los RED, matrices, resultados y bibliografía, con encuadres definidos por cada bloque.")
    for d in data:
        pa=doc.add_paragraph(style=None)
        rr=pa.add_run(d["title"]+"  ·  ");rr.bold=True
        add_link(pa,BASE+d["path"],d["path"])
    for d in data:
        doc.add_page_break()
        doc.add_heading(d["title"],1)
        pa=doc.add_paragraph("Enlace: ");add_link(pa,BASE+d["path"])
        doc.add_paragraph(d["summary"])
        for ix,(fname,description,sel) in enumerate(d["captures"],1):
            if ix>1:
                # Each screenshot starts in its own clearly labelled section, avoiding
                # collisions between captions, figures and following headings.
                doc.add_page_break()
                doc.add_heading(d["title"]+" — evidencia "+str(ix),2)
            cap=doc.add_paragraph()
            cap.paragraph_format.keep_with_next=True
            rr=cap.add_run(description+".");rr.bold=True;rr.font.color.rgb=RGBColor(21,57,46)
            im=Image.open(fname)
            maxw=Cm(16.7);maxh=Cm(15.8)
            shape=im.width/im.height
            width=min(maxw,maxh*shape)
            height=width/shape
            pa=doc.add_paragraph()
            pa.alignment=WD_ALIGN_PARAGRAPH.CENTER
            pa.paragraph_format.keep_together=True
            pa.add_run().add_picture(str(fname),width=width,height=height)
            note=doc.add_paragraph(f"Nota. Captura de un bloque completo de {d['path']}. Se muestra {description.lower()}.")
            note.runs[0].italic=True;note.runs[0].font.size=Pt(8)
    doc.add_page_break()
    doc.add_heading("Conclusiones",1)
    for txt in [
      "El portafolio web integra en una estructura accesible por páginas los modelos de evaluación de recursos educativos digitales, los instrumentos seleccionados y la caracterización de los materiales de Estadística destinados a grado sexto.",
      "Las evidencias visuales permiten reconocer la organización de las cinco páginas y dos subpáginas exigidas. Las imágenes muestran bloques completos de información representativa —y no fragmentos arbitrarios de la ventana— para facilitar la revisión académica del contenido.",
      "La aplicación del modelo contextualizado COdA-6 permitió valorar la calidad potencial de los dos recursos. PhET alcanzó 5,00/5 e INTEF/GeoGebra 4,80/5, resultados que describen el cumplimiento de indicadores y no constituyen mediciones de aprendizaje.",
      "La publicación digital facilita la consulta de las fuentes, el análisis comparativo y la presentación de evidencias que respaldan la toma de decisiones docentes sobre recursos para Matemáticas."
    ]:
        pa=doc.add_paragraph(txt);pa.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.add_heading("Referencias bibliográficas",1)
    for cite,url in [
       ("Asociación Española de Normalización. (2017). UNE 71362:2017. Calidad de los materiales educativos digitales.","https://intef.es/wp-content/uploads/2020/01/AnexoF_UNE_713622017.pdf"),
       ("Fernández-Pampillón Cesteros, A. M., Domínguez Romero, E., & de Armas Ranero, I. (2012). COdA: Herramienta de evaluación de la calidad de objetos de aprendizaje. Universidad Complutense de Madrid.","https://eprints.ucm.es/12533/1/COdAv1_1_07jul2012.pdf"),
       ("Baeza Rojas, O. (s. f.). Explorando la media aritmética [Applet de GeoGebra]. GeoGebra.","https://www.geogebra.org/m/WghmKJhY"),
       ("PhET Interactive Simulations. (s. f.). Media: Distribuye y Equilibra [Simulación interactiva]. Universidad de Colorado Boulder.","https://phet.colorado.edu/es/simulations/mean-share-and-balance"),
       ("Vergara Arroyo, J. F. (2026). Portafolio académico de evaluación de recursos educativos digitales [Sitio web]. GitHub Pages.",BASE)
    ]:
        pa=doc.add_paragraph();pa.paragraph_format.left_indent=Cm(.6);pa.paragraph_format.first_line_indent=Cm(-.6)
        pa.add_run(cite+" ");add_link(pa,url)
    OUT.parent.mkdir(parents=True,exist_ok=True)
    doc.core_properties.author="John Fernando Vergara Arroyo"
    doc.core_properties.title="Evidencias de páginas y subpáginas del portafolio web UDES — versión corregida"
    doc.save(OUT)
    check=Document(OUT)
    expected=sum(len(d["captures"]) for d in data)
    if len(check.inline_shapes)!=expected:raise RuntimeError(f"Faltan capturas {len(check.inline_shapes)}/{expected}")
    if OUT.stat().st_size<300000:raise RuntimeError("Word pequeño o incompleto")
    print(f"DOCX_OK|{OUT}|{OUT.stat().st_size}|{expected} capturas|7 secciones",flush=True)

if __name__=="__main__":
  create_doc(capture())
