#!/usr/bin/env python3
"""Exportar presentación COdA-6 (37 diapositivas) sin alterar el diseño aprobado.

Los gráficos de cada diapositiva se insertan como imagen completa para
garantizar fidelidad visual; el texto de la transcripción se añade en las notas
del PowerPoint. Las figuras incrustadas de las diapositivas modificadas
se renderizan desde los SVG integrados.
"""
from __future__ import annotations

import io
import json
import re
from pathlib import Path

import cairosvg
from PIL import Image
from pptx import Presentation
from pptx.util import Inches
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parent.parent
HTML = ROOT / "presentacion-coda6.html"
PPTX = ROOT / "COdA6_Informe_Implementacion_UDES.pptx"
PDF = ROOT / "COdA6_Informe_Implementacion_UDES.pdf"
SLIDE_IDS_CORREGIDOS = {10, 11, 12, 16, 24, 25, 27, 34, 35, 37}
N_SLIDES = 37
W, H = 1334, 750  # Diseño de fuente, proporción 16:9 aprox.
PAGE_W, PAGE_H = Inches(13.34), Inches(7.50)


def load_slides():
    html = HTML.read_text(encoding="utf-8")
    match = re.search(r"const slides=(\[.*?\]);", html, flags=re.S)
    if match is None:
        raise RuntimeError("No se encontró el listado de diapositivas en el HTML.")
    slides = json.loads(match.group(1))
    if len(slides) != N_SLIDES:
        raise RuntimeError(f"Se esperaban {N_SLIDES} diapositivas; hay {len(slides)}.")
    for number, item in enumerate(slides, 1):
        asset = ROOT / item["src"]
        if not asset.is_file():
            raise FileNotFoundError(f"Diapositiva {number}: {asset}")
        if number in SLIDE_IDS_CORREGIDOS and not str(asset).endswith("-integrada.svg"):
            raise RuntimeError(f"Diapositiva {number} no usa la versión visual aprobada.")
    return slides


def render_slide(asset):
    if asset.suffix.lower() == ".svg":
        raster = cairosvg.svg2png(
            url=str(asset),
            output_width=W,
            output_height=H,
        )
        im = Image.open(io.BytesIO(raster))
    else:
        im = Image.open(asset)
    im = im.convert("RGB")
    if im.size != (W, H):
        im = im.resize((W, H), Image.Resampling.LANCZOS)
    return im


def main():
    slides = load_slides()
    ppt = Presentation()
    ppt.slide_width, ppt.slide_height = PAGE_W, PAGE_H
    blank = ppt.slide_layouts[6]

    pdf_w = 13.34 * 72
    pdf_h = 7.50 * 72
    pdf = canvas.Canvas(str(PDF), pagesize=(pdf_w, pdf_h), pageCompression=1)
    pdf.setTitle("COdA-6 - Informe de implementación - UDES")
    pdf.setAuthor("John Fernando Vergara Arroyo")
    pdf.setSubject("Aplicación contextualizada de COdA-6 a PhET e INTEF/GeoGebra")

    for number, item in enumerate(slides, 1):
        path = ROOT / item["src"]
        img = render_slide(path)
        buf = io.BytesIO()
        # PNG sin pérdidas para mantener exactamente la apariencia aprobada
        img.save(buf, format="PNG", optimize=True)
        png = buf.getvalue()

        sl = ppt.slides.add_slide(blank)
        picture = sl.shapes.add_picture(io.BytesIO(png), 0, 0, width=PAGE_W, height=PAGE_H)
        try:
            picture._element._nvXxPr.cNvPr.set("descr", item.get("title", f"Diapositiva {number}"))
        except (AttributeError, TypeError):
            pass
        try:
            notes = sl.notes_slide.notes_text_frame
            if notes is not None:
                notes.text = item.get("transcript", "")
        except (AttributeError, TypeError):
            pass

        pdf.drawImage(ImageReader(img), 0, 0, width=pdf_w, height=pdf_h)
        pdf.showPage()
        print(f"[{number:02d}/{N_SLIDES}] {path.relative_to(ROOT)}")

    ppt.save(PPTX)
    pdf.save()

    # Preflight completo de los dos descargables.
    ppt_verified = Presentation(PPTX)
    pdf_verified = PdfReader(str(PDF))
    if len(ppt_verified.slides) != N_SLIDES or len(pdf_verified.pages) != N_SLIDES:
        raise RuntimeError("La exportación no contiene las 37 diapositivas.")
    if PPTX.stat().st_size < 500_000 or PDF.stat().st_size < 500_000:
        raise RuntimeError("Uno de los archivos generados parece incompleto.")
    print(f"LISTO: PPTX {PPTX.stat().st_size:,} bytes; PDF {PDF.stat().st_size:,} bytes")
    print(f"VERIFICADO: {len(ppt_verified.slides)} diapositivas PPTX; {len(pdf_verified.pages)} páginas PDF")


if __name__ == "__main__":
    main()
