# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import html
import re
import subprocess
import sys
import zipfile
import xml.etree.ElementTree as ET

# --- 1. ABRIR VENTANA PARA ELEGIR EL ARCHIVO ---
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(
    title="Selecciona un archivo Word de la Revista CICLOS",
    filetypes=[("Documentos de Word", "*.docx")]
)

if not file_path:
    print("No se seleccionó ningún archivo.")
    root.destroy()
    sys.exit()

DOCX = Path(file_path)

# --- 2. CREAR CARPETAS AUTOMÁTICAMENTE ---
OUT = DOCX.parent / DOCX.stem
IMG_DIR = OUT / "public" / "img"
ARTICLE_IMG_DIR = IMG_DIR / "article"

OUT.mkdir(parents=True, exist_ok=True)
ARTICLE_IMG_DIR.mkdir(parents=True, exist_ok=True)

# --- 3. EXTRACCIÓN DEL WORD ---
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

z = zipfile.ZipFile(DOCX)
rels = {}
try:
    rels_root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
    for rel in rels_root:
        rid = rel.attrib.get("Id")
        target = rel.attrib.get("Target")
        if rid and target:
            rels[rid] = target
except KeyError:
    print("Advertencia: No se encontró el archivo de relaciones (document.xml.rels)")

# Extraer imágenes y convertir TODOS los EMF a PNG dinámicamente
media_map = {}
for name in z.namelist():
    if name.startswith("word/media/"):
        src_name = Path(name).name
        data = z.read(name)
        out_path = ARTICLE_IMG_DIR / src_name
        out_path.write_bytes(data)
        
        # Si es un EMF, lo convertimos
        if src_name.lower().endswith(".emf"):
            png_name = src_name[:-4] + ".png"
            emf_png = ARTICLE_IMG_DIR / png_name
            
            ps_script = f"""
            Add-Type -AssemblyName System.Drawing
            $src = '{str(out_path).replace("'", "''")}'
            $dest = '{str(emf_png).replace("'", "''")}'
            $mf = New-Object System.Drawing.Imaging.Metafile($src)
            $bmp = New-Object System.Drawing.Bitmap($mf.Width, $mf.Height)
            $g = [System.Drawing.Graphics]::FromImage($bmp)
            $g.Clear([System.Drawing.Color]::White)
            $g.DrawImage($mf, 0, 0, $mf.Width, $mf.Height)
            $bmp.Save($dest, [System.Drawing.Imaging.ImageFormat]::Png)
            $g.Dispose(); $bmp.Dispose(); $mf.Dispose()
            """
            try:
                subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], check=True, capture_output=True, text=True, timeout=30)
                if emf_png.exists() and emf_png.stat().st_size > 0:
                    media_map["media/" + src_name] = f"public/img/article/{png_name}"
                    out_path.unlink()
            except Exception as e:
                print(f"Error convirtiendo {src_name}: {e}")
                media_map["media/" + src_name] = f"public/img/article/{src_name}"
        else:
            media_map["media/" + src_name] = f"public/img/article/{src_name}"

# Procesamiento de Notas al pie
root_xml = ET.fromstring(z.read("word/document.xml"))
body_node = root_xml.find(W + "body")
body = list(body_node)

footnote_text = {}
try:
    froot = ET.fromstring(z.read("word/footnotes.xml"))
    for fn in froot.findall(W + "footnote"):
        fid = fn.get(W + "id")
        if fid is not None and int(fid) >= 0:
            text = "".join((t.text or "") for t in fn.iter(W + "t"))
            text = re.sub(r"\s+", " ", text).strip()
            text = re.sub(r"^(\*+|\d+)\s*", "", text).strip()
            footnote_text[int(fid)] = text
except KeyError:
    pass

used_footnotes = set()
URL_RE = re.compile(r"(https?://[^\s<]+|www\.[^\s<]+)")

def autolink(escaped_text):
    def repl(match):
        label = match.group(0)
        href = label if label.startswith("http") else "https://" + label
        trailing = ""
        while href and href[-1] in ".,);":
            trailing = href[-1] + trailing
            href = href[:-1]
            label = label[:-1]
        return f'<a href="{html.escape(href)}" target="_blank" rel="noopener">{label}</a>{html.escape(trailing)}'
    return URL_RE.sub(repl, escaped_text)

def run_is_bold(run):
    rpr = run.find(W + "rPr")
    return rpr is not None and (rpr.find(W + "b") is not None or rpr.find(W + "bCs") is not None)

def run_is_italic(run):
    rpr = run.find(W + "rPr")
    return rpr is not None and (rpr.find(W + "i") is not None or rpr.find(W + "iCs") is not None)

def render_run(run):
    parts = []
    fnref = run.find(W + "footnoteReference")
    if fnref is not None:
        fid = int(fnref.get(W + "id"))
        used_footnotes.add(fid)
        parts.append(f'<sup class="footnote-ref"><a href="#fn{fid}" id="fnref{fid}" aria-label="Nota {fid}">{fid}</a></sup>')
        if fnref.get(W + "customMarkFollows") == "1": return "".join(parts)
    for child in run:
        if child.tag == W + "t": parts.append(autolink(html.escape(child.text or "")))
        elif child.tag == W + "tab": parts.append(" ")
        elif child.tag == W + "br": parts.append("<br>")
    text = "".join(parts)
    if not text: return ""
    if run_is_italic(run): text = f"<em>{text}</em>"
    if run_is_bold(run): text = f"<strong>{text}</strong>"
    return text

def render_inline(node):
    parts = []
    for child in node:
        if child.tag == W + "r":
            if list(child.iter(A + "blip")): continue
            parts.append(render_run(child))
        elif child.tag == W + "hyperlink":
            href = ""
            rid = child.get(R + "id")
            if rid and rid in rels: href = rels[rid]
            inner = "".join(render_run(r) for r in child.findall(W + "r"))
            if href: parts.append(f'<a href="{html.escape(href)}" target="_blank" rel="noopener">{inner}</a>')
            else: parts.append(inner)
    return "".join(parts).strip()

def plain_text(node):
    return re.sub(r"\s+", " ", "".join((t.text or "") for t in node.iter(W + "t"))).strip()

def paragraph_images(pnode):
    imgs = []
    for blip in pnode.iter(A + "blip"):
        rid = blip.get(R + "embed") or blip.get(R + "link")
        if rid:
            target = rels.get(rid, rid)
            imgs.append(media_map.get(target, target))
    return imgs

def child_kind(child): return child.tag.split("}", 1)[-1]

# --- 4. EXTRACCIÓN DINÁMICA DE METADATOS Y TEXTO ---
title_html, english_title, author_html, affiliation = "", "", "", ""
abstract_es, keywords_es, abstract_en, keywords_en = "", "", "", ""

article_parts = []
references_open = False
in_body = False

for i, c in enumerate(body):
    text = plain_text(c)
    imgs = paragraph_images(c)
    html_text = render_inline(c)
    kind = child_kind(c)
    
    if kind == "sectPr" or (not text and not imgs and kind != "tbl"):
        continue
        
    if not title_html and len(text) > 5 and not text.startswith("Resumen"):
        title_html = html_text
        continue
    if title_html and not english_title and len(text) > 5 and not text.startswith("Resumen"):
        english_title = html.escape(text)
        continue
    if english_title and not author_html and len(text) > 3 and not text.startswith("Resumen"):
        author_html = html_text
        affiliation = footnote_text.get(2, "Afiliación no encontrada")
        continue
        
    if text.startswith("Resumen"):
        abstract_es = render_inline(body[i+1]) if i+1 < len(body) else "Resumen no disponible."
        continue
    if text.startswith("Palabras clave"):
        keywords_es = html_text.replace("Palabras clave:", "").strip()
        continue
    if text.startswith("Abstract"):
        abstract_en = render_inline(body[i+1]) if i+1 < len(body) else "Abstract not available."
        continue
    if text.startswith("Key words") or text.startswith("Keywords"):
        keywords_en = html_text.replace("Key words:", "").replace("Keywords:", "").strip()
        in_body = True 
        continue

    if not in_body: continue 
    
    if text.lower() in ["listado de referencias", "referencias", "bibliografía"]:
        article_parts.append('<section class="references" id="referencias"><h3>Listado de referencias</h3>')
        references_open = True
        continue
        
    if len(text) < 150 and ("<strong>" in html_text or "<h3>" in html_text) and not references_open:
        clean_title = re.sub(r"<(/?strong|/?em)>", "", html_text)
        article_parts.append(f"<h3>{clean_title}</h3>")
        continue

    if imgs:
        out = ['<figure class="article-figure"><div class="figure-media">']
        for src in imgs: out.append(f'<img src="{html.escape(src)}" alt="Imagen del artículo" loading="lazy">')
        out.append("</div></figure>")
        article_parts.append("".join(out))
        continue
        
    cls = ' class="reference-item"' if references_open else ""
    if text.startswith('"') and text.endswith('"'):
        article_parts.append(f'<blockquote class="academic-quote">{html_text}</blockquote>')
    else:
        article_parts.append(f"<p{cls}>{html_text}</p>")

if references_open: article_parts.append("</section>")

footnote_items = []
for fid in sorted(used_footnotes):
    text_fn = autolink(html.escape(footnote_text.get(fid, "")))
    footnote_items.append(f'<li id="fn{fid}"><p>{text_fn} <a class="footnote-back" href="#fnref{fid}" aria-label="Volver a la referencia">↩</a></p></li>')

footnotes_html = '<section class="footnotes" id="notas"><h3>Notas</h3><ol>' + "".join(footnote_items) + "</ol></section>" if footnote_items else ""
article_html = "\n".join(article_parts) + "\n" + footnotes_html

# --- 5. LIMPIEZA DE METADATOS Y ARMADO DEL HTML ---

# Quitamos las etiquetas HTML (como <sup> o <a>) de las variables de título y autor para que no rompan las etiquetas <meta>
title_plain = re.sub(r'<[^>]+>', '', title_html).replace('"', '&quot;')
author_plain = re.sub(r'<[^>]+>', '', author_html).replace('"', '&quot;')

html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Language" content="ES">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CICLOS | {title_plain[:50]}...</title>
<meta name="description" content="CICLOS en la historia, la economía y la sociedad. Espacio académico para la publicación de trabajos de investigación.">
<meta name="robots" content="index,follow">
<meta name="author" content="{author_plain}">
<meta name="DC.identifier" content="ISSN 0327-4063">
<meta name="DC.creator" content="{author_plain}">
<meta name="DC.publisher" content="CIHESRI-IDEHESI, Facultad de Ciencias Económicas, Universidad de Buenos Aires">
<meta name="DC.Title" content="{title_plain}">
<meta name="DC.Rights" content="Creative commons Attribution-NonCommercial-ShareAlike 4.0 Internacional">
<meta name="DC.Language" content="es">
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Montserrat:wght@500;700;800&family=Merriweather:ital,wght@0,300;0,400;1,300&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.9.1/font/bootstrap-icons.css">
<link rel="icon" href="public/img/LogoCiclos.png" type="image/png">
<link rel="stylesheet" href="style.css">
<link rel="alternate" type="application/rss+xml" title="CICLOS RSS Atom Feed" href="https://ojs.economicas.uba.ar/revistaCICLOS/oai">
</head>
<body>
<header class="site-header">
<div class="header-inner">
<div class="logo-area">
<a href="#">
<img src="public/img/LogoCiclos.png" alt="Revista CICLOS" class="header-logo">
</a>
</div>
<nav class="top-nav">
<a href="#" class="active-link">Actual</a>
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/issue/archive">Archivos</a>
<div class="dropdown">
<a href="#" class="dropdown-toggle">Acerca de <i class="bi bi-caret-down-fill"></i></a>
<div class="dropdown-menu">
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/about">Sobre la revista</a>
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/about/submissions">Envíos</a>
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/about/editorialTeam">Equipo editorial</a>
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/about/privacy">Declaración de privacidad</a>
<a href="https://ojs.economicas.uba.ar/revistaCICLOS/about/contact">Contacto</a>
</div>
</div>
</nav>
</div>
</header>
<main class="article-wrapper">
<div class="metadata-bar">
<span class="badge">Artículos</span>
<span class="badge badge-outline">Vol. XXXII, Nro. 65, 2025</span>
<span class="sep">|</span>
<a href="https://portal.issn.org/resource/ISSN/1851-3735" class="doi-link" target="_blank">eISSN 1851-3735</a>
<span class="sep">|</span>
<span class="doi-link">ISSN 0327-4063</span>
<span class="sep">|</span>
<a href="https://doi.org/10.1234/ciclos.v32i65.12345" class="doi-link" target="_blank">DOI: 10.1234/ciclos.v32i65.12345</a>
</div>
<header class="article-heading">
<h1 class="title">{title_html}</h1>
<h2 class="subtitle">{english_title}</h2>
<div class="author-block">
<p class="author-name">
{author_html}
<a href="https://orcid.org/0000-0000-0000-0000" target="_blank" title="Perfil ORCID del autor" style="color: #a6ce39; margin-left: 6px; font-size: 1.1rem;">
<i class="bi bi-person-badge"></i>
</a>
</p>
<p class="author-affiliation">{affiliation}</p>
</div>
</header>
<section class="abstract-box">
<div class="abstract-lang">
<h3>Resumen</h3>
<p>{abstract_es}</p>
<p class="keywords"><strong>Palabras clave:</strong> {keywords_es}</p>
</div>
<hr class="abstract-divider">
<div class="abstract-lang">
<h3>Abstract</h3>
<p>{abstract_en}</p>
<p class="keywords"><strong>Key words:</strong> {keywords_en}</p>
</div>
</section>
<article class="article-content">
{article_html}
</article>
</main>
<footer class="site-footer">
<div class="footer-inner">
<div class="footer-col">
<h4>Revista CICLOS</h4>
<p><b>Director Fundador:</b> Mario Rapoport</p>
<p><b>Directora:</b> Noemí Brenta</p>
<p><b>Editor Técnico:</b> E. Nacusse | <b>Editor:</b> F. Lucietto</p>
<p class="footer-meta"><i class="bi bi-geo-alt"></i> Av. Córdoba 2122, 2do. Piso, CABA</p>
<p class="footer-meta"><i class="bi bi-envelope"></i> <a href="mailto:ciclos@economicas.uba.ar">ciclos@economicas.uba.ar</a></p>
</div>
<div class="footer-col">
<h4>Sponsors e Institucional</h4>
<div class="academic-logos">
<a href="https://www.economicas.uba.ar" target="_blank">
<img src="https://www.economicas.uba.ar/wp-content/uploads/2020/08/cropped-logo_FCE.png" alt="Logo FCE UBA" class="logo-fce">
</a>
</div>
<p class="institution-text">
<b>Propietario:</b> Universidad de Buenos Aires. Facultad de Ciencias Económicas. Centro de Investigaciones de Historia Económica, Social y de Relaciones Internacionales.
</p>
<p class="institution-text">
CIHESRI - IDEHESI (Unidad Ejecutora en Red del CONICET). Maestría en Historia Económica y de las Políticas Económicas, FCE-UBA.
</p>
</div>
<div class="footer-col">
<h4>Políticas y Canales</h4>
<div class="license-area">
<a href="http://creativecommons.org/licenses/by-nc-sa/4.0/" rel="license" target="_blank">
<img src="https://i.creativecommons.org/l/by-nc-sa/4.0/88x31.png" alt="Licencia CC" class="cc-badge">
</a>
<p>Acceso abierto inmediato bajo Licencia CC Atribución-NoComercial 4.0. Sistema de preservación digital LOCKSS (Stanford University).</p>
</div>
</div>
</div>
</footer>
</body>
</html>
"""

# Tu variable css_doc se mantiene igual (no la incluí aquí entera para no hacer la respuesta infinita, pegá tu variable css_doc abajo de esto)

css_doc = """/* =========================================
Variables de Diseño (Identidad CICLOS)
========================================= */
:root {
--blue-primary: #1e6292;
--blue-dark: #012662;
--blue-light: #eef5fb;
--orange-accent: #ee8001;
--text-main: #2b2b2b;
--text-muted: #5a6a7a;
--bg-page: #fdfdfc;
--white: #ffffff;
--border-color: #cbd5e1;
--font-heading: 'Playfair Display', Georgia, serif;
--font-ui: "Inter", sans-serif;
--font-body: "Merriweather", serif;
}
* { margin: 0; padding: 0; box-sizing: border-box; }
html { scroll-behavior: smooth; }
body { font-family: var(--font-body); background-color: var(--bg-page); color: var(--text-main); line-height: 1.85; font-size: 1.08rem; -webkit-font-smoothing: antialiased; }
a { color: var(--blue-primary); text-decoration: none; transition: color 0.2s ease; }
a:hover { color: var(--orange-accent); }
.article-content a { overflow-wrap: anywhere; word-break: break-word; }
.site-header { background-color: var(--blue-primary); border-top: 4px solid var(--orange-accent); padding: 0.8rem 0; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06); position: sticky; top: 0; z-index: 100; }
.header-inner { max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; }
.brand-link { display: inline-flex; align-items: center; color: var(--white); }
.brand-link:hover { color: var(--white); }
.brand-text { color: var(--white); font-family: var(--font-heading); font-size: 2rem; font-weight: 700; letter-spacing: 0; }
.brand-text:hover { color: var(--white); }
.header-logo { height: 75px; width: auto; display: block; object-fit: contain; }
.header-logo + .brand-text { display: none; }
.top-nav { display: flex; gap: 1.8rem; align-items: center; }
.top-nav a { color: #e2e8f0; font-family: var(--font-ui); font-size: 0.85rem; font-weight: 600; text-transform: uppercase; position: relative; padding-bottom: 4px; transition: color 0.2s ease; letter-spacing: 0.5px; }
.top-nav a::after { content: ""; position: absolute; width: 0; height: 2px; bottom: 0; left: 0; background-color: var(--orange-accent); transition: width 0.25s ease-in-out; }
.top-nav a:hover::after, .top-nav a.active-link::after { width: 100%; }
.dropdown { position: relative; display: flex; align-items: center; height: 100%; }
.dropdown-toggle { display: flex; align-items: center; gap: 5px; }
.dropdown-menu { display: none !important; position: absolute; background-color: var(--white); min-width: 240px; box-shadow: 0 8px 16px rgba(0, 0, 0, 0.15); z-index: 999; top: 100%; left: 0; padding: 0.5rem 0; border-top: 3px solid var(--orange-accent); flex-direction: column; }
.dropdown:hover .dropdown-menu { display: flex !important; }
.top-nav .dropdown-menu a { color: var(--blue-dark) !important; padding: 10px 16px !important; text-transform: none !important; font-weight: 500 !important; font-size: 0.95rem !important; width: 100%; display: block !important; }
.top-nav .dropdown-menu a::after { display: none !important; }
.top-nav .dropdown-menu a:hover { background-color: var(--blue-light) !important; color: var(--blue-primary) !important; }
.article-wrapper { max-width: 1000px; margin: 4rem auto; padding: 0 2rem; }
.metadata-bar { display: flex; align-items: center; flex-wrap: wrap; gap: 0.8rem; margin-bottom: 2rem; font-family: var(--font-ui); font-size: 0.75rem; font-weight: 700; }
.badge { background-color: var(--blue-primary); color: var(--white); padding: 0.3rem 0.8rem; border-radius: 4px; text-transform: uppercase; }
.badge-outline { background-color: transparent; color: var(--blue-primary); border: 1px solid var(--blue-primary); }
.doi-link { color: var(--text-muted); }
.title { font-family: var(--font-heading); font-size: 2.55rem; font-weight: 700; color: var(--blue-dark); line-height: 1.25; margin-bottom: 1.2rem; }
.title sup, .author-name sup { font-family: var(--font-ui); font-size: 0.55em; line-height: 0; }
.subtitle-group { display: grid; gap: 0.55rem; margin-bottom: 2rem; }
.subtitle { font-family: var(--font-heading); font-size: 1.22rem; color: var(--text-muted); font-style: italic; font-weight: 400; line-height: 1.4; }
.author-block { border-left: 4px solid var(--orange-accent); padding-left: 1rem; }
.author-name { font-family: var(--font-ui); font-size: 1.05rem; font-weight: 700; color: var(--blue-primary); }
.author-affiliation, .article-dates { font-family: var(--font-ui); font-size: 0.86rem; color: var(--text-muted); margin-top: 0.25rem; line-height: 1.55; }
.abstract-box { background-color: var(--blue-light); padding: 2.5rem; border-radius: 8px; margin: 3.5rem 0; border: 1px solid #d0e1f0; }
.abstract-lang h3 { font-family: var(--font-ui); color: var(--blue-dark); text-transform: uppercase; font-size: 0.85rem; margin-bottom: 0.8rem; font-weight: 700; letter-spacing: 0.5px; }
.abstract-lang p { font-size: 0.95rem; color: #334155; }
.abstract-divider { border: 0; height: 1px; background-color: #cbd5e1; margin: 2rem 0; }
.keywords { margin-top: 1rem; font-family: var(--font-ui); font-size: 0.82rem !important; color: var(--text-main); }
.article-content h3, .article-content h4 { font-family: var(--font-heading); color: var(--blue-dark); font-weight: 700; line-height: 1.3; }
.article-content h3 { font-size: 1.7rem; margin: 3.6rem 0 1.2rem; }
.article-content h4 { font-size: 1.35rem; margin: 2.7rem 0 1rem; }
.article-content p { margin-bottom: 1.55rem; text-align: justify; }
.footnote-ref a { font-family: var(--font-ui); font-weight: 700; padding: 0 0.08rem; }
.article-figure { margin: 2.6rem 0; padding: 1.5rem; background-color: var(--white); border: 1px solid var(--border-color); border-radius: 8px; }
.article-figure figcaption { font-family: var(--font-ui); color: var(--blue-dark); line-height: 1.45; margin-bottom: 1rem; }
.figure-label { display: block; font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; color: var(--orange-accent); margin-bottom: 0.25rem; }
.figure-title { display: block; font-size: 0.98rem; font-weight: 700; }
.figure-media { display: grid; justify-items: center; gap: 1rem; }
.figure-grid { grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); align-items: start; }
.figure-media img { display: block; max-width: 100%; height: auto; border: 1px solid #e2e8f0; background-color: var(--white); }
.figure-source { font-family: var(--font-ui); font-size: 0.78rem; color: var(--text-muted); line-height: 1.55; margin: 1rem 0 0 !important; text-align: left !important; }
.table-scroll { overflow-x: auto; margin-top: 1rem; }
.data-table { width: 100%; border-collapse: collapse; font-family: var(--font-ui); font-size: 0.86rem; line-height: 1.45; background-color: var(--white); }
.data-table th, .data-table td { border: 1px solid var(--border-color); padding: 0.65rem 0.75rem; vertical-align: top; text-align: right; }
.data-table th:first-child, .data-table td:first-child { text-align: left; }
.data-table thead th { background-color: var(--blue-light); color: var(--blue-dark); font-weight: 700; }
.references { margin-top: 3.5rem; }
.references .reference-item { font-size: 0.92rem; line-height: 1.65; margin-bottom: 0.85rem; text-align: left; }
.footnotes { border-top: 1px solid var(--border-color); margin-top: 3.5rem; padding-top: 1rem; }
.footnotes h3 { margin-top: 1rem; }
.footnotes ol { padding-left: 1.35rem; }
.footnotes li { padding-left: 0.25rem; margin-bottom: 0.75rem; }
.footnotes p { font-family: var(--font-ui); font-size: 0.82rem; line-height: 1.55; text-align: left; margin-bottom: 0; }
.footnote-back { font-weight: 700; }
.site-footer { background-color: var(--blue-primary); color: #cbd5e1; padding: 3rem 0; font-family: var(--font-ui); font-size: 0.8rem; margin-top: 6rem; border-top: 5px solid var(--blue-dark); }
.footer-inner { max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: grid; grid-template-columns: repeat(3, 1fr); gap: 3rem; }
.footer-col h4 { color: var(--white); margin-bottom: 1.2rem; font-size: 0.85rem; text-transform: uppercase; border-bottom: 2px solid rgba(255, 255, 255, 0.1); padding-bottom: 0.5rem; letter-spacing: 0.5px; }
.footer-col p { margin-bottom: 0.6rem; line-height: 1.6; }
.footer-col b { color: var(--white); }
.footer-col a { color: var(--white); text-decoration: underline; }
.academic-logos { display: flex; align-items: center; gap: 12px; margin-bottom: 1.2rem; }
.logo-fce { display: block; height: 42px; width: auto; background-color: white; padding: 6px 10px; border-radius: 4px; transition: transform 0.2s ease; }
.logo-fce:hover { transform: translateY(-2px); }
.institution-text { font-size: 0.75rem; color: #d4deea; line-height: 1.5; }
.license-area { margin-bottom: 1.2rem; }
.cc-badge { margin-bottom: 0.4rem; display: block; }
@media screen and (max-width: 900px) {
  .footer-inner { grid-template-columns: 1fr; gap: 2rem; }
  .header-inner { flex-direction: column; gap: 1.2rem; }
  .top-nav { width: 100%; justify-content: space-between; gap: 1rem; }
  .article-wrapper { margin: 2.5rem auto; padding: 0 1.2rem; }
  .title { font-size: 1.9rem; }
  .subtitle { font-size: 1.06rem; }
  .abstract-box { padding: 1.4rem; }
  .article-figure { padding: 1rem; }
  .figure-grid { grid-template-columns: 1fr; }
}
"""

(OUT / "index.html").write_text(html_doc, encoding="utf-8")
(OUT / "style.css").write_text(css_doc, encoding="utf-8")

root.destroy()
messagebox.showinfo("¡Éxito!", f"¡Listo! Se extrajo el artículo y las imágenes dinámicamente.\nTodo se guardó en:\n{OUT}")