import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET
import html
import re
import subprocess
import sys

# --- 1. SELECCIÓN DEL ARCHIVO MEDIANTE VENTANA VISUAL ---
root = tk.Tk()
root.withdraw() # Oculta la ventana de fondo
file_path = filedialog.askopenfilename(
    title="Selecciona el archivo Word de la Revista CICLOS",
    filetypes=[("Documentos de Word", "*.docx")]
)

if not file_path:
    print("No se seleccionó ningún archivo. Saliendo...")
    sys.exit()

DOCX = Path(file_path)

# --- 2. CREACIÓN DE LA CARPETA Y SU ESTRUCTURA ---
# Crea una carpeta con el mismo nombre del Word en el mismo directorio
OUT = DOCX.parent / DOCX.stem
IMG_DIR = OUT / "public" / "img"
ARTICLE_IMG_DIR = IMG_DIR / "article"

OUT.mkdir(parents=True, exist_ok=True)
ARTICLE_IMG_DIR.mkdir(parents=True, exist_ok=True)

# Copiar el LogoCiclos.png si es posible (Opcional, asume que está en tu PC, si no, el CSS lo maneja)
# shutil.copy("LogoCiclos.png", IMG_DIR / "LogoCiclos.png")

# --- 3. EXTRACCIÓN DEL WORD Y PROCESAMIENTO XML ---
W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

z = zipfile.ZipFile(DOCX)
rels = {}
rels_root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
for rel in rels_root:
    rid = rel.attrib.get("Id")
    target = rel.attrib.get("Target")
    if rid and target:
        rels[rid] = target

styles = {}
styles_root = ET.fromstring(z.read("word/styles.xml"))
for st in styles_root.findall(W + "style"):
    sid = st.get(W + "styleId")
    nm = st.find(W + "name")
    if sid and nm is not None:
        styles[sid] = nm.get(W + "val")

# Extracción de imágenes
media_map = {}
for name in z.namelist():
    if name.startswith("word/media/"):
        src_name = Path(name).name
        data = z.read(name)
        (ARTICLE_IMG_DIR / src_name).write_bytes(data)
        media_map["media/" + src_name] = f"public/img/article/{src_name}"

# Conversión de archivos .emf a .png usando PowerShell
emf_src = ARTICLE_IMG_DIR / "image21.emf"
emf_png = ARTICLE_IMG_DIR / "image21.png"
if emf_src.exists():
    ps_script = f"""
    Add-Type -AssemblyName System.Drawing
    $src = '{str(emf_src).replace("'", "''")}'
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
            media_map["media/image21.emf"] = "public/img/article/image21.png"
    except Exception:
        pass

# Procesamiento de textos y notas al pie
root_xml = ET.fromstring(z.read("word/document.xml"))
body = list(root_xml.find(W + "body"))

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

def render_table(tbl):
    rows = []
    for tr in tbl.findall(W + "tr"):
        cells = []
        for tc in tr.findall(W + "tc"):
            paras = [render_inline(p) for p in tc.findall(W + "p")]
            cells.append("<br>".join([p for p in paras if p]))
        if cells: rows.append(cells)
    if not rows: return ""
    html_rows = ['<div class="table-scroll"><table class="data-table"><thead><tr>']
    for cell in rows: html_rows.append(f"<th>{cell}</th>")
    html_rows.append("</tr></thead><tbody>")
    for row in rows[1:]:
        html_rows.append("<tr>")
        for cell in row: html_rows.append(f"<td>{cell}</td>")
        html_rows.append("</tr>")
    html_rows.append("</tbody></table></div>")
    return "".join(html_rows)

def next_nonempty_index(start):
    j = start
    while j < len(body):
        c = body[j]
        if child_kind(c) == "tbl" or (child_kind(c) == "p" and (plain_text(c) or paragraph_images(c))): return j
        j += 1
    return None

def is_figure_label(text): return re.match(r"^(Figura|Gráfico|Cuadro)\s+\d+\b", text or "") is not None

heading_texts = {
    "Introducción", "El Bocade de Tucumán: ¿por qué y cómo se emitió?", "Nacimiento y primera infancia del Bocade, 1985-1987",
    "Edad adulta: el Bocade en australes, 1987 - 1991", "Edad madura: el Bocade en peso dolarizado, 1992-2001",
    "Tercera etapa: el Bocade en pesos renacionalizados, 2 de enero de 2002 – 27 de marzo de 2003",
    "Muerte del Bocade: el Programa de Unificación Monetaria, del 28 de marzo al 31 de diciembre de 2003",
    "Conclusión: la experiencia del Bocade de Tucumán como ejemplo de un federalismo monetario reprimido",
    "Listado de referencias",
}
subheading_texts = {"El mecanismo establecido por el PUM"}

# Extracción de meta-datos del documento
try:
    title_html = render_inline(body[6])
    french_title = html.escape(plain_text(body[7]) + " " + plain_text(body[8]))
    english_title = html.escape(plain_text(body[9]) + " " + plain_text(body[10]))
    author_html = render_inline(body[11])
    affiliation = footnote_text.get(2, "")
    abstract_es = render_inline(body[12])
    keywords_es = render_inline(body[13])
    abstract_en = render_inline(body[14])
    keywords_en = render_inline(body[15])
except IndexError:
    title_html = "Título no encontrado"
    english_title, author_html, affiliation, abstract_es, keywords_es, abstract_en, keywords_en = "", "", "", "", "", "", ""

def render_figure(i):
    label = plain_text(body[i])
    figure_id = re.sub(r"[^a-z0-9]+", "-", label.lower()).strip("-")
    caption, notes, media, table_html = [], [], [], ""
    consumed = i + 1
    seen_media = False
    while consumed < len(body):
        c = body[consumed]
        kind = child_kind(c)
        if kind == "sectPr": break
        if kind == "tbl":
            table_html = render_table(c)
            seen_media = True
            consumed += 1
            continue
        text = plain_text(c) if kind == "p" else ""
        imgs = paragraph_images(c) if kind == "p" else []
        if not text and not imgs:
            consumed += 1
            nxt = next_nonempty_index(consumed)
            if seen_media and nxt is not None and (child_kind(body[nxt]) == "p" and paragraph_images(body[nxt])): continue
            if seen_media: break
            continue
        if text and is_figure_label(text): break
        if text in heading_texts or text in subheading_texts: break
        if imgs:
            if text and not seen_media and not caption: caption.append(render_inline(c))
            elif text and text.lower().startswith(("fuente:", "fuentes:")): notes.append(render_inline(c))
            media.extend(imgs)
            seen_media = True
            consumed += 1
            continue
        if text.lower().startswith(("fuente:", "fuentes:")):
            notes.append(render_inline(c))
            consumed += 1
            break
        if not seen_media:
            caption.append(render_inline(c))
            consumed += 1
            continue
        nxt = next_nonempty_index(consumed + 1)
        if nxt is not None and child_kind(body[nxt]) == "p" and paragraph_images(body[nxt]):
            notes.append(render_inline(c))
            consumed += 1
            continue
        break
    alt = html.escape(" ".join([label] + [re.sub("<[^>]+>", "", c) for c in caption]) or label)
    out = [f'<figure class="article-figure" id="{figure_id}"><figcaption><span class="figure-label">{html.escape(label)}</span>']
    if caption: out.append(f'<span class="figure-title">{" ".join(caption)}</span>')
    out.append("</figcaption>")
    if media:
        cls = " figure-grid" if len(media) > 1 else ""
        out.append(f'<div class="figure-media{cls}">')
        for src in media: out.append(f'<img src="{html.escape(src)}" alt="{alt}" loading="lazy">')
        out.append("</div>")
    if table_html: out.append(table_html)
    for note in notes: out.append(f'<p class="figure-source">{note}</p>')
    out.append("</figure>")
    return "".join(out), consumed

article_parts = []
references_open = False
i = 28
while i < len(body):
    c = body[i]
    kind = child_kind(c)
    if kind == "sectPr": i += 1; continue
    if kind == "tbl": article_parts.append(render_table(c)); i += 1; continue
    text = plain_text(c)
    imgs = paragraph_images(c)
    if not text and not imgs: i += 1; continue
    if is_figure_label(text):
        fig_html, new_i = render_figure(i)
        article_parts.append(fig_html)
        i = new_i
        continue
    if text in heading_texts:
        if text == "Listado de referencias":
            article_parts.append('<section class="references" id="referencias"><h3>Listado de referencias</h3>')
            references_open = True
        else:
            article_parts.append(f"<h3>{render_inline(c)}</h3>")
        i += 1
        continue
    if text in subheading_texts:
        article_parts.append(f"<h4>{render_inline(c)}</h4>")
        i += 1
        continue
    if imgs:
        out = ['<figure class="article-figure"><div class="figure-media">']
        for src in imgs: out.append(f'<img src="{html.escape(src)}" alt="Imagen del artículo" loading="lazy">')
        out.append("</div></figure>")
        article_parts.append("".join(out))
        i += 1
        continue
    
    cls = ' class="reference-item"' if references_open else ""
    paragraph = render_inline(c)
    if paragraph.startswith(".Incapaz"): paragraph = paragraph[1:]
    if paragraph.startswith("sta cita"): paragraph = "E" + paragraph
    
    # Manejar citas (blockquotes)
    if "derrame desarrollista" in paragraph:
        article_parts.append(f'<blockquote class="academic-quote">"{paragraph}"</blockquote>')
    else:
        article_parts.append(f"<p{cls}>{paragraph}</p>")
    i += 1

if references_open: article_parts.append("</section>")

footnote_items = []
for fid in sorted(used_footnotes):
    text = autolink(html.escape(footnote_text.get(fid, "")))
    footnote_items.append(f'<li id="fn{fid}"><p>{text} <a class="footnote-back" href="#fnref{fid}" aria-label="Volver a la referencia">↩</a></p></li>')

footnotes_html = ""
if footnote_items:
    footnotes_html = '<section class="footnotes" id="notas"><h3>Notas</h3><ol>' + "".join(footnote_items) + "</ol></section>"

article_html = "\n".join(article_parts) + "\n" + footnotes_html


# --- 4. TUS PLANTILLAS EXACTAS DE HTML Y CSS ---

html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Language" content="ES">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CICLOS | {title_html}</title>
<meta name="description" content="CICLOS en la historia, la economía y la sociedad. Espacio académico para la publicación de trabajos de investigación.">
<meta name="robots" content="index,follow">
<meta name="author" content="{author_html}">
<meta name="DC.identifier" content="ISSN 0327-4063">
<meta name="DC.creator" content="{author_html}">
<meta name="DC.publisher" content="CIHESRI-IDEHESI, Facultad de Ciencias Económicas, Universidad de Buenos Aires">
<meta name="DC.Title" content="{title_html}">
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
<a href="#" class="doi-link" target="_blank">DOI: 10.1234/ciclos.v32i65.12345</a>
</div>
<header class="article-heading">
<h1 class="title">{title_html}</h1>
<h2 class="subtitle">{english_title}</h2>
<div class="author-block">
<p class="author-name">
{author_html}
<a href="#" target="_blank" title="Perfil ORCID del autor" style="color: #a6ce39; margin-left: 6px; font-size: 1.1rem;">
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
<p class="keywords"><strong>Palabras clave:</strong> {keywords_es.replace("Palabras clave:", "").strip()}</p>
</div>
<hr class="abstract-divider">
<div class="abstract-lang">
<h3>Abstract</h3>
<p>{abstract_en}</p>
<p class="keywords"><strong>Key words:</strong> {keywords_en.replace("Key words:", "").strip()}</p>
</div>
</section>
<article class="article-content">
{article_html}
</article>
</main>
<footer class="site-footer">
<div class="footer-inner">
<!-- COLUMNA 1 -->
<div class="footer-col">
<h4>Revista CICLOS</h4>
<p><b>Director Fundador:</b> Mario Rapoport</p>
<p><b>Directora:</b> Noemí Brenta</p>
<p><b>Editor Técnico:</b> E. Nacusse | <b>Editor:</b> F. Lucietto</p>
<p class="footer-meta"><i class="bi bi-geo-alt"></i> Av. Córdoba 2122, 2do. Piso, CABA</p>
<p class="footer-meta"><i class="bi bi-envelope"></i> <a href="mailto:ciclos@economicas.uba.ar">ciclos@economicas.uba.ar</a></p>
</div>
<!-- COLUMNA 2 -->
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
<!-- COLUMNA 3 -->
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
* {
margin: 0;
padding: 0;
box-sizing: border-box;
}
body {
font-family: var(--font-body);
background-color: var(--bg-page);
color: var(--text-main);
line-height: 1.85;
font-size: 1.1rem;
-webkit-font-smoothing: antialiased;
}
a {
color: var(--blue-primary);
text-decoration: none;
transition: color 0.2s ease;
}
a:hover {
color: var(--orange-accent);
}
/* =========================================
Cabecera Horizontal con Proporción de Logo
========================================= */
.site-header {
background-color: var(--blue-primary);
border-top: 4px solid var(--orange-accent);
padding: 0.8rem 0;
box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06);
position: sticky;
top: 0;
z-index: 100;
}
.header-inner {
max-width: 1150px;
margin: 0 auto;
padding: 0 2rem;
display: flex;
justify-content: space-between;
align-items: center;
flex-wrap: wrap;
}
.header-logo {
height: 75px;
width: auto;
display: block;
object-fit: contain;
}
/* Enlaces y Efecto Hover Fluido */
.top-nav {
display: flex;
gap: 1.8rem;
align-items: center;
}
.top-nav a {
color: #e2e8f0;
font-family: var(--font-ui);
font-size: 0.85rem;
font-weight: 600;
text-transform: uppercase;
position: relative;
padding-bottom: 4px;
transition: color 0.2s ease;
letter-spacing: 0.5px;
}
.top-nav a::after {
content: "";
position: absolute;
width: 0;
height: 2px;
bottom: 0;
left: 0;
background-color: var(--orange-accent);
transition: width 0.25s ease-in-out;
}
.top-nav a.active-link:hover {
cursor: default !important;
color: var(--white);
}
.top-nav a.active-link:hover::after {
width: 100% !important;
}
.nav-highlight {
background-color: rgba(255, 255, 255, 0.15);
padding: 0.4rem 1rem !important;
border-radius: 4px;
}
.nav-highlight::after {
display: none !important;
}
.nav-highlight:hover {
background-color: var(--orange-accent) !important;
}
/* =========================================
Menú Desplegable (Dropdown)
========================================= */
.dropdown {
position: relative;
display: flex;
align-items: center;
height: 100%;
}
.dropdown-toggle {
display: flex;
align-items: center;
gap: 5px;
}
.dropdown-menu {
display: none !important;
position: absolute;
background-color: var(--white);
min-width: 240px;
box-shadow: 0px 8px 16px rgba(0, 0, 0, 0.15);
z-index: 999;
top: 100%;
left: 0;
padding: 0.5rem 0;
border-top: 3px solid var(--orange-accent);
flex-direction: column;
}
.dropdown:hover .dropdown-menu {
display: flex !important;
}
.top-nav .dropdown-menu a {
color: var(--blue-dark) !important;
padding: 10px 16px !important;
text-transform: none !important;
font-weight: 500 !important;
font-size: 0.95rem !important;
width: 100%;
display: block !important;
}
.top-nav .dropdown-menu a::after {
display: none !important;
}
.top-nav .dropdown-menu a:hover {
background-color: var(--blue-light) !important;
color: var(--blue-primary) !important;
}
/* =========================================
Cuerpo del Artículo
========================================= */
.article-wrapper {
max-width: 1000px;
margin: 4rem auto;
padding: 0 2rem;
}
.metadata-bar {
display: flex;
align-items: center;
gap: 0.8rem;
margin-bottom: 2rem;
font-family: var(--font-ui);
font-size: 0.75rem;
font-weight: 700;
}
.badge {
background-color: var(--blue-primary);
color: var(--white);
padding: 0.3rem 0.8rem;
border-radius: 4px;
text-transform: uppercase;
}
.badge-outline {
background-color: transparent;
color: var(--blue-primary);
border: 1px solid var(--blue-primary);
}
.doi-link {
color: var(--text-muted);
}
.title {
font-family: var(--font-heading);
font-size: 2.5rem;
font-weight: 700;
color: var(--blue-dark);
line-height: 1.25;
margin-bottom: 1.2rem;
}
.subtitle {
font-family: var(--font-heading);
font-size: 1.3rem;
color: var(--text-muted);
margin-bottom: 2rem;
font-style: italic;
font-weight: 400;
line-height: 1.4;
}
.author-block {
border-left: 4px solid var(--orange-accent);
padding-left: 1rem;
}
.author-name {
font-family: var(--font-ui);
font-size: 1.05rem;
font-weight: 700;
color: var(--blue-primary);
}
.author-affiliation {
font-family: var(--font-ui);
font-size: 0.85rem;
color: var(--text-muted);
margin-top: 0.2rem;
}
/* Cajas de Resumen */
.abstract-box {
background-color: var(--blue-light);
padding: 2.5rem;
border-radius: 12px;
margin: 3.5rem 0;
border: 1px solid #d0e1f0;
}
.abstract-lang h3 {
font-family: var(--font-ui);
color: var(--blue-dark);
text-transform: uppercase;
font-size: 0.85rem;
margin-bottom: 0.8rem;
font-weight: 700;
letter-spacing: 0.5px;
}
.abstract-lang p {
font-size: 0.95rem;
color: #334155;
}
.abstract-divider {
border: 0;
height: 1px;
background-color: #cbd5e1;
margin: 2rem 0;
}
.keywords {
margin-top: 1rem;
font-family: var(--font-ui);
font-size: 0.8rem;
color: var(--text-main);
}
.article-content h3 {
font-family: var(--font-heading);
font-size: 1.6rem;
color: var(--blue-dark);
margin: 3.5rem 0 1.2rem 0;
font-weight: 700;
}
.article-content p {
margin-bottom: 1.8rem;
text-align: justify;
}
blockquote, .academic-quote {
font-size: 1.35rem;
font-family: var(--font-heading);
font-style: italic;
color: var(--blue-primary);
margin: 3.5rem 0;
padding: 1.5rem 2.5rem;
border-left: 4px solid var(--orange-accent);
background-color: var(--white);
box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02);
border-radius: 0 8px 8px 0;
}
.article-figure {
margin: 2.6rem 0;
padding: 1.5rem;
background-color: var(--white);
border: 1px solid var(--border-color);
border-radius: 8px;
}
.article-figure figcaption {
font-family: var(--font-ui);
color: var(--blue-dark);
line-height: 1.45;
margin-bottom: 1rem;
}
.figure-label {
display: block;
font-size: 0.78rem;
font-weight: 800;
text-transform: uppercase;
letter-spacing: 0.5px;
color: var(--orange-accent);
margin-bottom: 0.25rem;
}
.figure-title { display: block; font-size: 0.98rem; font-weight: 700; }
.figure-media { display: grid; justify-items: center; gap: 1rem; }
.figure-grid { grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); align-items: start; }
.figure-media img {
display: block;
max-width: 100%;
height: auto;
border: 1px solid #e2e8f0;
background-color: var(--white);
}
.figure-source {
font-family: var(--font-ui);
font-size: 0.78rem;
color: var(--text-muted);
line-height: 1.55;
margin: 1rem 0 0 !important;
text-align: left !important;
}
.table-scroll { overflow-x: auto; margin-top: 1rem; }
.data-table {
width: 100%;
border-collapse: collapse;
font-family: var(--font-ui);
font-size: 0.86rem;
line-height: 1.45;
background-color: var(--white);
}
.data-table th, .data-table td {
border: 1px solid var(--border-color);
padding: 0.65rem 0.75rem;
vertical-align: top;
text-align: right;
}
.data-table th:first-child, .data-table td:first-child { text-align: left; }
.data-table thead th {
background-color: var(--blue-light);
color: var(--blue-dark);
font-weight: 700;
}
/* =========================================
Pie de página de 3 Columnas Fijas (Grid)
========================================= */
.site-footer {
background-color: var(--blue-primary);
color: #cbd5e1;
padding: 3rem 0;
font-family: var(--font-ui);
font-size: 0.8rem;
margin-top: 6rem;
border-top: 5px solid var(--blue-dark);
}
.footer-inner {
max-width: 1150px;
margin: 0 auto;
padding: 0 2rem;
display: grid;
grid-template-columns: repeat(3, 1fr);
gap: 3rem;
}
.footer-col h4 {
color: var(--white);
margin-bottom: 1.2rem;
font-size: 0.85rem;
text-transform: uppercase;
border-bottom: 2px solid rgba(255, 255, 255, 0.1);
padding-bottom: 0.5rem;
letter-spacing: 0.5px;
}
.footer-col p {
margin-bottom: 0.6rem;
line-height: 1.6;
}
.footer-col b {
color: var(--white);
}
.footer-col a {
color: var(--white);
text-decoration: underline;
}
/* =========================================
Logos Institucionales Footer
========================================= */
.academic-logos {
display: flex;
align-items: center;
gap: 12px;
margin-bottom: 1.2rem;
}
.logo-fce {
display: block;
height: 42px;
width: auto;
background-color: white;
padding: 6px 10px;
border-radius: 4px;
transition: transform 0.2s ease;
}
.logo-fce:hover {
transform: translateY(-2px);
}
.institution-text {
font-size: 0.75rem;
color: #94a3b8;
line-height: 1.5;
}
.license-area {
margin-bottom: 1.2rem;
}
.cc-badge {
margin-bottom: 0.4rem;
display: block;
}
.social-icons {
display: flex;
gap: 1.2rem;
font-size: 1.2rem;
margin-top: 1rem;
}
.social-icons a {
color: #94a3b8;
}
.social-icons a:hover {
color: var(--orange-accent);
}
/* Celulares y Tablets */
@media screen and (max-width: 900px) {
.footer-inner {
grid-template-columns: 1fr;
gap: 2rem;
}
.header-inner {
flex-direction: column;
gap: 1.2rem;
}
.top-nav {
width: 100%;
justify-content: space-between;
}
.title {
font-size: 1.9rem;
}
}
"""

# --- 5. GUARDADO FINAL DE LOS ARCHIVOS ---
(OUT / "index.html").write_text(html_doc, encoding="utf-8")
(OUT / "style.css").write_text(css_doc, encoding="utf-8")

print(f"¡Éxito! El artículo se guardó en: {OUT}")
# Muestra un cartel visual de éxito
messagebox.showinfo("¡Conversión Exitosa!", f"El archivo Word se ha convertido correctamente.\n\nRevisa la carpeta:\n{OUT}")