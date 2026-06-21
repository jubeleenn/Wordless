# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, colorchooser
from pathlib import Path
import html
import re
import subprocess
import zipfile
import xml.etree.ElementTree as ET
import shutil
import sys

class WordlessApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Wordless - Revistas Científicas")
        self.root.geometry("600x750")
        self.root.configure(padx=15, pady=15, bg="#fdfdfc")
        
        self.word_path = None
        self.logo_path = None

        # --- DICCIONARIO DE TEMAS ---
        self.temas = {
            "CICLOS": {"fondo": "#fdfdfc", "texto": "#2b2b2b", "primario": "#1e6292", "acento": "#ee8001", "fuente": "Merriweather"},
            "Erudita": {"fondo": "#ffffff", "texto": "#000000", "primario": "#111111", "acento": "#5c0000", "fuente": "Times New Roman"},
            "Prisma": {"fondo": "#f8f9fa", "texto": "#333333", "primario": "#005f73", "acento": "#0a9396", "fuente": "Roboto"},
            "Vanguardia": {"fondo": "#121212", "texto": "#e0e0e0", "primario": "#9d0208", "acento": "#ffba08", "fuente": "Inter"},
            "Lienzo": {"fondo": "#ffffff", "texto": "#1a1a1a", "primario": "#000000", "acento": "#a8a8a8", "fuente": "Georgia"}
        }

        tk.Label(root, text="Wordless App", font=("Arial", 16, "bold"), bg="#fdfdfc", fg="#012662").pack(pady=5)

        # 1. Archivos
        frame_archivos = tk.LabelFrame(root, text="1. Archivos", padx=10, pady=10, bg="#fdfdfc")
        frame_archivos.pack(fill="x", pady=5)

        self.lbl_word = tk.Label(frame_archivos, text="Ningún Word seleccionado", fg="red", bg="#fdfdfc")
        self.lbl_word.grid(row=0, column=1, sticky="w", padx=10)
        tk.Button(frame_archivos, text="Subir Word (.docx)", command=self.cargar_word).grid(row=0, column=0, pady=5)

        self.lbl_logo = tk.Label(frame_archivos, text="Ningún Logo seleccionado (Opcional)", fg="red", bg="#fdfdfc")
        self.lbl_logo.grid(row=1, column=1, sticky="w", padx=10)
        tk.Button(frame_archivos, text="Subir Logo", command=self.cargar_logo).grid(row=1, column=0, pady=5)

        # 2. Plantilla y Metadatos
        frame_datos = tk.LabelFrame(root, text="2. Estructura y Metadatos", padx=10, pady=10, bg="#fdfdfc")
        frame_datos.pack(fill="x", pady=5)

        tk.Label(frame_datos, text="Estructura Visual:", bg="#fdfdfc").grid(row=0, column=0, sticky="w")
        self.combo_revista = ttk.Combobox(frame_datos, values=list(self.temas.keys()), state="readonly", width=40)
        self.combo_revista.current(0)
        self.combo_revista.grid(row=0, column=1, pady=5, padx=10)
        self.combo_revista.bind("<<ComboboxSelected>>", self.aplicar_tema)

        tk.Label(frame_datos, text="Título (Opcional):", bg="#fdfdfc").grid(row=1, column=0, sticky="w")
        self.ent_titulo = tk.Entry(frame_datos, width=43)
        self.ent_titulo.grid(row=1, column=1, pady=5, padx=10)

        tk.Label(frame_datos, text="Autor (Opcional):", bg="#fdfdfc").grid(row=2, column=0, sticky="w")
        self.ent_autor = tk.Entry(frame_datos, width=43)
        self.ent_autor.grid(row=2, column=1, pady=5, padx=10)

        # 3. Diseño Manual
        frame_manual = tk.LabelFrame(root, text="3. Diseño Manual Personalizado", padx=10, pady=10, bg="#fdfdfc")
        frame_manual.pack(fill="x", pady=5)

        tk.Label(frame_manual, text="Tipografía:", bg="#fdfdfc").grid(row=0, column=0, sticky="w", pady=2)
        self.combo_fuente = ttk.Combobox(frame_manual, values=["Merriweather", "Georgia", "Times New Roman", "Arial", "Inter", "Roboto"], state="readonly", width=15)
        self.combo_fuente.set("Merriweather")
        self.combo_fuente.grid(row=0, column=1, pady=2, padx=10, sticky="w")

        tk.Label(frame_manual, text="Fondo / Texto:", bg="#fdfdfc").grid(row=1, column=0, sticky="w", pady=2)
        frame_colores_1 = tk.Frame(frame_manual, bg="#fdfdfc")
        frame_colores_1.grid(row=1, column=1, sticky="w", padx=10)
        self.btn_bg = tk.Button(frame_colores_1, bg="#fdfdfc", width=6, command=lambda: self.elegir_color(self.btn_bg))
        self.btn_bg.pack(side="left", padx=2)
        self.btn_text = tk.Button(frame_colores_1, bg="#2b2b2b", width=6, command=lambda: self.elegir_color(self.btn_text))
        self.btn_text.pack(side="left", padx=2)

        tk.Label(frame_manual, text="Principal / Acento:", bg="#fdfdfc").grid(row=2, column=0, sticky="w", pady=2)
        frame_colores_2 = tk.Frame(frame_manual, bg="#fdfdfc")
        frame_colores_2.grid(row=2, column=1, sticky="w", padx=10)
        self.btn_primario = tk.Button(frame_colores_2, bg="#1e6292", width=6, command=lambda: self.elegir_color(self.btn_primario))
        self.btn_primario.pack(side="left", padx=2)
        self.btn_acento = tk.Button(frame_colores_2, bg="#ee8001", width=6, command=lambda: self.elegir_color(self.btn_acento))
        self.btn_acento.pack(side="left", padx=2)

        # 4. Footer
        frame_footer = tk.LabelFrame(root, text="4. Personalizar Textos del Pie", padx=10, pady=10, bg="#fdfdfc")
        frame_footer.pack(fill="x", pady=5)
        self.txt_footer = tk.Text(frame_footer, height=3, width=65)
        self.txt_footer.insert("1.0", "Revista Académica\nDirector Fundador: Mario Rapoport | Directora: Noemí Brenta\nAcceso abierto bajo Licencia CC Atribución-NoComercial 4.0.")
        self.txt_footer.pack()

        # Botón
        tk.Button(root, text="¡GENERAR REVISTA!", bg="#1e6292", fg="white", font=("Arial", 14, "bold"), cursor="hand2", command=self.generar).pack(pady=15)

    def aplicar_tema(self, event=None):
        tema = self.combo_revista.get()
        colores = self.temas.get(tema, self.temas["CICLOS"])
        self.btn_bg.config(bg=colores["fondo"])
        self.btn_text.config(bg=colores["texto"])
        self.btn_primario.config(bg=colores["primario"])
        self.btn_acento.config(bg=colores["acento"])
        self.combo_fuente.set(colores["fuente"])

    def elegir_color(self, boton):
        color_elegido = colorchooser.askcolor(title="Elegir Color")
        if color_elegido and color_elegido[1]:
            boton.config(bg=color_elegido[1])

    def cargar_word(self):
        ruta = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        if ruta:
            self.word_path = Path(ruta)
            self.lbl_word.config(text=self.word_path.name, fg="green")

    def cargar_logo(self):
        ruta = filedialog.askopenfilename(filetypes=[("Imágenes", "*.png;*.jpg;*.jpeg")])
        if ruta:
            self.logo_path = Path(ruta)
            self.lbl_logo.config(text=self.logo_path.name, fg="green")

    def generar(self):
        if not self.word_path:
            messagebox.showwarning("Atención", "Debes subir un archivo Word.")
            return

        try:
            # 1. Directorios
            raw_name = self.ent_titulo.get() if self.ent_titulo.get() else self.word_path.stem
            safe_name = re.sub(r'[\\/*?:"<>|]', "", raw_name)
            nombre_carpeta = safe_name.replace(" ", "_")[:40]
            
            out_dir = self.word_path.parent / nombre_carpeta
            img_dir = out_dir / "public" / "img"
            article_img_dir = img_dir / "article"
            
            out_dir.mkdir(parents=True, exist_ok=True)
            article_img_dir.mkdir(parents=True, exist_ok=True)

            # 2. Copiar imágenes locales automáticamente
            base_script_dir = Path(__file__).parent.resolve()
            local_img_dir = base_script_dir / "public" / "img"
            if local_img_dir.exists():
                for item in local_img_dir.iterdir():
                    if item.is_file():
                        shutil.copy(item, img_dir / item.name)

            tema = self.combo_revista.get()
            
            # Lógica Transparente y Responsiva para el Logo
            logo_html = f'<span class="brand-text">{tema}</span>'
            if tema == "CICLOS" and (local_img_dir / "LogoCiclos.png").exists() and not self.logo_path:
                logo_html = '<img src="public/img/LogoCiclos.png" alt="Revista CICLOS" class="header-logo">'

            if self.logo_path:
                destino_logo = img_dir / self.logo_path.name
                shutil.copy(self.logo_path, destino_logo)
                logo_rel_path = f"public/img/{self.logo_path.name}"
                logo_html = f'<img src="{html.escape(logo_rel_path)}" alt="Logo Revista" class="header-logo">'

            # --- MOTOR DE EXTRACCIÓN ---
            W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
            R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
            A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

            z = zipfile.ZipFile(self.word_path)
            rels = {}
            rels_root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
            for rel in rels_root:
                rid = rel.attrib.get("Id")
                target = rel.attrib.get("Target")
                if rid and target:
                    rels[rid] = target

            media_map = {}
            for name in z.namelist():
                if name.startswith("word/media/"):
                    src_name = Path(name).name
                    data = z.read(name)
                    (article_img_dir / src_name).write_bytes(data)
                    media_map["media/" + src_name] = f"public/img/article/{src_name}"

            for emf_src in article_img_dir.glob("*.emf"):
                emf_png = emf_src.with_suffix(".png")
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
                        media_map["media/" + emf_src.name] = f"public/img/article/{emf_png.name}"
                except Exception:
                    pass

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

            def render_run(run):
                parts = []
                fnref = run.find(W + "footnoteReference")
                if fnref is not None:
                    fid = int(fnref.get(W + "id"))
                    used_footnotes.add(fid)
                    parts.append(f'<sup class="footnote-ref"><a href="#fn{fid}" id="fnref{fid}">{fid}</a></sup>')
                    if fnref.get(W + "customMarkFollows") == "1": return "".join(parts)
                for child in run:
                    if child.tag == W + "t": parts.append(autolink(html.escape(child.text or "")))
                    elif child.tag == W + "tab": parts.append(" ")
                    elif child.tag == W + "br": parts.append("<br>")
                text = "".join(parts)
                if not text: return ""
                rpr = run.find(W + "rPr")
                if rpr is not None and (rpr.find(W + "i") is not None or rpr.find(W + "iCs") is not None): text = f"<em>{text}</em>"
                if rpr is not None and (rpr.find(W + "b") is not None or rpr.find(W + "bCs") is not None): text = f"<strong>{text}</strong>"
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

            def render_table(tbl):
                try:
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
                except Exception:
                    return ""

            title_html, english_title, author_html, affiliation = "", "", "", ""
            abstract_es, keywords_es, abstract_en, keywords_en = "", "", "", ""

            article_parts = []
            toc_links = []
            references_open = False
            in_body = False

            for i, c in enumerate(body):
                text = plain_text(c)
                imgs = paragraph_images(c)
                html_text = render_inline(c)
                kind = c.tag.split("}", 1)[-1]
                
                if kind == "sectPr" or (not text and not imgs and kind != "tbl"): continue

                if kind == "tbl":
                    article_parts.append(render_table(c))
                    continue
                    
                if not title_html and len(text) > 10 and not text.startswith("Resumen"):
                    title_html = html_text
                    continue
                if title_html and not english_title and len(text) > 10 and not text.startswith("Resumen"):
                    english_title = html.escape(text)
                    continue
                if english_title and not author_html and len(text) > 3 and not text.startswith("Resumen"):
                    author_html = html_text
                    affiliation = footnote_text.get(2, "") 
                    continue
                    
                if text.startswith("Resumen"):
                    abstract_es = render_inline(body[i+1]) if i+1 < len(body) else ""
                    continue
                if text.startswith("Palabras clave"):
                    keywords_es = html_text.replace("Palabras clave:", "").strip()
                    continue
                if text.startswith("Abstract") or text.startswith("Résumé"):
                    abstract_en = render_inline(body[i+1]) if i+1 < len(body) else ""
                    continue
                if text.startswith("Key words") or text.startswith("Keywords") or text.startswith("Mots-clefs"):
                    keywords_en = html_text.replace("Key words:", "").replace("Keywords:", "").replace("Mots-clefs:", "").strip()
                    in_body = True 
                    continue

                if not in_body: continue 
                
                if text.lower() in ["listado de referencias", "referencias", "bibliografía"]:
                    article_parts.append('<section class="references" id="referencias"><h3>Listado de referencias</h3>')
                    references_open = True
                    toc_links.append('<a href="#referencias" class="toc-link">✦ Referencias</a>')
                    continue
                    
                if len(text) < 150 and ("<strong>" in html_text or "<h3>" in html_text) and not references_open:
                    clean_title = re.sub(r"<(/?strong|/?em)>", "", html_text)
                    sec_id = f"sec-{i}"
                    article_parts.append(f'<h3 id="{sec_id}">{clean_title}</h3>')
                    toc_text_clean = re.sub(r"<[^>]+>", "", clean_title)
                    toc_links.append(f'<a href="#{sec_id}" class="toc-link">✦ {toc_text_clean}</a>')
                    continue

                if imgs:
                    out = ['<figure class="article-figure"><div class="figure-media figure-grid">']
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
                footnote_items.append(f'<li id="fn{fid}"><p>{text_fn} <a class="footnote-back" href="#fnref{fid}">↩</a></p></li>')
            footnotes_html = '<section class="footnotes" id="notas"><h3>Notas</h3><ol>' + "".join(footnote_items) + "</ol></section>" if footnote_items else ""
            contenido_final = "\n".join(article_parts) + "\n" + footnotes_html

            titulo_final = self.ent_titulo.get() or title_html
            autor_final = self.ent_autor.get() or author_html
            titulo_head_limpio = re.sub(r"<[^>]+>", "", titulo_final)[:50]

            nav_html = """
            <nav class="top-nav">
                <a href="https://ojs.economicas.uba.ar/revistaCICLOS" class="active-link">Actual</a>
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
            """

            # ========================================================
            # SECCIÓN: PLANTILLA EXACTA DE CICLOS (Intocable, ancho 850px)
            # ========================================================
            if tema == "CICLOS":
                html_documento_final = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta http-equiv="Content-Language" content="ES">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>CICLOS | {html.escape(titulo_head_limpio)}...</title>
<meta name="robots" content="index,follow">
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.9.1/font/bootstrap-icons.css">
<link rel="stylesheet" href="style.css">
</head>
<body>
<header class="site-header">
<div class="header-inner">
<div class="logo-area">
<a href="https://ojs.economicas.uba.ar/revistaCICLOS" class="brand-link">
{logo_html}
</a>
</div>
{nav_html}
</div>
</header>
<main class="article-wrapper">
<div class="metadata-bar">
<span class="badge">Artículos</span>
<span class="badge badge-outline">2025</span>
<span class="sep">|</span>
<span class="doi-link">eISSN 1851-3735</span>
</div>
<header class="article-heading">
<h1 class="title">{titulo_final}</h1>
<div class="subtitle-group">
<h2 class="subtitle">{english_title}</h2>
</div>
<div class="author-block">
<p class="author-name">{autor_final}</p>
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
<div class="abstract-lang" lang="en">
<h3>Abstract</h3>
<p>{abstract_en}</p>
<p class="keywords"><strong>Key words:</strong> {keywords_en}</p>
</div>
</section>
<article class="article-content">
{contenido_final}
</article>
</main>
<footer class="site-footer">
<div class="footer-inner">
<div class="footer-col">
<h4>Revista CICLOS</h4>
<p style="white-space: pre-wrap; line-height: 1.8;">{html.escape(self.txt_footer.get("1.0", "end-1c"))}</p>
<p class="footer-meta"><i class="bi bi-geo-alt"></i> Av. Córdoba 2122, 2do. Piso, CABA</p>
<p class="footer-meta"><i class="bi bi-envelope"></i> <a href="mailto:ciclos@economicas.uba.ar">ciclos@economicas.uba.ar</a></p>
</div>
<div class="footer-col">
<h4>Sponsors e Institucional</h4>
<div class="academic-logos">
<a href="https://www.economicas.uba.ar" target="_blank" rel="noopener">
<img src="public/img/fce.png" alt="Logo FCE UBA" class="logo-fce" onerror="this.style.display='none'">
</a>
</div>
<p class="institution-text"><b>Propietario:</b> Universidad de Buenos Aires. Facultad de Ciencias Económicas. Centro de Investigaciones de Historia Económica, Social y de Relaciones Internacionales.</p>
<p class="institution-text">CIHESRI - IDEHESI (Unidad Ejecutora en Red del CONICET). Maestría en Historia Económica y de las Políticas Económicas, FCE-UBA.</p>
</div>
<div class="footer-col">
<h4>Políticas y Canales</h4>
<div class="license-area">
<a href="http://creativecommons.org/licenses/by-nc-sa/4.0/" rel="license" target="_blank">
<img src="public/img/licence.png" alt="Licencia CC" class="cc-badge" onerror="this.style.display='none'">
</a>
<p>Acceso abierto inmediato bajo Licencia CC Atribución-NoComercial 4.0. Sistema de preservación digital LOCKSS (Stanford University).</p>
</div>
</div>
</div>
</footer>
</body>
</html>"""

                css_documento_final = f"""/* =========================================
Variables de Diseño (Identidad CICLOS)
========================================= */
:root {{
--blue-primary: {self.btn_primario.cget('bg')};
--blue-dark: {self.btn_primario.cget('bg')};
--blue-light: #eef5fb;
--orange-accent: {self.btn_acento.cget('bg')};
--text-main: {self.btn_text.cget('bg')};
--text-muted: #5a6a7a;
--bg-page: {self.btn_bg.cget('bg')};
--white: #ffffff;
--border-color: #cbd5e1;
--font-heading: '{self.combo_fuente.get()}', Georgia, serif;
--font-ui: "Inter", sans-serif;
--font-body: '{self.combo_fuente.get()}', serif;
}}

* {{ margin: 0; padding: 0; box-sizing: border-box; }}
html {{ scroll-behavior: smooth; scroll-padding-top: 100px; }}
body {{ font-family: var(--font-body); background-color: var(--bg-page); color: var(--text-main); line-height: 1.85; font-size: 1.08rem; -webkit-font-smoothing: antialiased; }}
a {{ color: var(--blue-primary); text-decoration: none; transition: color 0.2s ease; }}
a:hover {{ color: var(--orange-accent); }}
.article-content a {{ overflow-wrap: anywhere; word-break: break-word; }}

.site-header {{ background-color: var(--blue-primary); border-top: 4px solid var(--orange-accent); padding: 0.8rem 0; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06); position: sticky; top: 0; z-index: 100; }}
.header-inner {{ max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; }}
.brand-link {{ display: inline-flex; align-items: center; color: var(--white); gap: 15px; }}
.brand-text {{ color: var(--white); font-family: var(--font-heading); font-size: 2rem; font-weight: 700; letter-spacing: 0; }}
.header-logo {{ height: 75px; width: auto; display: block; object-fit: contain; }}

.top-nav {{ display: flex; gap: 1.8rem; align-items: center; }}
.top-nav a {{ color: #e2e8f0; font-family: var(--font-ui); font-size: 0.85rem; font-weight: 600; text-transform: uppercase; position: relative; padding-bottom: 4px; transition: color 0.2s ease; letter-spacing: 0.5px; }}
.top-nav a::after {{ content: ""; position: absolute; width: 0; height: 2px; bottom: 0; left: 0; background-color: var(--orange-accent); transition: width 0.25s ease-in-out; }}
.top-nav a:hover::after, .top-nav a.active-link::after {{ width: 100%; }}
.dropdown {{ position: relative; display: flex; align-items: center; height: 100%; }}
.dropdown-toggle {{ display: flex; align-items: center; gap: 5px; }}
.dropdown-menu {{ display: none !important; position: absolute; background-color: var(--white); min-width: 240px; box-shadow: 0px 8px 16px rgba(0, 0, 0, 0.15); z-index: 999; top: 100%; left: 0; padding: 0.5rem 0; border-top: 3px solid var(--orange-accent); flex-direction: column; }}
.dropdown:hover .dropdown-menu {{ display: flex !important; }}
.top-nav .dropdown-menu a {{ color: var(--blue-dark) !important; padding: 10px 16px !important; text-transform: none !important; font-weight: 500 !important; font-size: 0.95rem !important; width: 100%; display: block !important; }}
.top-nav .dropdown-menu a::after {{ display: none !important; }}
.top-nav .dropdown-menu a:hover {{ background-color: var(--blue-light) !important; color: var(--blue-primary) !important; }}

/* Ancho super cómodo de 850px para lectura en PC */
.article-wrapper {{ max-width: 850px; margin: 4rem auto; padding: 0 2rem; }}

.metadata-bar {{ display: flex; align-items: center; flex-wrap: wrap; gap: 0.8rem; margin-bottom: 2rem; font-family: var(--font-ui); font-size: 0.75rem; font-weight: 700; }}
.badge {{ background-color: var(--blue-primary); color: var(--white); padding: 0.3rem 0.8rem; border-radius: 4px; text-transform: uppercase; }}
.badge-outline {{ background-color: transparent; color: var(--blue-primary); border: 1px solid var(--blue-primary); }}
.doi-link {{ color: var(--text-muted); }}

.title {{ font-family: var(--font-heading); font-size: 2.5rem; font-weight: 700; color: var(--blue-dark); line-height: 1.25; margin-bottom: 1.2rem; }}
.title sup, .author-name sup {{ font-family: var(--font-ui); font-size: 0.55em; line-height: 0; }}
.subtitle-group {{ display: grid; gap: 0.55rem; margin-bottom: 2rem; }}
.subtitle {{ font-family: var(--font-heading); font-size: 1.3rem; color: var(--text-muted); font-style: italic; font-weight: 400; line-height: 1.4; }}
.author-block {{ border-left: 4px solid var(--orange-accent); padding-left: 1rem; margin-bottom: 2rem; }}
.author-name {{ font-family: var(--font-ui); font-size: 1.05rem; font-weight: 700; color: var(--blue-primary); }}
.author-affiliation, .article-dates {{ font-family: var(--font-ui); font-size: 0.86rem; color: var(--text-muted); margin-top: 0.25rem; line-height: 1.55; }}

.abstract-box {{ background-color: var(--blue-light); padding: 2.5rem; border-radius: 12px; margin: 3.5rem 0; border: 1px solid #d0e1f0; }}
.abstract-lang h3 {{ font-family: var(--font-ui); color: var(--blue-dark); text-transform: uppercase; font-size: 0.85rem; margin-bottom: 0.8rem; font-weight: 700; letter-spacing: 0.5px; }}
.abstract-lang p {{ font-size: 0.95rem; color: #334155; }}
.abstract-divider {{ border: 0; height: 1px; background-color: #cbd5e1; margin: 2rem 0; }}
.keywords {{ margin-top: 1rem; font-family: var(--font-ui); font-size: 0.82rem !important; color: var(--text-main); }}

.article-content h3, .article-content h4 {{ font-family: var(--font-heading); color: var(--blue-dark); font-weight: 700; line-height: 1.3; }}
.article-content h3 {{ font-size: 1.7rem; margin: 3.6rem 0 1.2rem; }}
.article-content h4 {{ font-size: 1.35rem; margin: 2.7rem 0 1rem; }}
.article-content p {{ margin-bottom: 1.55rem; text-align: justify; }}
.footnote-ref a {{ font-family: var(--font-ui); font-weight: 700; padding: 0 0.08rem; }}

.article-figure {{ margin: 2.6rem 0; padding: 1.5rem; background-color: var(--white); border: 1px solid var(--border-color); border-radius: 8px; }}
.article-figure figcaption {{ font-family: var(--font-ui); color: var(--blue-dark); line-height: 1.45; margin-bottom: 1rem; }}
.figure-label {{ display: block; font-size: 0.78rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.5px; color: var(--orange-accent); margin-bottom: 0.25rem; }}
.figure-title {{ display: block; font-size: 0.98rem; font-weight: 700; }}
.figure-media {{ display: grid; justify-items: center; gap: 1rem; }}
.figure-grid {{ grid-template-columns: repeat(auto-fit, minmax(260px, 1fr)); align-items: start; }}
.figure-media img {{ display: block; max-width: 100%; height: auto; border: 1px solid #e2e8f0; background-color: var(--white); }}
.figure-source {{ font-family: var(--font-ui); font-size: 0.78rem; color: var(--text-muted); line-height: 1.55; margin: 1rem 0 0 !important; text-align: left !important; }}

.table-scroll {{ overflow-x: auto; margin-top: 1rem; }}
.data-table {{ width: 100%; border-collapse: collapse; font-family: var(--font-ui); font-size: 0.86rem; line-height: 1.45; background-color: var(--white); }}
.data-table th, .data-table td {{ border: 1px solid var(--border-color); padding: 0.65rem 0.75rem; vertical-align: top; text-align: right; }}
.data-table th:first-child, .data-table td:first-child {{ text-align: left; }}
.data-table thead th {{ background-color: var(--blue-light); color: var(--blue-dark); font-weight: 700; }}

.references {{ margin-top: 3.5rem; }}
.references .reference-item {{ font-size: 0.92rem; line-height: 1.65; margin-bottom: 0.85rem; text-align: left; }}
.footnotes {{ border-top: 1px solid var(--border-color); margin-top: 3.5rem; padding-top: 1rem; }}
.footnotes h3 {{ margin-top: 1rem; }}
.footnotes ol {{ padding-left: 1.35rem; }}
.footnotes li {{ padding-left: 0.25rem; margin-bottom: 0.75rem; }}
.footnotes p {{ font-family: var(--font-ui); font-size: 0.82rem; line-height: 1.55; text-align: left; margin-bottom: 0; }}
.footnote-back {{ font-weight: 700; }}

blockquote, .academic-quote {{ font-size: 1.35rem; font-family: var(--font-heading); font-style: italic; color: var(--blue-primary); margin: 3.5rem 0; padding: 1.5rem 2.5rem; border-left: 4px solid var(--orange-accent); background-color: var(--white); box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02); border-radius: 0 8px 8px 0; }}

.site-footer {{ background-color: var(--blue-primary); color: #cbd5e1; padding: 3rem 0; font-family: var(--font-ui); font-size: 0.8rem; margin-top: 6rem; border-top: 5px solid var(--blue-dark); }}
.footer-inner {{ max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: grid; grid-template-columns: repeat(3, 1fr); gap: 3rem; }}
.footer-col h4 {{ color: var(--white); margin-bottom: 1.2rem; font-size: 0.85rem; text-transform: uppercase; border-bottom: 2px solid rgba(255, 255, 255, 0.1); padding-bottom: 0.5rem; letter-spacing: 0.5px; }}
.footer-col p {{ margin-bottom: 0.6rem; line-height: 1.6; }}
.footer-col b, .footer-col a {{ color: var(--white); text-decoration: underline; }}
.academic-logos {{ display: flex; align-items: center; gap: 12px; margin-bottom: 1.2rem; }}
.logo-fce {{ display: block; height: 42px; width: auto; background-color: white; padding: 6px 10px; border-radius: 4px; transition: transform 0.2s ease; }}
.logo-fce:hover {{ transform: translateY(-2px); }}
.institution-text {{ font-size: 0.75rem; color: #d4deea; line-height: 1.5; }}
.license-area {{ margin-bottom: 1.2rem; }}
.cc-badge {{ margin-bottom: 0.4rem; display: block; }}

/* Diseño Responsivo Original Restaurado */
@media screen and (max-width: 900px) {{
  .footer-inner {{ grid-template-columns: 1fr; gap: 2rem; }}
  .header-inner {{ flex-direction: column; gap: 1.2rem; }}
  .top-nav {{ width: 100%; justify-content: space-between; gap: 1rem; flex-wrap: wrap; }}
  .article-wrapper {{ margin: 2.5rem auto; padding: 0 1.2rem; width: 100%; }}
  .title {{ font-size: 1.9rem; }}
  .subtitle {{ font-size: 1.06rem; }}
  .abstract-box {{ padding: 1.4rem; }}
  .article-figure {{ padding: 1rem; }}
  .figure-grid {{ grid-template-columns: 1fr; }}
}}
"""
            # ========================================================
            # SECCIÓN: PLANTILLAS SECUNDARIAS (Erudita, Prisma, etc.)
            # ========================================================
            else:
                footer_html = f"""
                <footer class="site-footer">
                    <div class="footer-inner">
                        <div class="footer-col">
                            <h4>Acerca de la Publicación</h4>
                            <p style="white-space: pre-wrap; line-height: 1.8;">{html.escape(self.txt_footer.get("1.0", "end-1c"))}</p>
                        </div>
                        <div class="footer-col">
                            <h4>Institucional</h4>
                            <div class="academic-logos">
                                <img src="public/img/fce.png" alt="Logo FCE" class="logo-fce" onerror="this.style.display='none'">
                            </div>
                            <p class="institution-text">Universidad / Facultad / Centro de Investigaciones.</p>
                        </div>
                        <div class="footer-col">
                            <h4>Políticas y Canales</h4>
                            <div class="license-area">
                                <img src="public/img/licence.png" alt="Licencia CC" class="cc-badge" onerror="this.style.display='none'">
                                <p>Sistema de preservación digital. Acceso abierto inmediato.</p>
                            </div>
                        </div>
                    </div>
                </footer>
                """

                cuerpo_html = ""
                css_extra = ""
                toc_html = "".join(toc_links) if toc_links else "<p>No hay secciones detectadas.</p>"

                # NUEVO REDISEÑO DE ERUDITA (Estilo SciELO/Redalyc Clásico Moderno)
                if tema == "Erudita":
                    cuerpo_html = f"""
                    <header class="site-header"><div class="header-inner"><div class="logo-area"><a href="#" class="brand-link">{logo_html}</a></div>{nav_html}</div></header>
                    <div class="erudita-page-bg">
                        <main class="erudita-paper article-content">
                            <header class="erudita-header-section">
                                <h1 class="erudita-title">{titulo_final}</h1>
                                <h2 class="erudita-subtitle">{english_title}</h2>
                                <div class="erudita-authors">
                                    <p class="erudita-author">{autor_final}</p>
                                    <p class="erudita-affil">{affiliation}</p>
                                </div>
                            </header>
                            <div class="erudita-abstracts">
                                <div class="erudita-abs"><h3>Resumen</h3><p>{abstract_es}</p><p><b>Palabras clave:</b> {keywords_es}</p></div>
                                <div class="erudita-abs"><h3>Abstract</h3><p>{abstract_en}</p><p><b>Keywords:</b> {keywords_en}</p></div>
                            </div>
                            <div class="erudita-body">
                                {contenido_final}
                            </div>
                        </main>
                    </div>
                    """
                    css_extra = """
                    .erudita-page-bg { background-color: #f3f4f6; padding: 3rem 1rem; }
                    .erudita-paper { background-color: var(--white); max-width: 850px; margin: 0 auto; padding: 4rem 5rem; box-shadow: 0 10px 25px rgba(0,0,0,0.08); border-top: 6px solid var(--blue-primary); }
                    .erudita-header-section { text-align: center; margin-bottom: 3rem; border-bottom: 1px solid var(--border-color); padding-bottom: 2rem; }
                    .erudita-title { font-family: var(--font-heading); font-size: 2.4rem; color: var(--blue-dark); line-height: 1.25; margin-bottom: 1rem; }
                    .erudita-subtitle { font-family: var(--font-heading); font-size: 1.2rem; color: var(--text-muted); font-style: italic; margin-bottom: 1.5rem; }
                    .erudita-author { font-size: 1.15rem; font-weight: bold; color: var(--blue-primary); margin-bottom: 0.3rem; }
                    .erudita-affil { font-size: 0.9rem; color: var(--text-muted); }
                    .erudita-abstracts { display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; background-color: var(--blue-light); padding: 2rem; margin-bottom: 3rem; border-radius: 6px; }
                    .erudita-abs h3 { font-size: 1rem; text-transform: uppercase; color: var(--blue-dark); margin-bottom: 0.8rem; border-bottom: 2px solid var(--blue-primary); display: inline-block; padding-bottom: 0.2rem; margin-top: 0; }
                    .erudita-abs p { font-size: 0.9rem; line-height: 1.6; margin-bottom: 0.5rem; text-align: justify; }
                    .erudita-body { text-align: justify; }
                    """
                elif tema == "Prisma":
                    cuerpo_html = f"""
                    <header class="site-header"><div class="header-inner"><div class="logo-area"><a href="#" class="brand-link">{logo_html}</a></div>{nav_html}</div></header>
                    <div class="prisma-hero">
                        <h1 class="prisma-title">{titulo_final}</h1><h2 class="prisma-subtitle">{english_title}</h2><p class="prisma-author">{autor_final}</p>
                        <div class="prisma-abstracts-grid">
                            <div class="prisma-box"><h3>Resumen</h3><p>{abstract_es}</p><p class="keywords"><b>Palabras clave:</b> {keywords_es}</p></div>
                            <div class="prisma-box"><h3>Abstract</h3><p>{abstract_en}</p><p class="keywords"><b>Keywords:</b> {keywords_en}</p></div>
                        </div>
                    </div>
                    <div class="prisma-layout">
                        <main class="prisma-content article-content">{contenido_final}</main>
                        <aside class="prisma-toc">
                            <div class="toc-box"><h3>Índice</h3>{toc_html}</div>
                        </aside>
                    </div>
                    """
                    css_extra = """
                    .prisma-hero { background: var(--blue-light); padding: 4rem 2rem; text-align: center; border-bottom: 4px solid var(--orange-accent); }
                    .prisma-title { font-family: var(--font-heading); font-size: 2.8rem; color: var(--blue-dark); margin-bottom: 1rem; max-width: 900px; margin-left: auto; margin-right: auto; }
                    .prisma-subtitle { font-family: var(--font-heading); font-size: 1.3rem; color: var(--text-muted); font-style: italic; margin-bottom: 2rem; }
                    .prisma-author { font-size: 1.2rem; font-weight: bold; color: var(--blue-primary); }
                    .prisma-abstracts-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 2rem; max-width: 1000px; margin: 3rem auto 0; text-align: left; }
                    .prisma-box { background: var(--white); padding: 2rem; border-radius: 8px; box-shadow: 0 4px 6px rgba(0,0,0,0.05); }
                    .prisma-layout { display: flex; max-width: 1100px; margin: 3rem auto; gap: 3rem; padding: 0 2rem; }
                    .prisma-content { flex: 3; max-width: 850px; }
                    .prisma-toc { flex: 1; position: sticky; top: 120px; height: max-content; }
                    .toc-box { background: var(--bg-page); border-left: 4px solid var(--blue-primary); padding: 1.5rem; }
                    .toc-link { display: block; color: var(--blue-dark); margin-bottom: 0.8rem; font-weight: 500; font-family: var(--font-ui); text-decoration: none; transition: color 0.2s; font-size: 0.9rem; }
                    .toc-link:hover { color: var(--orange-accent); }
                    """
                elif tema == "Vanguardia":
                    cuerpo_html = f"""
                    <div class="vanguardia-hero">
                        <header class="site-header vanguardia-nav"><div class="header-inner"><div class="logo-area"><a href="#" class="brand-link">{logo_html}</a></div>{nav_html}</div></header>
                        <div class="vanguardia-title-box"><h1>{titulo_final}</h1><h2>{english_title}</h2><p>{autor_final}</p></div>
                    </div>
                    <main class="vanguardia-main article-content">
                        <div class="vanguardia-abs"><h3>Resumen</h3><p>{abstract_es}</p></div>
                        <div class="vanguardia-abs"><h3>Abstract</h3><p>{abstract_en}</p></div>
                        {contenido_final}
                    </main>
                    """
                    css_extra = """
                    .vanguardia-hero { min-height: 70vh; background: var(--blue-primary); color: var(--white); display: flex; flex-direction: column; }
                    .vanguardia-nav { box-shadow: none; border: none; }
                    .vanguardia-title-box { flex: 1; display: flex; flex-direction: column; justify-content: center; align-items: center; text-align: center; padding: 3rem 2rem; }
                    .vanguardia-title-box h1 { font-family: var(--font-heading); font-size: 3.5rem; line-height: 1.1; margin-bottom: 1rem; max-width: 1000px; color: var(--white); }
                    .vanguardia-title-box h2 { font-size: 1.5rem; font-weight: 300; opacity: 0.8; margin-bottom: 2rem; color: var(--white); }
                    .vanguardia-title-box p { font-size: 1.3rem; font-weight: bold; color: var(--orange-accent); }
                    .vanguardia-main { max-width: 850px; margin: -4rem auto 4rem; background: var(--bg-page); padding: 4rem; border-radius: 12px; box-shadow: 0 10px 30px rgba(0,0,0,0.1); }
                    .vanguardia-abs { margin-bottom: 2rem; padding-bottom: 2rem; border-bottom: 1px solid var(--border-color); }
                    """
                elif tema == "Lienzo":
                    cuerpo_html = f"""
                    <header class="site-header"><div class="header-inner"><div class="logo-area"><a href="#" class="brand-link">{logo_html}</a></div>{nav_html}</div></header>
                    <main class="lienzo-main article-content">
                        <h1 class="lienzo-title">{titulo_final}</h1><p class="lienzo-author"><b>{autor_final}</b></p><p class="lienzo-affil">{affiliation}</p>
                        <hr class="lienzo-hr">
                        <h3>Resumen</h3><p>{abstract_es}</p><p class="keywords"><b>Palabras clave:</b> {keywords_es}</p><br>
                        <h3>Abstract</h3><p>{abstract_en}</p><p class="keywords"><b>Keywords:</b> {keywords_en}</p>
                        <hr class="lienzo-hr">
                        {contenido_final}
                    </main>
                    """
                    css_extra = """
                    .site-footer { background: var(--bg-page); color: var(--text-main); border-top: 1px solid #ddd; }
                    .site-footer h4, .site-footer p, .site-footer a, .site-footer b { color: var(--text-main); }
                    .lienzo-main { max-width: 850px; margin: 4rem auto; padding: 0 2rem 5rem; }
                    .lienzo-title { font-family: var(--font-heading); font-size: 2.6rem; color: var(--blue-primary); line-height: 1.2; margin-bottom: 1rem; text-align: center; }
                    .lienzo-author { text-align: center; font-size: 1.2rem; }
                    .lienzo-affil { text-align: center; font-style: italic; color: var(--text-muted); margin-bottom: 2rem; }
                    .lienzo-hr { border: 0; border-top: 1px solid #ddd; margin: 3rem 0; }
                    """

                html_documento_final = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{tema} | {html.escape(titulo_head_limpio)}...</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&family=Roboto:wght@400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.9.1/font/bootstrap-icons.css">
<link rel="stylesheet" href="style.css">
</head>
<body>
{cuerpo_html}
{footer_html}
</body>
</html>"""

                css_documento_final = f"""/* Variables Generales */
:root {{
--blue-primary: {self.btn_primario.cget('bg')};
--blue-dark: {self.btn_primario.cget('bg')};
--blue-light: #eef5fb;
--orange-accent: {self.btn_acento.cget('bg')};
--text-main: {self.btn_text.cget('bg')};
--text-muted: #5a6a7a;
--bg-page: {self.btn_bg.cget('bg')};
--white: #ffffff;
--border-color: #cbd5e1;
--font-heading: "{self.combo_fuente.get()}", serif;
--font-ui: "Inter", sans-serif;
--font-body: "{self.combo_fuente.get()}", serif;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
html {{ scroll-behavior: smooth; scroll-padding-top: 100px; }}
body {{ font-family: var(--font-body); background-color: var(--bg-page); color: var(--text-main); line-height: 1.85; font-size: 1.08rem; -webkit-font-smoothing: antialiased; }}
a {{ color: var(--blue-primary); text-decoration: none; transition: color 0.2s ease; }}
a:hover {{ color: var(--orange-accent); }}
.article-content a {{ overflow-wrap: anywhere; word-break: break-word; }}

/* REGLAS GLOBALES PARA ANCHO DE LECTURA ÓPTIMO */
.article-content {{
    max-width: 850px;
    margin: 0 auto;
}}
.article-content p {{
    font-size: 1.12rem;
    line-height: 1.8;
    margin-bottom: 1.55rem;
}}

/* Elementos comunes */
.article-content h3 {{ font-family: var(--font-heading); font-size: 1.7rem; margin: 3.6rem 0 1.2rem; color: var(--blue-primary); font-weight: 700; }}
.article-figure {{ margin: 2.6rem 0; padding: 1.5rem; background-color: var(--white); border: 1px solid var(--border-color); border-radius: 8px; }}
.figure-label {{ display: block; font-size: 0.78rem; font-weight: 800; text-transform: uppercase; color: var(--orange-accent); margin-bottom: 0.25rem; }}
.figure-title {{ display: block; font-size: 0.98rem; font-weight: 700; margin-bottom: 1rem; }}
.figure-media img {{ display: block; max-width: 100%; height: auto; }}
.data-table {{ width: 100%; border-collapse: collapse; font-family: var(--font-ui); font-size: 0.86rem; background-color: var(--white); margin: 2rem 0; }}
.data-table th, .data-table td {{ border: 1px solid var(--border-color); padding: 0.65rem 0.75rem; text-align: left; }}
.data-table thead th {{ background-color: var(--blue-light); color: var(--blue-dark); }}
.references .reference-item {{ font-size: 0.92rem; margin-bottom: 0.85rem; }}
.footnotes {{ border-top: 1px solid var(--border-color); margin-top: 3.5rem; padding-top: 1rem; }}
.footnotes ol {{ padding-left: 1.35rem; }}
.footnotes p {{ font-family: var(--font-ui); font-size: 0.82rem; margin-bottom: 0; }}
blockquote, .academic-quote {{ font-size: 1.35rem; font-family: var(--font-heading); font-style: italic; color: var(--blue-primary); margin: 3.5rem 0; padding: 1.5rem 2.5rem; border-left: 4px solid var(--orange-accent); background-color: var(--white); box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02); border-radius: 0 8px 8px 0; }}

/* Nav y Footer globales */
.site-header {{ background-color: var(--blue-primary); border-top: 4px solid var(--orange-accent); padding: 0.8rem 0; }}
.header-inner {{ max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: flex; justify-content: space-between; align-items: center; }}
.brand-link {{ display: inline-flex; align-items: center; color: var(--white); gap: 15px; }}
.brand-text {{ font-family: var(--font-heading); font-size: 2rem; font-weight: 700; color: var(--white); }}
/* Logo sin fondo blanco forzado */
.header-logo {{ height: 75px; width: auto; display: block; object-fit: contain; }}

.top-nav {{ display: flex; gap: 1.8rem; align-items: center; }}
.top-nav a {{ color: #e2e8f0; font-family: var(--font-ui); font-size: 0.85rem; font-weight: 600; text-transform: uppercase; position: relative; padding-bottom: 4px; }}
.top-nav a::after {{ content: ""; position: absolute; width: 0; height: 2px; bottom: 0; left: 0; background-color: var(--orange-accent); transition: width 0.25s ease-in-out; }}
.top-nav a:hover::after {{ width: 100%; }}
.dropdown {{ position: relative; display: flex; align-items: center; height: 100%; }}
.dropdown-menu {{ display: none !important; position: absolute; background-color: var(--white); min-width: 240px; box-shadow: 0 8px 16px rgba(0,0,0,0.15); top: 100%; left: 0; padding: 0.5rem 0; border-top: 3px solid var(--orange-accent); flex-direction: column; z-index: 999; }}
.dropdown:hover .dropdown-menu {{ display: flex !important; }}
.top-nav .dropdown-menu a {{ color: var(--blue-dark) !important; padding: 10px 16px !important; text-transform: none !important; width: 100%; }}

.site-footer {{ background-color: var(--blue-primary); color: #cbd5e1; padding: 3rem 0; font-family: var(--font-ui); font-size: 0.8rem; margin-top: 6rem; }}
.footer-inner {{ max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: grid; grid-template-columns: repeat(3, 1fr); gap: 3rem; }}
.footer-col h4 {{ color: var(--white); margin-bottom: 1.2rem; text-transform: uppercase; border-bottom: 2px solid rgba(255, 255, 255, 0.1); padding-bottom: 0.5rem; }}
.academic-logos img {{ height: 42px; background: white; padding: 6px; border-radius: 4px; }}

{css_extra}

/* =========================================
RESPONSIVE DESIGN (MÓVILES)
========================================= */
@media screen and (max-width: 900px) {{
  .footer-inner {{ grid-template-columns: 1fr; gap: 2rem; }}
  .header-inner {{ flex-direction: column; gap: 1.2rem; }}
  .top-nav {{ width: 100%; justify-content: center; gap: 1.5rem; flex-wrap: wrap; }}
  
  .article-wrapper {{ margin: 2.5rem auto; padding: 0 1.2rem; }}
  .title {{ font-size: 1.9rem; }}
  
  /* Nuevos ajustes de Erudita Móvil */
  .erudita-page-bg {{ padding: 1rem 0; background-color: var(--white); }}
  .erudita-paper {{ padding: 2rem 1.2rem; box-shadow: none; border-top: none; width: 100%; }}
  .erudita-title {{ font-size: 1.9rem; }}
  .erudita-abstracts {{ grid-template-columns: 1fr; padding: 1.5rem; }}

  .prisma-layout {{ flex-direction: column; padding: 0 1.2rem; margin: 1.5rem auto; }}
  .prisma-toc {{ display: none; }}
  .prisma-abstracts-grid {{ grid-template-columns: 1fr; padding: 0 1.2rem; }}
  .prisma-title {{ font-size: 2rem; }}

  .vanguardia-title-box h1 {{ font-size: 2.2rem; }}
  .vanguardia-main {{ margin: 0 auto; padding: 2rem 1.2rem; border-radius: 0; box-shadow: none; }}

  .lienzo-main {{ padding: 0 1.2rem 3rem; }}
}}
"""

            (out_dir / "index.html").write_text(html_documento_final, encoding="utf-8")
            (out_dir / "style.css").write_text(css_documento_final, encoding="utf-8")

            messagebox.showinfo("¡Éxito!", f"¡La revista se ha generado usando la estructura {tema}!\nSe guardó en:\n{out_dir}")

        except Exception as e:
            messagebox.showerror("Error al procesar", f"Ocurrió un error inesperado:\n{str(e)}")

if __name__ == "__main__":
    root = tk.Tk()
    app = WordlessApp(root)
    root.mainloop()