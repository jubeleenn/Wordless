# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, colorchooser, ttk
from pathlib import Path
import html
import re
import subprocess
import zipfile
import xml.etree.ElementTree as ET

# --- 1. DICCIONARIO DE TEMAS PREDETERMINADOS ---
temas = {
    "Clásico": {"fondo": "#fdfdfc", "texto": "#2b2b2b", "acento": "#1e6292", "fuente": "Merriweather"},
    "Modo Oscuro": {"fondo": "#121212", "texto": "#e0e0e0", "acento": "#ee8001", "fuente": "Inter"},
    "Moderno": {"fondo": "#f0f4f8", "texto": "#334155", "acento": "#10b981", "fuente": "Roboto"}
}

# --- 2. FUNCIONES DE LÓGICA VISUAL ---
def aplicar_tema(nombre_tema):
    tema = temas[nombre_tema]
    btn_color_fondo.config(bg=tema["fondo"])
    btn_color_texto.config(bg=tema["texto"])
    btn_color_acento.config(bg=tema["acento"])
    combo_fuentes.set(tema["fuente"])
    lbl_estado.config(text=f"Tema '{nombre_tema}' aplicado al panel.", fg="green")

def elegir_color(boton):
    color_elegido = colorchooser.askcolor(title="Elegí un color")[1]
    if color_elegido:
        boton.config(bg=color_elegido)
        lbl_estado.config(text="Color manual aplicado.", fg="blue")

# --- 3. LÓGICA DE EXTRACCIÓN Y GENERACIÓN (EL MOTOR) ---
def procesar_word():
    # 1. Leer las variables de diseño del Panel
    fuente_elegida = combo_fuentes.get()
    color_fondo = btn_color_fondo.cget("bg")
    color_texto = btn_color_texto.cget("bg")
    color_acento = btn_color_acento.cget("bg")

    # 2. Pedir el archivo al usuario
    file_path = filedialog.askopenfilename(
        title="Selecciona un archivo Word de la Revista CICLOS",
        filetypes=[("Documentos de Word", "*.docx")]
    )

    if not file_path:
        return

    lbl_estado.config(text="Procesando documento... aguarde.", fg="#ee8001")
    root.update()

    try:
        DOCX = Path(file_path)

        # Crear carpetas
        OUT = DOCX.parent / DOCX.stem
        IMG_DIR = OUT / "public" / "img"
        ARTICLE_IMG_DIR = IMG_DIR / "article"

        OUT.mkdir(parents=True, exist_ok=True)
        ARTICLE_IMG_DIR.mkdir(parents=True, exist_ok=True)

        # Extracción del Word
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
            pass

        media_map = {}
        for name in z.namelist():
            if name.startswith("word/media/"):
                src_name = Path(name).name
                data = z.read(name)
                out_path = ARTICLE_IMG_DIR / src_name
                out_path.write_bytes(data)
                
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
                    except Exception:
                        media_map["media/" + src_name] = f"public/img/article/{src_name}"
                else:
                    media_map["media/" + src_name] = f"public/img/article/{src_name}"

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
                out_imgs = ['<figure class="article-figure"><div class="figure-media">']
                for src in imgs: out_imgs.append(f'<img src="{html.escape(src)}" alt="Imagen del artículo" loading="lazy">')
                out_imgs.append("</div></figure>")
                article_parts.append("".join(out_imgs))
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

        title_plain = re.sub(r'<[^>]+>', '', title_html).replace('"', '&quot;')
        author_plain = re.sub(r'<[^>]+>', '', author_html).replace('"', '&quot;')

        # --- 4. INYECCIÓN DINÁMICA DE LA PLANTILLA ---
        # Formateamos la fuente para la URL de Google Fonts (ej. "Playfair Display" -> "Playfair+Display")
        fuente_url = fuente_elegida.replace(" ", "+")

        html_doc = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Artículo Dinámico | {title_plain[:40]}...</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family={fuente_url}:wght@400;500;700&display=swap" rel="stylesheet">
<link rel="stylesheet" href="style.css">
</head>
<body>
<main class="article-wrapper">
<header class="article-heading">
<h1 class="title">{title_html}</h1>
<h2 class="subtitle">{english_title}</h2>
<div class="author-block">
<p class="author-name">{author_html}</p>
<p class="author-affiliation">{affiliation}</p>
</div>
</header>
<article class="article-content">{article_html}</article>
</main>
</body>
</html>
"""

        # Acá sucede la magia: inyectamos los colores del panel al CSS
        css_doc = f"""
:root {{
    --bg-page: {color_fondo};
    --text-main: {color_texto};
    --blue-primary: {color_acento};
    --orange-accent: {color_acento};
    --blue-dark: {color_texto};
    --white: {color_fondo};
    --border-color: {color_acento};
    --font-heading: '{fuente_elegida}', serif;
    --font-ui: "Arial", sans-serif;
    --font-body: '{fuente_elegida}', sans-serif;
}}

* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ 
    font-family: var(--font-body); 
    background-color: var(--bg-page); 
    color: var(--text-main); 
    line-height: 1.85; 
    font-size: 1.1rem; 
    padding: 2rem;
}}
a {{ color: var(--blue-primary); text-decoration: none; }}
.article-wrapper {{ max-width: 1000px; margin: 0 auto; }}
.title {{ font-family: var(--font-heading); font-size: 2.5rem; font-weight: 700; color: var(--blue-primary); margin-bottom: 1rem; }}
.subtitle {{ font-family: var(--font-heading); font-size: 1.3rem; font-style: italic; opacity: 0.8; margin-bottom: 2rem; }}
.author-block {{ border-left: 4px solid var(--orange-accent); padding-left: 1rem; margin-bottom: 3rem; }}
.author-name {{ font-weight: 700; color: var(--blue-primary); }}
.article-content h3 {{ font-family: var(--font-heading); font-size: 1.7rem; color: var(--blue-primary); margin: 3rem 0 1rem; font-weight: 700; }}
.article-content p {{ margin-bottom: 1.5rem; text-align: justify; }}
.article-figure {{ margin: 2rem 0; padding: 1rem; border: 1px solid var(--border-color); }}
.figure-media img {{ display: block; max-width: 100%; height: auto; }}
.academic-quote {{ font-size: 1.2rem; font-style: italic; color: var(--text-main); margin: 2rem 0; padding: 1.5rem; border-left: 4px solid var(--orange-accent); }}
"""

        (OUT / "index.html").write_text(html_doc, encoding="utf-8")
        (OUT / "style.css").write_text(css_doc, encoding="utf-8")
        
        lbl_estado.config(text="¡Conversión dinámica finalizada!", fg="green")
        messagebox.showinfo("¡Éxito!", f"¡Listo! El artículo se adaptó al diseño y se guardó en:\n{OUT}")

    except Exception as e:
        lbl_estado.config(text="Error durante la conversión.", fg="red")
        messagebox.showerror("Error", f"Ocurrió un problema: {e}")

# --- 5. INTERFAZ GRÁFICA PRINCIPAL ---
root = tk.Tk()
root.title("Wordless | Generador Dinámico")
root.geometry("450x480")
root.configure(bg="#fdfdfc")

tk.Label(root, text="Personalizar Diseño", font=("Arial", 16, "bold"), bg="#fdfdfc", fg="#012662").pack(pady=(15, 5))

# --- SECCIÓN: Temas Rápidos ---
frame_temas = tk.LabelFrame(root, text="Temas Rápidos", bg="#fdfdfc", padx=10, pady=10)
frame_temas.pack(fill="x", padx=20, pady=10)

tk.Button(frame_temas, text="Clásico", command=lambda: aplicar_tema("Clásico")).pack(side="left", expand=True, padx=5)
tk.Button(frame_temas, text="Oscuro", command=lambda: aplicar_tema("Modo Oscuro")).pack(side="left", expand=True, padx=5)
tk.Button(frame_temas, text="Moderno", command=lambda: aplicar_tema("Moderno")).pack(side="left", expand=True, padx=5)

# --- SECCIÓN: Ajustes Manuales ---
frame_manual = tk.LabelFrame(root, text="Ajustes Manuales", bg="#fdfdfc", padx=10, pady=10)
frame_manual.pack(fill="x", padx=20, pady=10)

tk.Label(frame_manual, text="Tipografía:", bg="#fdfdfc").grid(row=0, column=0, sticky="w", pady=5)
combo_fuentes = ttk.Combobox(frame_manual, values=["Merriweather", "Inter", "Roboto", "Playfair Display"], state="readonly")
combo_fuentes.set("Merriweather")
combo_fuentes.grid(row=0, column=1, pady=5, padx=10, sticky="ew")

tk.Label(frame_manual, text="Color de Fondo:", bg="#fdfdfc").grid(row=1, column=0, sticky="w", pady=5)
btn_color_fondo = tk.Button(frame_manual, bg="#fdfdfc", width=15, command=lambda: elegir_color(btn_color_fondo))
btn_color_fondo.grid(row=1, column=1, pady=5, padx=10)

tk.Label(frame_manual, text="Color de Texto:", bg="#fdfdfc").grid(row=2, column=0, sticky="w", pady=5)
btn_color_texto = tk.Button(frame_manual, bg="#2b2b2b", width=15, command=lambda: elegir_color(btn_color_texto))
btn_color_texto.grid(row=2, column=1, pady=5, padx=10)

tk.Label(frame_manual, text="Color de Acento:", bg="#fdfdfc").grid(row=3, column=0, sticky="w", pady=5)
btn_color_acento = tk.Button(frame_manual, bg="#1e6292", width=15, command=lambda: elegir_color(btn_color_acento))
btn_color_acento.grid(row=3, column=1, pady=5, padx=10)

# --- BOTÓN FINAL ---
lbl_estado = tk.Label(root, text="", font=("Arial", 10, "italic"), bg="#fdfdfc")
lbl_estado.pack(pady=5)

btn_procesar = tk.Button(root, text="📂 Seleccionar Word y Generar", font=("Arial", 12, "bold"), bg="#1e6292", fg="white", pady=5, cursor="hand2", command=procesar_word)
btn_procesar.pack(fill="x", padx=20, pady=10)

root.mainloop()