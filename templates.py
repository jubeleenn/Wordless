import html

def generar_plantilla(revista, metadatos, contenido_html, logo_path, texto_footer):
    # Colores por defecto (CICLOS)
    c_primary = "#1e6292"
    c_dark = "#012662"
    c_accent = "#ee8001"
    
    bloque_especial = ""

    # Adaptación según la revista elegida
    if revista == "Mundo Agrario":
        c_primary = "#2d6a4f" # Verde oscuro
        c_dark = "#1b4332"
        c_accent = "#d8f3dc"  # Verde claro
        bloque_especial = """
        <div class="abstract-box" style="margin-top: 0; padding: 1.5rem; background-color: #f8f9fa; border-left: 4px solid #2d6a4f;">
            <h3 style="font-family: var(--font-ui); font-size: 0.9rem; color: #2d6a4f; text-transform: uppercase;">Transparencia y Ciencia Abierta</h3>
            <p style="font-size: 0.85rem; margin-bottom: 0.5rem;"><b>Roles de colaboración (CRediT):</b> Escritura, revisión y edición: <em>[Completar]</em></p>
            <p style="font-size: 0.85rem; margin-bottom: 0.5rem;"><b>Financiamiento:</b> <em>[Declarar fuentes de financiamiento]</em></p>
            <p style="font-size: 0.85rem; margin-bottom: 0;"><b>Fuentes documentales:</b> <em>[Archivos utilizados]</em></p>
        </div>"""
    
    elif revista == "ACADEMO":
        c_primary = "#780000" # Rojo oscuro
        c_dark = "#5c0000"
        c_accent = "#c1121f"  # Rojo brillante
        bloque_especial = """
        <div class="abstract-box" style="margin-top: 0; padding: 1.5rem; background-color: #fff0f3; border-left: 4px solid #c1121f;">
            <h3 style="font-family: var(--font-ui); font-size: 0.9rem; color: #780000; text-transform: uppercase;">Metadatos del Autor</h3>
            <p style="font-size: 0.85rem; margin-bottom: 0.5rem;"><b>ORCID:</b> <a href="#" style="color: #c1121f;">Vincular perfil visible</a></p>
            <p style="font-size: 0.85rem; margin-bottom: 0.5rem;"><b>Afiliación (ROR):</b> <em>[Nombre de la Universidad / Instituto]</em></p>
            <p style="font-size: 0.85rem; margin-bottom: 0;"><b>Estructura:</b> Introducción, Metodología, Resultados, Discusión (IMRyD).</p>
        </div>"""
    
    elif revista == "Atmósfera":
        c_primary = "#0077b6" # Azul cielo
        c_dark = "#03045e"
        c_accent = "#00b4d8"  # Celeste
        bloque_especial = """
        <div class="abstract-box" style="margin-top: 0; padding: 1.5rem; background-color: #caf0f8; border-left: 4px solid #0077b6;">
            <h3 style="font-family: var(--font-ui); font-size: 0.9rem; color: #03045e; text-transform: uppercase;">Declaración de Uso de Inteligencia Artificial</h3>
            <p style="font-size: 0.85rem; margin-bottom: 0;">Conforme a las políticas de la revista, se declara la forma en la cual se utilizaron (o no) herramientas de IA generativa en el proceso de investigación y redacción de este manuscrito.</p>
        </div>"""

    # Procesar logo y footer
    etiqueta_logo = f'<img src="{logo_path}" alt="Logo de la Revista" class="header-logo">' if logo_path else f'<span class="brand-text">{revista}</span>'
    
    html = f"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{revista} | {metadatos.get('title', 'Artículo')}</title>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&family=Playfair+Display:wght@600;700&family=Merriweather:ital,wght@0,300;0,400;0,700;1,300;1,400&display=swap" rel="stylesheet">
<link rel="stylesheet" href="style.css">
</head>
<body>
<header class="site-header">
<div class="header-inner">
<div class="logo-area">
<a href="#" class="brand-link">
{etiqueta_logo}
</a>
</div>
<nav class="top-nav">
<a href="#" class="active-link">Actual</a>
<a href="#">Archivos</a>
<a href="#">Acerca de</a>
</nav>
</div>
</header>
<main class="article-wrapper">
<div class="metadata-bar">
<span class="badge">Artículos</span>
<span class="badge badge-outline">2025</span>
<span class="sep">|</span> Revista {revista}
</div>
<header class="article-heading">
<h1 class="title">{metadatos.get('title', 'Sin Título')}</h1>
<div class="subtitle-group">
<h2 class="subtitle">{metadatos.get('subtitle', '')}</h2>
</div>
<div class="author-block">
<p class="author-name">{metadatos.get('author', 'Autor Desconocido')}</p>
<p class="author-affiliation">{metadatos.get('affiliation', '')}</p>
</div>
</header>
{bloque_especial}
<section class="abstract-box">
<div class="abstract-lang">
<h3>Resumen</h3>
<p>{metadatos.get('abstract_es', '')}</p>
<p class="keywords"><strong>Palabras clave:</strong> {metadatos.get('keywords_es', '')}</p>
</div>
<hr class="abstract-divider">
<div class="abstract-lang" lang="en">
<h3>Abstract</h3>
<p>{metadatos.get('abstract_en', '')}</p>
<p class="keywords"><strong>Key words:</strong> {metadatos.get('keywords_en', '')}</p>
</div>
</section>
<article class="article-content">
{contenido_html}
</article>
</main>
<footer class="site-footer">
<div class="footer-inner" style="display: block; text-align: center;">
<p style="white-space: pre-wrap; line-height: 1.8;">{html.escape(texto_footer)}</p>
</div>
</footer>
</body>
</html>
"""

    css = f"""
:root {{
--blue-primary: {c_primary};
--blue-dark: {c_dark};
--blue-light: #eef5fb;
--orange-accent: {c_accent};
--text-main: #2b2b2b;
--text-muted: #5a6a7a;
--bg-page: #fdfdfc;
--white: #ffffff;
--border-color: #cbd5e1;
--font-heading: 'Playfair Display', Georgia, serif;
--font-ui: "Inter", sans-serif;
--font-body: "Merriweather", serif;
}}
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
html {{ scroll-behavior: smooth; }}
body {{ font-family: var(--font-body); background-color: var(--bg-page); color: var(--text-main); line-height: 1.85; font-size: 1.08rem; -webkit-font-smoothing: antialiased; }}
a {{ color: var(--blue-primary); text-decoration: none; transition: color 0.2s ease; }}
a:hover {{ color: var(--orange-accent); }}
.article-content a {{ overflow-wrap: anywhere; word-break: break-word; }}
.site-header {{ background-color: var(--blue-primary); border-top: 4px solid var(--orange-accent); padding: 0.8rem 0; box-shadow: 0 4px 15px rgba(0, 0, 0, 0.06); position: sticky; top: 0; z-index: 100; }}
.header-inner {{ max-width: 1150px; margin: 0 auto; padding: 0 2rem; display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; }}
.brand-link {{ display: inline-flex; align-items: center; color: var(--white); }}
.brand-text {{ color: var(--white); font-family: var(--font-heading); font-size: 2rem; font-weight: 700; }}
.header-logo {{ height: 75px; width: auto; display: block; object-fit: contain; background: white; padding: 5px; border-radius: 4px; }}
.top-nav {{ display: flex; gap: 1.8rem; align-items: center; }}
.top-nav a {{ color: #e2e8f0; font-family: var(--font-ui); font-size: 0.85rem; font-weight: 600; text-transform: uppercase; position: relative; padding-bottom: 4px; transition: color 0.2s ease; letter-spacing: 0.5px; }}
.top-nav a::after {{ content: ""; position: absolute; width: 0; height: 2px; bottom: 0; left: 0; background-color: var(--orange-accent); transition: width 0.25s ease-in-out; }}
.top-nav a:hover::after, .top-nav a.active-link::after {{ width: 100%; }}
.article-wrapper {{ max-width: 1000px; margin: 4rem auto; padding: 0 2rem; }}
.metadata-bar {{ display: flex; align-items: center; flex-wrap: wrap; gap: 0.8rem; margin-bottom: 2rem; font-family: var(--font-ui); font-size: 0.75rem; font-weight: 700; }}
.badge {{ background-color: var(--blue-primary); color: var(--white); padding: 0.3rem 0.8rem; border-radius: 4px; text-transform: uppercase; }}
.badge-outline {{ background-color: transparent; color: var(--blue-primary); border: 1px solid var(--blue-primary); }}
.title {{ font-family: var(--font-heading); font-size: 2.55rem; font-weight: 700; color: var(--blue-dark); line-height: 1.25; margin-bottom: 1.2rem; }}
.subtitle-group {{ display: grid; gap: 0.55rem; margin-bottom: 2rem; }}
.subtitle {{ font-family: var(--font-heading); font-size: 1.22rem; color: var(--text-muted); font-style: italic; font-weight: 400; line-height: 1.4; }}
.author-block {{ border-left: 4px solid var(--orange-accent); padding-left: 1rem; margin-bottom: 2rem; }}
.author-name {{ font-family: var(--font-ui); font-size: 1.05rem; font-weight: 700; color: var(--blue-primary); }}
.author-affiliation, .article-dates {{ font-family: var(--font-ui); font-size: 0.86rem; color: var(--text-muted); margin-top: 0.25rem; line-height: 1.55; }}
.abstract-box {{ background-color: var(--blue-light); padding: 2.5rem; border-radius: 8px; margin: 3.5rem 0; border: 1px solid #d0e1f0; }}
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
.site-footer {{ background-color: var(--blue-primary); color: #cbd5e1; padding: 3rem 0; font-family: var(--font-ui); font-size: 0.8rem; margin-top: 6rem; border-top: 5px solid var(--blue-dark); }}
blockquote, .academic-quote {{ font-size: 1.35rem; font-family: var(--font-heading); font-style: italic; color: var(--blue-primary); margin: 3.5rem 0; padding: 1.5rem 2.5rem; border-left: 4px solid var(--orange-accent); background-color: var(--white); box-shadow: 0 4px 12px rgba(0, 0, 0, 0.02); border-radius: 0 8px 8px 0; }}
"""
    return html, css