# 1. Definimos nuestros "Temas Rápidos" en un diccionario
temas = {
    "Clasico": {
        "fondo": "#fdfdfc",
        "texto": "#2b2b2b",
        "acento": "#1e6292",
        "fuente": "Merriweather"
    },
    "Oscuro": {
        "fondo": "#1a1a1a",
        "texto": "#e2e8f0",
        "acento": "#ee8001",
        "fuente": "Inter"
    }
}

# 2. Simulamos que el usuario eligió un tema en el panel
tema_elegido = temas["Oscuro"]

# 3. Nuestro "Molde" CSS (fijate que usamos variables entre llaves {})
css_dinamico = f"""
:root {{
    --bg-page: {tema_elegido['fondo']};
    --text-main: {tema_elegido['texto']};
    --blue-primary: {tema_elegido['acento']};
    --font-body: '{tema_elegido['fuente']}', sans-serif;
}}

body {{
    background-color: var(--bg-page);
    color: var(--text-main);
    font-family: var(--font-body);
}}

a {{
    color: var(--blue-primary);
}}
"""

# Vemos el resultado
print(css_dinamico)