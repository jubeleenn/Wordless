import tkinter as tk
from tkinter import colorchooser, ttk

# --- 1. DICCIONARIO DE TEMAS PREDETERMINADOS ---
temas = {
    "Clásico": {"fondo": "#fdfdfc", "texto": "#2b2b2b", "acento": "#1e6292", "fuente": "Merriweather"},
    "Modo Oscuro": {"fondo": "#121212", "texto": "#e0e0e0", "acento": "#ee8001", "fuente": "Inter"},
    "Moderno": {"fondo": "#f0f4f8", "texto": "#334155", "acento": "#10b981", "fuente": "Roboto"}
}

# --- 2. FUNCIONES DE LÓGICA VISUAL ---
def aplicar_tema(nombre_tema):
    tema = temas[nombre_tema]
    
    # Actualizamos los colores de los botones visuales para que coincidan
    btn_color_fondo.config(bg=tema["fondo"])
    btn_color_texto.config(bg=tema["texto"])
    btn_color_acento.config(bg=tema["acento"])
    
    # Actualizamos el menú desplegable de la fuente
    combo_fuentes.set(tema["fuente"])
    
    lbl_estado.config(text=f"Tema '{nombre_tema}' seleccionado", fg="green")

def elegir_color(boton):
    # Abre la paleta nativa de colores. [1] guarda el código Hexadecimal (ej: #ffffff)
    color_elegido = colorchooser.askcolor(title="Elegí un color")[1]
    
    if color_elegido:
        boton.config(bg=color_elegido)
        lbl_estado.config(text="Color manual aplicado", fg="blue")

# --- 3. INTERFAZ GRÁFICA ---
root = tk.Tk()
root.title("Wordless | Panel de Control")
root.geometry("450x450")
root.configure(bg="#fdfdfc")

tk.Label(root, text="Personalizar Diseño", font=("Arial", 16, "bold"), bg="#fdfdfc", fg="#012662").pack(pady=(15, 5))

# --- SECCIÓN: Temas Rápidos ---
frame_temas = tk.LabelFrame(root, text="Temas Rápidos", bg="#fdfdfc", padx=10, pady=10)
frame_temas.pack(fill="x", padx=20, pady=10)

tk.Button(frame_temas, text="Clásico", command=lambda: aplicar_tema("Clásico")).pack(side="left", expand=True, padx=5)
tk.Button(frame_temas, text="Modo Oscuro", command=lambda: aplicar_tema("Modo Oscuro")).pack(side="left", expand=True, padx=5)
tk.Button(frame_temas, text="Moderno", command=lambda: aplicar_tema("Moderno")).pack(side="left", expand=True, padx=5)

# --- SECCIÓN: Ajustes Manuales ---
frame_manual = tk.LabelFrame(root, text="Ajustes Manuales", bg="#fdfdfc", padx=10, pady=10)
frame_manual.pack(fill="x", padx=20, pady=10)

# Fila: Tipografía
tk.Label(frame_manual, text="Tipografía:", bg="#fdfdfc").grid(row=0, column=0, sticky="w", pady=5)
combo_fuentes = ttk.Combobox(frame_manual, values=["Merriweather", "Inter", "Roboto", "Playfair Display"], state="readonly")
combo_fuentes.set("Merriweather")
combo_fuentes.grid(row=0, column=1, pady=5, padx=10, sticky="ew")

# Fila: Color de Fondo
tk.Label(frame_manual, text="Color de Fondo:", bg="#fdfdfc").grid(row=1, column=0, sticky="w", pady=5)
btn_color_fondo = tk.Button(frame_manual, bg="#fdfdfc", width=15, command=lambda: elegir_color(btn_color_fondo))
btn_color_fondo.grid(row=1, column=1, pady=5, padx=10)

# Fila: Color de Texto
tk.Label(frame_manual, text="Color de Texto:", bg="#fdfdfc").grid(row=2, column=0, sticky="w", pady=5)
btn_color_texto = tk.Button(frame_manual, bg="#2b2b2b", width=15, command=lambda: elegir_color(btn_color_texto))
btn_color_texto.grid(row=2, column=1, pady=5, padx=10)

# Fila: Color de Acento
tk.Label(frame_manual, text="Color de Acento:", bg="#fdfdfc").grid(row=3, column=0, sticky="w", pady=5)
btn_color_acento = tk.Button(frame_manual, bg="#1e6292", width=15, command=lambda: elegir_color(btn_color_acento))
btn_color_acento.grid(row=3, column=1, pady=5, padx=10)

# --- BOTÓN FINAL ---
lbl_estado = tk.Label(root, text="", font=("Arial", 10, "italic"), bg="#fdfdfc")
lbl_estado.pack(pady=5)

btn_procesar = tk.Button(root, text="📂 Seleccionar Word y Generar Web", font=("Arial", 12, "bold"), bg="#1e6292", fg="white", pady=5)
btn_procesar.pack(fill="x", padx=20, pady=10)

root.mainloop()