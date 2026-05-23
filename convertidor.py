import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import threading
import os
import sys
import subprocess
import tempfile
import shutil

try:
    from markitdown import MarkItDown
except ImportError:
    messagebox.showerror(
        "Error",
        "MarkItDown no está instalado.\nEjecutá: pip install markitdown[pdf,docx,xlsx,pptx]"
    )
    sys.exit(1)

TIPOS = [
    ("Documentos soportados",
     "*.pdf *.docx *.doc *.xlsx *.xls *.pptx *.ppt *.png *.jpg *.jpeg *.webp *.bmp"),
    ("PDF", "*.pdf"),
    ("Word", "*.docx *.doc"),
    ("Excel", "*.xlsx *.xls"),
    ("PowerPoint", "*.pptx *.ppt"),
    ("Imágenes", "*.png *.jpg *.jpeg *.webp *.bmp"),
    ("Todos los archivos", "*.*"),
]

# Formatos viejos que necesitan conversión previa con LibreOffice
NECESITAN_LIBREOFFICE = {
    ".doc":  ".docx",
    ".xls":  ".xlsx",
    ".ppt":  ".pptx",
}

BG      = "#F7F5F0"
PANEL   = "#FFFFFF"
ACCENT  = "#2563EB"
ACCENT2 = "#1D4ED8"
TEXT    = "#1C1917"
MUTED   = "#78716C"
BORDER  = "#E7E5E0"
SUCCESS = "#16A34A"
ERROR_C = "#DC2626"
WARN    = "#D97706"


def buscar_libreoffice():
    """Devuelve la ruta al ejecutable de LibreOffice o None si no está."""
    # Windows: rutas típicas de instalación
    candidatos_win = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidatos_win:
        if os.path.isfile(c):
            return c
    # Linux / Mac: en el PATH
    for nombre in ("libreoffice", "soffice"):
        ruta = shutil.which(nombre)
        if ruta:
            return ruta
    return None


LIBREOFFICE = buscar_libreoffice()


def convertir_con_libreoffice(path_entrada, ext_salida, carpeta_tmp):
    """
    Convierte un archivo viejo (.doc/.xls/.ppt) al formato nuevo usando LibreOffice.
    Devuelve la ruta del archivo convertido.
    """
    if not LIBREOFFICE:
        raise RuntimeError(
            "LibreOffice no está instalado.\n"
            "Descargalo desde https://www.libreoffice.org\n"
            "o convertí el archivo a .docx/.xlsx/.pptx manualmente."
        )

    # Mapeo de extensión a filtro de LibreOffice
    filtros = {
        ".docx": "MS Word 2007 XML",
        ".xlsx": "Calc MS Excel 2007 XML",
        ".pptx": "Impress MS PowerPoint 2007 XML",
    }
    filtro = filtros.get(ext_salida, "")
    cmd = [
        LIBREOFFICE, "--headless", "--convert-to",
        ext_salida.lstrip(".") + (f':{filtro}' if filtro else ""),
        "--outdir", carpeta_tmp,
        path_entrada,
    ]
    resultado = subprocess.run(cmd, capture_output=True, text=True, timeout=60)
    if resultado.returncode != 0:
        raise RuntimeError(f"LibreOffice falló:\n{resultado.stderr or resultado.stdout}")

    nombre_base = os.path.splitext(os.path.basename(path_entrada))[0]
    archivo_conv = os.path.join(carpeta_tmp, nombre_base + ext_salida)
    if not os.path.isfile(archivo_conv):
        raise RuntimeError(
            f"LibreOffice no generó el archivo esperado:\n{archivo_conv}"
        )
    return archivo_conv


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Convertidor a Markdown")
        self.geometry("720x640")
        self.minsize(600, 520)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.archivo_actual = None
        self.resultado_md   = None
        self._tmp_dir       = None
        self._construir_ui()

    def _construir_ui(self):
        # Encabezado
        hdr = tk.Frame(self, bg=ACCENT, pady=18)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Convertidor a Markdown",
                 font=("Georgia", 20, "bold"), bg=ACCENT, fg="white").pack()
        tk.Label(hdr, text="PDF · Word · Excel · PowerPoint · Imágenes",
                 font=("Helvetica", 11), bg=ACCENT, fg="#BFDBFE").pack(pady=(2, 0))

        # Aviso si LibreOffice no está
        if not LIBREOFFICE:
            aviso = tk.Frame(self, bg="#FEF3C7", pady=6)
            aviso.pack(fill="x")
            tk.Label(
                aviso,
                text="⚠️  LibreOffice no encontrado — los archivos .doc/.xls/.ppt no podrán convertirse.",
                font=("Helvetica", 10), bg="#FEF3C7", fg=WARN,
            ).pack()

        # Zona de selección
        zona = tk.Frame(self, bg=BG, pady=20, padx=30)
        zona.pack(fill="x")

        self.lbl_archivo = tk.Label(
            zona,
            text="  Ningún archivo seleccionado",
            font=("Helvetica", 12),
            bg=BORDER, fg=MUTED,
            relief="flat", padx=16, pady=14,
            anchor="w", cursor="hand2",
        )
        self.lbl_archivo.pack(fill="x", ipady=2)
        self.lbl_archivo.bind("<Button-1>", lambda e: self.seleccionar())

        btn_row = tk.Frame(zona, bg=BG, pady=12)
        btn_row.pack(fill="x")

        self._btn(btn_row, "📂  Elegir archivo", self.seleccionar,
                  bg=ACCENT, fg="white", hover=ACCENT2).pack(side="left")
        self.btn_conv = self._btn(btn_row, "✨  Convertir", self.convertir,
                                  bg=SUCCESS, fg="white", hover="#15803D")
        self.btn_conv.pack(side="left", padx=12)
        self.btn_conv.config(state="disabled")

        # Barra de progreso (oculta hasta que se use)
        self.progreso = ttk.Progressbar(zona, mode="indeterminate", length=200)

        # Etiqueta de estado de conversión (ej: "Convirtiendo .doc → .docx…")
        self.lbl_paso = tk.Label(zona, text="", font=("Helvetica", 10, "italic"),
                                 bg=BG, fg=MUTED)
        self.lbl_paso.pack(anchor="w")

        # Resultado
        res_frame = tk.Frame(self, bg=BG, padx=30)
        res_frame.pack(fill="both", expand=True)

        tk.Label(res_frame, text="Resultado", font=("Helvetica", 11, "bold"),
                 bg=BG, fg=TEXT).pack(anchor="w", pady=(0, 6))

        txt_wrap = tk.Frame(res_frame, bg=BORDER, bd=1, relief="flat")
        txt_wrap.pack(fill="both", expand=True)

        self.txt = tk.Text(
            txt_wrap,
            font=("Courier New", 11),
            bg=PANEL, fg=TEXT,
            relief="flat", padx=14, pady=12,
            wrap="word", state="disabled",
            selectbackground=ACCENT, selectforeground="white",
            insertbackground=ACCENT,
        )
        scroll = tk.Scrollbar(txt_wrap, command=self.txt.yview, bg=BG)
        self.txt.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.txt.pack(fill="both", expand=True)

        # Pie
        pie = tk.Frame(self, bg=BG, padx=30, pady=14)
        pie.pack(fill="x")

        self.btn_guardar = self._btn(pie, "💾  Guardar .md", self.guardar,
                                     bg="#7C3AED", fg="white", hover="#6D28D9")
        self.btn_guardar.pack(side="left")
        self.btn_guardar.config(state="disabled")

        self.btn_copiar = self._btn(pie, "📋  Copiar texto", self.copiar,
                                    bg=MUTED, fg="white", hover="#57534E")
        self.btn_copiar.pack(side="left", padx=10)
        self.btn_copiar.config(state="disabled")

        self.lbl_estado = tk.Label(pie, text="", font=("Helvetica", 10),
                                   bg=BG, fg=MUTED)
        self.lbl_estado.pack(side="right")

    def _btn(self, parent, texto, cmd, bg=ACCENT, fg="white", hover=ACCENT2):
        b = tk.Button(
            parent, text=texto, command=cmd,
            font=("Helvetica", 11, "bold"),
            bg=bg, fg=fg, activebackground=hover, activeforeground=fg,
            relief="flat", padx=18, pady=8, cursor="hand2", bd=0,
        )
        b.bind("<Enter>", lambda e, _b=b, _h=hover: _b.config(bg=_h))
        b.bind("<Leave>", lambda e, _b=b, _bg=bg: _b.config(bg=_bg))
        return b

    # ── Lógica ──────────────────────────────────────────────────────────────

    def seleccionar(self):
        path = filedialog.askopenfilename(
            title="Elegir archivo para convertir",
            filetypes=TIPOS,
        )
        if path:
            self.archivo_actual = path
            nombre = os.path.basename(path)
            ext = os.path.splitext(nombre)[1].lower()
            icono = "📄"
            if ext == ".pdf":             icono = "📕"
            elif ext in (".doc", ".docx"): icono = "📘"
            elif ext in (".xls", ".xlsx"): icono = "📗"
            elif ext in (".ppt", ".pptx"): icono = "📙"
            elif ext in (".png", ".jpg", ".jpeg", ".webp", ".bmp"): icono = "🖼️"
            self.lbl_archivo.config(text=f"  {icono}  {nombre}", fg=TEXT, bg="#EFF6FF")
            self.btn_conv.config(state="normal")
            self.lbl_paso.config(text="")
            self._estado("")

    def convertir(self):
        if not self.archivo_actual:
            return
        self._bloquear(True)
        self.progreso.pack(pady=(4, 0))
        self.progreso.start(10)
        threading.Thread(target=self._convertir_hilo, daemon=True).start()

    def _convertir_hilo(self):
        tmp_dir = None
        try:
            path = self.archivo_actual
            ext  = os.path.splitext(path)[1].lower()

            # Si es formato viejo, convertir primero con LibreOffice
            if ext in NECESITAN_LIBREOFFICE:
                ext_nueva = NECESITAN_LIBREOFFICE[ext]
                self.after(0, self.lbl_paso.config,
                           {"text": f"Convirtiendo {ext} → {ext_nueva} con LibreOffice…"})
                tmp_dir = tempfile.mkdtemp(prefix="markitdown_")
                path = convertir_con_libreoffice(path, ext_nueva, tmp_dir)
                self.after(0, self.lbl_paso.config,
                           {"text": "Extrayendo texto…"})
            else:
                self.after(0, self.lbl_paso.config, {"text": "Extrayendo texto…"})

            md = MarkItDown()
            resultado = md.convert(path)
            texto = resultado.text_content or ""
            self.after(0, self._mostrar_resultado, texto)

        except Exception as e:
            self.after(0, self._mostrar_error, str(e))
        finally:
            if tmp_dir and os.path.isdir(tmp_dir):
                shutil.rmtree(tmp_dir, ignore_errors=True)

    def _mostrar_resultado(self, texto):
        self.progreso.stop()
        self.progreso.pack_forget()
        self.lbl_paso.config(text="")
        self.resultado_md = texto
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.insert("1.0", texto)
        self.txt.config(state="disabled")
        self._bloquear(False)
        self.btn_guardar.config(state="normal")
        self.btn_copiar.config(state="normal")
        lineas = texto.count("\n")
        self._estado(f"✅  Listo — {lineas} líneas", SUCCESS)

    def _mostrar_error(self, err):
        self.progreso.stop()
        self.progreso.pack_forget()
        self.lbl_paso.config(text="")
        self._bloquear(False)
        self._estado(f"❌  Error", ERROR_C)
        messagebox.showerror("Error al convertir", err)

    def guardar(self):
        if not self.resultado_md:
            return
        nombre_base = os.path.splitext(os.path.basename(self.archivo_actual))[0]
        destino = filedialog.asksaveasfilename(
            defaultextension=".md",
            initialfile=f"{nombre_base}.md",
            filetypes=[("Markdown", "*.md"), ("Texto", "*.txt")],
            title="Guardar como…",
        )
        if destino:
            with open(destino, "w", encoding="utf-8") as f:
                f.write(self.resultado_md)
            self._estado(f"💾  Guardado: {os.path.basename(destino)}", SUCCESS)

    def copiar(self):
        if self.resultado_md:
            self.clipboard_clear()
            self.clipboard_append(self.resultado_md)
            self._estado("📋  Copiado al portapapeles", SUCCESS)

    def _bloquear(self, bloqueado):
        estado = "disabled" if bloqueado else "normal"
        self.btn_conv.config(state=estado)

    def _estado(self, msg, color=None):
        self.lbl_estado.config(text=msg, fg=color or MUTED)


if __name__ == "__main__":
    app = App()
    app.mainloop()
