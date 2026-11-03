"""
Genera enlaces de OneTimeSecret a partir de un Excel.

Columna A: texto secreto
Columna B: frase de contraseña
Columna C: aquí se escribe el enlace generado
"""
import json
import threading
import time
import tkinter as tk
import urllib.parse
import urllib.request
from tkinter import filedialog, messagebox, ttk

import openpyxl

# selenium solo hace falta para el modo "navegador"
try:
    from selenium import webdriver
    from selenium.webdriver.chrome.options import Options
    from selenium.webdriver.common.by import By
    from selenium.webdriver.support import expected_conditions as EC
    from selenium.webdriver.support.ui import WebDriverWait
    HAY_SELENIUM = True
except ImportError:
    HAY_SELENIUM = False

URL = "https://onetimesecret.com/es"
API = "https://us.onetimesecret.com/api/v1/share"
SIETE_DIAS = 7 * 24 * 3600


def crear_enlace_directo(texto, clave):
    """Crea el secreto con la API de la página (no necesita navegador)."""
    datos = {"secret": texto, "ttl": SIETE_DIAS}
    if clave:
        datos["passphrase"] = clave
    req = urllib.request.Request(
        API, data=urllib.parse.urlencode(datos).encode("utf-8"),
        headers={"User-Agent": "Mozilla/5.0"}, method="POST")
    with urllib.request.urlopen(req, timeout=30) as resp:
        r = json.loads(resp.read().decode("utf-8"))
    return f"https://us.onetimesecret.com/secret/{r['secret_key']}"


def abrir_navegador(oculto):
    """Abre Chrome; si no está, usa Edge."""
    if not HAY_SELENIUM:
        raise RuntimeError("El modo navegador necesita la librería selenium. "
                           "Usa el modo 'Directo, sin navegador'.")
    errores = []
    for nombre in ("chrome", "edge"):
        try:
            if nombre == "chrome":
                opts = Options()
            else:
                opts = webdriver.EdgeOptions()
            if oculto:
                opts.add_argument("--headless=new")
            opts.add_argument("--window-size=1300,950")
            return webdriver.Chrome(options=opts) if nombre == "chrome" else webdriver.Edge(options=opts)
        except Exception as e:
            errores.append(f"{nombre}: {str(e).splitlines()[0] if str(e) else type(e).__name__}")
    raise RuntimeError("No se pudo abrir Chrome ni Edge.\n" + "\n".join(errores))


def crear_enlace(driver, texto, clave):
    """Llena el formulario de la web y devuelve el enlace generado."""
    wait = WebDriverWait(driver, 30)
    driver.get(URL)

    area = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "textarea")))
    area.clear()
    area.send_keys(texto)

    if clave:
        boton_frase = wait.until(EC.presence_of_element_located(
            (By.XPATH, "//button[contains(., 'Frase de contrase')]")))
        driver.execute_script("arguments[0].click()", boton_frase)
        campo = wait.until(EC.visibility_of_element_located(
            (By.CSS_SELECTOR, "input[type='password']")))
        campo.clear()
        campo.send_keys(clave)

    boton_crear = wait.until(EC.presence_of_element_located(
        (By.XPATH, "//button[contains(., 'Crear enlace')]")))
    driver.execute_script("arguments[0].click()", boton_crear)

    resultado = wait.until(EC.presence_of_element_located(
        (By.CSS_SELECTOR, "input[aria-label='Secret URL']")))
    wait.until(lambda d: (resultado.get_attribute("value") or "").startswith("http"))
    return resultado.get_attribute("value")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → OneTimeSecret")
        self.geometry("680x580")
        self.minsize(560, 400)

        self.ruta = tk.StringVar()
        self.encabezado = tk.BooleanVar(value=True)
        self.saltar_llenas = tk.BooleanVar(value=True)
        self.oculto = tk.BooleanVar(value=False)
        self.modo = tk.StringVar(value="directo")
        self.detener = False

        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)

        ttk.Label(marco, text="Archivo Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(marco, textvariable=self.ruta).grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(marco, text="Buscar…", command=self.elegir).grid(row=1, column=1)

        ttk.Label(marco, text="Hoja:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.hoja = ttk.Combobox(marco, state="readonly")
        self.hoja.grid(row=3, column=0, sticky="w")

        opciones = ttk.Frame(marco)
        opciones.grid(row=4, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(opciones, text="La primera fila es encabezado",
                        variable=self.encabezado).pack(anchor="w")
        ttk.Checkbutton(opciones, text="Saltar filas que ya tienen enlace en la columna C",
                        variable=self.saltar_llenas).pack(anchor="w")
        ttk.Label(opciones, text="Método:").pack(anchor="w", pady=(6, 0))
        ttk.Radiobutton(opciones, text="Directo, sin navegador (recomendado)",
                        variable=self.modo, value="directo").pack(anchor="w")
        ttk.Radiobutton(opciones, text="Usando el navegador (Chrome o Edge)",
                        variable=self.modo, value="navegador").pack(anchor="w")
        ttk.Checkbutton(opciones, text="    Ocultar el navegador mientras trabaja",
                        variable=self.oculto).pack(anchor="w")

        botones = ttk.Frame(marco)
        botones.grid(row=5, column=0, columnspan=2, sticky="w")
        self.btn_iniciar = ttk.Button(botones, text="Iniciar", command=self.iniciar)
        self.btn_iniciar.pack(side="left")
        self.btn_detener = ttk.Button(botones, text="Detener", command=self.pedir_detener,
                                      state="disabled")
        self.btn_detener.pack(side="left", padx=6)

        self.progreso = ttk.Progressbar(marco, mode="determinate")
        self.progreso.grid(row=6, column=0, columnspan=2, sticky="ew", pady=8)

        self.log = tk.Text(marco, height=12, state="disabled")
        self.log.grid(row=7, column=0, columnspan=2, sticky="nsew")

        marco.columnconfigure(0, weight=1)
        marco.rowconfigure(7, weight=1)

    # ---------- interfaz ----------
    def escribir(self, msg):
        def _():
            self.log.configure(state="normal")
            self.log.insert("end", msg + "\n")
            self.log.see("end")
            self.log.configure(state="disabled")
        self.after(0, _)

    def elegir(self):
        ruta = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not ruta:
            return
        self.ruta.set(ruta)
        try:
            wb = openpyxl.load_workbook(ruta, read_only=True)
            self.hoja["values"] = wb.sheetnames
            self.hoja.set(wb.sheetnames[0])
            wb.close()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    def pedir_detener(self):
        self.detener = True
        self.escribir("Deteniendo después de la fila actual…")

    def iniciar(self):
        if not self.ruta.get():
            messagebox.showwarning("Falta archivo", "Elige un archivo Excel primero.")
            return
        self.detener = False
        self.btn_iniciar.configure(state="disabled")
        self.btn_detener.configure(state="normal")
        threading.Thread(target=self.procesar, daemon=True).start()

    def terminar(self):
        self.btn_iniciar.configure(state="normal")
        self.btn_detener.configure(state="disabled")

    # ---------- trabajo ----------
    def guardar(self, wb, ruta):
        while True:
            try:
                wb.save(ruta)
                return True
            except PermissionError:
                ok = messagebox.askretrycancel(
                    "Archivo en uso",
                    "No se puede guardar porque el Excel está abierto.\n"
                    "Ciérralo y presiona Reintentar.")
                if not ok:
                    return False

    def procesar(self):
        ruta = self.ruta.get()
        driver = None
        try:
            wb = openpyxl.load_workbook(ruta)
            ws = wb[self.hoja.get()] if self.hoja.get() else wb.active
            inicio = 2 if self.encabezado.get() else 1

            filas = []
            for r in range(inicio, ws.max_row + 1):
                texto = ws.cell(r, 1).value
                if texto is None or str(texto).strip() == "":
                    continue
                if self.saltar_llenas.get() and ws.cell(r, 3).value:
                    continue
                filas.append(r)

            if not filas:
                self.escribir("No hay filas para procesar.")
                return

            self.after(0, lambda: self.progreso.configure(maximum=len(filas), value=0))
            self.escribir(f"Filas a procesar: {len(filas)}.")
            if self.modo.get() == "navegador":
                self.escribir("Abriendo el navegador…")
                driver = abrir_navegador(self.oculto.get())

            hechas = 0
            for i, r in enumerate(filas, 1):
                if self.detener:
                    break
                texto = str(ws.cell(r, 1).value)
                clave = ws.cell(r, 2).value
                clave = "" if clave is None else str(clave)
                try:
                    if driver:
                        enlace = crear_enlace(driver, texto, clave)
                    else:
                        enlace = crear_enlace_directo(texto, clave)
                    ws.cell(r, 3).value = enlace
                    hechas += 1
                    self.escribir(f"Fila {r}: {enlace}")
                    if not self.guardar(wb, ruta):
                        self.escribir("Guardado cancelado. Se detiene el proceso.")
                        break
                except Exception as e:
                    self.escribir(f"Fila {r}: ERROR - {type(e).__name__}: {str(e).splitlines()[0] if str(e) else ''}")
                self.after(0, lambda v=i: self.progreso.configure(value=v))
                time.sleep(1)

            self.escribir(f"Listo. Enlaces generados: {hechas} de {len(filas)}.")
            self.after(0, lambda: messagebox.showinfo(
                "Terminado", f"Se generaron {hechas} enlaces.\nArchivo guardado:\n{ruta}"))
        except Exception as e:
            self.escribir(f"ERROR: {e}")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            if driver:
                driver.quit()
            self.after(0, self.terminar)


if __name__ == "__main__":
    App().mainloop()
