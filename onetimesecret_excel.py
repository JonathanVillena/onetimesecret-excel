"""
Genera enlaces de OneTimeSecret a partir de un Excel.

Columna A: texto secreto
Columna B: frase de contraseña
Columna C: aquí se escribe el enlace generado

Solo usa librerías que vienen con Python (no necesita openpyxl).
"""
import json
import os
import posixpath
import re
import tempfile
import threading
import time
import tkinter as tk
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from tkinter import filedialog, messagebox, ttk
from xml.sax.saxutils import escape

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


def numero_columna(letras):
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def letras_columna(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class LibroExcel:
    """Lee y escribe un .xlsx usando solo zipfile y xml (sin openpyxl)."""

    def __init__(self, ruta):
        self.ruta = ruta
        with zipfile.ZipFile(ruta) as z:
            self.infos = z.infolist()
            self.archivos = {i.filename: z.read(i.filename) for i in self.infos}

        wb = ET.fromstring(self.archivos["xl/workbook.xml"])
        ns = wb.tag.split("}")[0] + "}"
        rels = ET.fromstring(self.archivos["xl/_rels/workbook.xml.rels"])
        destinos = {r.get("Id"): r.get("Target") for r in rels}
        self.hojas = {}
        for s in wb.iter(ns + "sheet"):
            rid = next(v for k, v in s.attrib.items() if k.endswith("}id"))
            t = destinos[rid]
            ruta_hoja = t.lstrip("/") if t.startswith("/") else posixpath.normpath(posixpath.join("xl", t))
            self.hojas[s.get("name")] = ruta_hoja
        self.nombres = list(self.hojas)

        self.compartidos = []
        if "xl/sharedStrings.xml" in self.archivos:
            sst = ET.fromstring(self.archivos["xl/sharedStrings.xml"])
            m = sst.tag.split("}")[0] + "}"
            for si in sst.findall(m + "si"):
                partes = si.findall(m + "t") + [t for r in si.findall(m + "r") for t in r.findall(m + "t")]
                self.compartidos.append("".join(t.text or "" for t in partes))

    def abrir_hoja(self, nombre):
        self.ruta_hoja = self.hojas[nombre]
        self.xml = self.archivos[self.ruta_hoja].decode("utf-8")
        p = re.search(r"<(\w+:)?worksheet\b", self.xml)
        self.p = (p.group(1) or "") if p else ""

        raiz = ET.fromstring(self.archivos[self.ruta_hoja])
        m = raiz.tag.split("}")[0] + "}"
        self.valores = {}
        self.max_fila = 0
        fila = 0
        for row in raiz.iter(m + "row"):
            fila = int(row.get("r")) if row.get("r") else fila + 1
            self.max_fila = max(self.max_fila, fila)
            col = 0
            for c in row.findall(m + "c"):
                ref = c.get("r")
                col = numero_columna(re.match(r"[A-Z]+", ref).group()) if ref else col + 1
                if col > 3:
                    continue
                tipo = c.get("t")
                v = c.find(m + "v")
                if tipo == "s" and v is not None:
                    valor = self.compartidos[int(v.text)]
                elif tipo == "inlineStr":
                    valor = "".join(t.text or "" for t in c.iter(m + "t"))
                elif v is not None:
                    valor = v.text
                else:
                    valor = None
                if valor is not None and str(valor) != "":
                    self.valores[(fila, col)] = valor

    def celda(self, fila, col):
        return self.valores.get((fila, col))

    def poner_c(self, fila, texto):
        """Escribe texto en la columna C de la fila indicada."""
        p = re.escape(self.p)
        fila_re = re.compile(rf'<{p}row\b(?P<attrs>[^>]*?\br="{fila}"[^>]*)(?<!/)>(?P<body>.*?)</{p}row>', re.S)
        m = fila_re.search(self.xml)
        if not m:
            raise RuntimeError(f"No se encontró la fila {fila} en el archivo. "
                               "Abre el Excel, guárdalo y vuelve a intentar.")
        body = m.group("body")
        celda_re = re.compile(rf'<{p}c\b[^>]*?\br="C{fila}"[^>]*?(?:/>|>.*?</{p}c>)', re.S)
        vieja = celda_re.search(body)
        estilo = ""
        if vieja:
            s = re.search(r'\bs="(\d+)"', vieja.group())
            estilo = f' s="{s.group(1)}"' if s else ""
        q = self.p
        nueva = (f'<{q}c r="C{fila}"{estilo} t="inlineStr"><{q}is><{q}t>{escape(texto)}'
                 f'</{q}t></{q}is></{q}c>')
        if vieja:
            body = body[:vieja.start()] + nueva + body[vieja.end():]
        else:
            pos = len(body)
            for c in re.finditer(rf'<{p}c\b[^>]*?\br="([A-Z]+)\d+"', body):
                if numero_columna(c.group(1)) > 3:
                    pos = c.start()
                    break
            body = body[:pos] + nueva + body[pos:]
        attrs = re.sub(r'\s+spans="[^"]*"', "", m.group("attrs"))
        self.xml = self.xml[:m.start()] + f"<{q}row{attrs}>{body}</{q}row>" + self.xml[m.end():]

        # Ajustar el rango de la hoja para que incluya la columna C
        def ajustar(d):
            ref = d.group(2)
            fin = ref.split(":")[-1]
            mm = re.match(r"([A-Z]+)(\d+)", fin)
            if not mm:
                return d.group(0)
            col = max(numero_columna(mm.group(1)), 3)
            fil = max(int(mm.group(2)), fila)
            ini = ref.split(":")[0] if ":" in ref else "A1"
            return f'{d.group(1)}{ini}:{letras_columna(col)}{fil}"'
        self.xml = re.sub(rf'(<{p}dimension\b[^>]*?\bref=")([^"]*)"', ajustar, self.xml, count=1)
        self.valores[(fila, 3)] = texto

    def guardar(self):
        self.archivos[self.ruta_hoja] = self.xml.encode("utf-8")
        carpeta = os.path.dirname(os.path.abspath(self.ruta))
        fd, temporal = tempfile.mkstemp(suffix=".xlsx", dir=carpeta)
        os.close(fd)
        try:
            with zipfile.ZipFile(temporal, "w") as z:
                for info in self.infos:
                    z.writestr(info, self.archivos[info.filename], compress_type=zipfile.ZIP_DEFLATED)
            os.replace(temporal, self.ruta)
        finally:
            if os.path.exists(temporal):
                os.remove(temporal)


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
            nombres = LibroExcel(ruta).nombres
            self.hoja["values"] = nombres
            self.hoja.set(nombres[0])
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
    def guardar(self, wb):
        while True:
            try:
                wb.guardar()
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
            wb = LibroExcel(ruta)
            wb.abrir_hoja(self.hoja.get() or wb.nombres[0])
            inicio = 2 if self.encabezado.get() else 1

            filas = []
            for r in range(inicio, wb.max_fila + 1):
                texto = wb.celda(r, 1)
                if texto is None or str(texto).strip() == "":
                    continue
                if self.saltar_llenas.get() and wb.celda(r, 3):
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
                texto = str(wb.celda(r, 1))
                clave = wb.celda(r, 2)
                clave = "" if clave is None else str(clave)
                try:
                    if driver:
                        enlace = crear_enlace(driver, texto, clave)
                    else:
                        enlace = crear_enlace_directo(texto, clave)
                    wb.poner_c(r, enlace)
                    hechas += 1
                    self.escribir(f"Fila {r}: {enlace}")
                    if not self.guardar(wb):
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
