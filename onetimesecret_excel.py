"""
Genera enlaces de OneTimeSecret y correos .msg a partir de un Excel.

Hoja "Datos":
  A: Texto secreto
  B: Frase de contraseña
  C: Enlace (lo escribe el programa)
  D: Correo del contacto
  E: Nombre del contacto
  F: Archivo .msg (lo escribe el programa)

Hoja "Plantilla":
  B1: Asunto      B2: Cuerpo del correo
  Variables: {nombre}  {correo}  {url}

Solo usa librerías que vienen con Python.
"""
import base64
import json
import os
import posixpath
import re
import struct
import subprocess
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

COL_TEXTO, COL_CLAVE, COL_ENLACE, COL_CORREO, COL_NOMBRE, COL_MSG = 1, 2, 3, 4, 5, 6

ASUNTO_DEFECTO = "Información de acceso para {nombre}"
CUERPO_DEFECTO = (
    "Hola {nombre},\n\n"
    "Te comparto la información de acceso en el siguiente enlace seguro:\n\n"
    "{url}\n\n"
    "El enlace solo puede abrirse una vez y caduca en 7 días.\n"
    "La frase de contraseña para abrirlo te la enviaré por otro medio.\n\n"
    "Saludos,"
)


# ======================================================================
#  Excel (.xlsx) sin librerías externas
# ======================================================================
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


def _limpiar_texto_excel(s):
    # Excel guarda algunos caracteres como _x000D_
    return re.sub(r"_x([0-9A-Fa-f]{4})_", lambda m: chr(int(m.group(1), 16)), s)


class LibroExcel:
    """Lee y escribe un .xlsx usando solo zipfile y xml."""

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
                self.compartidos.append(_limpiar_texto_excel("".join(t.text or "" for t in partes)))

    def leer_hoja(self, nombre):
        """Devuelve ({(fila, col): valor}, última fila)."""
        raiz = ET.fromstring(self.archivos[self.hojas[nombre]])
        m = raiz.tag.split("}")[0] + "}"
        valores = {}
        max_fila = 0
        fila = 0
        for row in raiz.iter(m + "row"):
            fila = int(row.get("r")) if row.get("r") else fila + 1
            max_fila = max(max_fila, fila)
            col = 0
            for c in row.findall(m + "c"):
                ref = c.get("r")
                col = numero_columna(re.match(r"[A-Z]+", ref).group()) if ref else col + 1
                tipo = c.get("t")
                v = c.find(m + "v")
                if tipo == "s" and v is not None:
                    valor = self.compartidos[int(v.text)]
                elif tipo == "inlineStr":
                    valor = _limpiar_texto_excel("".join(t.text or "" for t in c.iter(m + "t")))
                elif v is not None:
                    valor = v.text
                else:
                    valor = None
                if valor is not None and str(valor) != "":
                    valores[(fila, col)] = valor
        return valores, max_fila

    def abrir_hoja(self, nombre):
        self.ruta_hoja = self.hojas[nombre]
        self.xml = self.archivos[self.ruta_hoja].decode("utf-8")
        p = re.search(r"<(\w+:)?worksheet\b", self.xml)
        self.p = (p.group(1) or "") if p else ""
        self.valores, self.max_fila = self.leer_hoja(nombre)

    def celda(self, fila, col):
        v = self.valores.get((fila, col))
        return None if v is None else str(v)

    def poner(self, fila, col, texto):
        """Escribe texto en la celda (fila, col) de la hoja abierta."""
        letra = letras_columna(col)
        p = re.escape(self.p)
        fila_re = re.compile(rf'<{p}row\b(?P<attrs>[^>]*?\br="{fila}"[^>]*)(?<!/)>(?P<body>.*?)</{p}row>', re.S)
        m = fila_re.search(self.xml)
        if not m:
            raise RuntimeError(f"No se encontró la fila {fila} en el archivo. "
                               "Abre el Excel, guárdalo y vuelve a intentar.")
        body = m.group("body")
        celda_re = re.compile(rf'<{p}c\b[^>]*?\br="{letra}{fila}"[^>]*?(?:/>|>.*?</{p}c>)', re.S)
        vieja = celda_re.search(body)
        estilo = ""
        if vieja:
            s = re.search(r'\bs="(\d+)"', vieja.group())
            estilo = f' s="{s.group(1)}"' if s else ""
        q = self.p
        nueva = (f'<{q}c r="{letra}{fila}"{estilo} t="inlineStr"><{q}is><{q}t xml:space="preserve">'
                 f'{escape(texto)}</{q}t></{q}is></{q}c>')
        if vieja:
            body = body[:vieja.start()] + nueva + body[vieja.end():]
        else:
            pos = len(body)
            for c in re.finditer(rf'<{p}c\b[^>]*?\br="([A-Z]+)\d+"', body):
                if numero_columna(c.group(1)) > col:
                    pos = c.start()
                    break
            body = body[:pos] + nueva + body[pos:]
        attrs = re.sub(r'\s+spans="[^"]*"', "", m.group("attrs"))
        self.xml = self.xml[:m.start()] + f"<{q}row{attrs}>{body}</{q}row>" + self.xml[m.end():]

        # Ajustar el rango de la hoja para que incluya la celda
        def ajustar(d):
            ref = d.group(2)
            fin = ref.split(":")[-1]
            mm = re.match(r"([A-Z]+)(\d+)", fin)
            if not mm:
                return d.group(0)
            c2 = max(numero_columna(mm.group(1)), col)
            f2 = max(int(mm.group(2)), fila)
            ini = ref.split(":")[0] if ":" in ref else "A1"
            return f'{d.group(1)}{ini}:{letras_columna(c2)}{f2}"'
        self.xml = re.sub(rf'(<{p}dimension\b[^>]*?\bref=")([^"]*)"', ajustar, self.xml, count=1)
        self.valores[(fila, col)] = texto

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


# ======================================================================
#  Correo .msg (formato de Outlook) sin librerías externas
# ======================================================================
_FREE, _FIN, _FATSECT, _SIN = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF


def _cfb(hijos_raiz, clsid_raiz):
    """Arma un archivo 'Compound File Binary' (el contenedor de los .msg)."""
    nodos = [{"nombre": "Root Entry", "tipo": 5, "datos": b"", "hijos": hijos_raiz, "clsid": clsid_raiz}]
    i = 0
    while i < len(nodos):
        for h in nodos[i].get("hijos", []):
            nodos.append(h)
        i += 1
    for n, nodo in enumerate(nodos):
        nodo["id"] = n
        nodo.setdefault("izq", _SIN)
        nodo.setdefault("der", _SIN)
        nodo.setdefault("hijo", _SIN)

    def arbol(lista):
        if not lista:
            return _SIN
        m = len(lista) // 2
        lista[m]["izq"] = arbol(lista[:m])
        lista[m]["der"] = arbol(lista[m + 1:])
        return lista[m]["id"]

    for nodo in nodos:
        if nodo.get("hijos"):
            orden = sorted(nodo["hijos"], key=lambda h: (len(h["nombre"]), h["nombre"].upper()))
            nodo["hijo"] = arbol(orden)

    sectores, fat = [], []

    def agregar(datos):
        if not datos:
            return _FIN
        inicio = len(sectores)
        cant = (len(datos) + 511) // 512
        datos = datos + b"\0" * (cant * 512 - len(datos))
        for k in range(cant):
            sectores.append(datos[k * 512:(k + 1) * 512])
            fat.append(inicio + k + 1 if k < cant - 1 else _FIN)
        return inicio

    mini, minifat = bytearray(), []
    for nodo in nodos[1:]:
        if nodo["tipo"] != 2:
            nodo["inicio"], nodo["tam"] = 0, 0
            continue
        d = nodo["datos"]
        nodo["tam"] = len(d)
        if not d:
            nodo["inicio"] = _FIN
        elif len(d) < 4096:
            inicio = len(minifat)
            cant = (len(d) + 63) // 64
            mini += d + b"\0" * (cant * 64 - len(d))
            minifat += [inicio + j + 1 if j < cant - 1 else _FIN for j in range(cant)]
            nodo["inicio"] = inicio
        else:
            nodo["inicio"] = agregar(d)
    raiz = nodos[0]
    raiz["inicio"] = agregar(bytes(mini)) if mini else _FIN
    raiz["tam"] = len(mini)

    if minifat:
        minifat += [_FREE] * (-len(minifat) % 128)
        minifat_inicio = agregar(struct.pack(f"<{len(minifat)}I", *minifat))
        minifat_cant = len(minifat) // 128
    else:
        minifat_inicio, minifat_cant = _FIN, 0

    directorio = bytearray()
    for nodo in nodos:
        nombre = nodo["nombre"].encode("utf-16-le") + b"\0\0"
        directorio += struct.pack(
            "<64sHBBIII16sIQQIQ", nombre, len(nombre), nodo["tipo"], 1,
            nodo["izq"], nodo["der"], nodo["hijo"], nodo.get("clsid", b"\0" * 16),
            0, 0, 0, nodo["inicio"], nodo["tam"])
    vacia = struct.pack("<64sHBBIII16sIQQIQ", b"", 0, 0, 0, _SIN, _SIN, _SIN, b"\0" * 16, 0, 0, 0, 0, 0)
    while len(directorio) % 512:
        directorio += vacia
    dir_inicio = agregar(bytes(directorio))

    nfat = 1
    while nfat * 128 < len(fat) + nfat:
        nfat += 1
    if nfat > 109:
        raise RuntimeError("El correo es demasiado grande.")
    pos_fat = list(range(len(sectores), len(sectores) + nfat))
    fat += [_FATSECT] * nfat
    fat += [_FREE] * (nfat * 128 - len(fat))
    datos_fat = struct.pack(f"<{len(fat)}I", *fat)
    for k in range(nfat):
        sectores.append(datos_fat[k * 512:(k + 1) * 512])

    difat = pos_fat + [_FREE] * (109 - nfat)
    cabecera = struct.pack(
        "<8s16sHHHHH6sIIIIIIIII", bytes.fromhex("D0CF11E0A1B11AE1"), b"\0" * 16,
        0x3E, 3, 0xFFFE, 9, 6, b"\0" * 6,
        0, nfat, dir_inicio, 0, 4096, minifat_inicio, minifat_cant, _FIN, 0)
    cabecera += struct.pack("<109I", *difat)
    return cabecera + b"".join(sectores)


_PT_LONG, _PT_TEXTO, _PT_BIN = 0x0003, 0x001F, 0x0102


def _propiedades(lista, cabecera):
    flujo = bytearray(cabecera)
    nodos = []
    for pid, tipo, valor in lista:
        tag = (pid << 16) | tipo
        if tipo == _PT_TEXTO:
            d = valor.encode("utf-16-le")
            flujo += struct.pack("<IIII", tag, 6, len(d) + 2, 0)
        elif tipo == _PT_BIN:
            d = valor
            flujo += struct.pack("<IIII", tag, 6, len(d), 0)
        else:
            flujo += struct.pack("<IIQ", tag, 6, valor)
            continue
        nodos.append({"nombre": f"__substg1.0_{tag:08X}", "tipo": 2, "datos": d})
    return [{"nombre": "__properties_version1.0", "tipo": 2, "datos": bytes(flujo)}] + nodos


def crear_msg(ruta, para, nombre, asunto, texto, html):
    """Crea un borrador .msg (sin enviar) que Outlook abre listo para editar y enviar."""
    nombre_mostrar = nombre or para
    destinatario = _propiedades([
        (0x0C15, _PT_LONG, 1),             # tipo: Para
        (0x0FFE, _PT_LONG, 6),             # usuario de correo
        (0x3900, _PT_LONG, 0),
        (0x3000, _PT_LONG, 0),
        (0x3001, _PT_TEXTO, nombre_mostrar),
        (0x3002, _PT_TEXTO, "SMTP"),
        (0x3003, _PT_TEXTO, para),
        (0x39FE, _PT_TEXTO, para),
        (0x5FF6, _PT_TEXTO, nombre_mostrar),
    ], b"\0" * 8)
    mensaje = _propiedades([
        (0x001A, _PT_TEXTO, "IPM.Note"),
        (0x0037, _PT_TEXTO, asunto),
        (0x0E1D, _PT_TEXTO, asunto),
        (0x0070, _PT_TEXTO, asunto),
        (0x0E04, _PT_TEXTO, nombre_mostrar),
        (0x1000, _PT_TEXTO, texto),
        (0x1013, _PT_BIN, html.encode("utf-8")),
        (0x3FDE, _PT_LONG, 65001),         # HTML en UTF-8
        (0x3FFD, _PT_LONG, 65001),
        (0x0017, _PT_LONG, 1),             # importancia normal
        (0x0E07, _PT_LONG, 8),             # sin enviar (borrador)
        (0x340D, _PT_LONG, 0x00040000),    # admite Unicode
    ], struct.pack("<8sIIII8s", b"\0" * 8, 1, 0, 1, 0, b"\0" * 8))
    nameid = {"nombre": "__nameid_version1.0", "tipo": 1, "hijos": [
        {"nombre": "__substg1.0_00020102", "tipo": 2, "datos": b""},
        {"nombre": "__substg1.0_00030102", "tipo": 2, "datos": b""},
        {"nombre": "__substg1.0_00040102", "tipo": 2, "datos": b""},
    ]}
    recip = {"nombre": "__recip_version1.0_#00000000", "tipo": 1, "hijos": destinatario}
    clsid = bytes.fromhex("0B0D020000000000C000000000000046")
    with open(ruta, "wb") as f:
        f.write(_cfb(mensaje + [nameid, recip], clsid))


PS_OUTLOOK = r'''
$ErrorActionPreference = "Stop"
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$items = Get-Content -Raw -Encoding UTF8 -LiteralPath $env:OTS_JSON | ConvertFrom-Json
$abierto = @(Get-Process OUTLOOK -ErrorAction SilentlyContinue).Count -gt 0
$ol = New-Object -ComObject Outlook.Application
$i = 0
foreach ($it in $items) {
  try {
    $m = $ol.CreateItem(0)
    $m.To = $it.para
    $m.Subject = $it.asunto
    $m.HTMLBody = $it.html
    $m.SaveAs($it.ruta, 9)
    $m.Close(1)
    "OK`t$i"
  } catch { "ERR`t$i`t" + $_.Exception.Message }
  $i++
}
if (-not $abierto) { $ol.Quit() }
'''


def crear_msgs_con_outlook(items):
    """Crea los .msg usando el Outlook instalado. Devuelve {indice: error o None}."""
    fd, ruta_json = tempfile.mkstemp(suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(items, f, ensure_ascii=False)
    try:
        cmd = base64.b64encode(PS_OUTLOOK.encode("utf-16-le")).decode()
        env = dict(os.environ, OTS_JSON=ruta_json)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", cmd],
            capture_output=True, text=True, encoding="utf-8", errors="replace", env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=600)
    finally:
        os.remove(ruta_json)
    resultado = {}
    for linea in r.stdout.splitlines():
        partes = linea.split("\t")
        if partes[0] in ("OK", "ERR") and len(partes) > 1:
            resultado[int(partes[1])] = None if partes[0] == "OK" else (partes[2] if len(partes) > 2 else "error")
    if not resultado:
        raise RuntimeError("No se pudo usar Outlook: " + (r.stderr.strip().splitlines() or ["sin detalle"])[-1])
    return resultado


def rellenar(plantilla, datos):
    """Reemplaza {nombre}, {correo}, {url} (sin importar mayúsculas)."""
    def cambio(m):
        clave = m.group(1).lower()
        return datos[clave] if clave in datos else m.group(0)
    return re.sub(r"\{(\w+)\}", cambio, plantilla)


def texto_a_html(texto):
    partes = re.split(r"(https?://[^\s<>\"]+)", texto)
    salida = []
    for i, parte in enumerate(partes):
        if i % 2:
            salida.append(f'<a href="{escape(parte, {chr(34): "&quot;"})}">{escape(parte)}</a>')
        else:
            salida.append(escape(parte).replace("\n", "<br>\n"))
    return ('<html><head><meta charset="utf-8"></head>'
            '<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt">'
            + "".join(salida) + "</body></html>")


def nombre_archivo(fila, nombre, correo):
    base = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", nombre or correo or "correo").strip(" .")[:60]
    return f"{fila:03d} - {base}.msg"


# ======================================================================
#  OneTimeSecret
# ======================================================================
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
            opts = Options() if nombre == "chrome" else webdriver.EdgeOptions()
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


# ======================================================================
#  Ventana
# ======================================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → OneTimeSecret")
        self.geometry("720x680")
        self.minsize(580, 500)

        self.ruta = tk.StringVar()
        self.encabezado = tk.BooleanVar(value=True)
        self.saltar_llenas = tk.BooleanVar(value=True)
        self.hacer_enlaces = tk.BooleanVar(value=True)
        self.hacer_correos = tk.BooleanVar(value=True)
        self.usar_outlook = tk.BooleanVar(value=False)
        self.oculto = tk.BooleanVar(value=False)
        self.modo = tk.StringVar(value="directo")
        self.detener = False
        self.carpeta_correos = None

        marco = ttk.Frame(self, padding=12)
        marco.pack(fill="both", expand=True)

        ttk.Label(marco, text="Archivo Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(marco, textvariable=self.ruta).grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(marco, text="Buscar…", command=self.elegir).grid(row=1, column=1)

        ttk.Label(marco, text="Hoja con los datos:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.hoja = ttk.Combobox(marco, state="readonly")
        self.hoja.grid(row=3, column=0, sticky="w")

        op = ttk.Frame(marco)
        op.grid(row=4, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(op, text="La primera fila es encabezado", variable=self.encabezado).pack(anchor="w")
        ttk.Checkbutton(op, text="Saltar filas ya procesadas (enlace o correo ya generado)",
                        variable=self.saltar_llenas).pack(anchor="w")

        ttk.Label(op, text="¿Qué hacer?", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(op, text="Crear enlaces OneTimeSecret (columna C)",
                        variable=self.hacer_enlaces).pack(anchor="w")
        ttk.Checkbutton(op, text="Crear correos .msg (carpeta \"Correos\" junto al Excel)",
                        variable=self.hacer_correos).pack(anchor="w")
        ttk.Checkbutton(op, text="    Crear los .msg con el Outlook instalado (opcional)",
                        variable=self.usar_outlook).pack(anchor="w")

        ttk.Label(op, text="Método para los enlaces:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Radiobutton(op, text="Directo, sin navegador (recomendado)",
                        variable=self.modo, value="directo").pack(anchor="w")
        ttk.Radiobutton(op, text="Usando el navegador (Chrome o Edge)",
                        variable=self.modo, value="navegador").pack(anchor="w")
        ttk.Checkbutton(op, text="    Ocultar el navegador mientras trabaja",
                        variable=self.oculto).pack(anchor="w")

        botones = ttk.Frame(marco)
        botones.grid(row=5, column=0, columnspan=2, sticky="w")
        self.btn_iniciar = ttk.Button(botones, text="Iniciar", command=self.iniciar)
        self.btn_iniciar.pack(side="left")
        self.btn_detener = ttk.Button(botones, text="Detener", command=self.pedir_detener, state="disabled")
        self.btn_detener.pack(side="left", padx=6)
        self.btn_carpeta = ttk.Button(botones, text="Abrir carpeta de correos",
                                      command=self.abrir_carpeta, state="disabled")
        self.btn_carpeta.pack(side="left")

        self.progreso = ttk.Progressbar(marco, mode="determinate")
        self.progreso.grid(row=6, column=0, columnspan=2, sticky="ew", pady=8)

        self.log = tk.Text(marco, height=10, state="disabled")
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
            datos = [n for n in nombres if n.lower() != "plantilla"] or nombres
            self.hoja["values"] = datos
            self.hoja.set("Datos" if "Datos" in datos else datos[0])
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir el archivo:\n{e}")

    def abrir_carpeta(self):
        if self.carpeta_correos and os.path.isdir(self.carpeta_correos):
            os.startfile(self.carpeta_correos)

    def pedir_detener(self):
        self.detener = True
        self.escribir("Deteniendo después de la fila actual…")

    def iniciar(self):
        if not self.ruta.get():
            messagebox.showwarning("Falta archivo", "Elige un archivo Excel primero.")
            return
        if not (self.hacer_enlaces.get() or self.hacer_correos.get()):
            messagebox.showwarning("Nada que hacer", "Marca al menos una opción en \"¿Qué hacer?\".")
            return
        self.detener = False
        self.btn_iniciar.configure(state="disabled")
        self.btn_detener.configure(state="normal")
        threading.Thread(target=self.procesar, daemon=True).start()

    def terminar(self):
        self.btn_iniciar.configure(state="normal")
        self.btn_detener.configure(state="disabled")
        if self.carpeta_correos and os.path.isdir(self.carpeta_correos):
            self.btn_carpeta.configure(state="normal")

    def avanzar(self, valor):
        self.after(0, lambda: self.progreso.configure(value=valor))

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

    def leer_plantilla(self, wb):
        asunto, cuerpo = ASUNTO_DEFECTO, CUERPO_DEFECTO
        hoja = next((n for n in wb.nombres if n.lower() == "plantilla"), None)
        if hoja:
            v, _ = wb.leer_hoja(hoja)
            asunto = str(v.get((1, 2)) or asunto)
            cuerpo = str(v.get((2, 2)) or cuerpo)
        else:
            self.escribir("No hay hoja \"Plantilla\"; se usa el asunto y cuerpo por defecto.")
        return asunto.replace("\r\n", "\n").replace("\n", " ").strip(), cuerpo.replace("\r\n", "\n")

    def procesar(self):
        ruta = self.ruta.get()
        driver = None
        try:
            wb = LibroExcel(ruta)
            wb.abrir_hoja(self.hoja.get() or wb.nombres[0])
            inicio = 2 if self.encabezado.get() else 1
            saltar = self.saltar_llenas.get()
            todas = range(inicio, wb.max_fila + 1)

            filas_enlace = []
            if self.hacer_enlaces.get():
                filas_enlace = [r for r in todas
                                if (wb.celda(r, COL_TEXTO) or "").strip()
                                and not (saltar and wb.celda(r, COL_ENLACE))]

            def filas_para_correo():
                return [r for r in todas
                        if wb.celda(r, COL_ENLACE) and (wb.celda(r, COL_CORREO) or "").strip()
                        and not (saltar and wb.celda(r, COL_MSG))]

            total = len(filas_enlace) + (len(filas_para_correo()) if self.hacer_correos.get() else 0)
            self.after(0, lambda: self.progreso.configure(maximum=max(total, 1), value=0))
            paso = 0
            enlaces_ok = correos_ok = 0

            # ----- 1) Enlaces -----
            if filas_enlace:
                self.escribir(f"Enlaces a crear: {len(filas_enlace)}.")
                if self.modo.get() == "navegador":
                    self.escribir("Abriendo el navegador…")
                    driver = abrir_navegador(self.oculto.get())
                for r in filas_enlace:
                    if self.detener:
                        break
                    texto = wb.celda(r, COL_TEXTO)
                    clave = wb.celda(r, COL_CLAVE) or ""
                    try:
                        enlace = crear_enlace(driver, texto, clave) if driver else crear_enlace_directo(texto, clave)
                        wb.poner(r, COL_ENLACE, enlace)
                        enlaces_ok += 1
                        self.escribir(f"Fila {r}: enlace {enlace}")
                        if not self.guardar(wb):
                            self.escribir("Guardado cancelado. Se detiene el proceso.")
                            self.detener = True
                            break
                    except Exception as e:
                        self.escribir(f"Fila {r}: ERROR enlace - {str(e).splitlines()[0] if str(e) else type(e).__name__}")
                    paso += 1
                    self.avanzar(paso)
                    time.sleep(1)
            elif self.hacer_enlaces.get():
                self.escribir("No hay enlaces nuevos para crear.")

            # ----- 2) Correos -----
            if self.hacer_correos.get() and not self.detener:
                filas_correo = filas_para_correo()
                if not filas_correo:
                    self.escribir("No hay correos para crear (se necesita enlace en C y correo en D).")
                else:
                    asunto_t, cuerpo_t = self.leer_plantilla(wb)
                    carpeta = os.path.join(os.path.dirname(os.path.abspath(ruta)), "Correos")
                    os.makedirs(carpeta, exist_ok=True)
                    self.carpeta_correos = carpeta
                    self.escribir(f"Correos a crear: {len(filas_correo)} en {carpeta}")

                    items = []
                    for r in filas_correo:
                        correo = wb.celda(r, COL_CORREO).strip()
                        nombre = (wb.celda(r, COL_NOMBRE) or "").strip()
                        datos = {"nombre": nombre, "correo": correo,
                                 "url": wb.celda(r, COL_ENLACE), "enlace": wb.celda(r, COL_ENLACE)}
                        cuerpo = rellenar(cuerpo_t, datos)
                        archivo = nombre_archivo(r, nombre, correo)
                        items.append({"fila": r, "para": correo, "nombre": nombre,
                                      "asunto": rellenar(asunto_t, datos), "texto": cuerpo,
                                      "html": texto_a_html(cuerpo), "archivo": archivo,
                                      "ruta": os.path.join(carpeta, archivo)})

                    errores = {}
                    if self.usar_outlook.get():
                        self.escribir("Creando los .msg con Outlook…")
                        try:
                            errores = crear_msgs_con_outlook(items)
                        except Exception as e:
                            self.escribir(f"{e}\nSe usará el método integrado.")
                            errores = None
                    else:
                        errores = None

                    for i, it in enumerate(items):
                        if self.detener:
                            break
                        try:
                            if errores is None:
                                crear_msg(it["ruta"], it["para"], it["nombre"], it["asunto"], it["texto"], it["html"])
                            elif errores.get(i, "Outlook no respondió"):
                                raise RuntimeError(errores.get(i, "Outlook no respondió"))
                            wb.poner(it["fila"], COL_MSG, it["archivo"])
                            correos_ok += 1
                            self.escribir(f"Fila {it['fila']}: correo {it['archivo']}")
                        except Exception as e:
                            self.escribir(f"Fila {it['fila']}: ERROR correo - {e}")
                        paso += 1
                        self.avanzar(paso)
                    if correos_ok and not self.guardar(wb):
                        self.escribir("No se guardó la columna F en el Excel.")

            resumen = []
            if self.hacer_enlaces.get():
                resumen.append(f"Enlaces creados: {enlaces_ok}")
            if self.hacer_correos.get():
                resumen.append(f"Correos creados: {correos_ok}")
            self.escribir("Listo. " + " | ".join(resumen))
            self.after(0, lambda: messagebox.showinfo("Terminado", "\n".join(resumen)))
        except Exception as e:
            self.escribir(f"ERROR: {e}")
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            if driver:
                driver.quit()
            self.after(0, self.terminar)


if __name__ == "__main__":
    App().mainloop()
