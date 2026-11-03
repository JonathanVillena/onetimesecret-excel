"""
Genera enlaces de OneTimeSecret y 2 correos .msg por fila a partir de un Excel.

  Correo 1: credenciales, con el enlace de OneTimeSecret.
  Correo 2: la contraseña para abrir ese enlace.

Las plantillas de los correos (asunto y cuerpo con formato) se editan en el
programa y se guardan en "config_correos.json" junto a este archivo.
Las columnas del Excel también se configuran en el programa.

Variables para asunto y cuerpo:
  [Nombre]  [URL]  [Contraseña]  [Correo]  [Ticket]   (también {nombre} {url} ...)

Solo usa librerías que vienen con Python.
"""
import base64
import ctypes
import html as htmlmod
import json
import os
import posixpath
import re
import struct
import subprocess
import sys
import tempfile
import threading
import time
import tkinter as tk
import unicodedata
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from html.parser import HTMLParser
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

CARPETA_PROGRAMA = os.path.dirname(os.path.abspath(sys.executable if getattr(sys, "frozen", False) else __file__))
RUTA_CONFIG = os.path.join(CARPETA_PROGRAMA, "config_correos.json")

# Columnas por defecto (letras). Vacío = no se usa.
COLUMNAS = [
    ("texto",  "Texto secreto (lo que va en OneTimeSecret)", "A"),
    ("clave",  "Contraseña / frase de contraseña",           "B"),
    ("enlace", "Enlace generado (lo escribe el programa)",   "C"),
    ("para1",  "Para - correo 1 (credenciales)",             "D"),
    ("nombre", "Nombre (para [Nombre])",                     "E"),
    ("cc1",    "CC - correo 1 (opcional)",                   "F"),
    ("para2",  "Para - correo 2 (vacío = mismo que correo 1)", "G"),
    ("cc2",    "CC - correo 2 (opcional)",                   ""),
    ("estado", "Estado (lo escribe el programa)",            "H"),
    ("ticket", "Ticket de referencia (para [Ticket])",       "I"),
]

_ESTILO = "font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000"
_CELDA = "border:1px solid #808080;padding:6px 8px;vertical-align:top"

ASUNTO1_DEFECTO = ("[Acción requerida] [Nombre] – Nuevas credenciales de producción "
                   "y actualización del PeSignature")
CUERPO1_DEFECTO = f"""<html><head><meta charset="utf-8"></head>
<body style="{_ESTILO}">
<p>Hola Equipo,</p>
<p>Les compartimos las nuevas credenciales de producción.</p>
<p><b><span style="color:#C00000">&#9888; ACCIÓN REQUERIDA:</span> las nuevas credenciales tienen 2 impactos</b></p>
<table style="border-collapse:collapse;{_ESTILO}">
<tr style="background:#1F4E78;color:#FFFFFF">
<th style="{_CELDA};text-align:left">Módulo</th>
<th style="{_CELDA};text-align:left">Qué cambiar</th>
<th style="{_CELDA};text-align:left">Si no se configura</th>
</tr>
<tr>
<td style="{_CELDA}"><b>1. Generación de pagos (CIP)</b></td>
<td style="{_CELDA}">AccessKey y SecretKey</td>
<td style="{_CELDA}">No podrán crear transacciones</td>
</tr>
<tr>
<td style="{_CELDA}"><b>2. Servicio que recibe las notificaciones</b></td>
<td style="{_CELDA}">SecretKey con la que calculan el PeSignature para validar cada notificación</td>
<td style="{_CELDA}">El cliente realiza el pago, pero su sistema rechaza la notificación enviada por PagoEfectivo y la orden queda pendiente.</td>
</tr>
</table>
<p>&nbsp;</p>
<p><b>Para el responsable técnico.</b><br>
A continuación, encontrará la url con las credenciales. El enlace vence en 7 días y solo puede abrirse una vez después de haber puesto la contraseña.</p>
<p><b>One-Time Secret:</b> [URL]<br>
En otro correo recibirá la contraseña para abrir ese enlace.</p>
<p><b>El flujo se cierra con la validación de la Notificación:</b><br>
Con la primera transacción pagada después del cambio, verifiquen que su sistema validó la notificación y que la orden quedó como pagada.<br>
Confírmenlo respondiendo a este correo, con todos en copia.</p>
<p><b>TICKET REFERENCIA:</b> [Ticket]</p>
<p>Quedamos atentos.</p>
</body></html>"""

ASUNTO2_DEFECTO = "[Nombre] – Contraseña del enlace de credenciales de producción"
CUERPO2_DEFECTO = f"""<html><head><meta charset="utf-8"></head>
<body style="{_ESTILO}">
<p>Hola,</p>
<p>Comparto la contraseña correspondiente para acceder al Link de credenciales de producción enviadas en el correo principal:</p>
<p><b>Contraseña:</b> [Contraseña]</p>
<p>Por favor, cualquier consulta o comunicación relacionada deberá mantenerse exclusivamente en la línea principal de correo, copiando a todos los involucrados.</p>
<p><b>TICKET REFERENCIA:</b> [Ticket]</p>
<p>Saludos,</p>
</body></html>"""

DATOS_EJEMPLO = {"nombre": "Comercio Ejemplo SAC", "url": "https://us.onetimesecret.com/secret/EJEMPLO123",
                 "clave": "Clave-Ejemplo-123", "correo": "contacto@ejemplo.com", "ticket": "NC123"}

LINEA_TICKET = "<p><b>TICKET REFERENCIA:</b> [Ticket]</p>\n"


def asegurar_ticket(html):
    """Si la plantilla no tiene [Ticket], agrega la línea antes de la despedida."""
    if any(_ALIAS.get(_normalizar(m.group(1) or m.group(2))) == "ticket" for m in _PATRON.finditer(html)):
        return html
    pos = None
    for m in re.finditer(r"<(p|div)\b[^>]*>", html, re.I):
        fin = html.find("</", m.end())
        trozo = _normalizar(html[m.end():fin if fin > 0 else len(html)])
        if any(p in trozo for p in ("saludos", "quedamos atentos", "atentamente", "cordialmente")):
            pos = m.start()
    if pos is None:
        i = html.lower().rfind("</body")
        pos = i if i >= 0 else len(html)
    return html[:pos] + LINEA_TICKET + html[pos:]


def config_defecto():
    return {
        "columnas": {k: v for k, _, v in COLUMNAS},
        "correo1": {"asunto": ASUNTO1_DEFECTO, "html": CUERPO1_DEFECTO},
        "correo2": {"asunto": ASUNTO2_DEFECTO, "html": CUERPO2_DEFECTO},
    }


def cargar_config():
    cfg = config_defecto()
    try:
        with open(RUTA_CONFIG, encoding="utf-8") as f:
            guardada = json.load(f)
        for k in ("correo1", "correo2"):
            cfg[k].update(guardada.get(k, {}))
            cfg[k]["html"] = asegurar_ticket(cfg[k]["html"])
        cfg["columnas"].update(guardada.get("columnas", {}))
    except FileNotFoundError:
        pass
    except Exception:
        messagebox.showwarning("Configuración", "No se pudo leer config_correos.json; se usan los valores por defecto.")
    return cfg


def guardar_config(cfg):
    with open(RUTA_CONFIG, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


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



def rtf_con_html(html):
    """Envuelve el HTML en RTF (\\fromhtml1), como hace Outlook, y lo empaqueta
    en el formato 'RTF comprimido' de los .msg (variante sin comprimir 'MELA')."""
    # Todo a ASCII: los caracteres especiales pasan a entidades HTML (&#233;)
    seguro = html.encode("ascii", "xmlcharrefreplace").decode("ascii")
    seguro = seguro.replace("\\", "\\\\").replace("{", "\\{").replace("}", "\\}")
    seguro = seguro.replace("\r\n", "\n").replace("\n", "\\par\n")
    rtf = ("{\\rtf1\\ansi\\ansicpg1252\\fromhtml1 \\deff0"
           "{\\fonttbl{\\f0\\fswiss Arial;}}"
           "{\\*\\htmltag1 " + seguro + "}}").encode("ascii")
    return struct.pack("<IIII", len(rtf) + 12, len(rtf), 0x414C454D, 0) + rtf


def crear_msg(ruta, para, cc, asunto, html):
    """Crea un borrador .msg (sin enviar) que Outlook abre listo para editar y enviar.
    para y cc son listas de direcciones."""
    destinatarios = [(d, 1) for d in para] + [(d, 2) for d in cc]
    nodos_recip = []
    for i, (direccion, tipo) in enumerate(destinatarios):
        props = _propiedades([
            (0x0C15, _PT_LONG, tipo),          # 1 = Para, 2 = CC
            (0x0FFE, _PT_LONG, 6),             # usuario de correo
            (0x3900, _PT_LONG, 0),
            (0x3000, _PT_LONG, i),
            (0x3001, _PT_TEXTO, direccion),
            (0x3002, _PT_TEXTO, "SMTP"),
            (0x3003, _PT_TEXTO, direccion),
            (0x39FE, _PT_TEXTO, direccion),
            (0x5FF6, _PT_TEXTO, direccion),
        ], b"\0" * 8)
        nodos_recip.append({"nombre": f"__recip_version1.0_#{i:08X}", "tipo": 1, "hijos": props})
    n = len(destinatarios)
    mensaje = _propiedades([
        (0x001A, _PT_TEXTO, "IPM.Note"),
        (0x0037, _PT_TEXTO, asunto),
        (0x0E1D, _PT_TEXTO, asunto),
        (0x0070, _PT_TEXTO, asunto),
        (0x0E04, _PT_TEXTO, "; ".join(para)),
        (0x0E03, _PT_TEXTO, "; ".join(cc)),
        (0x1000, _PT_TEXTO, html_a_texto(html)),
        (0x1013, _PT_BIN, html.encode("utf-8")),
        (0x1009, _PT_BIN, rtf_con_html(html)),  # RTF con el HTML dentro (Outlook lo prefiere)
        (0x0E1F, 0x000B, 1),               # el RTF está sincronizado con el cuerpo
        (0x1016, _PT_LONG, 3),             # el cuerpo principal es HTML
        (0x3FDE, _PT_LONG, 65001),         # HTML en UTF-8
        (0x3FFD, _PT_LONG, 65001),
        (0x0017, _PT_LONG, 1),             # importancia normal
        (0x0E07, _PT_LONG, 8),             # sin enviar (borrador)
        (0x340D, _PT_LONG, 0x00040000),    # admite Unicode
    ], struct.pack("<8sIIII8s", b"\0" * 8, n, 0, n, 0, b"\0" * 8))
    nameid = {"nombre": "__nameid_version1.0", "tipo": 1, "hijos": [
        {"nombre": "__substg1.0_00020102", "tipo": 2, "datos": b""},
        {"nombre": "__substg1.0_00030102", "tipo": 2, "datos": b""},
        {"nombre": "__substg1.0_00040102", "tipo": 2, "datos": b""},
    ]}
    clsid = bytes.fromhex("0B0D020000000000C000000000000046")
    with open(ruta, "wb") as f:
        f.write(_cfb(mensaje + [nameid] + nodos_recip, clsid))


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
    if ($it.cc) { $m.CC = $it.cc }
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
    datos = [{"para": "; ".join(it["para"]), "cc": "; ".join(it["cc"]), "asunto": it["asunto"],
              "html": it["html"], "ruta": it["ruta"]} for it in items]
    fd, ruta_json = tempfile.mkstemp(suffix=".json")
    with os.fdopen(fd, "w", encoding="utf-8") as f:
        json.dump(datos, f, ensure_ascii=False)
    try:
        cmd = base64.b64encode(PS_OUTLOOK.encode("utf-16-le")).decode()
        env = dict(os.environ, OTS_JSON=ruta_json)
        r = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-EncodedCommand", cmd],
            capture_output=True, text=True, encoding="utf-8", errors="replace", env=env,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=900)
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


# ----------------------------------------------------------------------
#  Plantillas: variables y formato
# ----------------------------------------------------------------------
_ALIAS = {
    "nombre": "nombre", "name": "nombre", "comercio": "nombre", "empresa": "nombre", "cliente": "nombre",
    "url": "url", "enlace": "url", "link": "url", "one time secret": "url", "one time secre": "url",
    "onetimesecret": "url", "onetime secret": "url", "secret": "url",
    "contrasena": "clave", "password": "clave", "clave": "clave", "frase": "clave",
    "frase de contrasena": "clave",
    "correo": "correo", "email": "correo", "mail": "correo",
    "ticket": "ticket", "ticket referencia": "ticket", "referencia": "ticket", "nro ticket": "ticket",
    "numero ticket": "ticket", "numero de ticket": "ticket", "n ticket": "ticket", "nc": "ticket",
}

# [ ... ] o { ... }, permitiendo etiquetas HTML dentro (Word suele partir las palabras)
_PATRON = re.compile(r"\[((?:[^\[\]<>]|<[^>]*>){1,300}?)\]|\{((?:[^{}<>]|<[^>]*>){1,300}?)\}")


def _normalizar(k):
    k = re.sub(r"<[^>]*>", "", k)
    k = htmlmod.unescape(k)
    k = unicodedata.normalize("NFKD", k)
    k = "".join(c for c in k if not unicodedata.combining(c))
    return re.sub(r"[\s\-_:]+", " ", k.lower()).strip()


def rellenar(plantilla, datos, es_html=False):
    """Reemplaza [Nombre], [URL], [Contraseña], [Correo] (o {nombre}, ...).
    Lo que no es una variable conocida, como [Acción requerida], se deja igual."""
    def cambio(m):
        clave = _ALIAS.get(_normalizar(m.group(1) or m.group(2)))
        if not clave:
            return m.group(0)
        valor = datos.get(clave, "")
        if not es_html:
            return valor
        if clave == "url" and valor:
            return f'<a href="{escape(valor, {chr(34): "&quot;"})}">{escape(valor)}</a>'
        return escape(valor)

    if not es_html:
        return _PATRON.sub(cambio, plantilla)
    # en HTML solo se reemplaza dentro del <body> (no en los estilos)
    i = plantilla.lower().find("<body")
    i = 0 if i < 0 else i
    return plantilla[:i] + _PATRON.sub(cambio, plantilla[i:])


class _HtmlATexto(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.partes, self.oculto = [], 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head", "title"):
            self.oculto += 1
        elif tag == "br":
            self.partes.append("\n")
        elif tag in ("p", "div", "tr", "li", "table", "h1", "h2", "h3"):
            self.partes.append("\n")
        elif tag in ("td", "th"):
            self.partes.append("\t")

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head", "title"):
            self.oculto = max(0, self.oculto - 1)
        elif tag in ("p", "div", "table", "h1", "h2", "h3"):
            self.partes.append("\n")

    def handle_data(self, d):
        if not self.oculto:
            self.partes.append(re.sub(r"[ \t\r\n]+", " ", d))


def html_a_texto(html):
    p = _HtmlATexto()
    p.feed(html)
    lineas = [l.strip(" ") for l in "".join(p.partes).replace("\xa0", " ").split("\n")]
    texto = "\n".join(lineas)
    return re.sub(r"\n{3,}", "\n\n", texto).strip()


def leer_html_portapapeles():
    """Devuelve el HTML del portapapeles (lo que se copia de Outlook/Word/navegador) o None."""
    user32, kernel32 = ctypes.windll.user32, ctypes.windll.kernel32
    user32.GetClipboardData.restype = ctypes.c_void_p
    kernel32.GlobalLock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalLock.restype = ctypes.c_void_p
    kernel32.GlobalUnlock.argtypes = [ctypes.c_void_p]
    kernel32.GlobalSize.argtypes = [ctypes.c_void_p]
    kernel32.GlobalSize.restype = ctypes.c_size_t
    formato = user32.RegisterClipboardFormatW("HTML Format")
    if not user32.OpenClipboard(None):
        return None
    try:
        h = user32.GetClipboardData(formato)
        if not h:
            return None
        p = kernel32.GlobalLock(h)
        try:
            datos = ctypes.string_at(p, kernel32.GlobalSize(h))
        finally:
            kernel32.GlobalUnlock(h)
    finally:
        user32.CloseClipboard()
    datos = datos.split(b"\0")[0]
    cab = datos[:400].decode("ascii", "ignore")
    ini = re.search(r"StartHTML:(-?\d+)", cab)
    fin = re.search(r"EndHTML:(-?\d+)", cab)
    if ini and fin and int(ini.group(1)) >= 0:
        datos = datos[int(ini.group(1)):int(fin.group(1))]
    else:
        ini = re.search(r"StartFragment:(\d+)", cab)
        fin = re.search(r"EndFragment:(\d+)", cab)
        if ini and fin:
            datos = datos[int(ini.group(1)):int(fin.group(1))]
    html = datos.decode("utf-8", "replace")
    html = re.sub(r"<!--\s*(Start|End)Fragment\s*-->", "", html)
    if "charset" not in html.lower()[:2000]:
        html = '<meta charset="utf-8">' + html
    return html


def separar_correos(valor):
    if not valor:
        return []
    return [d for d in re.split(r"[;,\s]+", str(valor)) if "@" in d]


def nombre_archivo(fila, nombre, sufijo):
    base = re.sub(r'[\\/:*?"<>|\r\n\t]+', "_", nombre or "correo").strip(" .")[:50]
    return f"{fila:03d} - {base} - {sufijo}.msg"


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
class EditorPlantilla(ttk.Frame):
    """Pestaña para editar asunto y cuerpo (HTML con formato) de un correo."""

    def __init__(self, padre, titulo, asunto_defecto, html_defecto):
        super().__init__(padre, padding=10)
        self.asunto_defecto, self.html_defecto = asunto_defecto, html_defecto
        self.asunto = tk.StringVar()

        ttk.Label(self, text=titulo, font=("Segoe UI", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="w")
        ttk.Label(self, text="Asunto:").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(self, textvariable=self.asunto).grid(row=2, column=0, columnspan=2, sticky="ew")

        ayuda = ("Cuerpo con formato: escribe el correo en Outlook o Word (tablas, negritas, colores), "
                 "selecciónalo todo, cópialo (Ctrl+C) y pulsa \"Pegar con formato\".\n"
                 "Variables: [Nombre]  [URL]  [Contraseña]  [Correo]  [Ticket]. "
                 "Usa \"Vista previa\" para comprobar cómo queda.")
        ttk.Label(self, text=ayuda, wraplength=700, foreground="#444").grid(
            row=3, column=0, columnspan=2, sticky="w", pady=(10, 4))

        botones = ttk.Frame(self)
        botones.grid(row=4, column=0, columnspan=2, sticky="w", pady=(0, 6))
        ttk.Button(botones, text="Pegar con formato", command=self.pegar).pack(side="left")
        ttk.Button(botones, text="Vista previa", command=self.vista_previa).pack(side="left", padx=6)
        ttk.Button(botones, text="Restaurar ejemplo", command=self.restaurar).pack(side="left")

        ttk.Label(self, text="Código HTML del cuerpo (se puede editar a mano):").grid(row=5, column=0, sticky="w")
        self.texto = tk.Text(self, wrap="word", font=("Consolas", 9), undo=True)
        barra = ttk.Scrollbar(self, orient="vertical", command=self.texto.yview)
        self.texto.configure(yscrollcommand=barra.set)
        self.texto.grid(row=6, column=0, sticky="nsew")
        barra.grid(row=6, column=1, sticky="ns")
        self.columnconfigure(0, weight=1)
        self.rowconfigure(6, weight=1)

    def poner(self, asunto, html):
        self.asunto.set(asunto)
        self.texto.delete("1.0", "end")
        self.texto.insert("1.0", html)

    def obtener(self):
        return {"asunto": self.asunto.get().strip(), "html": self.texto.get("1.0", "end-1c")}

    def pegar(self):
        try:
            html = leer_html_portapapeles()
        except Exception as e:
            html = None
            messagebox.showerror("Portapapeles", f"No se pudo leer el portapapeles:\n{e}")
            return
        if not html:
            messagebox.showwarning(
                "Sin formato", "El portapapeles no tiene texto con formato.\n"
                "Copia el correo desde Outlook, Word o el navegador y vuelve a intentar.")
            return
        self.texto.delete("1.0", "end")
        self.texto.insert("1.0", html)
        messagebox.showinfo("Listo", "Cuerpo pegado con formato. Revisa la vista previa.")

    def vista_previa(self):
        d = self.obtener()
        asunto = rellenar(d["asunto"], DATOS_EJEMPLO)
        cuerpo = rellenar(d["html"], DATOS_EJEMPLO, es_html=True)
        encabezado = (f'<div style="font-family:Segoe UI,Arial;font-size:10pt;background:#F2F2F2;'
                      f'border-bottom:1px solid #CCC;padding:8px 12px;margin-bottom:12px">'
                      f'<b>Vista previa (datos de ejemplo)</b><br>Asunto: {escape(asunto)}</div>')
        i = cuerpo.lower().find("<body")
        if i >= 0:
            j = cuerpo.find(">", i) + 1
            cuerpo = cuerpo[:j] + encabezado + cuerpo[j:]
        else:
            cuerpo = '<meta charset="utf-8">' + encabezado + cuerpo
        ruta = os.path.join(tempfile.gettempdir(), f"vista_previa_{id(self)}.html")
        with open(ruta, "w", encoding="utf-8") as f:
            f.write(cuerpo)
        os.startfile(ruta)

    def restaurar(self):
        if messagebox.askyesno("Restaurar", "¿Reemplazar el asunto y el cuerpo por el ejemplo original?"):
            self.poner(self.asunto_defecto, self.html_defecto)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel → OneTimeSecret")
        self.geometry("820x720")
        self.minsize(640, 560)

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

        pestanas = ttk.Notebook(self)
        pestanas.pack(fill="both", expand=True, padx=6, pady=6)

        # ----- Pestaña Proceso -----
        marco = ttk.Frame(pestanas, padding=12)
        pestanas.add(marco, text="Proceso")

        ttk.Label(marco, text="Archivo Excel:").grid(row=0, column=0, sticky="w")
        ttk.Entry(marco, textvariable=self.ruta).grid(row=1, column=0, sticky="ew", padx=(0, 6))
        ttk.Button(marco, text="Buscar…", command=self.elegir).grid(row=1, column=1)

        ttk.Label(marco, text="Hoja con los datos:").grid(row=2, column=0, sticky="w", pady=(10, 0))
        self.hoja = ttk.Combobox(marco, state="readonly")
        self.hoja.grid(row=3, column=0, sticky="w")

        op = ttk.Frame(marco)
        op.grid(row=4, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Checkbutton(op, text="La primera fila es encabezado", variable=self.encabezado).pack(anchor="w")
        ttk.Checkbutton(op, text="Saltar filas ya procesadas (con enlace o con estado)",
                        variable=self.saltar_llenas).pack(anchor="w")

        ttk.Label(op, text="¿Qué hacer?", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Checkbutton(op, text="Crear enlaces OneTimeSecret", variable=self.hacer_enlaces).pack(anchor="w")
        ttk.Checkbutton(op, text="Crear los 2 correos .msg por fila (carpeta \"Correos\" junto al Excel)",
                        variable=self.hacer_correos).pack(anchor="w")
        ttk.Checkbutton(op, text="    Crear los .msg con el Outlook instalado (opcional)",
                        variable=self.usar_outlook).pack(anchor="w")

        ttk.Label(op, text="Método para los enlaces:", font=("Segoe UI", 9, "bold")).pack(anchor="w", pady=(8, 0))
        ttk.Radiobutton(op, text="Directo, sin navegador (recomendado)",
                        variable=self.modo, value="directo").pack(anchor="w")
        ttk.Radiobutton(op, text="Usando el navegador (Chrome o Edge)",
                        variable=self.modo, value="navegador").pack(anchor="w")
        ttk.Checkbutton(op, text="    Ocultar el navegador mientras trabaja", variable=self.oculto).pack(anchor="w")

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

        # ----- Pestañas de correos -----
        self.editor1 = EditorPlantilla(pestanas, "Correo 1 – Credenciales (lleva el enlace [URL])",
                                       ASUNTO1_DEFECTO, CUERPO1_DEFECTO)
        self.editor2 = EditorPlantilla(pestanas, "Correo 2 – Contraseña (lleva [Contraseña])",
                                       ASUNTO2_DEFECTO, CUERPO2_DEFECTO)
        pestanas.add(self.editor1, text="Correo 1: credenciales")
        pestanas.add(self.editor2, text="Correo 2: contraseña")

        # ----- Pestaña Columnas -----
        cols = ttk.Frame(pestanas, padding=12)
        pestanas.add(cols, text="Columnas del Excel")
        ttk.Label(cols, text="Escribe la letra de la columna del Excel para cada dato (A, B, … K). "
                             "Deja vacío lo que no uses.", wraplength=700).grid(
            row=0, column=0, columnspan=2, sticky="w", pady=(0, 10))
        self.vars_col = {}
        for i, (clave, etiqueta, _) in enumerate(COLUMNAS, 1):
            ttk.Label(cols, text=etiqueta + ":").grid(row=i, column=0, sticky="w", pady=2)
            v = tk.StringVar()
            ttk.Entry(cols, textvariable=v, width=6).grid(row=i, column=1, sticky="w", padx=8)
            self.vars_col[clave] = v
        ttk.Label(cols, text="En \"Para\" y \"CC\" se pueden poner varios correos separados por ; o ,",
                  foreground="#444").grid(row=len(COLUMNAS) + 1, column=0, columnspan=2, sticky="w", pady=(10, 0))

        ttk.Button(self, text="Guardar configuración", command=self.guardar_cfg).pack(side="right", padx=8, pady=(0, 8))
        ttk.Label(self, text=f"Configuración: {RUTA_CONFIG}", foreground="#666").pack(side="left", padx=8, pady=(0, 8))

        cfg = cargar_config()
        self.editor1.poner(cfg["correo1"]["asunto"], cfg["correo1"]["html"])
        self.editor2.poner(cfg["correo2"]["asunto"], cfg["correo2"]["html"])
        for k, v in self.vars_col.items():
            v.set(cfg["columnas"].get(k, ""))
        self.protocol("WM_DELETE_WINDOW", self.cerrar)

    # ---------- configuración ----------
    def config_actual(self):
        return {"columnas": {k: v.get().strip().upper() for k, v in self.vars_col.items()},
                "correo1": self.editor1.obtener(), "correo2": self.editor2.obtener()}

    def guardar_cfg(self, avisar=True):
        try:
            guardar_config(self.config_actual())
            if avisar:
                messagebox.showinfo("Guardado", "Configuración guardada.")
            return True
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la configuración:\n{e}")
            return False

    def cerrar(self):
        self.guardar_cfg(avisar=False)
        self.destroy()

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
            self.hoja.set("Datos" if "Datos" in nombres else nombres[0])
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
        cfg = self.config_actual()
        malas = [k for k, v in cfg["columnas"].items() if v and not re.fullmatch(r"[A-Z]{1,3}", v)]
        faltan = [k for k in ("texto", "enlace", "para1", "estado") if not cfg["columnas"][k]]
        if malas or faltan:
            nombres = {k: e for k, e, _ in COLUMNAS}
            msg = ""
            if malas:
                msg += "Letras no válidas en: " + ", ".join(nombres[k] for k in malas) + "\n"
            if faltan:
                msg += "Faltan columnas obligatorias: " + ", ".join(nombres[k] for k in faltan)
            messagebox.showwarning("Columnas", msg + "\n\nRevisa la pestaña \"Columnas del Excel\".")
            return
        self.guardar_cfg(avisar=False)
        self.cfg = cfg
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

    def procesar(self):
        ruta = self.ruta.get()
        cfg = self.cfg
        col = {k: (numero_columna(v) if v else None) for k, v in cfg["columnas"].items()}
        driver = None
        try:
            wb = LibroExcel(ruta)
            wb.abrir_hoja(self.hoja.get() or wb.nombres[0])
            inicio = 2 if self.encabezado.get() else 1
            saltar = self.saltar_llenas.get()
            todas = range(inicio, wb.max_fila + 1)

            def valor(r, clave):
                return (wb.celda(r, col[clave]) or "").strip() if col[clave] else ""

            filas_enlace = []
            if self.hacer_enlaces.get():
                filas_enlace = [r for r in todas if valor(r, "texto") and not (saltar and valor(r, "enlace"))]

            def filas_para_correo():
                return [r for r in todas if valor(r, "enlace") and separar_correos(valor(r, "para1"))
                        and not (saltar and valor(r, "estado"))]

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
                    texto, clave = valor(r, "texto"), valor(r, "clave")
                    try:
                        enlace = crear_enlace(driver, texto, clave) if driver else crear_enlace_directo(texto, clave)
                        wb.poner(r, col["enlace"], enlace)
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
                    self.escribir("No hay correos para crear (se necesita enlace y correo \"Para\").")
                else:
                    carpeta = os.path.join(os.path.dirname(os.path.abspath(ruta)), "Correos")
                    os.makedirs(carpeta, exist_ok=True)
                    self.carpeta_correos = carpeta
                    self.escribir(f"Filas con correos a crear: {len(filas_correo)} → {carpeta}")
                    c1, c2 = cfg["correo1"], cfg["correo2"]

                    items = []
                    for r in filas_correo:
                        para1 = separar_correos(valor(r, "para1"))
                        nombre, clave = valor(r, "nombre"), valor(r, "clave")
                        datos = {"nombre": nombre, "url": valor(r, "enlace"), "clave": clave,
                                 "correo": "; ".join(para1), "ticket": valor(r, "ticket")}
                        if not datos["ticket"]:
                            self.escribir(f"Fila {r}: aviso, no tiene número de ticket.")
                        items.append({"fila": r, "n": 1, "para": para1, "cc": separar_correos(valor(r, "cc1")),
                                      "asunto": " ".join(rellenar(c1["asunto"], datos).split()),
                                      "html": rellenar(c1["html"], datos, es_html=True),
                                      "archivo": nombre_archivo(r, nombre, "1 Credenciales")})
                        if clave:
                            para2 = separar_correos(valor(r, "para2")) or para1
                            items.append({"fila": r, "n": 2, "para": para2, "cc": separar_correos(valor(r, "cc2")),
                                          "asunto": " ".join(rellenar(c2["asunto"], datos).split()),
                                          "html": rellenar(c2["html"], datos, es_html=True),
                                          "archivo": nombre_archivo(r, nombre, "2 Contraseña")})
                        else:
                            self.escribir(f"Fila {r}: sin contraseña, no se crea el correo 2.")
                    for it in items:
                        it["ruta"] = os.path.join(carpeta, it["archivo"])

                    errores = None
                    if self.usar_outlook.get():
                        self.escribir("Creando los .msg con Outlook…")
                        try:
                            errores = crear_msgs_con_outlook(items)
                        except Exception as e:
                            self.escribir(f"{e}\nSe usará el método integrado.")
                            errores = None

                    hechos = {}
                    for i, it in enumerate(items):
                        if self.detener:
                            break
                        try:
                            if errores is None:
                                crear_msg(it["ruta"], it["para"], it["cc"], it["asunto"], it["html"])
                            elif errores.get(i, "Outlook no respondió"):
                                raise RuntimeError(errores.get(i, "Outlook no respondió"))
                            hechos.setdefault(it["fila"], []).append(it["n"])
                            correos_ok += 1
                            self.escribir(f"Fila {it['fila']}: {it['archivo']}")
                        except Exception as e:
                            self.escribir(f"Fila {it['fila']}: ERROR correo {it['n']} - {e}")
                        if it["n"] == 1:
                            paso += 1
                            self.avanzar(paso)

                    for r, ns in hechos.items():
                        if ns == [1, 2]:
                            estado = "Correos 1 y 2 creados"
                        elif ns == [1]:
                            estado = "Solo correo 1 creado"
                        else:
                            estado = "Solo correo 2 creado"
                        wb.poner(r, col["estado"], estado)
                    if hechos and not self.guardar(wb):
                        self.escribir("No se guardó la columna de estado en el Excel.")

            resumen = []
            if self.hacer_enlaces.get():
                resumen.append(f"Enlaces creados: {enlaces_ok}")
            if self.hacer_correos.get():
                resumen.append(f"Correos .msg creados: {correos_ok}")
            self.escribir("Listo. " + " | ".join(resumen))
            self.after(0, lambda: messagebox.showinfo("Terminado", "\n".join(resumen)))
        except Exception as e:
            texto_error = str(e)
            self.escribir(f"ERROR: {texto_error}")
            self.after(0, lambda: messagebox.showerror("Error", texto_error))
        finally:
            if driver:
                driver.quit()
            self.after(0, self.terminar)


if __name__ == "__main__":
    App().mainloop()
