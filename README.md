# Excel → OneTimeSecret + correos .msg

Programa con ventana que, a partir de un Excel:

1. Crea un enlace de [onetimesecret.com](https://onetimesecret.com/es) por cada fila (texto + frase de contraseña).
2. Genera un correo **.msg** por contacto, con el asunto y el cuerpo armados desde una plantilla.
   Solo abres el .msg en Outlook, agregas tu firma y envías.

No necesita instalar librerías: solo Python. No necesita Outlook para crear los .msg.

## Plantilla de Excel

Descarga [`plantilla.xlsx`](plantilla.xlsx).

**Hoja `Datos`**

| Col | Contenido |
|---|---|
| A | Texto secreto |
| B | Frase de contraseña (puede ir vacía) |
| C | Enlace (lo escribe el programa) |
| D | Correo del contacto |
| E | Nombre del contacto |
| F | Archivo .msg (lo escribe el programa) |

**Hoja `Plantilla`**

| Celda | Contenido |
|---|---|
| B1 | Asunto |
| B2 | Cuerpo del correo (Alt+Enter para saltos de línea) |

Variables que se reemplazan en asunto y cuerpo: `{nombre}`, `{correo}`, `{url}`.

## Uso desde VS Code

1. Descarga [`onetimesecret_excel.py`](onetimesecret_excel.py) y `plantilla.xlsx`.
2. Abre el `.py` en VS Code y pulsa ▶ (o `python onetimesecret_excel.py`).
3. Elige el Excel, marca lo que quieres hacer y pulsa **Iniciar**.

Los correos quedan en la carpeta `Correos`, junto al Excel.

## Notas

- Cierra el Excel antes de iniciar; el programa guarda en el mismo archivo.
- Los enlaces caducan a los 7 días.
- Con "Saltar filas ya procesadas" no se repiten enlaces ni correos; desmárcalo para volver a generar.
- Opción "Crear los .msg con el Outlook instalado": usa Outlook (clásico) si está disponible.
