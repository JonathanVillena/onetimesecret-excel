# Excel → OneTimeSecret + 2 correos .msg

Programa con ventana que, a partir de un Excel:

1. Crea un enlace de [onetimesecret.com](https://onetimesecret.com/es) por fila (texto secreto + contraseña).
2. Genera **2 correos .msg por fila**, listos para abrir en Outlook, agregar la firma y enviar:
   - **Correo 1 – Credenciales:** lleva el enlace.
   - **Correo 2 – Contraseña:** lleva la contraseña para abrir el enlace.

Solo necesita Python (sin librerías). No necesita Outlook para crear los .msg.

## Uso (VS Code)

1. Descarga [`onetimesecret_excel.py`](onetimesecret_excel.py) y [`plantilla.xlsx`](plantilla.xlsx).
2. Abre el `.py` en VS Code y pulsa ▶ (o `python onetimesecret_excel.py`).
3. Pestaña **Proceso**: elige el Excel y pulsa **Iniciar**.
4. Los correos quedan en la carpeta `Correos`, junto al Excel.

## Plantillas de correo (en el programa)

Pestañas **Correo 1** y **Correo 2**: asunto y cuerpo **con formato** (tablas, negritas, colores).

- Escribe el correo en Outlook o Word, cópialo y pulsa **Pegar con formato**.
- **Vista previa** lo abre en el navegador con datos de ejemplo.
- Variables: `[Nombre]`, `[URL]`, `[Contraseña]`, `[Correo]`, `[Ticket]` (también `{nombre}`, `{url}`…).
  Lo que no es variable, como `[Acción requerida]`, se deja igual.

Se guarda en `config_correos.json` junto al programa (cópialo para llevar tus plantillas a otra PC).

## Columnas del Excel (configurables)

Pestaña **Columnas del Excel**. Valores por defecto (`plantilla.xlsx`):

| Col | Dato |
|---|---|
| A | Texto secreto |
| B | Contraseña |
| C | Enlace (lo escribe el programa) |
| D | Para – correo 1 |
| E | Nombre (para `[Nombre]`) |
| F | CC – correo 1 (opcional) |
| G | Para – correo 2 (opcional; vacío = mismo que D) |
| H | Estado (lo escribe el programa) |
| I | Ticket de referencia (para `[Ticket]`) |

Varios correos en una celda: sepáralos con `;` o `,`.

## Notas

- Cierra el Excel antes de iniciar; el programa guarda en el mismo archivo.
- Los enlaces caducan a los 7 días.
- Si una fila no tiene contraseña, solo se crea el correo 1.
- "Saltar filas ya procesadas" evita repetir enlaces y correos.
