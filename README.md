# Excel → OneTimeSecret

Programa con ventana que lee un Excel y genera un enlace de [onetimesecret.com](https://onetimesecret.com/es) por cada fila.

| Columna | Contenido |
|---|---|
| A | Texto secreto |
| B | Frase de contraseña (puede ir vacía) |
| C | Aquí el programa escribe el enlace generado |

Ver `ejemplo.xlsx`.

## Descargar

Botón verde **Code → Download ZIP**, o en la pestaña **Releases**.

## Requisitos

1. **Python 3.10+**: https://www.python.org/downloads/ (al instalar, marca **"Add Python to PATH"**).
2. **Google Chrome**.

## Uso

1. Descomprime la carpeta.
2. Doble clic en `instalar.bat` (solo la primera vez, necesita internet).
3. Doble clic en `abrir_programa.bat`.
4. Pulsa **Buscar…**, elige tu Excel y pulsa **Iniciar**.

## Notas

- Cierra el Excel antes de iniciar; el programa guarda después de cada fila.
- Los enlaces caducan a los 7 días (valor por defecto de la página).
- Si la página cambia su diseño, el programa puede dejar de encontrar los campos.
