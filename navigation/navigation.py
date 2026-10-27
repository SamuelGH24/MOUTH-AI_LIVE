"""
Módulo de Navegación Asistida y Dictado - Mouth AI Live

Responsabilidades:
1. Navegación por teclado (arriba/abajo/Tab/Enter/atrás), tal como lo haría
   una persona con teclado.
2. Lectura en voz alta del elemento enfocado (TTS + UI Automation). En Word,
   Excel y PowerPoint lee por COM (selección, celda o diapositiva).
3. Búsqueda en Google escribiendo en la barra de direcciones.
4. Dictado: escribir texto dictado por voz en cualquier campo de texto
   (Word, Excel, Bloc de notas, cajas de búsqueda, etc.) y comandos básicos
   de edición (nueva línea, borrar, deshacer, guardar, seleccionar todo).
5. Documentos de Word, libros de Excel y presentaciones de PowerPoint:
   crear, guardar con nombre y abrir por voz.

NOTA sobre el dictado: en Word y Excel el texto se escribe directo por COM.
En el resto de aplicaciones se pega desde el portapapeles (Ctrl+V) en vez de
simular cada tecla, porque es más confiable con español (tildes, eñes). El
portapapeles del usuario se restaura después de pegar.

Requiere (instalar en el entorno de Windows):
    pip install pyautogui pyttsx3 uiautomation pyperclip pywin32
"""

import atexit
import ctypes
import difflib
import os
import re
import time
import unicodedata
from pathlib import Path

import pyautogui
import pyperclip
import pyttsx3

try:
    import uiautomation as auto
    UIAUTOMATION_DISPONIBLE = True
except ImportError:
    UIAUTOMATION_DISPONIBLE = False

try:
    import win32api
    import win32com.client
    import win32gui
    import win32process
    WIN32_DISPONIBLE = True
except ImportError:
    WIN32_DISPONIBLE = False


TECLAS_NAVEGACION = {
    "abajo": "down",
    "arriba": "up",
    "siguiente": "tab",
    "anterior": ["shift", "tab"],
    "entrar": "enter",
    "atras_pagina": ["alt", "left"],
    "adelante_pagina": ["alt", "right"],
    "borrar": "backspace",
    "borrar_palabra": ["ctrl", "backspace"],
    "guardar": ["ctrl", "s"],
    "deshacer": ["ctrl", "z"],
    "seleccionar_todo": ["ctrl", "a"],
}

# ---------------------------------------------------------------------------
# Edición de texto por voz: copiar / pegar / cortar / borrar / seleccionar /
# mover el cursor del texto.
#
# Cada comando es (mensaje_hablado, pasos). Cada paso es una combinación de
# teclas que se pulsa en orden (los comandos de varios pasos, como
# "seleccionar_linea", pulsan Inicio y luego Shift+Fin).
#
# Los atajos de copiar/pegar/cortar/deshacer/rehacer y de mover/seleccionar
# con flechas son estándar de Windows y funcionan igual en Word, Bloc de
# notas, navegadores, etc.
#
# ATENCIÓN: guardar, seleccionar todo y el formato (negrita, cursiva,
# subrayado) NO son iguales en todos los idiomas/versiones de Office (en
# Word en español pueden ser Ctrl+N, Ctrl+G, Ctrl+E...). Por eso en Word,
# Excel y PowerPoint se resuelven por COM (ver _editar_en_word,
# _editar_en_excel y _editar_en_ppt) y estas teclas quedan solo como
# respaldo para las demás aplicaciones.
# ---------------------------------------------------------------------------
TECLAS_EDICION = {
    # Portapapeles
    "copiar": ("Copiado", [["ctrl", "c"]]),
    "pegar": ("Pegado", [["ctrl", "v"]]),
    "cortar": ("Cortado", [["ctrl", "x"]]),
    "rehacer": ("Rehecho", [["ctrl", "y"]]),
    # Borrar
    "borrar_seleccion": ("Borrado", [["delete"]]),
    "borrar_palabra_siguiente": ("Palabra borrada", [["ctrl", "delete"]]),
    "borrar_linea": ("Línea borrada", [["home"], ["shift", "end"], ["delete"]]),
    # Seleccionar
    "seleccionar_linea": ("Línea seleccionada", [["home"], ["shift", "end"]]),
    "seleccionar_palabra_anterior": ("Palabra seleccionada", [["ctrl", "shift", "left"]]),
    "seleccionar_palabra_siguiente": ("Palabra seleccionada", [["ctrl", "shift", "right"]]),
    "seleccionar_hasta_inicio": ("Seleccionado hasta el inicio", [["shift", "home"]]),
    "seleccionar_hasta_fin": ("Seleccionado hasta el final", [["shift", "end"]]),
    # Mover el cursor del texto
    "cursor_inicio_linea": ("Inicio de línea", [["home"]]),
    "cursor_fin_linea": ("Fin de línea", [["end"]]),
    "cursor_inicio_documento": ("Inicio del documento", [["ctrl", "home"]]),
    "cursor_fin_documento": ("Fin del documento", [["ctrl", "end"]]),
    "cursor_palabra_anterior": ("Palabra anterior", [["ctrl", "left"]]),
    "cursor_palabra_siguiente": ("Palabra siguiente", [["ctrl", "right"]]),
    "cursor_letra_anterior": ("Letra anterior", [["left"]]),
    "cursor_letra_siguiente": ("Letra siguiente", [["right"]]),
    # Respaldo fuera de Office (en Word, Excel y PowerPoint se usa COM)
    "guardar": ("Guardando", [["ctrl", "s"]]),
    "seleccionar_todo": ("Todo seleccionado", [["ctrl", "a"]]),
    "negrita": ("Negrita", [["ctrl", "b"]]),
    "cursiva": ("Cursiva", [["ctrl", "i"]]),
    "subrayado": ("Subrayado", [["ctrl", "u"]]),
}

# Constantes de Word (wdToggle / wdUndefined) para el formato por COM.
WD_TOGGLE = 9999998
WD_INDEFINIDO = 9999999

# Constantes de Word para mover/extender la selección por COM.
WD_PALABRA = 2
WD_LINEA = 5
WD_MOVER = 0
WD_EXTENDER = 1

# Constantes de Excel y PowerPoint.
XL_FORMATO_XLSX = 51        # xlOpenXMLWorkbook
PPT_FORMATO_PPTX = 24       # ppSaveAsOpenXMLPresentation
XL_ARRIBA = -4162           # xlUp
XL_IZQUIERDA = -4159        # xlToLeft
XL_SIN_SUBRAYADO = -4142    # xlUnderlineStyleNone
XL_SUBRAYADO_SIMPLE = 2     # xlUnderlineStyleSingle

# ---------------------------------------------------------------------------
# Envío de teclas con la bandera de "tecla extendida".
#
# PyAutoGUI manda las flechas, Inicio, Fin y Suprimir SIN esa bandera, y
# Windows las confunde con las del teclado numérico. Con Shift presionado
# (y Bloq Num activo) el cursor se mueve pero NO selecciona nada, que es
# justo lo que pasaba con "selecciona palabra" / "selecciona linea".
# Enviándolas como teclas extendidas, Shift+flecha sí extiende la selección.
# ---------------------------------------------------------------------------
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002

VK_TECLAS = {
    "ctrl": 0x11, "shift": 0x10, "alt": 0x12,
    "left": 0x25, "up": 0x26, "right": 0x27, "down": 0x28,
    "home": 0x24, "end": 0x23, "delete": 0x2E,
    "backspace": 0x08, "enter": 0x0D, "tab": 0x09,
}
VK_TECLAS.update({letra: ord(letra.upper()) for letra in "abcdefghijklmnopqrstuvwxyz"})
TECLAS_EXTENDIDAS = {"left", "up", "right", "down", "home", "end", "delete"}

# ---------------------------------------------------------------------------
# Dónde se guardan los documentos hechos por voz.
#
# El usuario dice solo el nombre ("guarda como informe de tesis"); el sistema
# agrega la extensión (.docx, .xlsx o .pptx según la aplicación) y elige la
# carpeta. Por defecto es una carpeta "Mouth AI Live" dentro de la carpeta
# Documentos del computador (se crea sola). Para usar otra carpeta, escribe
# aquí su ruta, por ejemplo: r"C:\Users\Samuel\Documents\Tesis"
# ---------------------------------------------------------------------------
NOMBRE_CARPETA_APP = "Mouth AI Live"
CARPETA_GUARDADO = None

WD_RUTA_DOCUMENTOS = 0    # WdDefaultFilePath.wdDocumentsPath
WD_FORMATO_DOCX = 16      # WdSaveFormat.wdFormatDocumentDefault
EXTENSIONES_WORD = (".docx", ".doc", ".docm")
EXTENSIONES_DOCUMENTOS = EXTENSIONES_WORD + (".xlsx", ".xlsm", ".pptx")


# El reconocedor escribe los números con letras ("cuatro"); los nombres de
# archivo suelen tener dígitos ("Taller4"). Se igualan antes de comparar.
NUMEROS_HABLADOS = {
    "cero": "0", "uno": "1", "dos": "2", "tres": "3", "cuatro": "4", "cinco": "5",
    "seis": "6", "siete": "7", "ocho": "8", "nueve": "9", "diez": "10",
}
PALABRAS_VACIAS = {"de", "del", "la", "el", "lo", "los", "las", "un", "una", "y", "en", "a", "mi", "con"}

# Números en letras hasta 999 999 para dictar valores en Excel ("ciento cincuenta").
_UNIDADES = {
    "cero": 0, "uno": 1, "un": 1, "una": 1, "dos": 2, "tres": 3, "cuatro": 4,
    "cinco": 5, "seis": 6, "siete": 7, "ocho": 8, "nueve": 9, "diez": 10,
    "once": 11, "doce": 12, "trece": 13, "catorce": 14, "quince": 15,
    "dieciseis": 16, "diecisiete": 17, "dieciocho": 18, "diecinueve": 19,
    "veinte": 20, "veintiuno": 21, "veintiun": 21, "veintidos": 22,
    "veintitres": 23, "veinticuatro": 24, "veinticinco": 25, "veintiseis": 26,
    "veintisiete": 27, "veintiocho": 28, "veintinueve": 29,
    "treinta": 30, "cuarenta": 40, "cincuenta": 50, "sesenta": 60,
    "setenta": 70, "ochenta": 80, "noventa": 90,
    "cien": 100, "ciento": 100, "doscientos": 200, "trescientos": 300,
    "cuatrocientos": 400, "quinientos": 500, "seiscientos": 600,
    "setecientos": 700, "ochocientos": 800, "novecientos": 900,
}

# Cómo suena cada letra de columna cuando el reconocedor la escribe.
_LETRAS_HABLADAS = {
    "a": "A", "be": "B", "b": "B", "ce": "C", "c": "C", "de": "D", "d": "D",
    "e": "E", "efe": "F", "f": "F", "ge": "G", "g": "G", "hache": "H", "h": "H",
    "i": "I", "jota": "J", "j": "J", "ka": "K", "k": "K", "ele": "L", "l": "L",
    "eme": "M", "m": "M", "ene": "N", "n": "N", "o": "O", "pe": "P", "p": "P",
    "cu": "Q", "q": "Q", "erre": "R", "r": "R", "ese": "S", "s": "S",
    "te": "T", "t": "T", "u": "U", "uve": "V", "ve": "V", "v": "V",
    "w": "W", "equis": "X", "x": "X", "ye": "Y", "y": "Y",
    "ze": "Z", "zeta": "Z", "z": "Z",
}


def _normalizar(texto: str) -> str:
    """Minúsculas, sin tildes, sin signos, con letras y dígitos separados
    ("Taller4_SG" -> "taller 4 sg") y números hablados como dígitos, para
    comparar nombres de archivo con lo que entendió el reconocedor."""
    texto = unicodedata.normalize("NFD", texto.lower())
    texto = "".join(c for c in texto if unicodedata.category(c) != "Mn")
    texto = re.sub(r"(?<=[a-z])(?=\d)|(?<=\d)(?=[a-z])", " ", texto)
    texto = re.sub(r"[^a-z0-9]+", " ", texto)
    return " ".join(NUMEROS_HABLADOS.get(p, p) for p in texto.split())


def _nombre_de_archivo_seguro(texto: str) -> str:
    """Convierte lo dictado en un nombre de archivo válido en Windows."""
    limpio = re.sub(r'[\\/:*?"<>|]', " ", texto or "")
    limpio = re.sub(r"\s+", " ", limpio).strip(" .")
    limpio = limpio[:80].strip()
    return limpio[:1].upper() + limpio[1:]


def _sin_tildes(texto: str) -> str:
    """Minúsculas, sin tildes ni signos (sin convertir números hablados)."""
    t = unicodedata.normalize("NFD", (texto or "").lower())
    t = "".join(c for c in t if unicodedata.category(c) != "Mn")
    return " ".join(re.sub(r"[^a-z0-9 ]", " ", t).split())


def _numero_hablado(texto: str):
    """'ciento cincuenta' -> 150, 'tres coma cinco' -> 3.5, '42' -> 42.
    Devuelve None si el texto no es un número. Rechaza cosas ambiguas como
    'uno dos' (que sumaría 3 por error)."""
    tokens = [p for p in _sin_tildes(texto).split() if p != "y"]
    if not tokens:
        return None

    if "coma" in tokens:
        i = tokens.index("coma")
        entero = _numero_hablado(" ".join(tokens[:i]))
        digitos = []
        for p in tokens[i + 1:]:
            v = int(p) if p.isdigit() and len(p) == 1 else _UNIDADES.get(p)
            if v is None or v > 9:
                return None
            digitos.append(str(v))
        if not isinstance(entero, int) or not digitos:
            return None
        return float(f"{entero}.{''.join(digitos)}")

    total = actual = 0
    maximo = 999999   # lo máximo que puede venir después del token anterior
    for p in tokens:
        if p == "mil":
            total += (actual or 1) * 1000
            actual, maximo = 0, 999
            continue
        v = int(p) if p.isdigit() else _UNIDADES.get(p)
        if v is None or v > maximo:
            return None
        actual += v
        maximo = 99 if v >= 100 else (9 if v >= 30 else 0)
    return total + actual


def _celda_hablada(texto: str):
    """'be cuatro' -> 'B4', 'ce veinte' -> 'C20', 'b4' -> 'B4'.
    Devuelve None si no se entiende."""
    t = _sin_tildes(texto)
    palabras = t.split()
    if palabras[:2] == ["doble", "ve"]:
        letra, resto = "W", palabras[2:]
    elif palabras and palabras[0] in _LETRAS_HABLADAS:
        letra, resto = _LETRAS_HABLADAS[palabras[0]], palabras[1:]
    else:
        m = re.fullmatch(r"([a-z]{1,3})(\d+)", t.replace(" ", ""))
        if not m:
            return None
        letra, resto = m.group(1).upper(), [m.group(2)]
    fila = _numero_hablado(" ".join(resto))
    if not isinstance(fila, int) or not 1 <= fila <= 1048576:
        return None
    return f"{letra}{fila}"


def _carpeta_documentos_windows() -> Path:
    """Carpeta Documentos real de Windows (incluye el caso de OneDrive)."""
    try:
        buf = ctypes.create_unicode_buffer(260)
        ctypes.windll.shell32.SHGetFolderPathW(None, 5, None, 0, buf)  # CSIDL_PERSONAL
        if buf.value:
            return Path(buf.value)
    except Exception:
        pass
    return Path.home() / "Documents"


class NavegadorAsistido:
    def __init__(self):
        self._activar_lector_de_pantalla()

    @staticmethod
    def _activar_lector_de_pantalla():
        """Avisa a Windows de que hay un lector de pantalla activo (lo mismo que
        hacen NVDA o JAWS). Al detectarlo, Chrome y Edge exponen el contenido de
        la página (enlaces, resultados de búsqueda) por UI Automation; sin este
        aviso, 'lee' no encuentra texto en el navegador. Debe hacerse antes de
        abrir el navegador. Se desactiva al cerrar el programa."""
        SPI_SETSCREENREADER = 0x0047
        SPIF_SENDCHANGE = 0x0002
        try:
            ctypes.windll.user32.SystemParametersInfoW(
                SPI_SETSCREENREADER, 1, None, SPIF_SENDCHANGE)
            atexit.register(
                lambda: ctypes.windll.user32.SystemParametersInfoW(
                    SPI_SETSCREENREADER, 0, None, SPIF_SENDCHANGE))
        except Exception:
            pass

    # ------------------------------------------------------------------
    # Detección de la aplicación en primer plano y acceso a Office por COM
    # ------------------------------------------------------------------
    def _exe_activo(self) -> str:
        """Ruta (en minúsculas) del programa que tiene la ventana en primer plano."""
        if not WIN32_DISPONIBLE:
            return ""
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            PROCESS_QUERY_INFORMATION = 0x0400
            PROCESS_VM_READ = 0x0010
            handle = win32api.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
            try:
                return win32process.GetModuleFileNameEx(handle, 0).lower()
            finally:
                win32api.CloseHandle(handle)
        except Exception:
            return ""

    def _app_activa_es_word(self) -> bool:
        return self._exe_activo().endswith("winword.exe")

    def _obtener_office(self, prog_id: str, crear: bool = False):
        """Word/Excel/PowerPoint por COM. Si no está abierto y crear=True, lo
        inicia. Devuelve None si no se puede."""
        if not WIN32_DISPONIBLE:
            return None
        try:
            return win32com.client.GetActiveObject(prog_id)
        except Exception:
            if not crear:
                return None
        try:
            app = win32com.client.Dispatch(prog_id)
            app.Visible = True
            return app
        except Exception:
            return None

    # ------------------------------------------------------------------
    # Dictado
    # ------------------------------------------------------------------
    def _dictar_en_word(self, texto: str):
        """Escribe directo en el documento activo de Word vía COM, sin
        pasar por el portapapeles ni simular teclas. Más confiable con
        tildes/eñes y no interfiere con lo que el usuario tenga copiado."""
        word_app = win32com.client.GetActiveObject("Word.Application")
        sel = word_app.Selection
        # Si el cursor queda pegado a la palabra anterior, agrega un espacio
        # (si no, "escribe hola" + "escribe mundo" daría "holamundo").
        if sel.Start == sel.End and sel.Start > 0 and texto[:1] not in ",.;:!?)":
            anterior = word_app.ActiveDocument.Range(sel.Start - 1, sel.Start).Text
            if anterior and anterior not in " \t\r\n\x0b\x0c\x07([{¿¡\"'«":
                texto = " " + texto
        sel.TypeText(texto)

    def _escribir_en_excel(self, texto: str):
        """Escribe en la celda activa (como número si el texto es un número)
        y baja una celda, para poder seguir dictando hacia abajo."""
        app = win32com.client.GetActiveObject("Excel.Application")
        celda = app.ActiveCell
        numero = _numero_hablado(texto)
        celda.Value = numero if numero is not None else texto[:1].upper() + texto[1:]
        direccion = celda.Address(False, False)
        escrito = celda.Value
        try:
            celda.Offset(1, 0).Select()
        except Exception:
            pass
        return True, f"Escrito en {direccion}: {escrito}"

    def escribir_texto(self, texto: str):
        if not texto or not texto.strip():
            return False, "No se especificó qué escribir"

        exe = self._exe_activo()
        if exe.endswith("winword.exe"):
            try:
                self._dictar_en_word(texto)
                return True, f"Texto dictado en Word: {texto}"
            except Exception:
                pass  # si falla la vía COM, seguimos con el método genérico
        elif exe.endswith("excel.exe"):
            try:
                return self._escribir_en_excel(texto)
            except Exception:
                pass  # celda en modo edición u otro error: seguimos con el pegado

        try:
            portapapeles_anterior = None
            try:
                portapapeles_anterior = pyperclip.paste()
            except Exception:
                pass

            pyperclip.copy(texto)
            time.sleep(0.1)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.1)

            if portapapeles_anterior is not None:
                pyperclip.copy(portapapeles_anterior)

            return True, f"Texto escrito: {texto}"
        except Exception as e:
            return False, f"Error al escribir el texto: {e}"

    def _pulsar_combinacion(self, teclas):
        """Pulsa una combinación (ej. ["shift", "end"]) enviando las teclas
        de navegación como extendidas. Si no estamos en Windows o hay una
        tecla desconocida, usa pyautogui como respaldo."""
        if not hasattr(ctypes, "windll") or any(t not in VK_TECLAS for t in teclas):
            pyautogui.hotkey(*teclas)
            return

        user32 = ctypes.windll.user32

        def evento(tecla, soltar):
            flags = KEYEVENTF_EXTENDEDKEY if tecla in TECLAS_EXTENDIDAS else 0
            if soltar:
                flags |= KEYEVENTF_KEYUP
            user32.keybd_event(VK_TECLAS[tecla], 0, flags, 0)

        for tecla in teclas:
            evento(tecla, soltar=False)
        time.sleep(0.02)
        for tecla in reversed(teclas):
            evento(tecla, soltar=True)
        time.sleep(0.05)  # deja que la app procese antes del siguiente paso

    # ------------------------------------------------------------------
    # Edición de texto por COM en Word, Excel y PowerPoint
    # ------------------------------------------------------------------
    def _editar_en_word(self, comando: str):
        """Versión de los comandos de edición que dependen del idioma de
        Office, hecha por COM (no usa atajos de teclado, así que funciona
        igual en Word en español, inglés, etc.).

        Devuelve (exito, mensaje) si manejó el comando, o None si el comando
        no tiene versión COM y debe resolverse con teclas."""
        if not WIN32_DISPONIBLE:
            return None

        word_app = win32com.client.GetActiveObject("Word.Application")
        sel = word_app.Selection

        # Selección por COM: no depende de teclas ni del estado de Bloq Num.
        if comando == "seleccionar_linea":
            sel.HomeKey(WD_LINEA, WD_MOVER)
            sel.EndKey(WD_LINEA, WD_EXTENDER)
            return True, "Línea seleccionada"

        if comando == "seleccionar_palabra_anterior":
            sel.MoveLeft(WD_PALABRA, 1, WD_EXTENDER)
            return True, "Palabra seleccionada"

        if comando == "seleccionar_palabra_siguiente":
            sel.MoveRight(WD_PALABRA, 1, WD_EXTENDER)
            return True, "Palabra seleccionada"

        if comando == "seleccionar_hasta_inicio":
            sel.HomeKey(WD_LINEA, WD_EXTENDER)
            return True, "Seleccionado hasta el inicio"

        if comando == "seleccionar_hasta_fin":
            sel.EndKey(WD_LINEA, WD_EXTENDER)
            return True, "Seleccionado hasta el final"

        if comando == "borrar_linea":
            sel.HomeKey(WD_LINEA, WD_MOVER)
            sel.EndKey(WD_LINEA, WD_EXTENDER)
            if sel.Start != sel.End:  # en una línea vacía, Delete borraría el salto de línea
                sel.Delete()
            return True, "Línea borrada"

        if comando in ("negrita", "cursiva"):
            propiedad = "Bold" if comando == "negrita" else "Italic"
            setattr(sel.Font, propiedad, WD_TOGGLE)
            valor = getattr(sel.Font, propiedad)
            if valor == 0:
                estado = "desactivada"
            elif valor == WD_INDEFINIDO:
                estado = "aplicada"
            else:
                estado = "activada"
            return True, f"{comando.capitalize()} {estado}"

        if comando == "subrayado":
            activar = sel.Font.Underline == 0
            sel.Font.Underline = 1 if activar else 0
            return True, "Subrayado activado" if activar else "Subrayado desactivado"

        if comando in ("letra_grande", "letra_pequena"):
            if comando == "letra_grande":
                sel.Font.Grow()
            else:
                sel.Font.Shrink()
            tamano = sel.Font.Size
            if isinstance(tamano, (int, float)) and 0 < tamano < 1000:
                return True, f"Tamaño {tamano:g}"
            return True, "Tamaño de letra cambiado"

        if comando == "seleccionar_todo":
            sel.WholeStory()
            return True, "Todo seleccionado"

        if comando == "guardar":
            documento = word_app.ActiveDocument
            if documento.Path:
                documento.Save()
                return True, "Documento guardado"
            # Documento nuevo: no abrimos el diálogo "Guardar como" porque
            # no se puede llenar por voz. Se guía al usuario a "guarda como".
            return True, "Es un documento nuevo. Di: guarda como, y el nombre"

        return None

    def _editar_en_excel(self, comando: str):
        """Formato y guardado por COM en Excel. None = resolver con teclas."""
        if not WIN32_DISPONIBLE:
            return None

        app = win32com.client.GetActiveObject("Excel.Application")
        sel = app.Selection

        if comando in ("negrita", "cursiva"):
            propiedad = "Bold" if comando == "negrita" else "Italic"
            actual = getattr(sel.Font, propiedad)
            nuevo = True if actual is None else not bool(actual)
            setattr(sel.Font, propiedad, nuevo)
            return True, f"{comando.capitalize()} {'activada' if nuevo else 'desactivada'}"

        if comando == "subrayado":
            actual = sel.Font.Underline
            activar = actual is None or actual == XL_SIN_SUBRAYADO
            sel.Font.Underline = XL_SUBRAYADO_SIMPLE if activar else XL_SIN_SUBRAYADO
            return True, "Subrayado activado" if activar else "Subrayado desactivado"

        if comando in ("letra_grande", "letra_pequena"):
            tamano = sel.Font.Size
            if not isinstance(tamano, (int, float)):
                return True, "Selecciona celdas con el mismo tamaño de letra"
            nuevo = tamano + 1 if comando == "letra_grande" else max(1, tamano - 1)
            sel.Font.Size = nuevo
            return True, f"Tamaño {nuevo:g}"

        if comando == "seleccionar_todo":
            app.Cells.Select()
            return True, "Todo seleccionado"

        if comando == "guardar":
            libro = app.ActiveWorkbook
            if libro.Path:
                libro.Save()
                return True, "Libro guardado"
            return True, "Es un libro nuevo. Di: guarda como, y el nombre"

        return None

    def _editar_en_ppt(self, comando: str):
        """Formato y guardado por COM en PowerPoint. None = resolver con teclas."""
        if not WIN32_DISPONIBLE:
            return None

        app = win32com.client.GetActiveObject("PowerPoint.Application")

        if comando == "guardar":
            pres = app.ActivePresentation
            if pres.Path:
                pres.Save()
                return True, "Presentación guardada"
            return True, "Es una presentación nueva. Di: guarda como, y el nombre"

        if comando in ("negrita", "cursiva", "subrayado"):
            # Falla si no hay texto seleccionado; entonces se usan las teclas.
            fuente = app.ActiveWindow.Selection.TextRange.Font
            propiedad = {"negrita": "Bold", "cursiva": "Italic", "subrayado": "Underline"}[comando]
            activar = getattr(fuente, propiedad) != -1   # -1 = msoTrue
            setattr(fuente, propiedad, -1 if activar else 0)
            return True, f"{comando.capitalize()} {'activada' if activar else 'desactivada'}"

        if comando in ("letra_grande", "letra_pequena"):
            fuente = app.ActiveWindow.Selection.TextRange.Font
            tamano = fuente.Size
            if not isinstance(tamano, (int, float)) or tamano <= 0:
                return True, "Selecciona texto con el mismo tamaño de letra"
            nuevo = tamano + 2 if comando == "letra_grande" else max(1, tamano - 2)
            fuente.Size = nuevo
            return True, f"Tamaño {nuevo:g}"

        return None

    def editar(self, comando: str):
        """Comandos de edición de texto (copiar, pegar, cortar, borrar,
        seleccionar, mover el cursor, formato). Ver TECLAS_EDICION."""
        exe = self._exe_activo()
        try:
            resultado = None
            if exe.endswith("winword.exe"):
                resultado = self._editar_en_word(comando)
            elif exe.endswith("excel.exe"):
                resultado = self._editar_en_excel(comando)
            elif exe.endswith("powerpnt.exe"):
                resultado = self._editar_en_ppt(comando)
            if resultado is not None:
                return resultado
        except Exception:
            pass  # si falla COM, seguimos con las teclas

        entrada = TECLAS_EDICION.get(comando)
        if entrada is None:
            if comando in ("letra_grande", "letra_pequena"):
                return False, "Cambiar el tamaño de letra solo funciona en Word, Excel y PowerPoint"
            return False, f"Comando de edición no reconocido: {comando}"

        mensaje, pasos = entrada
        try:
            for paso in pasos:
                self._pulsar_combinacion(paso)
            return True, mensaje
        except Exception as e:
            return False, f"Error al editar ({comando}): {e}"

    # ------------------------------------------------------------------
    # Documentos de Word: crear, guardar con nombre y abrir por voz
    # ------------------------------------------------------------------
    def _obtener_word(self, crear: bool = False):
        """Devuelve Word (COM). Si no está abierto y crear=True, lo inicia.
        Devuelve None si no se puede."""
        return self._obtener_office("Word.Application", crear)

    def _traer_word_al_frente(self, word_app):
        """Mejor esfuerzo: que Word (o Excel/PowerPoint) quede al frente para
        seguir dictando."""
        try:
            word_app.Activate()
        except Exception:
            pass
        try:
            hwnd = word_app.ActiveWindow.Hwnd
            if win32gui.IsIconic(hwnd):
                win32gui.ShowWindow(hwnd, 9)  # SW_RESTORE, solo si está minimizada
            win32gui.SetForegroundWindow(hwnd)
        except Exception:
            pass

    def _carpeta_base_documentos(self, word_app=None) -> Path:
        try:
            ruta = word_app.Options.DefaultFilePath(WD_RUTA_DOCUMENTOS)  # solo existe en Word
        except Exception:
            ruta = ""
        return Path(ruta) if ruta else _carpeta_documentos_windows()

    def _carpeta_para_guardar(self, word_app=None) -> Path:
        if CARPETA_GUARDADO:
            carpeta = Path(CARPETA_GUARDADO)
        else:
            carpeta = self._carpeta_base_documentos(word_app) / NOMBRE_CARPETA_APP
        carpeta.mkdir(parents=True, exist_ok=True)
        return carpeta

    def _listar_documentos(self, word_app=None):
        """Documentos de Word, Excel y PowerPoint en la carpeta de la app y
        en Documentos."""
        carpetas = [self._carpeta_para_guardar(word_app)]
        if not CARPETA_GUARDADO:
            carpetas.append(self._carpeta_base_documentos(word_app))
        vistos, documentos = set(), []
        for carpeta in carpetas:
            if not carpeta.exists():
                continue
            for ruta in carpeta.iterdir():
                if (ruta.suffix.lower() in EXTENSIONES_DOCUMENTOS and not ruta.name.startswith("~$")
                        and ruta not in vistos):
                    vistos.add(ruta)
                    documentos.append(ruta)
        return documentos

    @staticmethod
    def _elegir_documento(consulta: str, documentos):
        """El documento que mejor coincide con lo dicho, o None."""
        q = _normalizar(consulta)
        if not q or not documentos:
            return None
        palabras = [p for p in q.split() if p not in PALABRAS_VACIAS] or q.split()

        def coincide(nombre: str) -> bool:
            # Cada palabra dicha debe aparecer en el nombre (en cualquier
            # orden). Los números deben coincidir completos: "4" no es "40".
            return all(
                re.search(rf"\b{p}\b", nombre) if p.isdigit() else p in nombre
                for p in palabras
            )

        # 1) el nombre contiene todas las palabras dichas; gana el más corto
        contienen = [d for d in documentos if coincide(_normalizar(d.stem))]
        if contienen:
            return min(contienen, key=lambda d: len(d.stem))
        # 2) último recurso: parecido aproximado, por si el reconocedor se
        #    equivocó en una letra. Umbral alto para no abrir uno equivocado.
        #    Los números dichos deben estar exactos en el nombre: "reporte 4"
        #    es casi igual a "reporte 40" como texto, pero no es el mismo.
        numeros = set(re.findall(r"\d+", q))
        por_nombre = {}
        for d in documentos:
            nombre = _normalizar(d.stem)
            if numeros <= set(re.findall(r"\d+", nombre)):
                por_nombre[nombre] = d
        cercanos = difflib.get_close_matches(q, list(por_nombre), n=1, cutoff=0.8)
        return por_nombre[cercanos[0]] if cercanos else None

    def documento_nuevo(self):
        """Documento en blanco. Evita la pantalla de inicio de Word, donde
        habría que elegir "Documento en blanco" con el mouse."""
        word_app = self._obtener_word(crear=True)
        if word_app is None:
            return False, "No se pudo controlar Word"
        try:
            word_app.Documents.Add()
            self._traer_word_al_frente(word_app)
            return True, "Documento nuevo en blanco"
        except Exception as e:
            return False, f"No se pudo crear el documento: {e}"

    def _app_office_activa(self):
        """(tipo, app) de la aplicación de Office en primer plano; si ninguna
        lo está, la primera que esté abierta. (None, None) si no hay ninguna."""
        exe = self._exe_activo()
        orden = [("word", "Word.Application", "winword.exe"),
                 ("excel", "Excel.Application", "excel.exe"),
                 ("ppt", "PowerPoint.Application", "powerpnt.exe")]
        for tipo, prog, nombre in orden:
            if exe.endswith(nombre):
                app = self._obtener_office(prog)
                if app is not None:
                    return tipo, app
        for tipo, prog, _ in orden:
            app = self._obtener_office(prog)
            if app is not None:
                return tipo, app
        return None, None

    def guardar_como(self, nombre: str):
        """Guarda el documento, libro o presentación activo con el nombre
        dicho, en la carpeta configurada y con la extensión que corresponda
        (.docx, .xlsx o .pptx); el usuario no la menciona."""
        tipo, app = self._app_office_activa()
        if app is None:
            return False, "No hay Word, Excel ni PowerPoint abierto"
        nombre_limpio = _nombre_de_archivo_seguro(nombre)
        if not nombre_limpio:
            return False, "No entendí el nombre del documento"
        alertas_antes = None
        try:
            if tipo == "word":
                if app.Documents.Count == 0:
                    return False, "No hay ningún documento abierto"
                objeto, ext = app.ActiveDocument, ".docx"
            elif tipo == "excel":
                if app.Workbooks.Count == 0:
                    return False, "No hay ningún libro abierto"
                objeto, ext = app.ActiveWorkbook, ".xlsx"
            else:
                if app.Presentations.Count == 0:
                    return False, "No hay ninguna presentación abierta"
                objeto, ext = app.ActivePresentation, ".pptx"

            carpeta = self._carpeta_para_guardar(app)
            actual = os.path.normcase(str(objeto.FullName))

            ruta = carpeta / f"{nombre_limpio}{ext}"
            contador = 2
            # No pisar otro archivo: si el nombre existe, agrega (2), (3)...
            while ruta.exists() and os.path.normcase(str(ruta)) != actual:
                ruta = carpeta / f"{nombre_limpio} ({contador}){ext}"
                contador += 1

            if tipo == "word":
                objeto.SaveAs2(str(ruta), WD_FORMATO_DOCX)
            elif tipo == "excel":
                alertas_antes = app.DisplayAlerts
                app.DisplayAlerts = False
                objeto.SaveAs(str(ruta), XL_FORMATO_XLSX)
            else:
                objeto.SaveAs(str(ruta), PPT_FORMATO_PPTX)
            return True, f"Guardado como {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo guardar el documento: {e}"
        finally:
            if alertas_antes is not None:
                try:
                    app.DisplayAlerts = alertas_antes
                except Exception:
                    pass

    def _abrir_ruta(self, ruta: Path):
        """Abre un documento: Word por COM; Excel y PowerPoint con su
        programa asociado de Windows."""
        if ruta.suffix.lower() in EXTENSIONES_WORD:
            word_app = self._obtener_word(crear=True)
            if word_app is None:
                raise RuntimeError("No se pudo controlar Word")
            word_app.Documents.Open(str(ruta))
            self._traer_word_al_frente(word_app)
        else:
            os.startfile(str(ruta))

    def abrir_documento(self, consulta: str):
        """Abre un documento guardado buscándolo por nombre (aproximado)."""
        try:
            ruta = self._elegir_documento(consulta, self._listar_documentos())
            if ruta is None:
                return False, f"No encontré un documento parecido a {consulta}"
            self._abrir_ruta(ruta)
            return True, f"Abriendo {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo abrir el documento: {e}"

    def abrir_ultimo_documento(self):
        """Abre el documento modificado más recientemente."""
        try:
            documentos = self._listar_documentos()
            if not documentos:
                return False, "No hay documentos guardados"
            ruta = max(documentos, key=lambda d: d.stat().st_mtime)
            self._abrir_ruta(ruta)
            return True, f"Abriendo {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo abrir el documento: {e}"

    # ------------------------------------------------------------------
    # Excel por voz
    # ------------------------------------------------------------------
    def _describir_celda(self, celda) -> str:
        valor = celda.Value
        texto = "vacía" if valor is None else str(valor)
        return f"{celda.Address(False, False)}, {texto}"

    def excel_nuevo(self):
        """Libro en blanco. Evita la pantalla de inicio de Excel."""
        app = self._obtener_office("Excel.Application", crear=True)
        if app is None:
            return False, "No se pudo controlar Excel"
        try:
            app.Workbooks.Add()
            self._traer_word_al_frente(app)
            return True, "Libro nuevo en blanco"
        except Exception as e:
            return False, f"No se pudo crear el libro: {e}"

    def excel_mover(self, direccion: str):
        """Mueve la celda activa y narra su dirección y contenido."""
        desplazamientos = {"arriba": (-1, 0), "abajo": (1, 0), "izquierda": (0, -1), "derecha": (0, 1)}
        d = desplazamientos.get(direccion)
        if d is None:
            return False, f"Dirección inválida: {direccion}"
        app = self._obtener_office("Excel.Application")
        if app is None or app.Workbooks.Count == 0:
            return False, "Excel no está abierto"
        try:
            app.ActiveCell.Offset(d[0], d[1]).Select()
        except Exception:
            return False, "No hay más celdas en esa dirección"
        return True, self._describir_celda(app.ActiveCell)

    def excel_ir_celda(self, texto: str):
        """Va a una celda dicha en voz alta: 'be cuatro' -> B4."""
        direccion = _celda_hablada(texto)
        if direccion is None:
            return False, f"No entendí la celda: {texto}"
        app = self._obtener_office("Excel.Application")
        if app is None or app.Workbooks.Count == 0:
            return False, "Excel no está abierto"
        try:
            app.Range(direccion).Select()
            return True, self._describir_celda(app.ActiveCell)
        except Exception as e:
            return False, f"No se pudo ir a la celda {direccion}: {e}"

    def excel_sumar(self, direccion: str = "arriba"):
        """Pone =SUM(...) en la celda activa con los números de arriba o de
        la izquierda."""
        app = self._obtener_office("Excel.Application")
        if app is None or app.Workbooks.Count == 0:
            return False, "Excel no está abierto"
        try:
            celda = app.ActiveCell
            if direccion == "arriba":
                if celda.Row == 1 or celda.Offset(-1, 0).Value is None:
                    return False, "No hay números arriba de esta celda"
                fin = celda.Offset(-1, 0)
                inicio = fin.End(XL_ARRIBA)
            else:
                if celda.Column == 1 or celda.Offset(0, -1).Value is None:
                    return False, "No hay números a la izquierda de esta celda"
                fin = celda.Offset(0, -1)
                inicio = fin.End(XL_IZQUIERDA)
            rango = app.Range(inicio, fin)
            celda.Formula = f"=SUM({rango.Address(False, False)})"  # .Formula siempre usa nombres en inglés
            return True, f"Suma en {celda.Address(False, False)}: {celda.Text}"
        except Exception as e:
            return False, f"No se pudo sumar: {e}"

    # ------------------------------------------------------------------
    # PowerPoint por voz
    # ------------------------------------------------------------------
    def _ppt_activo(self):
        app = self._obtener_office("PowerPoint.Application")
        if app is None or app.Presentations.Count == 0:
            return None, None
        return app, app.ActivePresentation

    def _ppt_diapositiva_actual(self, app, pres):
        try:
            return app.ActiveWindow.View.Slide
        except Exception:
            return pres.Slides(pres.Slides.Count)

    def ppt_nuevo(self):
        """Presentación en blanco con una diapositiva de título."""
        app = self._obtener_office("PowerPoint.Application", crear=True)
        if app is None:
            return False, "No se pudo controlar PowerPoint"
        try:
            pres = app.Presentations.Add()
            pres.Slides.AddSlide(1, pres.SlideMaster.CustomLayouts(1))
            self._traer_word_al_frente(app)
            return True, "Presentación nueva. Di título y el texto para la primera diapositiva"
        except Exception as e:
            return False, f"No se pudo crear la presentación: {e}"

    def ppt_nueva_diapositiva(self):
        """Diapositiva de 'Título y contenido' después de la actual."""
        app, pres = self._ppt_activo()
        if app is None:
            return False, "PowerPoint no está abierto"
        try:
            try:
                indice = self._ppt_diapositiva_actual(app, pres).SlideIndex + 1
            except Exception:
                indice = pres.Slides.Count + 1
            pres.Slides.AddSlide(indice, pres.SlideMaster.CustomLayouts(2))
            try:
                app.ActiveWindow.View.GotoSlide(indice)
            except Exception:
                pass
            return True, f"Diapositiva {indice} de {pres.Slides.Count}"
        except Exception as e:
            return False, f"No se pudo crear la diapositiva: {e}"

    def ppt_titulo(self, texto: str):
        app, pres = self._ppt_activo()
        if app is None:
            return False, "PowerPoint no está abierto"
        try:
            diapositiva = self._ppt_diapositiva_actual(app, pres)
            diapositiva.Shapes.Title.TextFrame.TextRange.Text = texto[:1].upper() + texto[1:]
            return True, f"Título: {texto}"
        except Exception:
            return False, "Esta diapositiva no tiene espacio para título"

    def ppt_contenido(self, texto: str):
        """Agrega un punto al cuerpo de la diapositiva actual."""
        app, pres = self._ppt_activo()
        if app is None:
            return False, "PowerPoint no está abierto"
        try:
            diapositiva = self._ppt_diapositiva_actual(app, pres)
            for forma in diapositiva.Shapes:
                # 14 = marcador de posición; tipos: 2 = cuerpo, 4 = subtítulo, 7 = objeto
                if forma.Type == 14 and forma.PlaceholderFormat.Type in (2, 4, 7) and forma.HasTextFrame:
                    rango = forma.TextFrame.TextRange
                    nuevo = texto[:1].upper() + texto[1:]
                    if rango.Text.strip():
                        rango.InsertAfter("\r" + nuevo)
                    else:
                        rango.Text = nuevo
                    return True, f"Agregado: {texto}"
            return False, "Esta diapositiva no tiene espacio para contenido"
        except Exception as e:
            return False, f"No se pudo agregar el contenido: {e}"

    def ppt_diapositiva(self, direccion: str):
        """Pasa a la diapositiva siguiente o anterior (también en pantalla completa)."""
        app, pres = self._ppt_activo()
        if app is None:
            return False, "PowerPoint no está abierto"
        if direccion not in ("siguiente", "anterior"):
            return False, f"Dirección inválida: {direccion}"
        try:
            if app.SlideShowWindows.Count > 0:   # presentación en pantalla completa
                vista = app.SlideShowWindows(1).View
                if direccion == "siguiente":
                    vista.Next()
                else:
                    vista.Previous()
                return True, f"Diapositiva {vista.CurrentShowPosition}"
            vista = app.ActiveWindow.View
            actual = vista.Slide.SlideIndex
            nuevo = actual + 1 if direccion == "siguiente" else actual - 1
            if nuevo < 1 or nuevo > pres.Slides.Count:
                return False, "No hay más diapositivas en esa dirección"
            vista.GotoSlide(nuevo)
            return True, f"Diapositiva {nuevo} de {pres.Slides.Count}"
        except Exception as e:
            return False, f"No se pudo cambiar de diapositiva: {e}"

    def ppt_presentar(self, modo: str):
        """Inicia o termina la presentación en pantalla completa."""
        app, pres = self._ppt_activo()
        if app is None:
            return False, "PowerPoint no está abierto"
        try:
            if modo == "inicio":
                pres.SlideShowSettings.Run()
                return True, "Presentación iniciada"
            if app.SlideShowWindows.Count > 0:
                app.SlideShowWindows(1).View.Exit()
            return True, "Presentación terminada"
        except Exception as e:
            return False, f"No se pudo controlar la presentación: {e}"

    # ------------------------------------------------------------------
    # Búsqueda, navegación, lectura, mouse y voz
    # ------------------------------------------------------------------
    def buscar_google(self, consulta: str):
        """Escribe la consulta en la barra de direcciones. Se pega desde el
        portapapeles (no se teclea) para que las tildes y eñes salgan bien."""
        if not consulta or not consulta.strip():
            return False, "No se especificó qué buscar"
        anterior = None
        try:
            try:
                anterior = pyperclip.paste()
            except Exception:
                pass
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.3)
            pyperclip.copy(consulta)
            time.sleep(0.1)
            pyautogui.hotkey("ctrl", "v")
            time.sleep(0.15)
            pyautogui.press("enter")
            time.sleep(0.2)
            return True, f"Buscando en Google: {consulta}"
        except Exception as e:
            return False, f"Error al realizar la búsqueda: {e}"
        finally:
            if anterior is not None:
                try:
                    pyperclip.copy(anterior)
                except Exception:
                    pass

    def navegar(self, comando: str):
        tecla = TECLAS_NAVEGACION.get(comando)
        if tecla is None:
            return False, f"Comando de navegación no reconocido: {comando}"
        try:
            if isinstance(tecla, list):
                pyautogui.hotkey(*tecla)
            else:
                pyautogui.press(tecla)
            return True, f"Navegación ejecutada: {comando}"
        except Exception as e:
            return False, f"Error al navegar ({comando}): {e}"

    @staticmethod
    def _intentar(funcion) -> str:
        """Ejecuta funcion() y devuelve su resultado como texto; '' si falla."""
        try:
            valor = funcion()
            return str(valor).strip() if valor is not None else ""
        except Exception:
            return ""

    def _texto_de_control(self, control) -> str:
        """Prueba varias fuentes de UI Automation hasta encontrar texto legible."""
        i = self._intentar
        nombre = i(lambda: control.Name)
        valor = (i(lambda: control.GetValuePattern().Value)
                 or i(lambda: control.GetLegacyIAccessiblePattern().Value))
        if nombre and valor and valor != nombre:
            return f"{nombre}: {valor}"[:600]
        if nombre or valor:
            return (nombre or valor)[:600]

        for fuente in (
            lambda: control.GetTextPattern().DocumentRange.GetText(500),
            lambda: control.GetWindowText(),
            lambda: control.GetLegacyIAccessiblePattern().Name,
            lambda: control.GetLegacyIAccessiblePattern().Description,
            lambda: control.HelpText,
            lambda: " ".join(h.Name.strip() for h in control.GetChildren()[:8]
                             if h.Name and h.Name.strip()),
        ):
            texto = i(fuente)
            if texto:
                return texto[:600]
        return ""

    def _leer_office(self):
        """Texto a leer cuando la app activa es Word, Excel o PowerPoint (por
        COM). Devuelve None si no aplica o si COM falla (por ejemplo, con un
        diálogo abierto); entonces se usa UI Automation."""
        exe = self._exe_activo()
        try:
            if exe.endswith("winword.exe"):
                sel = win32com.client.GetActiveObject("Word.Application").Selection
                if sel.Start != sel.End:
                    return sel.Text.strip()[:600]
                parrafo = sel.Paragraphs(1).Range.Text.strip("\r\x07\x0b \t\n")
                return parrafo[:600] if parrafo else "Párrafo vacío"

            if exe.endswith("excel.exe"):
                app = win32com.client.GetActiveObject("Excel.Application")
                return self._describir_celda(app.ActiveCell)

            if exe.endswith("powerpnt.exe"):
                app = win32com.client.GetActiveObject("PowerPoint.Application")
                try:
                    sel = app.ActiveWindow.Selection
                    if sel.Type == 3 and sel.TextRange.Text.strip():  # ppSelectionText
                        return sel.TextRange.Text.strip()[:600]
                except Exception:
                    pass
                pres = app.ActivePresentation
                diapositiva = self._ppt_diapositiva_actual(app, pres)
                partes = []
                for forma in diapositiva.Shapes:
                    try:
                        if forma.HasTextFrame and forma.TextFrame.HasText:
                            partes.append(forma.TextFrame.TextRange.Text.replace("\r", ". ").strip())
                    except Exception:
                        continue
                cuerpo = " ".join(partes) if partes else "sin texto"
                return f"Diapositiva {diapositiva.SlideIndex}. {cuerpo}"[:600]
        except Exception:
            return None
        return None

    def _es_navegador(self) -> bool:
        return self._exe_activo().endswith(
            ("chrome.exe", "msedge.exe", "firefox.exe", "brave.exe", "opera.exe"))

    @staticmethod
    def _limpiar_texto_copiado(texto: str, limite: int) -> str:
        """Une las líneas con puntos, quita líneas vacías o de 1-2 caracteres
        (menús, iconos) y recorta al límite sin cortar a mitad de palabra."""
        lineas = [l.strip() for l in texto.replace("\r", "\n").split("\n")]
        unido = ". ".join(l for l in lineas if len(l) > 2)
        if len(unido) <= limite:
            return unido
        corte = unido[:limite].rsplit(" ", 1)[0]
        return corte

    def _leer_seleccion_portapapeles(self, todo: bool = False, limite: int = 1500) -> str:
        """Devuelve el texto seleccionado en el navegador (Ctrl+C) o ''. Con
        todo=True selecciona primero toda la página (Ctrl+A). Usa una marca en el
        portapapeles para saber si de verdad se copió algo, y deja el portapapeles
        del usuario como estaba."""
        anterior = None
        marca = f"__mouth_ai_live_{time.time()}__"
        try:
            try:
                anterior = pyperclip.paste()
            except Exception:
                pass
            pyperclip.copy(marca)
            time.sleep(0.05)
            if todo:
                pyautogui.hotkey("ctrl", "a")
                time.sleep(0.2)
            pyautogui.hotkey("ctrl", "c")
            time.sleep(0.3 if todo else 0.2)
            texto = pyperclip.paste()
            if texto and texto != marca:
                return self._limpiar_texto_copiado(texto, limite)
            return ""
        except Exception:
            return ""
        finally:
            if anterior is not None:
                try:
                    pyperclip.copy(anterior)
                except Exception:
                    pass

    def leer_elemento_enfocado(self, pagina: bool = False):
        texto_office = self._leer_office()
        if texto_office:
            self._hablar(texto_office)
            return True, texto_office

        # En el navegador: 'lee la pagina' selecciona y lee el texto de la página;
        # 'lee' lee el texto seleccionado, si lo hay.
        if self._es_navegador():
            seleccion = self._leer_seleccion_portapapeles(todo=pagina)
            if seleccion:
                self._hablar(seleccion)
                return True, seleccion

        if not UIAUTOMATION_DISPONIBLE:
            mensaje = "uiautomation no está instalado - no se puede leer la pantalla"
            self._hablar("La función de lectura no está disponible")
            return False, mensaje

        try:
            control = auto.GetFocusedControl()
            if control is None:
                self._hablar("No se pudo identificar el elemento actual")
                return False, "No se encontró elemento enfocado"

            texto = self._texto_de_control(control)
            if not texto:
                self._hablar("Este elemento no tiene texto para leer")
                return False, "Elemento sin texto legible"

            self._hablar(texto)
            return True, texto
        except Exception as e:
            mensaje = f"Error al leer el elemento enfocado: {e}"
            self._hablar("Ocurrió un error al intentar leer")
            return False, mensaje

    def mover_mouse(self, direccion: str, paso: int = 40):
        offsets = {
            "arriba": (0, -paso),
            "abajo": (0, paso),
            "izquierda": (-paso, 0),
            "derecha": (paso, 0),
        }
        offset = offsets.get(direccion)
        if offset is None:
            return False, f"Dirección de mouse inválida: {direccion}"
        try:
            ancho, alto = pyautogui.size()
            x_actual, y_actual = pyautogui.position()
            # Margen de 2px: nunca tocar exactamente (0,0) ni el borde opuesto,
            # porque eso dispara el failsafe de PyAutoGUI y bloquea la próxima acción.
            x_nuevo = max(2, min(x_actual + offset[0], ancho - 3))
            y_nuevo = max(2, min(y_actual + offset[1], alto - 3))
            pyautogui.moveTo(x_nuevo, y_nuevo, duration=0)
            return True, f"Mouse movido: {direccion} -> ({x_nuevo}, {y_nuevo})"
        except Exception as e:
            return False, f"Error al mover el mouse: {e}"

    def clic_mouse(self, tipo: str):
        try:
            if tipo == "izquierdo":
                pyautogui.click(button="left")
            elif tipo == "derecho":
                pyautogui.click(button="right")
            elif tipo == "doble":
                pyautogui.doubleClick(button="left")
            else:
                return False, f"Tipo de clic inválido: {tipo}"
            return True, f"Clic {tipo} ejecutado"
        except Exception as e:
            return False, f"Error al hacer clic: {e}"

    def hablar(self, texto: str):
        """Método público para que otros módulos (ej. actions.py) pidan narrar algo."""
        self._hablar(texto)

    def _hablar(self, texto: str):
        try:
            motor = pyttsx3.init()
            motor.setProperty("rate", 175)
            motor.setProperty("volume", 1.0)
            motor.say(texto)
            motor.runAndWait()
            motor.stop()
            del motor
        except Exception as e:
            print(f"[TTS] Error al reproducir voz: {e}")


if __name__ == "__main__":
    nav = NavegadorAsistido()
    print(nav.navegar("abajo"))
    print(nav.leer_elemento_enfocado())