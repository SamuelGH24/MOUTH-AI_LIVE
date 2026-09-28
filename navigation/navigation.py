"""
Módulo de Navegación Asistida y Dictado - Mouth AI Live

Responsabilidades:
1. Navegación por teclado (arriba/abajo/Tab/Enter/atrás), tal como lo haría
   una persona con teclado.
2. Lectura en voz alta del elemento enfocado (TTS + UI Automation).
3. Búsqueda en Google escribiendo en la barra de direcciones.
4. Dictado: escribir texto dictado por voz en cualquier campo de texto
   (Word, Bloc de notas, cajas de búsqueda, etc.) y comandos básicos de
   edición (nueva línea, borrar, deshacer, guardar, seleccionar todo).

NOTA sobre el dictado: el texto se escribe pegándolo desde el portapapeles
(Ctrl+V) en vez de simular cada tecla una por una. Esto es más confiable
con español (tildes, eñes) que la escritura carácter por carácter, que
puede fallar con acentos según la configuración de teclado de Windows.
Efecto secundario a tener en cuenta: esto reemplaza temporalmente lo que
el usuario tuviera copiado en el portapapeles.

Requiere (instalar en el entorno de Windows):
    pip install pyautogui pyttsx3 uiautomation pyperclip
"""

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
# Word en español pueden ser Ctrl+N, Ctrl+G, Ctrl+E...). Por eso en Word se
# resuelven por COM (ver _editar_en_word) y estas teclas quedan solo como
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
    # Respaldo fuera de Word (en Word se usa COM)
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
# agrega la extensión (.docx) y elige la carpeta. Por defecto es una carpeta
# "Mouth AI Live" dentro de la carpeta Documentos que use Word en ese
# computador (se crea sola). Para usar otra carpeta, escribe aquí su ruta,
# por ejemplo: r"C:\Users\Samuel\Documents\Tesis"
# ---------------------------------------------------------------------------
NOMBRE_CARPETA_APP = "Mouth AI Live"
CARPETA_GUARDADO = None

WD_RUTA_DOCUMENTOS = 0    # WdDefaultFilePath.wdDocumentsPath
WD_FORMATO_DOCX = 16      # WdSaveFormat.wdFormatDocumentDefault
EXTENSIONES_WORD = (".docx", ".doc", ".docm")


# El reconocedor escribe los números con letras ("cuatro"); los nombres de
# archivo suelen tener dígitos ("Taller4"). Se igualan antes de comparar.
NUMEROS_HABLADOS = {
    "cero": "0", "uno": "1", "dos": "2", "tres": "3", "cuatro": "4", "cinco": "5",
    "seis": "6", "siete": "7", "ocho": "8", "nueve": "9", "diez": "10",
}
PALABRAS_VACIAS = {"de", "del", "la", "el", "lo", "los", "las", "un", "una", "y", "en", "a", "mi", "con"}


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


class NavegadorAsistido:
    def __init__(self):
        pass

    def _app_activa_es_word(self) -> bool:
        if not WIN32_DISPONIBLE:
            return False
        try:
            hwnd = win32gui.GetForegroundWindow()
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            PROCESS_QUERY_INFORMATION = 0x0400
            PROCESS_VM_READ = 0x0010
            handle = win32api.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, pid)
            try:
                ruta = win32process.GetModuleFileNameEx(handle, 0)
            finally:
                win32api.CloseHandle(handle)
            return ruta.lower().endswith("winword.exe")
        except Exception:
            return False

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

    def escribir_texto(self, texto: str):
        if not texto or not texto.strip():
            return False, "No se especificó qué escribir"

        if self._app_activa_es_word():
            try:
                self._dictar_en_word(texto)
                return True, f"Texto dictado en Word: {texto}"
            except Exception:
                pass  # si falla la vía COM, seguimos con el método genérico

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

    def editar(self, comando: str):
        """Comandos de edición de texto (copiar, pegar, cortar, borrar,
        seleccionar, mover el cursor, formato). Ver TECLAS_EDICION."""
        if self._app_activa_es_word():
            try:
                resultado = self._editar_en_word(comando)
                if resultado is not None:
                    return resultado
            except Exception:
                pass  # si falla COM, seguimos con las teclas

        entrada = TECLAS_EDICION.get(comando)
        if entrada is None:
            if comando in ("letra_grande", "letra_pequena"):
                return False, "Cambiar el tamaño de letra solo funciona en Word"
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
        if not WIN32_DISPONIBLE:
            return None
        try:
            return win32com.client.GetActiveObject("Word.Application")
        except Exception:
            if not crear:
                return None
        try:
            word_app = win32com.client.Dispatch("Word.Application")
            word_app.Visible = True
            return word_app
        except Exception:
            return None

    def _traer_word_al_frente(self, word_app):
        """Mejor esfuerzo: que Word quede al frente para seguir dictando."""
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

    def _carpeta_base_documentos(self, word_app) -> Path:
        try:
            ruta = word_app.Options.DefaultFilePath(WD_RUTA_DOCUMENTOS)
        except Exception:
            ruta = ""
        return Path(ruta) if ruta else Path.home() / "Documents"

    def _carpeta_para_guardar(self, word_app) -> Path:
        if CARPETA_GUARDADO:
            carpeta = Path(CARPETA_GUARDADO)
        else:
            carpeta = self._carpeta_base_documentos(word_app) / NOMBRE_CARPETA_APP
        carpeta.mkdir(parents=True, exist_ok=True)
        return carpeta

    def _listar_documentos(self, word_app):
        """Documentos de Word en la carpeta de la app y en Documentos."""
        carpetas = [self._carpeta_para_guardar(word_app)]
        if not CARPETA_GUARDADO:
            carpetas.append(self._carpeta_base_documentos(word_app))
        vistos, documentos = set(), []
        for carpeta in carpetas:
            if not carpeta.exists():
                continue
            for ruta in carpeta.iterdir():
                if (ruta.suffix.lower() in EXTENSIONES_WORD and not ruta.name.startswith("~$")
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

    def guardar_como(self, nombre: str):
        """Guarda el documento activo con el nombre dicho, en la carpeta
        configurada y con extensión .docx (el usuario no la menciona)."""
        word_app = self._obtener_word()
        if word_app is None:
            return False, "Word no está abierto"
        nombre_limpio = _nombre_de_archivo_seguro(nombre)
        if not nombre_limpio:
            return False, "No entendí el nombre del documento"
        try:
            if word_app.Documents.Count == 0:
                return False, "No hay ningún documento abierto"
            documento = word_app.ActiveDocument
            carpeta = self._carpeta_para_guardar(word_app)
            actual = os.path.normcase(str(documento.FullName))

            ruta = carpeta / f"{nombre_limpio}.docx"
            contador = 2
            # No pisar otro archivo: si el nombre existe, agrega (2), (3)...
            while ruta.exists() and os.path.normcase(str(ruta)) != actual:
                ruta = carpeta / f"{nombre_limpio} ({contador}).docx"
                contador += 1

            documento.SaveAs2(str(ruta), WD_FORMATO_DOCX)
            return True, f"Guardado como {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo guardar el documento: {e}"

    def abrir_documento(self, consulta: str):
        """Abre un documento guardado buscándolo por nombre (aproximado)."""
        word_app = self._obtener_word(crear=True)
        if word_app is None:
            return False, "No se pudo controlar Word"
        try:
            ruta = self._elegir_documento(consulta, self._listar_documentos(word_app))
            if ruta is None:
                return False, f"No encontré un documento parecido a {consulta}"
            word_app.Documents.Open(str(ruta))
            self._traer_word_al_frente(word_app)
            return True, f"Abriendo {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo abrir el documento: {e}"

    def abrir_ultimo_documento(self):
        """Abre el documento modificado más recientemente."""
        word_app = self._obtener_word(crear=True)
        if word_app is None:
            return False, "No se pudo controlar Word"
        try:
            documentos = self._listar_documentos(word_app)
            if not documentos:
                return False, "No hay documentos guardados"
            ruta = max(documentos, key=lambda d: d.stat().st_mtime)
            word_app.Documents.Open(str(ruta))
            self._traer_word_al_frente(word_app)
            return True, f"Abriendo {ruta.stem}"
        except Exception as e:
            return False, f"No se pudo abrir el documento: {e}"

    def buscar_google(self, consulta: str):
        if not consulta or not consulta.strip():
            return False, "No se especificó qué buscar"
        try:
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.3)
            pyautogui.typewrite(consulta, interval=0.02)
            pyautogui.press("enter")
            return True, f"Buscando en Google: {consulta}"
        except Exception as e:
            return False, f"Error al realizar la búsqueda: {e}"

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

    def leer_elemento_enfocado(self):
        if not UIAUTOMATION_DISPONIBLE:
            mensaje = "uiautomation no está instalado - no se puede leer la pantalla"
            self._hablar("La función de lectura no está disponible")
            return False, mensaje

        try:
            control = auto.GetFocusedControl()
            if control is None:
                self._hablar("No se pudo identificar el elemento actual")
                return False, "No se encontró elemento enfocado"

            texto = (control.Name or "").strip()
            if not texto:
                try:
                    texto = (control.GetWindowText() or "").strip()
                except Exception:
                    texto = ""

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