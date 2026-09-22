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

import time

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
        word_app.Selection.TypeText(texto)

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