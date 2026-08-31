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

    def escribir_texto(self, texto: str):
        if not texto or not texto.strip():
            return False, "No se especificó qué escribir"
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