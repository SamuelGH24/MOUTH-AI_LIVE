"""
Modulo de Navegacion Asistida, Dictado y Control de Mouse - Mouth AI Live
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
    "negrita": ["ctrl", "b"],
    "cursiva": ["ctrl", "i"],
    "subrayado": ["ctrl", "u"],
    "letra_grande": ["ctrl", "shift", "."],
    "letra_pequena": ["ctrl", "shift", ","],
}

# Distancia en pixeles que se mueve el cursor por cada comando de voz.
PASO_MOUSE_PIXELES = 40


class NavegadorAsistido:
    def __init__(self):
        pass

    def escribir_texto(self, texto: str):
        if not texto or not texto.strip():
            return False, "No se especifico que escribir"
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
            return False, "No se especifico que buscar"
        try:
            pyautogui.hotkey("ctrl", "l")
            time.sleep(0.3)
            pyautogui.typewrite(consulta, interval=0.02)
            pyautogui.press("enter")
            return True, f"Buscando en Google: {consulta}"
        except Exception as e:
            return False, f"Error al realizar la busqueda: {e}"

    def navegar(self, comando: str):
        tecla = TECLAS_NAVEGACION.get(comando)
        if tecla is None:
            return False, f"Comando de navegacion no reconocido: {comando}"
        try:
            if isinstance(tecla, list):
                pyautogui.hotkey(*tecla)
            else:
                pyautogui.press(tecla)
            return True, f"Navegacion ejecutada: {comando}"
        except Exception as e:
            return False, f"Error al navegar ({comando}): {e}"

    def leer_elemento_enfocado(self):
        if not UIAUTOMATION_DISPONIBLE:
            mensaje = "uiautomation no esta instalado - no se puede leer la pantalla"
            self._hablar("La funcion de lectura no esta disponible")
            return False, mensaje
        try:
            control = auto.GetFocusedControl()
            if control is None:
                self._hablar("No se pudo identificar el elemento actual")
                return False, "No se encontro elemento enfocado"
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
            self._hablar("Ocurrio un error al intentar leer")
            return False, mensaje

    def mover_mouse(self, direccion: str):
        """
        Mueve el cursor una distancia fija (PASO_MOUSE_PIXELES) en la
        direccion indicada. Repetir el comando sigue moviendolo mas lejos.
        """
        deltas = {
            "arriba": (0, -PASO_MOUSE_PIXELES),
            "abajo": (0, PASO_MOUSE_PIXELES),
            "izquierda": (-PASO_MOUSE_PIXELES, 0),
            "derecha": (PASO_MOUSE_PIXELES, 0),
        }
        delta = deltas.get(direccion)
        if delta is None:
            return False, f"Direccion de mouse no reconocida: {direccion}"
        try:
            pyautogui.moveRel(delta[0], delta[1], duration=0.1)
            return True, f"Cursor movido: {direccion}"
        except Exception as e:
            return False, f"Error moviendo el cursor: {e}"

    def clic_mouse(self, tipo: str):
        try:
            if tipo == "izquierdo":
                pyautogui.click()
                return True, "Clic izquierdo realizado"
            elif tipo == "derecho":
                pyautogui.rightClick()
                return True, "Clic derecho realizado"
            elif tipo == "doble":
                pyautogui.doubleClick()
                return True, "Doble clic realizado"
            else:
                return False, f"Tipo de clic no reconocido: {tipo}"
        except Exception as e:
            return False, f"Error al hacer clic: {e}"

    def hablar(self, texto: str):
        """Metodo publico para que otros modulos (ej. actions.py) narren algo."""
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
    print(nav.mover_mouse("derecha"))
    print(nav.clic_mouse("izquierdo"))