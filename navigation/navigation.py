"""
Módulo de Navegación Asistida - Mouth AI Live

Dos responsabilidades:
1. Simular teclas de navegación (arriba/abajo/Tab/Enter/atrás) para moverse
   por el navegador o cualquier ventana, igual que lo haría una persona con
   teclado.
2. Leer en voz alta (texto a voz) el elemento que está enfocado actualmente
   en pantalla, usando UI Automation de Windows - la misma tecnología base
   que usan lectores de pantalla reales (NVDA, JAWS).

LIMITACIÓN HONESTA: la lectura depende de que la página/aplicación exponga
correctamente la información de accesibilidad. Funciona bien con enlaces,
botones y texto normal bien etiquetado. Puede fallar o no leer nada en
elementos sin descripción (ej. imágenes sin texto alternativo). Esto no es
un error del código, es una limitación inherente de cómo está construida
la página que se está usando.

Requiere (instalar en el entorno de Windows):
    pip install pyautogui pyttsx3 uiautomation
"""

import pyautogui
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
}


class NavegadorAsistido:
    def __init__(self):
        # NOTA: no se guarda un único motor pyttsx3 reutilizable a propósito.
        # Es un bug conocido de esta librería: reutilizar el mismo objeto
        # engine en llamadas sucesivas hace que, después de una o dos veces,
        # se quede "mudo" sin lanzar ningún error. La solución confiable es
        # crear un motor nuevo cada vez que se necesita hablar.
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