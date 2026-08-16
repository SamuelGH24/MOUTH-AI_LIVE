"""
Módulo de Ejecución de Acciones - Mouth AI Live
Recibe el resultado del intérprete de comandos y ejecuta la acción
correspondiente en el sistema operativo (Windows).

IMPORTANTE - MODO_PRUEBA:
Las acciones críticas/irreversibles (bloquear sesión, apagar el equipo)
están protegidas por MODO_PRUEBA = True por defecto. En ese modo, solo
se imprime en consola lo que haría, sin ejecutarlo. Cuando estés seguro
de que el reconocimiento es confiable, cambia MODO_PRUEBA a False.
"""

import ctypes
import subprocess
import webbrowser

# ---------------------------------------------------------------------------
MODO_PRUEBA = True  # ⚠️ Cambia a False solo cuando confíes en el sistema
# ---------------------------------------------------------------------------

# Mapeo de nombres lógicos de apps a comandos reales de Windows
APPS_DISPONIBLES = {
    "chrome": "start chrome",
    "notepad": "notepad.exe",
    "calc": "calc.exe",
}

# Códigos de tecla virtual de Windows para control de volumen (multimedia)
VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002


class EjecutorAcciones:
    """
    Ejecuta acciones del sistema operativo a partir de una intención
    ya interpretada (accion + parametro).
    """

    def __init__(self, modo_prueba: bool = MODO_PRUEBA):
        self.modo_prueba = modo_prueba

    # -------------------------------------------------------------------
    def ejecutar(self, accion: str, parametro):
        """
        Punto de entrada único. Devuelve (exito: bool, mensaje: str).
        """
        try:
            if accion == "abrir_app":
                return self._abrir_app(parametro)
            elif accion == "abrir_url":
                return self._abrir_url(parametro)
            elif accion == "control_volumen":
                return self._control_volumen(parametro)
            elif accion == "bloquear_sesion":
                return self._bloquear_sesion()
            elif accion == "apagar_equipo":
                return self._apagar_equipo()
            elif accion == "detener":
                return True, "DETENER_SISTEMA"  # señal especial para el orquestador
            else:
                return False, f"Acción no reconocida: {accion}"
        except Exception as e:
            return False, f"Error ejecutando '{accion}': {e}"

    # -------------------------------------------------------------------
    def _abrir_app(self, nombre_app: str):
        comando = APPS_DISPONIBLES.get(nombre_app)
        if not comando:
            return False, f"App no configurada: {nombre_app}"
        subprocess.Popen(comando, shell=True)
        return True, f"Abriendo aplicación: {nombre_app}"

    def _abrir_url(self, url: str):
        webbrowser.open(url)
        return True, f"Abriendo URL: {url}"

    def _control_volumen(self, direccion: str):
        codigo = {
            "subir": VK_VOLUME_UP,
            "bajar": VK_VOLUME_DOWN,
            "mute": VK_VOLUME_MUTE,
        }.get(direccion)
        if codigo is None:
            return False, f"Dirección de volumen inválida: {direccion}"

        ctypes.windll.user32.keybd_event(codigo, 0, KEYEVENTF_EXTENDEDKEY, 0)
        ctypes.windll.user32.keybd_event(
            codigo, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0
        )
        return True, f"Volumen: {direccion}"

    def _bloquear_sesion(self):
        if self.modo_prueba:
            return True, "[MODO PRUEBA] Se habría bloqueado la sesión (no ejecutado)"
        ctypes.windll.user32.LockWorkStation()
        return True, "Sesión bloqueada"

    def _apagar_equipo(self):
        if self.modo_prueba:
            return True, "[MODO PRUEBA] Se habría apagado el equipo (no ejecutado)"
        subprocess.run(["shutdown", "/s", "/t", "10"], shell=True)
        return True, "Apagando el equipo en 10 segundos (shutdown /a para cancelar)"


if __name__ == "__main__":
    # Prueba manual rápida desde consola
    ejecutor = EjecutorAcciones()
    print(ejecutor.ejecutar("abrir_url", "https://www.google.com"))
    print(ejecutor.ejecutar("control_volumen", "subir"))
    print(ejecutor.ejecutar("bloquear_sesion", None))