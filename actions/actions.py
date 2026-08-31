import ctypes
import json
import subprocess
import sys
import webbrowser
from pathlib import Path

MODO_PRUEBA = True

APPS_DISPONIBLES = {
    "chrome": "start chrome",
    "notepad": "notepad.exe",
    "calc": "calc.exe",
    "word": "start winword",
}

VK_VOLUME_MUTE = 0xAD
VK_VOLUME_DOWN = 0xAE
VK_VOLUME_UP = 0xAF
KEYEVENTF_EXTENDEDKEY = 0x0001
KEYEVENTF_KEYUP = 0x0002

RUTA_CONFIG_POR_DEFECTO = Path(__file__).parent.parent / "config" / "comandos.json"


class EjecutorAcciones:
    def __init__(self, modo_prueba: bool = MODO_PRUEBA, ruta_config: str = None):
        self.modo_prueba = modo_prueba
        self.ruta_config = Path(ruta_config) if ruta_config else RUTA_CONFIG_POR_DEFECTO
        self._navegador_asistido = None

    def _obtener_navegador_asistido(self):
        if self._navegador_asistido is None:
            ruta_navigation = str(Path(__file__).parent.parent / "navigation")
            if ruta_navigation not in sys.path:
                sys.path.insert(0, ruta_navigation)
            from navigation import NavegadorAsistido
            self._navegador_asistido = NavegadorAsistido()
        return self._navegador_asistido

    def ejecutar(self, accion: str, parametro):
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
                return True, "PAUSAR_SISTEMA"
            elif accion == "reanudar":
                return True, "REANUDAR_SISTEMA"
            elif accion == "buscar_google":
                return self._buscar_google(parametro)
            elif accion == "navegar":
                return self._navegar(parametro)
            elif accion == "leer_pantalla":
                return self._leer_pantalla()
            elif accion == "escribir_texto":
                return self._escribir_texto(parametro)
            elif accion == "listar_comandos":
                return self._listar_comandos()
            else:
                return False, f"Acción no reconocida: {accion}"
        except Exception as e:
            return False, f"Error ejecutando '{accion}': {e}"

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
        codigo = {"subir": VK_VOLUME_UP, "bajar": VK_VOLUME_DOWN, "mute": VK_VOLUME_MUTE}.get(direccion)
        if codigo is None:
            return False, f"Dirección de volumen inválida: {direccion}"
        ctypes.windll.user32.keybd_event(codigo, 0, KEYEVENTF_EXTENDEDKEY, 0)
        ctypes.windll.user32.keybd_event(codigo, 0, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0)
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
        return True, "Apagando el equipo en 10 segundos"

    def _navegar(self, comando: str):
        try:
            navegador = self._obtener_navegador_asistido()
        except Exception as e:
            return False, f"No se pudo cargar el módulo de navegación. Detalle: {e}"
        return navegador.navegar(comando)

    def _buscar_google(self, consulta: str):
        try:
            navegador = self._obtener_navegador_asistido()
        except Exception as e:
            return False, f"No se pudo cargar el módulo de navegación. Detalle: {e}"
        return navegador.buscar_google(consulta)

    def _leer_pantalla(self):
        try:
            navegador = self._obtener_navegador_asistido()
        except Exception as e:
            return False, f"No se pudo cargar el módulo de navegación. Detalle: {e}"
        return navegador.leer_elemento_enfocado()

    def _escribir_texto(self, texto: str):
        try:
            navegador = self._obtener_navegador_asistido()
        except Exception as e:
            return False, f"No se pudo cargar el módulo de navegación. ¿Instalaste pyperclip? Detalle: {e}"
        return navegador.escribir_texto(texto)

    def _listar_comandos(self):
        try:
            navegador = self._obtener_navegador_asistido()
        except Exception as e:
            return False, f"No se pudo cargar el módulo de voz. Detalle: {e}"

        try:
            with open(self.ruta_config, "r", encoding="utf-8") as f:
                data = json.load(f)
            texto_ayuda = data.get("ayuda_hablada", "No hay ayuda configurada todavía.")
        except Exception as e:
            texto_ayuda = "No se pudo cargar la lista de comandos."
            navegador.hablar(texto_ayuda)
            return False, f"Error leyendo comandos.json: {e}"

        navegador.hablar(texto_ayuda)
        return True, "Lista de comandos narrada por voz"