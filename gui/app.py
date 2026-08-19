"""
Interfaz Gráfica - Mouth AI Live

Permite al usuario iniciar, detener y monitorear el sistema de reconocimiento
de voz sin usar la terminal. El reconocimiento corre en un hilo separado
para no congelar la ventana mientras escucha.

Ejecutar desde la raíz del proyecto:
    python gui/app.py
"""

import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, scrolledtext

RAIZ = Path(__file__).parent.parent
sys.path.insert(0, str(RAIZ / "asr"))
sys.path.insert(0, str(RAIZ / "interpreter"))
sys.path.insert(0, str(RAIZ / "actions"))
sys.path.insert(0, str(RAIZ / "metrics"))

from vosk_asr import CapturaASR  # noqa: E402
from interpreter import InterpretadorComandos  # noqa: E402
from actions import EjecutorAcciones  # noqa: E402
from logger import RegistradorMetricas  # noqa: E402

RUTA_CONFIG_COMANDOS = str(RAIZ / "config" / "comandos.json")


class MouthAILiveGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Mouth AI Live")
        self.root.geometry("640x480")
        self.root.resizable(False, False)

        self.cola_mensajes = queue.Queue()
        self.evento_detener = threading.Event()
        self.hilo_escucha = None

        self.asr = None
        self.interprete = None
        self.ejecutor = None
        self.registrador = None
        self.modelo_listo = False

        self._construir_interfaz()
        self._cargar_modelo_en_segundo_plano()
        self._revisar_cola_mensajes()

    def _construir_interfaz(self):
        marco_superior = tk.Frame(self.root, pady=10)
        marco_superior.pack(fill=tk.X)

        tk.Label(marco_superior, text="Mouth AI Live", font=("Segoe UI", 16, "bold")).pack()

        self.etiqueta_estado = tk.Label(
            marco_superior, text="Cargando modelo de voz...",
            font=("Segoe UI", 11, "bold"), fg="#CC7A00",
        )
        self.etiqueta_estado.pack(pady=(5, 0))

        marco_botones = tk.Frame(self.root, pady=10)
        marco_botones.pack()

        self.boton_iniciar = tk.Button(
            marco_botones, text="▶ Iniciar", width=15, height=2,
            command=self.iniciar, state=tk.DISABLED, bg="#DFF5DF",
        )
        self.boton_iniciar.grid(row=0, column=0, padx=8)

        self.boton_detener = tk.Button(
            marco_botones, text="■ Detener", width=15, height=2,
            command=self.detener, state=tk.DISABLED, bg="#F5DFDF",
        )
        self.boton_detener.grid(row=0, column=1, padx=8)

        tk.Label(self.root, text="Registro de actividad:", font=("Segoe UI", 10, "bold")).pack(
            anchor="w", padx=15
        )

        self.area_log = scrolledtext.ScrolledText(
            self.root, width=76, height=18, state=tk.DISABLED,
            font=("Consolas", 9), bg="#1e1e1e", fg="#e0e0e0",
        )
        self.area_log.pack(padx=15, pady=(5, 10))

        marco_inferior = tk.Frame(self.root)
        marco_inferior.pack(pady=(0, 10))

        tk.Button(marco_inferior, text="Limpiar registro", command=self._limpiar_log).grid(
            row=0, column=0, padx=5
        )
        tk.Button(marco_inferior, text="📊 Ver resumen de métricas", command=self._mostrar_resumen).grid(
            row=0, column=1, padx=5
        )

        self.root.protocol("WM_DELETE_WINDOW", self._al_cerrar)

    def _cargar_modelo_en_segundo_plano(self):
        def tarea():
            try:
                self.asr = CapturaASR()
                self.interprete = InterpretadorComandos(RUTA_CONFIG_COMANDOS)
                self.ejecutor = EjecutorAcciones()
                self.registrador = RegistradorMetricas()
                self.modelo_listo = True
                self.cola_mensajes.put(("estado_listo", None))
            except Exception as e:
                self.cola_mensajes.put(("error", f"No se pudo cargar el modelo: {e}"))

        threading.Thread(target=tarea, daemon=True).start()

    def iniciar(self):
        if not self.modelo_listo:
            return

        self.evento_detener.clear()
        self.boton_iniciar.config(state=tk.DISABLED)
        self.boton_detener.config(state=tk.NORMAL)
        self.etiqueta_estado.config(text="🎤 Escuchando...", fg="#1A8F1A")

        self.hilo_escucha = threading.Thread(target=self._bucle_escucha, daemon=True)
        self.hilo_escucha.start()

    def detener(self):
        self.evento_detener.set()
        self.boton_detener.config(state=tk.DISABLED)
        self.etiqueta_estado.config(text="Deteniendo...", fg="#CC7A00")

    def _bucle_escucha(self):
        try:
            self.asr.iniciar_microfono()

            def callback(texto):
                self.registrador.iniciar_medicion()
                self.cola_mensajes.put(("log", f"🎤 Reconocido: '{texto}'"))
                resultado = self.interprete.interpretar(texto)

                if not resultado.reconocido:
                    self.cola_mensajes.put(
                        ("log", f"   ⚠️ Sin comando válido (confianza {resultado.confianza:.2f})")
                    )
                    self.registrador.finalizar_y_registrar(texto, resultado, False, "")
                    return

                self.cola_mensajes.put(
                    ("log", f"   ✅ Intención: {resultado.intencion} "
                            f"(confianza {resultado.confianza:.2f})")
                )
                exito, mensaje = self.ejecutor.ejecutar(resultado.accion, resultado.parametro)
                self.registrador.finalizar_y_registrar(texto, resultado, exito, mensaje)

                if mensaje == "DETENER_SISTEMA":
                    self.cola_mensajes.put(("log", "   🛑 Comando de detener recibido por voz."))
                    self.evento_detener.set()
                    return

                simbolo = "✔️" if exito else "❌"
                self.cola_mensajes.put(("log", f"   {simbolo} {mensaje}"))

            self.asr.escuchar_y_transcribir(
                callback_texto=callback,
                evento_detener=self.evento_detener,
            )
        except Exception as e:
            self.cola_mensajes.put(("error", f"Error durante la escucha: {e}"))
        finally:
            self.cola_mensajes.put(("detenido", None))

    def _revisar_cola_mensajes(self):
        try:
            while True:
                tipo, contenido = self.cola_mensajes.get_nowait()

                if tipo == "log":
                    self._escribir_log(contenido)
                elif tipo == "estado_listo":
                    self.etiqueta_estado.config(text="Listo. Presiona Iniciar.", fg="#1A6FCC")
                    self.boton_iniciar.config(state=tk.NORMAL)
                    self._escribir_log("[Sistema] Modelo cargado. Listo para iniciar.")
                elif tipo == "error":
                    self._escribir_log(f"❌ {contenido}")
                    messagebox.showerror("Mouth AI Live - Error", contenido)
                elif tipo == "detenido":
                    self.boton_iniciar.config(state=tk.NORMAL)
                    self.boton_detener.config(state=tk.DISABLED)
                    self.etiqueta_estado.config(text="Detenido. Presiona Iniciar.", fg="#CC7A00")
                    self._escribir_log("[Sistema] Escucha detenida.")
        except queue.Empty:
            pass

        self.root.after(150, self._revisar_cola_mensajes)

    def _escribir_log(self, texto: str):
        self.area_log.config(state=tk.NORMAL)
        self.area_log.insert(tk.END, texto + "\n")
        self.area_log.see(tk.END)
        self.area_log.config(state=tk.DISABLED)

    def _limpiar_log(self):
        self.area_log.config(state=tk.NORMAL)
        self.area_log.delete("1.0", tk.END)
        self.area_log.config(state=tk.DISABLED)

    def _mostrar_resumen(self):
        if self.registrador is None:
            messagebox.showinfo("Mouth AI Live", "El sistema todavía no ha cargado. Espera un momento.")
            return

        resumen = self.registrador.generar_resumen()
        if resumen is None:
            messagebox.showinfo("Mouth AI Live", "Todavía no hay pruebas registradas.")
            return

        texto = (
            f"Total de intentos: {resumen['total_intentos']}\n"
            f"Reconocidos: {resumen['reconocidos']} | No reconocidos: {resumen['no_reconocidos']}\n"
            f"Tasa de reconocimiento exitoso: {resumen['tasa_reconocimiento_pct']}%\n"
            f"Tasa de error: {resumen['tasa_error_pct']}%\n"
            f"Tasa de éxito en ejecución: {resumen['tasa_exito_ejecucion_pct']}%\n"
        )
        if resumen["tiempo_respuesta_promedio_ms"] is not None:
            texto += (
                f"Tiempo de respuesta promedio: {resumen['tiempo_respuesta_promedio_ms']} ms\n"
                f"Tiempo de respuesta mediana: {resumen['tiempo_respuesta_mediana_ms']} ms\n"
            )
        texto += f"\nArchivo CSV: {self.registrador.ruta_csv}"

        messagebox.showinfo("Resumen de métricas", texto)

    def _al_cerrar(self):
        self.evento_detener.set()
        self.root.destroy()


def main():
    root = tk.Tk()
    MouthAILiveGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()