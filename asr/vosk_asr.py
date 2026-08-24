"""
Módulo de Captura de Audio + ASR (Reconocimiento Automático del Habla)
Mouth AI Live

Usa Vosk para reconocer voz en tiempo real desde el micrófono.
Requiere: pip install vosk pyaudio
"""

import json
import time

import pyaudio
from vosk import KaldiRecognizer, Model

RUTA_MODELO = r"C:\Users\X1504VA\Desktop\proyecto\Mouth-ai-live\models\vosk-model-es-0.42"
FRECUENCIA_MUESTREO = 16000
TAMANO_BLOQUE = 8000
MAX_ERRORES_SEGUIDOS = 5  # si el micrófono falla más de esto seguido, se avisa y se detiene


class CapturaASR:
    def __init__(self, ruta_modelo: str = RUTA_MODELO):
        print(f"[ASR] Cargando modelo desde: {ruta_modelo}")
        self.modelo = Model(ruta_modelo)
        self.reconocedor = KaldiRecognizer(self.modelo, FRECUENCIA_MUESTREO)
        self.audio_interface = None
        self.stream = None
        print("[ASR] Modelo cargado correctamente.")

    def iniciar_microfono(self):
        self.audio_interface = pyaudio.PyAudio()
        self.stream = self.audio_interface.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=FRECUENCIA_MUESTREO,
            input=True,
            frames_per_buffer=TAMANO_BLOQUE,
        )
        self.stream.start_stream()
        print("[ASR] Micrófono activo.")

    def detener_microfono(self):
        if self.stream is not None:
            try:
                self.stream.stop_stream()
                self.stream.close()
            except Exception as e:
                print(f"[ASR] Aviso al cerrar el stream: {e}")
            self.stream = None
        if self.audio_interface is not None:
            try:
                self.audio_interface.terminate()
            except Exception as e:
                print(f"[ASR] Aviso al cerrar la interfaz de audio: {e}")
            self.audio_interface = None
        print("[ASR] Micrófono detenido.")

    def _reiniciar_microfono(self):
        """Intenta cerrar y volver a abrir el micrófono tras un fallo."""
        print("[ASR] Intentando reiniciar el micrófono...")
        self.detener_microfono()
        time.sleep(1.0)
        self.iniciar_microfono()

    def escuchar_y_transcribir(self, callback_texto, evento_detener=None,
                                preprocesador=None, callback_error=None):
        """
        Bucle principal de escucha. Es tolerante a fallos: si leer del
        micrófono falla (se desconectó, otro programa lo tomó, etc.), no
        se cae todo el programa - se reintenta. Solo se detiene si el
        fallo persiste MAX_ERRORES_SEGUIDOS veces seguidas.

        callback_error: función opcional que recibe un mensaje de texto
        cuando ocurre un problema recuperable, para poder informarlo en
        la interfaz sin interrumpir la escucha.
        """
        errores_seguidos = 0

        while True:
            if evento_detener is not None and evento_detener.is_set():
                break

            try:
                datos = self.stream.read(TAMANO_BLOQUE, exception_on_overflow=False)
                errores_seguidos = 0  # se resetea el contador tras una lectura exitosa

                if preprocesador is not None:
                    datos, tiene_voz, energia = preprocesador.procesar_bloque(datos)
                    if not tiene_voz:
                        continue

                if self.reconocedor.AcceptWaveform(datos):
                    resultado = json.loads(self.reconocedor.Result())
                    texto = resultado.get("text", "").strip()
                    if texto:
                        callback_texto(texto)
                else:
                    parcial = json.loads(self.reconocedor.PartialResult())
                    texto_parcial = parcial.get("partial", "").strip()
                    if texto_parcial:
                        print(f"   (escuchando...) {texto_parcial}", end="\r")

            except KeyboardInterrupt:
                print("\n[ASR] Detenido por el usuario (Ctrl+C).")
                break

            except Exception as e:
                errores_seguidos += 1
                mensaje = f"Problema temporal con el micrófono ({errores_seguidos}/{MAX_ERRORES_SEGUIDOS}): {e}"
                print(f"[ASR] {mensaje}")
                if callback_error is not None:
                    callback_error(mensaje)

                if errores_seguidos >= MAX_ERRORES_SEGUIDOS:
                    mensaje_final = "El micrófono falló repetidamente. Deteniendo la escucha."
                    print(f"[ASR] {mensaje_final}")
                    if callback_error is not None:
                        callback_error(mensaje_final)
                    break

                try:
                    self._reiniciar_microfono()
                except Exception as e2:
                    print(f"[ASR] No se pudo reiniciar el micrófono: {e2}")
                    time.sleep(1.0)

        self.detener_microfono()


def _prueba_manual(texto_reconocido: str):
    print(f"\n✅ TEXTO RECONOCIDO: '{texto_reconocido}'\n")


if __name__ == "__main__":
    asr = CapturaASR()
    asr.iniciar_microfono()
    asr.escuchar_y_transcribir(callback_texto=_prueba_manual)