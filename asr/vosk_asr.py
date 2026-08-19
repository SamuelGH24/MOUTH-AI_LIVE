"""
Módulo de Captura de Audio + ASR (Reconocimiento Automático del Habla)
Mouth AI Live
"""

import json

import pyaudio
from vosk import KaldiRecognizer, Model

RUTA_MODELO = r"C:\Users\X1504VA\Desktop\proyecto\Mouth-ai-live\models\vosk-model-es-0.42"
FRECUENCIA_MUESTREO = 16000
TAMANO_BLOQUE = 8000


class CapturaASR:
    def __init__(self, ruta_modelo: str = RUTA_MODELO):
        print(f"[ASR] Cargando modelo desde: {ruta_modelo}")
        print("[ASR] Esto puede tardar varios segundos con el modelo grande...")
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
            self.stream.stop_stream()
            self.stream.close()
            self.stream = None
        if self.audio_interface is not None:
            self.audio_interface.terminate()
            self.audio_interface = None
        print("[ASR] Micrófono detenido.")

    def escuchar_y_transcribir(self, callback_texto, evento_detener=None, preprocesador=None):
        try:
            while True:
                if evento_detener is not None and evento_detener.is_set():
                    break

                datos = self.stream.read(TAMANO_BLOQUE, exception_on_overflow=False)

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
        finally:
            self.detener_microfono()


def _prueba_manual(texto_reconocido: str):
    print(f"\n✅ TEXTO RECONOCIDO: '{texto_reconocido}'\n")


if __name__ == "__main__":
    asr = CapturaASR()
    asr.iniciar_microfono()
    asr.escuchar_y_transcribir(callback_texto=_prueba_manual)