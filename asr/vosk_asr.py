"""
Módulo de Captura de Audio + ASR (Reconocimiento Automático del Habla)
Mouth AI Live

Usa Vosk para reconocer voz en tiempo real desde el micrófono.
Este módulo se ejecuta en Windows y requiere:
    pip install vosk pyaudio
    y el modelo descargado en la carpeta configurada en RUTA_MODELO.
"""

import json
import queue
import sys

import pyaudio
from vosk import KaldiRecognizer, Model

# ---------------------------------------------------------------------------
# CONFIGURACIÓN - ajusta esta ruta según donde descomprimiste el modelo
# ---------------------------------------------------------------------------
RUTA_MODELO = r"C:\Users\X1504VA\Desktop\proyecto\Mouth-ai-live\models\vosk-model-es-0.42"
FRECUENCIA_MUESTREO = 16000  # Hz, requerido por los modelos de Vosk
TAMANO_BLOQUE = 8000         # tamaño del bloque de audio leído por ciclo


class CapturaASR:
    """
    Encapsula la captura de audio del micrófono y el reconocimiento
    de voz en streaming usando Vosk.
    """

    def __init__(self, ruta_modelo: str = RUTA_MODELO):
        print(f"[ASR] Cargando modelo desde: {ruta_modelo}")
        print("[ASR] Esto puede tardar varios segundos con el modelo grande...")
        self.modelo = Model(ruta_modelo)
        self.reconocedor = KaldiRecognizer(self.modelo, FRECUENCIA_MUESTREO)
        self.audio_interface = None
        self.stream = None
        print("[ASR] Modelo cargado correctamente.")

    def iniciar_microfono(self):
        """Abre el flujo de audio desde el micrófono predeterminado."""
        self.audio_interface = pyaudio.PyAudio()
        self.stream = self.audio_interface.open(
            format=pyaudio.paInt16,
            channels=1,
            rate=FRECUENCIA_MUESTREO,
            input=True,
            frames_per_buffer=TAMANO_BLOQUE,
        )
        self.stream.start_stream()
        print("[ASR] Micrófono activo. Habla ahora (Ctrl+C para detener)...")

    def detener_microfono(self):
        if self.stream is not None:
            self.stream.stop_stream()
            self.stream.close()
        if self.audio_interface is not None:
            self.audio_interface.terminate()
        print("[ASR] Micrófono detenido.")

    def escuchar_y_transcribir(self, callback_texto, preprocesador=None):
        """
        Bucle principal: lee audio del micrófono, opcionalmente lo preprocesa
        (normalización + detección de voz), y llama a callback_texto(texto)
        cada vez que se detecta una frase completa reconocida.
        """
        try:
            while True:
                datos = self.stream.read(TAMANO_BLOQUE, exception_on_overflow=False)

                if preprocesador is not None:
                    datos, tiene_voz, energia = preprocesador.procesar_bloque(datos)
                    if not tiene_voz:
                        continue  # se descarta el bloque de silencio/ruido

                if self.reconocedor.AcceptWaveform(datos):
                    resultado = json.loads(self.reconocedor.Result())
                    texto = resultado.get("text", "").strip()
                    if texto:
                        callback_texto(texto)
                else:
                    # Resultado parcial (mientras la persona sigue hablando)
                    parcial = json.loads(self.reconocedor.PartialResult())
                    texto_parcial = parcial.get("partial", "").strip()
                    if texto_parcial:
                        print(f"   (escuchando...) {texto_parcial}", end="\r")

        except KeyboardInterrupt:
            print("\n[ASR] Detenido por el usuario.")
        finally:
            self.detener_microfono()


def _prueba_manual(texto_reconocido: str):
    """Función de prueba: solo imprime lo que se reconoció."""
    print(f"\n✅ TEXTO RECONOCIDO: '{texto_reconocido}'\n")


if __name__ == "__main__":
    asr = CapturaASR()
    asr.iniciar_microfono()
    asr.escuchar_y_transcribir(callback_texto=_prueba_manual)