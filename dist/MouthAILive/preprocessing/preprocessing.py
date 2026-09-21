"""
Módulo de Preprocesamiento de Señal - Mouth AI Live
Se ubica entre la Captura de Audio y el ASR. Procesa cada bloque de audio
crudo antes de enviarlo al reconocedor de voz, para:
    1. Normalizar el volumen (independiente de qué tan fuerte/bajo hable el usuario).
    2. Detectar si el bloque contiene voz o es silencio/ruido de fondo (VAD simple
       basado en energía RMS), evitando procesar bloques vacíos innecesariamente.

Trabaja sobre audio PCM de 16 bits, mono, tal como lo entrega PyAudio.
"""

import numpy as np

# ---------------------------------------------------------------------------
UMBRAL_ENERGIA_RMS = 300  # ajustar según nivel de ruido ambiente real
NIVEL_OBJETIVO_NORMALIZACION = 8000  # amplitud objetivo tras normalizar (de 32768 máx)
# ---------------------------------------------------------------------------


class PreprocesadorAudio:
    """
    Recibe bloques de audio crudo (bytes PCM int16) y devuelve:
        - el audio normalizado (bytes, listo para el ASR)
        - un booleano indicando si se detectó actividad de voz
    """

    def __init__(self, umbral_rms: int = UMBRAL_ENERGIA_RMS):
        self.umbral_rms = umbral_rms

    @staticmethod
    def _bytes_a_array(datos_pcm: bytes) -> np.ndarray:
        return np.frombuffer(datos_pcm, dtype=np.int16)

    @staticmethod
    def _array_a_bytes(array: np.ndarray) -> bytes:
        return array.astype(np.int16).tobytes()

    def calcular_energia_rms(self, datos_pcm: bytes) -> float:
        """Calcula la energía RMS de un bloque de audio (indicador de volumen)."""
        array = self._bytes_a_array(datos_pcm)
        if array.size == 0:
            return 0.0
        # Se castea a float64 para evitar overflow al elevar al cuadrado
        return float(np.sqrt(np.mean(array.astype(np.float64) ** 2)))

    def hay_actividad_de_voz(self, datos_pcm: bytes) -> bool:
        """Determina si el bloque de audio probablemente contiene voz (vs silencio/ruido)."""
        return self.calcular_energia_rms(datos_pcm) >= self.umbral_rms

    def normalizar(self, datos_pcm: bytes) -> bytes:
        """
        Escala la amplitud de la señal para que el pico máximo se acerque
        a NIVEL_OBJETIVO_NORMALIZACION, sin generar recorte (clipping).
        """
        array = self._bytes_a_array(datos_pcm)
        if array.size == 0:
            return datos_pcm

        pico_actual = np.max(np.abs(array))
        if pico_actual == 0:
            return datos_pcm  # silencio total, nada que normalizar

        factor = NIVEL_OBJETIVO_NORMALIZACION / pico_actual
        # Evitamos amplificar demasiado ruido de fondo si la señal ya es muy débil
        factor = min(factor, 10.0)

        array_normalizado = array.astype(np.float64) * factor
        array_normalizado = np.clip(array_normalizado, -32768, 32767)
        return self._array_a_bytes(array_normalizado)

    def procesar_bloque(self, datos_pcm: bytes):
        """
        Punto de entrada principal usado por el orquestador.
        Devuelve (audio_procesado: bytes, tiene_voz: bool, energia_rms: float)
        """
        energia = self.calcular_energia_rms(datos_pcm)
        tiene_voz = energia >= self.umbral_rms
        audio_procesado = self.normalizar(datos_pcm) if tiene_voz else datos_pcm
        return audio_procesado, tiene_voz, energia


if __name__ == "__main__":
    # ------------------------------------------------------------------
    # Prueba con datos sintéticos (sin necesidad de micrófono real)
    # ------------------------------------------------------------------
    preprocesador = PreprocesadorAudio()

    # Caso 1: silencio total
    silencio = np.zeros(1600, dtype=np.int16).tobytes()
    _, voz, energia = preprocesador.procesar_bloque(silencio)
    print(f"[Silencio total]      energía={energia:.1f}  voz_detectada={voz}")

    # Caso 2: ruido de fondo bajo (simulado como ruido aleatorio pequeño)
    ruido_bajo = (np.random.randint(-50, 50, 1600)).astype(np.int16).tobytes()
    _, voz, energia = preprocesador.procesar_bloque(ruido_bajo)
    print(f"[Ruido de fondo bajo]  energía={energia:.1f}  voz_detectada={voz}")

    # Caso 3: voz simulada (señal fuerte tipo onda)
    t = np.linspace(0, 1, 1600)
    voz_simulada = (np.sin(2 * np.pi * 200 * t) * 5000).astype(np.int16).tobytes()
    audio_proc, voz, energia = preprocesador.procesar_bloque(voz_simulada)
    array_final = np.frombuffer(audio_proc, dtype=np.int16)
    print(f"[Voz simulada]         energía={energia:.1f}  voz_detectada={voz}  "
          f"pico_original=5000  pico_normalizado={np.max(np.abs(array_final))}")

    # Caso 4: voz simulada muy débil (habla bajito)
    voz_debil = (np.sin(2 * np.pi * 200 * t) * 500).astype(np.int16).tobytes()
    audio_proc, voz, energia = preprocesador.procesar_bloque(voz_debil)
    array_final = np.frombuffer(audio_proc, dtype=np.int16)
    print(f"[Voz débil]            energía={energia:.1f}  voz_detectada={voz}  "
          f"pico_original=500  pico_normalizado={np.max(np.abs(array_final))}")