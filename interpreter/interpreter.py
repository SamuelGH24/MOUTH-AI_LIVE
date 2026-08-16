"""
Módulo de Interpretación de Comandos - Mouth AI Live
Convierte texto transcrito (proveniente del módulo ASR) en una intención
y acción ejecutable, usando coincidencia de patrones (normalización +
similitud de texto), sin necesidad de un modelo de ML entrenado.
"""

import json
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional


@dataclass
class ResultadoInterpretacion:
    reconocido: bool
    intencion: Optional[str] = None
    accion: Optional[str] = None
    parametro: Optional[str] = None
    frase_detectada: Optional[str] = None
    confianza: float = 0.0


class InterpretadorComandos:
    """
    Carga comandos desde un archivo JSON de configuración y determina,
    a partir de un texto de entrada, cuál es la intención más probable.
    """

    UMBRAL_CONFIANZA = 0.75  # ajustable: qué tan estricta es la coincidencia

    def __init__(self, ruta_config: str):
        self.ruta_config = Path(ruta_config)
        self.comandos = []
        self._cargar_configuracion()

    def _cargar_configuracion(self):
        if not self.ruta_config.exists():
            raise FileNotFoundError(f"No se encontró el archivo de configuración: {self.ruta_config}")
        with open(self.ruta_config, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.comandos = data.get("comandos", [])
        if not self.comandos:
            raise ValueError("El archivo de configuración no contiene comandos válidos.")

    @staticmethod
    def _normalizar(texto: str) -> str:
        """Quita tildes, pasa a minúsculas y limpia espacios/puntuación."""
        texto = texto.lower().strip()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        texto = re.sub(r"[^\w\s]", "", texto)
        texto = re.sub(r"\s+", " ", texto)
        return texto

    @staticmethod
    def _similitud(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio()

    def interpretar(self, texto_reconocido: str) -> ResultadoInterpretacion:
        """
        Recibe el texto transcrito por el ASR y devuelve la mejor coincidencia
        encontrada entre las frases configuradas, si supera el umbral de confianza.
        """
        if not texto_reconocido or not texto_reconocido.strip():
            return ResultadoInterpretacion(reconocido=False)

        texto_norm = self._normalizar(texto_reconocido)

        mejor_resultado = ResultadoInterpretacion(reconocido=False)
        mejor_score = 0.0

        for comando in self.comandos:
            for frase in comando["frases"]:
                frase_norm = self._normalizar(frase)

                # 1. Coincidencia exacta o de subcadena (prioridad máxima)
                if frase_norm == texto_norm or frase_norm in texto_norm:
                    return ResultadoInterpretacion(
                        reconocido=True,
                        intencion=comando["intencion"],
                        accion=comando["accion"],
                        parametro=comando.get("parametro"),
                        frase_detectada=frase,
                        confianza=1.0,
                    )

                # 2. Coincidencia por similitud (tolera errores de transcripción)
                score = self._similitud(frase_norm, texto_norm)
                if score > mejor_score:
                    mejor_score = score
                    mejor_resultado = ResultadoInterpretacion(
                        reconocido=score >= self.UMBRAL_CONFIANZA,
                        intencion=comando["intencion"] if score >= self.UMBRAL_CONFIANZA else None,
                        accion=comando["accion"] if score >= self.UMBRAL_CONFIANZA else None,
                        parametro=comando.get("parametro") if score >= self.UMBRAL_CONFIANZA else None,
                        frase_detectada=frase if score >= self.UMBRAL_CONFIANZA else None,
                        confianza=score,
                    )

        return mejor_resultado

    def listar_comandos_disponibles(self):
        """Útil para depuración o para mostrar en la interfaz gráfica."""
        return [
            {"intencion": c["intencion"], "frases": c["frases"]}
            for c in self.comandos
        ]