"""
Módulo de Interpretación de Comandos - Mouth AI Live (v3)
Soporta comandos FIJOS (frases exactas/similares) y comandos DINÁMICOS
(texto variable extraído por expresiones regulares).
"""

import json
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Optional

NUMEROS_TEXTO = {
    "uno": 1, "primero": 1, "primera": 1, "dos": 2, "segundo": 2, "segunda": 2,
    "tres": 3, "tercero": 3, "tercera": 3, "cuatro": 4, "cuarto": 4,
    "cinco": 5, "quinto": 5, "seis": 6, "sexto": 6, "siete": 7, "septimo": 7,
    "ocho": 8, "octavo": 8, "nueve": 9, "noveno": 9, "diez": 10, "decimo": 10,
}


@dataclass
class ResultadoInterpretacion:
    reconocido: bool
    intencion: Optional[str] = None
    accion: Optional[str] = None
    parametro: Optional[str] = None
    frase_detectada: Optional[str] = None
    confianza: float = 0.0


class InterpretadorComandos:
    UMBRAL_CONFIANZA = 0.75

    def __init__(self, ruta_config: str):
        self.ruta_config = Path(ruta_config)
        self.comandos = []
        self.comandos_dinamicos = []
        self._cargar_configuracion()

    def _cargar_configuracion(self):
        if not self.ruta_config.exists():
            raise FileNotFoundError(f"No se encontró el archivo de configuración: {self.ruta_config}")
        with open(self.ruta_config, "r", encoding="utf-8") as f:
            data = json.load(f)
        self.comandos = data.get("comandos", [])
        self.comandos_dinamicos = data.get("comandos_dinamicos", [])

    @staticmethod
    def _normalizar(texto: str) -> str:
        texto = texto.lower().strip()
        texto = unicodedata.normalize("NFKD", texto)
        texto = "".join(c for c in texto if not unicodedata.combining(c))
        texto = re.sub(r"[^\w\s]", "", texto)
        texto = re.sub(r"\s+", " ", texto)
        return texto

    @staticmethod
    def _similitud(a: str, b: str) -> float:
        return SequenceMatcher(None, a, b).ratio()

    def _resolver_numero(self, texto: str) -> Optional[int]:
        texto = texto.strip()
        if texto.isdigit():
            return int(texto)
        return NUMEROS_TEXTO.get(self._normalizar(texto))

    @staticmethod
    def _normalizar_patron(patron: str) -> str:
        texto = unicodedata.normalize("NFKD", patron)
        return "".join(c for c in texto if not unicodedata.combining(c))

    def _intentar_comandos_dinamicos(self, texto_original):
        texto_norm = self._normalizar(texto_original)
        for comando in self.comandos_dinamicos:
            for patron in comando["patrones"]:
                patron_norm = self._normalizar_patron(patron)
                match = re.match(patron_norm, texto_norm)
                if match:
                    parametro_crudo = match.group(1).strip() if match.groups() else None
                    if comando.get("tipo_parametro") == "numero":
                        numero = self._resolver_numero(parametro_crudo)
                        if numero is None:
                            continue
                        parametro_final = numero
                    else:
                        parametro_final = parametro_crudo
                    return ResultadoInterpretacion(
                        reconocido=True, intencion=comando["intencion"],
                        accion=comando["accion"], parametro=parametro_final,
                        frase_detectada=texto_original, confianza=1.0,
                    )
        return None

    def _intentar_comandos_fijos(self, texto_reconocido):
        texto_norm = self._normalizar(texto_reconocido)

        # Primera pasada: coincidencia EXACTA únicamente (máxima prioridad,
        # sin importar el orden en que aparezcan los comandos en el JSON).
        for comando in self.comandos:
            for frase in comando["frases"]:
                if self._normalizar(frase) == texto_norm:
                    return ResultadoInterpretacion(
                        reconocido=True, intencion=comando["intencion"],
                        accion=comando["accion"], parametro=comando.get("parametro"),
                        frase_detectada=frase, confianza=1.0,
                    )

        # Segunda pasada: coincidencia parcial (frase contenida en el texto).
        # Se elige la frase MÁS LARGA/específica entre todas las que calcen,
        # para evitar que una frase corta ("borra") gane sobre una más
        # específica ("borra palabra") solo por aparecer primero en la lista.
        mejor_parcial = None
        mejor_longitud = -1
        for comando in self.comandos:
            for frase in comando["frases"]:
                frase_norm = self._normalizar(frase)
                if frase_norm in texto_norm and len(frase_norm) > mejor_longitud:
                    mejor_longitud = len(frase_norm)
                    mejor_parcial = ResultadoInterpretacion(
                        reconocido=True, intencion=comando["intencion"],
                        accion=comando["accion"], parametro=comando.get("parametro"),
                        frase_detectada=frase, confianza=1.0,
                    )
        if mejor_parcial is not None:
            return mejor_parcial

        # Tercera pasada: similitud aproximada (tolera errores de transcripción).
        mejor_resultado = ResultadoInterpretacion(reconocido=False)
        mejor_score = 0.0
        for comando in self.comandos:
            for frase in comando["frases"]:
                frase_norm = self._normalizar(frase)
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

    def interpretar(self, texto_reconocido):
        if not texto_reconocido or not texto_reconocido.strip():
            return ResultadoInterpretacion(reconocido=False)
        resultado_dinamico = self._intentar_comandos_dinamicos(texto_reconocido)
        if resultado_dinamico is not None:
            return resultado_dinamico
        return self._intentar_comandos_fijos(texto_reconocido)

    def listar_comandos_disponibles(self):
        fijos = [{"intencion": c["intencion"], "frases": c["frases"]} for c in self.comandos]
        dinamicos = [{"intencion": c["intencion"], "patrones": c["patrones"]} for c in self.comandos_dinamicos]
        return {"fijos": fijos, "dinamicos": dinamicos}