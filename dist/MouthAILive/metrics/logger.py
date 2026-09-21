"""
Módulo de Métricas - Mouth AI Live

Registra cada interacción (comando dicho -> interpretado -> ejecutado) en un
archivo CSV, y calcula las métricas de desempeño que pide el Incremento 4
del proyecto: tasa de reconocimiento exitoso, tiempo de respuesta y tasa
de error.

NOTA HONESTA sobre "tiempo de respuesta": se mide desde que el ASR entrega
el texto final reconocido hasta que la acción termina de ejecutarse. NO
incluye el tiempo que Vosk tarda internamente en procesar el audio (eso
ya lo observamos manualmente antes: el modelo grande introduce un retraso
notorio). Si se quiere medir el pipeline completo de principio a fin, habría
que añadir una marca de tiempo en el momento exacto en que la persona
empieza a hablar, lo cual Vosk no expone directamente.
"""

import csv
import statistics
import time
from datetime import datetime
from pathlib import Path

RUTA_CSV_POR_DEFECTO = Path(__file__).parent / "registro_pruebas.csv"
ENCABEZADOS = [
    "timestamp",
    "texto_reconocido",
    "intencion",
    "confianza",
    "reconocido",
    "accion",
    "parametro",
    "exito_ejecucion",
    "mensaje",
    "tiempo_respuesta_ms",
]


class RegistradorMetricas:
    """
    Registra cada evento de interacción en un CSV. Se puede usar como
    cronómetro (iniciar_medicion / finalizar_y_registrar) o registrar
    manualmente con un tiempo ya calculado.
    """

    def __init__(self, ruta_csv: str = None):
        self.ruta_csv = Path(ruta_csv) if ruta_csv else RUTA_CSV_POR_DEFECTO
        self._inicio_medicion = None
        self._asegurar_archivo_con_encabezados()

    def _asegurar_archivo_con_encabezados(self):
        if not self.ruta_csv.exists():
            with open(self.ruta_csv, "w", newline="", encoding="utf-8") as f:
                csv.writer(f).writerow(ENCABEZADOS)

    def iniciar_medicion(self):
        """Llamar justo cuando el ASR entrega el texto reconocido."""
        self._inicio_medicion = time.perf_counter()

    def finalizar_y_registrar(self, texto_reconocido, resultado_interpretacion, exito_ejecucion, mensaje_ejecucion):
        """
        Llamar justo después de ejecutar (o intentar ejecutar) la acción.
        Calcula el tiempo transcurrido desde iniciar_medicion() y guarda
        la fila completa en el CSV.
        """
        tiempo_ms = None
        if self._inicio_medicion is not None:
            tiempo_ms = round((time.perf_counter() - self._inicio_medicion) * 1000, 2)
            self._inicio_medicion = None

        self._registrar_fila(
            texto_reconocido=texto_reconocido,
            intencion=resultado_interpretacion.intencion,
            confianza=resultado_interpretacion.confianza,
            reconocido=resultado_interpretacion.reconocido,
            accion=resultado_interpretacion.accion,
            parametro=resultado_interpretacion.parametro,
            exito_ejecucion=exito_ejecucion,
            mensaje=mensaje_ejecucion,
            tiempo_respuesta_ms=tiempo_ms,
        )

    def _registrar_fila(self, texto_reconocido, intencion, confianza, reconocido,
                         accion, parametro, exito_ejecucion, mensaje, tiempo_respuesta_ms):
        with open(self.ruta_csv, "a", newline="", encoding="utf-8") as f:
            csv.writer(f).writerow([
                datetime.now().isoformat(timespec="seconds"),
                texto_reconocido,
                intencion,
                round(confianza, 3) if confianza is not None else "",
                reconocido,
                accion,
                parametro,
                exito_ejecucion,
                mensaje,
                tiempo_respuesta_ms if tiempo_respuesta_ms is not None else "",
            ])

    # -------------------------------------------------------------------
    def generar_resumen(self):
        """
        Lee todo el CSV acumulado y calcula las métricas del Incremento 4.
        Devuelve un diccionario con los resultados.
        """
        if not self.ruta_csv.exists():
            return None

        with open(self.ruta_csv, "r", newline="", encoding="utf-8") as f:
            filas = list(csv.DictReader(f))

        if not filas:
            return None

        total = len(filas)
        reconocidos = [f for f in filas if f["reconocido"] == "True"]
        no_reconocidos = [f for f in filas if f["reconocido"] == "False"]
        exitosos = [f for f in filas if f["exito_ejecucion"] == "True"]

        tiempos = [
            float(f["tiempo_respuesta_ms"])
            for f in filas
            if f["tiempo_respuesta_ms"] not in ("", None)
        ]

        tasa_reconocimiento = len(reconocidos) / total * 100
        tasa_error = len(no_reconocidos) / total * 100
        tasa_exito_ejecucion = (len(exitosos) / len(reconocidos) * 100) if reconocidos else 0.0

        return {
            "total_intentos": total,
            "reconocidos": len(reconocidos),
            "no_reconocidos": len(no_reconocidos),
            "tasa_reconocimiento_pct": round(tasa_reconocimiento, 2),
            "tasa_error_pct": round(tasa_error, 2),
            "tasa_exito_ejecucion_pct": round(tasa_exito_ejecucion, 2),
            "tiempo_respuesta_promedio_ms": round(statistics.mean(tiempos), 2) if tiempos else None,
            "tiempo_respuesta_mediana_ms": round(statistics.median(tiempos), 2) if tiempos else None,
            "tiempo_respuesta_max_ms": round(max(tiempos), 2) if tiempos else None,
            "tiempo_respuesta_min_ms": round(min(tiempos), 2) if tiempos else None,
        }

    def imprimir_resumen(self):
        resumen = self.generar_resumen()
        if resumen is None:
            print("No hay datos registrados todavía.")
            return

        print("=" * 55)
        print("RESUMEN DE MÉTRICAS - Mouth AI Live")
        print("=" * 55)
        print(f"Total de intentos:              {resumen['total_intentos']}")
        print(f"Reconocidos:                     {resumen['reconocidos']}")
        print(f"No reconocidos:                  {resumen['no_reconocidos']}")
        print(f"Tasa de reconocimiento exitoso:  {resumen['tasa_reconocimiento_pct']}%")
        print(f"Tasa de error:                   {resumen['tasa_error_pct']}%")
        print(f"Tasa de éxito en ejecución:      {resumen['tasa_exito_ejecucion_pct']}%")
        if resumen["tiempo_respuesta_promedio_ms"] is not None:
            print(f"Tiempo de respuesta promedio:    {resumen['tiempo_respuesta_promedio_ms']} ms")
            print(f"Tiempo de respuesta mediana:     {resumen['tiempo_respuesta_mediana_ms']} ms")
            print(f"Tiempo de respuesta mínimo:      {resumen['tiempo_respuesta_min_ms']} ms")
            print(f"Tiempo de respuesta máximo:      {resumen['tiempo_respuesta_max_ms']} ms")
        print("=" * 55)


if __name__ == "__main__":
    registrador = RegistradorMetricas()
    registrador.imprimir_resumen()