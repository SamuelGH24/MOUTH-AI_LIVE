"""
Mouth AI Live - Prueba de integración: ASR + Interpretación de comandos + Acciones
Este script conecta el reconocimiento de voz con el intérprete de comandos
y ejecuta las acciones correspondientes en el sistema operativo.

Ejecutar desde la raíz del proyecto:
    python main_test.py
"""

import sys
from pathlib import Path

RAIZ = Path(__file__).parent
sys.path.insert(0, str(RAIZ / "asr"))
sys.path.insert(0, str(RAIZ / "interpreter"))
sys.path.insert(0, str(RAIZ / "actions"))

from vosk_asr import CapturaASR  # noqa: E402
from interpreter import InterpretadorComandos  # noqa: E402
from actions import EjecutorAcciones  # noqa: E402

RUTA_CONFIG_COMANDOS = str(RAIZ / "config" / "comandos.json")


def main():
    print("=" * 60)
    print("MOUTH AI LIVE - Prueba de integración completa")
    print("=" * 60)

    interprete = InterpretadorComandos(RUTA_CONFIG_COMANDOS)
    ejecutor = EjecutorAcciones()
    print(f"[OK] {len(interprete.comandos)} comandos cargados.\n")

    asr = CapturaASR()

    def procesar_texto_reconocido(texto: str):
        print(f"\n🎤 Reconocido: '{texto}'")
        resultado = interprete.interpretar(texto)

        if not resultado.reconocido:
            print(f"   ⚠️  Sin comando válido (confianza: {resultado.confianza:.2f})")
            return

        print(f"   ✅ Intención: {resultado.intencion} (confianza {resultado.confianza:.2f})")
        exito, mensaje = ejecutor.ejecutar(resultado.accion, resultado.parametro)

        if mensaje == "DETENER_SISTEMA":
            print("   🛑 Comando de detener recibido. Cerrando sistema...")
            raise KeyboardInterrupt

        simbolo = "✔️" if exito else "❌"
        print(f"   {simbolo} {mensaje}")

    asr.iniciar_microfono()
    asr.escuchar_y_transcribir(
        callback_texto=procesar_texto_reconocido,
    )


if __name__ == "__main__":
    main()