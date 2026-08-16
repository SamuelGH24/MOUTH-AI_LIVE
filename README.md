# Mouth AI Live

Prototipo de software modular basado en inteligencia artificial que permite
controlar un computador mediante comandos de voz. Proyecto de grado —
Universidad Antonio Nariño, Ingeniería de Sistemas y Computación.

Autor: Samuel Gerena Huelgos
Director: Fabio Gonzales

## Estado actual del proyecto

### ✅ Funcionando y probado (con micrófono real en Windows)
- **Captura de audio + ASR** (`asr/vosk_asr.py`): reconocimiento de voz en
  tiempo real usando Vosk (modelo `vosk-model-es-0.42`).
- **Intérprete de comandos** (`interpreter/interpreter.py`): soporta comandos
  fijos (frases exactas/similares) y comandos dinámicos (texto variable
  extraído por expresiones regulares, ej. "busca [lo que sea]").
- **Ejecución de acciones** (`actions/actions.py`): abrir aplicaciones
  (Chrome, Notepad, Calculadora), abrir URLs, controlar volumen, bloquear
  sesión y apagar equipo (estas dos últimas en modo prueba por defecto).
- **Integrador** (`main_test.py`): conecta los tres módulos anteriores.

### ⚠️ Construido pero NO integrado
- **Preprocesamiento de señal** (`preprocessing/preprocessing.py`): normaliza
  audio y detecta actividad de voz (VAD). Se desactivó porque el umbral de
  energía no está calibrado con hardware real y estaba bloqueando voz válida
  del usuario. Pendiente: calibrar con datos reales antes de reactivar.

### 🚧 Definido pero sin código aún
- Control de volumen por niveles específicos (tipo Alexa).
- Navegación dentro de resultados de búsqueda / control de Chrome (comandos
  ya existen en `config/comandos.json`, pero `actions.py` responde que aún
  no están implementados).
- Dictado de texto hacia Word.
- Interfaz gráfica (Tkinter).

## Instalación

1. Crear y activar entorno virtual: