# Mouth AI Live

Prototipo de software modular basado en inteligencia artificial que permite
controlar un computador mediante comandos de voz, con enfoque en accesibilidad
para personas con limitaciones motrices o visuales. Proyecto de grado —
Universidad Antonio Nariño, Ingeniería de Sistemas y Computación.

Autor: Samuel Gerena Huelgos
Director: Fabio Gonzales

## Estado actual del proyecto

### ✅ Funcionando y probado (con micrófono real en Windows)

- **Captura de audio + ASR** (`asr/vosk_asr.py`): reconocimiento de voz en
  tiempo real usando Vosk (modelo `vosk-model-es-0.42`).
- **Intérprete de comandos** (`interpreter/interpreter.py`): comandos fijos
  y comandos dinámicos (texto variable extraído por expresiones regulares).
- **Ejecución de acciones** (`actions/actions.py`): abrir aplicaciones,
  abrir URLs, controlar volumen, bloquear sesión/apagar equipo (modo prueba).
- **Interfaz gráfica** (`gui/app.py`): iniciar, detener y monitorear el
  sistema con botones, sin depender de la terminal.
- **Sistema de métricas** (`metrics/logger.py`): registra cada interacción
  en un CSV y calcula tasas de reconocimiento/error automáticamente.
- **Navegación asistida por voz** (`navigation/navigation.py`) — replicando
  cómo ya navega alguien con un lector de pantalla (Tab/flechas/Enter):
  - `sube` / `baja`: desplazamiento de página.
  - `siguiente` / `anterior`: moverse entre elementos (Tab / Shift+Tab).
  - `entra`: activa el elemento enfocado (Enter).
  - `atras` / `pagina siguiente`: navegación de historial.
  - `lee`: lee en voz alta (TTS vía `pyttsx3`) el elemento enfocado, usando
    UI Automation de Windows (`uiautomation`) — la misma base que usan
    lectores de pantalla reales como NVDA o JAWS.
  - `busca [lo que sea]`: escribe la búsqueda en la barra de direcciones
    y presiona Enter (mismo método que una persona con teclado).

  **Flujo de uso real:** "busca clima Bogotá" → el sistema escribe y
  busca → la persona dice "siguiente" y "lee" para explorar resultados →
  dice "entra" cuando encuentra lo que buscaba.

  **Limitación honesta:** la lectura depende de que la página exponga bien
  su información de accesibilidad. Puede fallar en elementos sin descripción.

### ⚠️ Construido pero NO integrado
- **Preprocesamiento de señal** (`preprocessing/preprocessing.py`): sin
  calibrar con hardware real, desactivado por ahora.

### 🚧 Pendiente (parte central del alcance, no opcional)
- Dictado de texto hacia Word/editores de texto.
- Interfaz gráfica: mejoras visuales pendientes.

## Estructura del proyecto

Mouth-ai-live/
├── asr/
├── interpreter/
├── actions/
├── navigation/
├── metrics/
├── gui/
├── preprocessing/
├── config/comandos.json
├── tests/
├── models/ (ignorado por git)
├── requirements.txt
└── .gitignore

## Instalación

1. `python -m venv venv` y `venv\Scripts\activate`
2. `pip install -r requirements.txt`
3. Descargar el modelo desde [alphacephei.com/vosk/models](https://alphacephei.com/vosk/models)
   (`vosk-model-es-0.42`) y descomprimir en `models/vosk-model-es-0.42/`
4. Ajustar `RUTA_MODELO` en `asr/vosk_asr.py` si es necesario.

## Uso

python gui/app.py


Presiona **Iniciar** y di comandos como "abre google", "busca clima Bogotá",
"siguiente", "lee", "entra", "sube el volumen".

## Comandos disponibles

Ver/editar `config/comandos.json`