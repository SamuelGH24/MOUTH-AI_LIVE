# Mouth AI Live

Prototipo de software modular basado en inteligencia artificial que permite
controlar un computador mediante comandos de voz, con enfoque en accesibilidad
para personas con discapacidad motriz y visual. Proyecto de grado —
Universidad Antonio Nariño, Ingeniería de Sistemas y Computación.

Autor: Samuel Gerena Huelgos
Director: Fabio Gonzales

## Estado actual del proyecto

### Funcionando y probado (con micrófono real en Windows)

- **Captura de audio + ASR** (`asr/vosk_asr.py`): reconocimiento de voz en
  tiempo real con Vosk. Corta frases largas automáticamente (3.5s) para
  evitar que dos comandos dichos con poca pausa se mezclen en un solo texto.
  Tolerante a fallos del micrófono (se reintenta solo, no cae el programa).
- **Intérprete de comandos** (`interpreter/interpreter.py`): comandos fijos
  y dinámicos (texto libre extraído por expresiones regulares).
- **Ejecución de acciones** (`actions/actions.py`): abrir apps/URLs, volumen,
  bloquear/apagar (modo prueba), pausar/reanudar el sistema.
- **Interfaz gráfica** (`gui/app.py`): arranca y empieza a escuchar sola, sin
  necesitar clic (indispensable para discapacidad motriz). El micrófono
  nunca se apaga durante la sesión; "pausar" solo ignora comandos hasta
  que se diga "reanuda" — así la reanudación también es por voz.
- **Métricas** (`metrics/logger.py`): registra cada interacción en CSV y
  calcula tasas de reconocimiento/error.
- **Navegación asistida** (`navigation/navigation.py`):
  - `sube` / `baja` / `siguiente` / `anterior` / `entra`: navegación tipo
    lector de pantalla (Tab/flechas/Enter).
  - `lee`: lee en voz alta el elemento enfocado (TTS + UI Automation).
  - `busca [algo]`: búsqueda en el navegador escribiendo en la barra de
    direcciones (sin Selenium).
  - `escribe [algo]` / `dicta [algo]`: dictado de texto en cualquier campo
    (Word, Bloc de notas, etc.), vía portapapeles para que tildes y eñes
    salgan bien.
  - Edición: `nueva linea`, `borra`, `borra palabra`, `guarda`, `deshacer`,
    `selecciona todo`.
  - Formato de texto: `negrita`, `cursiva`, `subrayado`, `letra mas grande`,
    `letra mas pequena`.
  - `que puedo decir` / `ayuda`: narra por voz un resumen de comandos
    disponibles (configurable en `comandos.json`, campo `ayuda_hablada`).

### Mantenimiento reciente
- La ruta del modelo de Vosk ya no está fija en el código: se lee desde
  `config/ajustes.json`. Si cambias de computador o mueves la carpeta,
  solo se edita ese archivo.
- Mensajes de error más claros cuando falta configuración o el modelo no
  se encuentra.

### Construido pero NO integrado
- **Preprocesamiento de señal** (`preprocessing/preprocessing.py`): sin
  calibrar con hardware real.

### Pendiente
- Soporte de idioma inglés (Vosk lo permite con otro modelo, falta integrarlo).
- Control de mouse por voz (para discapacidad motriz que sí puede ver pantalla).
- Confirmación por voz activable/desactivable según perfil (motriz vs visual).
- Seguridad: por ahora cualquier voz cercana al micrófono puede dar órdenes.
- Registro de errores en archivo persistente (`logs/`) para diagnóstico
  posterior sin depender de estar mirando la pantalla en el momento del fallo.

## Estructura del proyecto
Mouth-ai-live/
├── asr/
├── interpreter/
├── actions/
├── navigation/
├── metrics/
├── gui/
├── preprocessing/
├── config/
│ ├── comandos.json # comandos reconocidos + texto de ayuda hablada
│ └── ajustes.json # ruta del modelo de Vosk (edítalo si cambias de PC)
├── tests/
├── models/ # ignorado por git, se descarga aparte
├── requirements.txt
└── .gitignore

## Instalación

1. `python -m venv venv` y `venv\Scripts\activate`
2. `pip install -r requirements.txt`
3. Descargar el modelo desde [alphacephei.com/vosk/models](https://alphacephei.com/vosk/models)
   y descomprimirlo donde prefieras.
4. Editar `config/ajustes.json` con la ruta real donde quedó el modelo.

## Uso


