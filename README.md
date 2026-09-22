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

- **Dictado de texto** (`navigation/navigation.py`, método `escribir_texto`):
  - Si la ventana activa es **Word**, el texto se escribe directo en el
    documento vía COM (`win32com`, `Selection.TypeText`) — sin tocar el
    portapapeles ni simular teclas, más confiable con tildes/eñes.
  - En cualquier otra app (Notepad, navegador, etc.), usa el método genérico
    de portapapeles + Ctrl+V, que sigue funcionando como respaldo.
  - Requiere `pywin32` instalado y registrado
    (`python venv\Scripts\pywin32_postinstall.py -install`).

- **Control de mouse por voz** (`navigation/navigation.py`, `actions/actions.py`):
  - `mueve arriba/abajo/izquierda/derecha`: desplaza el cursor en pasos
    fijos, con límite de pantalla (nunca se sale del monitor ni dispara el
    "failsafe" de `pyautogui` en la esquina).
  - `clic` / `doble clic` / `clic derecho`: ejecuta el clic correspondiente.
  - **Nota de diseño:** el mouse a nudges es el método de último recurso.
    Para la mayoría de apps normales, moverse por foco de teclado
    (`siguiente`/`entra`, ya implementado) es más rápido y confiable por
    voz que apuntar con el cursor. El mouse queda pensado como respaldo
    para lienzos/imágenes sin elementos nombrados.

- **Confirmación por voz de cada acción** (`actions/actions.py`): el sistema
  narra el resultado de lo que acaba de hacer ("Abriendo aplicación: chrome",
  "Volumen: subir", etc.), no solo lo escribe en el log visual — importante
  para usuarios con limitación visual. Se puede silenciar/reactivar con los
  comandos `desactiva la voz` / `activa la voz`.

### ⚠️ Construido pero NO integrado
- **Preprocesamiento de señal** (`preprocessing/preprocessing.py`): sin
  calibrar con hardware real, desactivado por ahora.

### 🚧 Pendiente
- Interfaz gráfica: mejoras visuales pendientes.
- `cambia a [app]`: enfocar una ventana ya abierta por su título en vez de
  lanzar una instancia nueva (hoy `abre chrome` siempre abre una copia más,
  aunque Chrome ya esté abierto).
- Clic por nombre de elemento (`clic en Aceptar`) vía UI Automation, sin
  depender de coordenadas de mouse.
- Navegación por dirección lógica en Excel/PowerPoint vía COM (celda,
  diapositiva) en vez de mouse/grid, para cuando el contenido esté fuera
  del área visible en pantalla.

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
3. Registrar los módulos COM de `pywin32` (necesario para el dictado en Word):
   `python venv\Scripts\pywin32_postinstall.py -install`
4. Descargar el modelo desde [alphacephei.com/vosk/models](https://alphacephei.com/vosk/models)
   (`vosk-model-es-0.42`) y descomprimir en `models/vosk-model-es-0.42/`
5. Ajustar `RUTA_MODELO` en `asr/vosk_asr.py` si es necesario.

## Uso

python gui/app.py


El micrófono arranca solo al abrir la app (diseño manos libres, ver
`gui/app.py`) — no hay botón "Iniciar". Usa **Pausar** / **Reanudar** si
necesitas que el sistema ignore comandos temporalmente (o di "detente" /
"reanuda"). Di comandos como "abre google", "busca clima Bogotá",
"siguiente", "lee", "entra", "sube el volumen", "mueve derecha", "clic",
"escribe [lo que quieras]", "desactiva la voz".

## Comandos disponibles

Ver/editar `config/comandos.json`