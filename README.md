# Mouth AI Live

Prototipo de software modular basado en inteligencia artificial que permite
controlar un computador con Windows mediante comandos de voz en español, **sin
conexión a internet**, con enfoque en accesibilidad para personas con
limitaciones motrices o visuales.

Proyecto de grado — Universidad Antonio Nariño, Ingeniería de Sistemas y Computación (2026).

**Autor:** Samuel Gerena Huelgos
**Director:** Fabio Antonio González Mendieta

---

## Cómo funciona

```
Micrófono → ASR (Vosk) → Intérprete (comandos.json) → Acciones → Windows / Word / Chrome
                                                          │
                                                          └→ Narración por voz (TTS)
```

1. **ASR:** Vosk transcribe la voz en tiempo real (`vosk-model-es-0.42`), todo en local.
2. **Intérprete:** normaliza el texto (minúsculas, sin tildes) y lo compara por similitud
   con las frases de `config/comandos.json`, con un umbral de confianza configurable.
   Los comandos con texto variable usan expresiones regulares.
3. **Acciones:** ejecutan la orden y devuelven `(éxito, mensaje)`.
4. **Narración:** el resultado se lee en voz alta, para confirmar sin mirar la pantalla.

El micrófono se abre **una sola vez al iniciar** y el sistema empieza a escuchar sin
pulsar ningún botón. "Pausar" no cierra el micrófono: solo hace que se ignoren las
órdenes hasta decir "reanuda", así que todo el ciclo es accesible por voz.

## Estado actual del proyecto

### ✅ Funcionando y probado (con micrófono real en Windows)

El archivo `config/comandos.json` define **63 comandos fijos** (215 frases
equivalentes) y **4 comandos dinámicos** (búsqueda, dictado, guardar como y abrir
archivo por nombre).

- **Aplicaciones y web:** abrir Chrome, Bloc de notas, Calculadora, Word y Google.
- **Navegación asistida:** replica lo que ya hace una persona con lector de pantalla
  (Tab, flechas, Enter, atrás/adelante de página).
  - `lee`: lee en voz alta el elemento enfocado con UI Automation (`uiautomation`,
    la misma base de NVDA o JAWS) y `pyttsx3`.
  - `busca [lo que sea]`: escribe en la barra de direcciones (Ctrl+L) y pulsa Enter.
- **Dictado:** `escribe [texto]`.
  - **En Word:** escribe directo en el documento por COM (`pywin32`), sin usar el
    portapapeles, con tildes y eñes correctas y espacios automáticos entre dictados.
  - **En otras aplicaciones:** pega con Ctrl+V y luego restaura el portapapeles del usuario.
- **Edición por voz:** copiar, pegar, cortar, deshacer, rehacer; borrar letra,
  palabra, línea o selección; seleccionar todo, línea o palabra; mover el cursor por
  letra, palabra, línea o documento; nueva línea.
- **Formato (Word):** negrita, cursiva, subrayado, letra más grande o más pequeña.
- **Documentos de Word por voz:**
  - `documento nuevo`: abre un documento en blanco sin pasar por la pantalla de inicio.
  - `guarda como [nombre]`: agrega `.docx`, guarda en `Documentos/Mouth AI Live/`
    (configurable con `CARPETA_GUARDADO`) y nunca pisa un archivo existente
    (agrega `(2)`, `(3)`...).
  - `abre el archivo [nombre]`: búsqueda aproximada, tolerante a tildes, mayúsculas y
    números hablados ("taller cuatro" encuentra `Taller4_SG.docx`).
  - `abre el último documento`.
- **Mouse por voz:** mover en cuatro direcciones, clic, doble clic y clic derecho.
- **Sonido y sistema:** subir, bajar y silenciar el volumen; bloquear sesión y apagar
  el equipo (protegidos por modo prueba, ver abajo).
- **Control del propio sistema:** pausar, reanudar, detener, activar/desactivar la voz
  de confirmación y `ayuda` (narra los comandos disponibles).
- **Interfaz gráfica** (`gui/app.py`, Tkinter): botones Pausar y Reanudar, registro de
  actividad, limpiar registro y resumen de métricas. La escucha corre en un hilo aparte.
- **Métricas** (`metrics/logger.py`): cada interacción se guarda en un CSV (texto
  reconocido, intención, confianza, éxito, tiempo de respuesta en ms) y se calculan
  tasas de reconocimiento, error y tiempos.

> **Detalle técnico:** en Word, la selección, el formato y el guardado se hacen por
> COM en lugar de atajos de teclado, así funcionan igual en Word en español o inglés
> (donde los atajos cambian) y no dependen de Bloq Num. Fuera de Word se envían teclas
> como *teclas extendidas*, para que Shift+flecha seleccione correctamente.

### ⚠️ Construido pero NO integrado

- **Preprocesamiento de señal** (`preprocessing/preprocessing.py`): sin calibrar con
  hardware real, desactivado por ahora.

### 🚧 Fuera del alcance final (trabajo futuro)

El alcance se delimitó a **navegador web y Microsoft Word**. Quedan pendientes:

- Comandos para Microsoft Excel y PowerPoint.
- Comando para cerrar aplicaciones.
- Comparar la precisión de Vosk con la de Whisper.
- Calibrar e integrar el preprocesamiento de señal y medir su efecto con ruido.
- Pruebas de usabilidad con un grupo mayor de usuarios y un instrumento estandarizado.
- Compatibilidad con aplicaciones con información de accesibilidad incompleta.

## Resultados de las pruebas

Métricas registradas por el propio sistema: **928 interacciones en 12 sesiones**
(19 de agosto a 5 de octubre de 2026), incluyendo el uso durante el desarrollo y
grabaciones de una persona con discapacidad.

| Indicador | Valor |
| --- | --- |
| Interacciones identificadas como comando | 202 |
| Comandos ejecutados con éxito | 180 (**89,1 %**) |
| Confianza promedio del reconocimiento | 0,97 |
| Tiempo de respuesta (mediana / percentil 90) | 1.637 ms / 3.931 ms |

| Categoría | Identificados | Tasa de éxito |
| --- | --- | --- |
| Búsqueda y navegador web | 48 | 100 % |
| Dictado, edición y formato | 70 | 97,1 % |
| Navegación asistida | 28 | 100 % |
| Aplicaciones y documentos | 17 | 100 % |
| Control del sistema de voz | 5 | 100 % |
| Lectura en voz alta | 19 | 36,8 % |
| Control del mouse | 15 | 46,7 % |

**Cómo leer estos datos:**

- Las otras 726 interacciones no se identificaron como comando (habla del entorno,
  comentarios u órdenes no reconocidas); el registro no permite separarlas, así que el
  21,8 % de interacciones identificadas **no** es la proporción de órdenes acertadas.
- La lectura en voz alta depende de que cada aplicación o página exponga bien su
  información de accesibilidad.
- Los fallos del mouse ocurrieron todos en una sola sesión (21 de septiembre), cuando la
  acción aún no estaba disponible; falta repetir la prueba.
- El registro no guarda variables como ruido ambiental o distancia al micrófono, ni
  identifica al participante de cada sesión.

## Modo prueba

`MODO_PRUEBA = True` (en `actions/actions.py`) es el valor por defecto. Con él activo,
**bloquear sesión** y **apagar equipo** solo registran lo que habrían hecho, sin
ejecutarlo. Cámbialo a `False` únicamente si quieres que se ejecuten de verdad.

## Estructura del proyecto

```
MOUTH-AI_LIVE/
├── actions/          # Ejecutor de acciones (apps, volumen, sesión, narración)
├── asr/              # Captura de audio y reconocimiento con Vosk
├── config/           # comandos.json (comandos, patrones y ayuda hablada)
├── gui/              # Interfaz gráfica (punto de entrada: app.py)
├── interpreter/      # Texto reconocido → intención
├── metrics/          # Registro de interacciones y tasas de error
├── navigation/       # Navegación, dictado, edición, Word, mouse, TTS
├── preprocessing/    # Preprocesamiento de señal (no integrado)
├── models/           # Modelo de Vosk (ignorado por git)
├── main_test.py
├── requirements.txt
└── .gitignore
```

## Requisitos

- **Windows de 64 bits** (usa UI Automation, COM y la API de teclado de Windows).
- **Python 3.10** (desarrollado con 3.10.6; solo necesario al instalar desde el código).
- **Microsoft Word**, para el dictado y los documentos por COM.
- Micrófono (obligatorio) y parlantes o audífonos para las confirmaciones.
- Recomendado: 4 núcleos, 8 GB de RAM y ~3 GB libres (el modelo pesa ~1,4 GB).
- Librerías: `vosk`, `pyaudio`, `numpy`, `pyautogui`, `pyttsx3`, `uiautomation`,
  `pyperclip`, `pywin32`.

## Instalación

1. `python -m venv venv` y `venv\Scripts\activate`
2. `pip install -r requirements.txt`
3. Descargar `vosk-model-es-0.42` desde [alphacephei.com/vosk/models](https://alphacephei.com/vosk/models)
   y descomprimirlo en `models/vosk-model-es-0.42/`
4. Ajustar `RUTA_MODELO` en `asr/vosk_asr.py` si el modelo está en otra ubicación.
   Debe apuntar a la carpeta que contiene directamente los archivos del modelo.

## Uso

```
python gui/app.py
```

Espera el mensaje "Cargando modelo de voz..." y, cuando termine, **empieza a hablar**:
no hay botón de inicio. Haz una pausa de un par de segundos entre órdenes.

**Buscar y explorar la web**
"abre chrome" → "busca clima Bogotá" → "siguiente" → "lee" → "entra"

**Escribir un documento**
"abre word" → "escribe hola cómo estás" → "selecciona línea" → "negrita" → "guarda como informe final"

**Retomar un trabajo**
"abre el último documento" o "abre el archivo informe final"

Di "ayuda" o "qué puedo decir" para que el sistema narre los comandos disponibles.

### Ejecutable (opcional)

Se puede empaquetar con PyInstaller (Vosk y pyttsx3 incluyen binarios):

```
pip install pyinstaller
pyinstaller --noconfirm --collect-all vosk --collect-all pyttsx3 gui/app.py
```

Copia junto al ejecutable las carpetas `config/` y la del modelo de voz.

## Personalizar comandos

Edita `config/comandos.json` (haz una copia antes), agrega la frase nueva a la lista
`frases` del comando correspondiente, en minúsculas y sin tildes, y reinicia el programa.
Las frases muy cortas o parecidas entre sí aumentan los errores de reconocimiento.

## Limitaciones conocidas

- Solo Windows y solo español.
- Solo entiende los comandos definidos en `comandos.json`; no interpreta lenguaje abierto.
- La lectura en voz alta depende de la información de accesibilidad de cada aplicación.
- El dictado no funciona en ventanas abiertas con permisos de administrador.
- El tamaño de letra solo funciona en Word.
- El reconocimiento depende del micrófono, el ruido ambiental y la pronunciación; no hay
  cancelación activa de ruido.

## Documentación

El repositorio acompaña el trabajo de grado, el Manual Técnico y el Manual de Usuario.
