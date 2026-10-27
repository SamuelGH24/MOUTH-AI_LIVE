# Mouth AI Live

Prototipo de software modular basado en inteligencia artificial que permite
controlar un computador con Windows mediante comandos de voz en español, **sin
conexión a internet**. Está pensado para personas con limitaciones motrices o
visuales, y también sirve a quien necesite trabajar sin usar las manos.

Proyecto de grado, Universidad Antonio Nariño, Ingeniería de Sistemas y
Computación (Bogotá, 2026).

- **Autor:** Samuel Gerena Huelgos
- **Director:** Fabio Antonio González Mendieta

Tu voz se procesa en tu propio computador: no se envía a ningún servidor.

## Qué puede hacer

- Abrir programas y páginas, y controlar el volumen.
- Moverse por la pantalla como con un lector de pantalla y **escuchar** el
  elemento enfocado (`lee`).
- Buscar en internet y **dictar** texto en Word, Bloc de notas, formularios, etc.
- Editar: copiar, pegar, cortar, borrar, seleccionar y mover el cursor.
- Dar formato (negrita, cursiva, subrayado, tamaño de letra).
- Crear, guardar con nombre y abrir documentos.
- Mover el mouse y hacer clic.
- Pausar, reanudar y pedir ayuda hablada (`ayuda`).

Hay **77 comandos fijos** (261 frases equivalentes) y **7 comandos dinámicos**
con texto variable (`busca [texto]`, `escribe [texto]`, `guarda como [nombre]`,
`abre el archivo [nombre]`, `ve a celda [celda]`, `titulo [texto]`,
`contenido [texto]`). La lista completa y buscable está en la página de
referencia visual del proyecto.

| Categoría | Ejemplos |
|---|---|
| Aplicaciones y sistema | abre chrome, abre word, abre la calculadora, abre google, sube el volumen, silenciar, bloquear sesion, apagar el equipo |
| Navegación y lectura | sube, baja, siguiente, anterior, entra, atras, lee |
| Búsqueda y dictado | busca [texto], escribe [texto] |
| Edición | copia, pega, corta, borra palabra, borra linea, deshacer, rehacer, selecciona linea, fin de linea |
| Formato | negrita, cursiva, subrayado, letra mas grande |
| Documentos | documento nuevo, guarda, guarda como [nombre], abre el ultimo documento |
| Mouse | mueve arriba, clic, doble clic, clic derecho |
| Control del programa | detente, reanuda, activa la voz, desactiva la voz, ayuda |
| Excel y PowerPoint (extensión) | libro nuevo, celda abajo, suma arriba, diapositiva nueva, siguiente diapositiva, iniciar presentacion |

> **Excel y PowerPoint** se agregaron después del cierre formal del trabajo de
> grado. El alcance evaluado fue el navegador web y Microsoft Word; los
> comandos de Excel y PowerPoint no hacen parte de las pruebas de desempeño.

## Cómo funciona

```
Micrófono → [Preprocesamiento] → ASR (Vosk) → Intérprete → Acciones → Windows / Office
             (opcional, apagado)                                 ↓
                                                        Voz de confirmación (TTS)
```

| Módulo | Qué hace |
|---|---|
| `asr/` | Captura el audio por bloques y lo reconoce con Vosk (`vosk-model-es-0.42`). Corta la frase si pasan más de 3,5 s sin que el motor la dé por terminada. Reintenta abrir el micrófono si falla. |
| `interpreter/` | Normaliza el texto (minúsculas, sin tildes). Prueba primero los comandos dinámicos (expresiones regulares) y luego los fijos en tres pasadas: coincidencia exacta, frase más larga contenida en lo dicho, y similitud aproximada (umbral 0,75). |
| `actions/` | Ejecuta la acción y narra el resultado. Delega la navegación y la edición. |
| `navigation/` | Teclado simulado, lectura con UI Automation y pyttsx3, dictado, mouse, y control de Word, Excel y PowerPoint por COM. |
| `gui/` | Interfaz Tkinter. La escucha corre en un hilo aparte para que la ventana no se congele. Es el punto de entrada. |
| `metrics/` | Registra cada interacción en un CSV y calcula tasas de reconocimiento y error, y tiempos de respuesta. |
| `preprocessing/` | Detección de voz y normalización de volumen. Construido, pero no integrado. |
| `config/comandos.json` | Todos los comandos y el texto de la ayuda hablada. Se edita sin tocar el código. |

Detalles útiles:

- **Lectura:** `lee` usa UI Automation de Windows, la misma base de NVDA y JAWS.
  En Word, Excel y PowerPoint lee por COM: la selección, la celda activa o la
  diapositiva.
- **Dictado:** en Word y Excel el texto se escribe directo por COM. En otras
  aplicaciones se pega desde el portapapeles, que es más fiable con tildes y
  eñes, y el portapapeles del usuario se restaura después.
- **Formato y guardado:** en Office se hacen por COM, así que no dependen del
  idioma de la instalación.
- **Documentos:** se guardan en `Documentos\Mouth AI Live` con la extensión que
  corresponda. Si el nombre ya existe, se agrega `(2)`, `(3)`…
- **Seguridad:** `bloquear sesion` y `apagar el equipo` están en **modo prueba**
  por defecto: solo se registran, no se ejecutan.

## Requisitos

| | |
|---|---|
| Sistema operativo | Windows de 64 bits (10 u 11). Usa funciones propias de Windows. |
| Python | 3.10 (se desarrolló con 3.10.6). No hace falta con el ejecutable. |
| Procesador y memoria | 2 núcleos o más (recomendado 4), 8 GB de RAM recomendados. |
| Disco | Unos 3 GB libres (el modelo descomprimido pesa cerca de 1,4 GB). |
| Hardware | Micrófono (obligatorio; mejor de diadema o de escritorio) y parlantes o audífonos. |
| Aplicaciones | Chrome, Bloc de notas, calculadora y Microsoft Word de escritorio. Excel y PowerPoint, solo para sus comandos. |
| Internet | No hace falta para usarlo. Solo para descargar el modelo al instalar. |

Los valores de RAM y disco son de referencia y no se midieron formalmente.

Librerías principales: `vosk 0.3.45`, `pyaudio 0.2.14`, `numpy`, `pyautogui`,
`pyttsx3`, `uiautomation`. El código también usa **`pyperclip`** (dictado y
búsqueda) y **`pywin32`** (control de Office por COM); asegúrate de que estén
en `requirements.txt`.

## Instalación

### Desde el código fuente

1. Instala Python 3.10 de 64 bits y marca "Add Python to PATH".
2. Descarga el repositorio en una carpeta, por ejemplo `C:\MouthAILive`.
3. `python -m venv venv` y `venv\Scripts\activate`
4. `pip install -r requirements.txt`
5. Descarga `vosk-model-es-0.42` desde
   [alphacephei.com/vosk/models](https://alphacephei.com/vosk/models) y
   descomprímelo en `models/vosk-model-es-0.42/`.
6. Si el modelo quedó en otra ruta, ajusta `RUTA_MODELO` en `asr/vosk_asr.py`.
   Debe apuntar a la carpeta que contiene directamente los archivos del modelo.

Si falla `pyaudio`: usa Python 3.10 de 64 bits, actualiza pip
(`python -m pip install --upgrade pip`) y ejecuta `pip install pyaudio==0.2.14`.

### Ejecutable

Copia la carpeta completa del programa (sin mover el `.exe` fuera de ella) y
comprueba que incluye `config/` y la carpeta del modelo. Para generarlo:

```
pip install pyinstaller
pyinstaller --noconfirm --collect-all vosk --collect-all pyttsx3 gui/app.py
```

Luego copia junto al ejecutable las carpetas `config` y la del modelo, y pruébalo
en un equipo sin Python.

## Uso

```
python gui/app.py
```

O doble clic en el ejecutable. La ventana muestra "Cargando modelo de voz..."
unos segundos; luego el micrófono queda abierto y el programa **escucha solo**,
sin pulsar nada.

- Habla con claridad, en tono normal, y haz una pausa de un par de segundos
  entre una orden y la siguiente.
- **Pausar:** botón *Pausar* (sigue escuchando para saber cuándo reanudar).
- **Reanudar:** di "reanuda" o pulsa *Reanudar*.
- **Terminar:** di "detente".
- Si no sabes qué decir, di "ayuda".
- El registro de actividad muestra lo que entendió y lo que hizo.

Flujos de ejemplo:

- "abre word", "escribe hola, este es mi informe", "negrita", "guarda como informe final".
- "busca clima Bogotá", "siguiente", "lee", "entra".
- "abre excel", "escribe cien", "suma arriba".
- "abre powerpoint", "titulo mi proyecto", "diapositiva nueva", "iniciar presentacion".

## Configuración

- **Comandos:** edita `config/comandos.json` (frases en minúsculas y sin tildes)
  y reinicia el programa. Haz una copia antes. Para una acción nueva, además hay
  que implementarla en `actions` y, si corresponde, en `navigation`. Las frases
  muy cortas o parecidas entre sí aumentan los errores.
- **Modelo de voz:** `RUTA_MODELO` en `asr/vosk_asr.py`.
- **Umbral de confianza:** en `interpreter`. Más alto, menos falsos positivos
  pero más frases mal pronunciadas rechazadas.
- **Modo prueba:** `MODO_PRUEBA` en `actions/actions.py`. Ponlo en `False` solo
  si quieres que apagar y bloquear funcionen de verdad.
- **Carpeta de documentos:** `CARPETA_GUARDADO` en `navigation/navigation.py`.
  Con `None` se usa `Documentos\Mouth AI Live`.

El sistema es de un solo usuario por sesión de Windows, no maneja cuentas y no
necesita permisos de administrador. Para respaldar, copia `config/comandos.json`
y el CSV de métricas (puede contener las frases que dijo el usuario; guárdalo
con cuidado). Para desinstalar basta eliminar la carpeta.

## Resultados de las pruebas

Sobre 928 interacciones registradas en 12 sesiones (19 de agosto al 5 de octubre
de 2026):

| Indicador | Valor |
|---|---|
| Interacciones identificadas como comando | 202 |
| Ejecutadas con éxito | 180 (89,1 %) |
| Confianza promedio del reconocimiento | 0,97 |
| Tiempo de respuesta, mediana / percentil 90 | 1.637 ms / 3.931 ms |

Por categoría: búsqueda y navegador, navegación asistida, aplicaciones y control
de la voz, 100 %; dictado, edición y formato, 97,1 %; **lectura en voz alta,
36,8 %**; **mouse, 46,7 %**.

Cómo leerlos con cuidado:

- Las 726 interacciones sin comando mezclan habla del entorno y órdenes no
  reconocidas, así que el 21,8 % no es la proporción de órdenes acertadas.
- Los fallos de lectura vienen de aplicaciones que no exponen bien su
  información de accesibilidad. Los del mouse ocurrieron en una sola sesión,
  cuando la acción aún no estaba disponible, y falta repetir la prueba.
- El registro no guarda ruido, distancia al micrófono ni quién habló, y la
  evaluación con usuarios se limita a grabaciones de una persona con
  discapacidad. Falta medir usabilidad con un grupo mayor.

## Preprocesamiento de señal (no integrado)

`preprocessing/preprocessing.py` va entre la captura y el ASR y hace dos cosas:
detecta voz por energía RMS (umbral 300) y normaliza el volumen (pico a 8000,
con tope de 10 veces de amplificación). No incluye reducción de ruido ni
filtros. Está apagado porque no se ha calibrado con un micrófono real.

Riesgos si se activa tal cual: una voz muy débil (energía menor de 300) se
trataría como silencio, y descartar los bloques de silencio puede impedir que
Vosk detecte el final de las frases. Antes de integrarlo, mide la tasa de error
con y sin él, en el mismo ambiente y con las mismas frases. Para probarlo sin
micrófono: `python preprocessing/preprocessing.py`.

## Limitaciones

- Solo Windows y solo español. No hay versión móvil, web ni multiplataforma.
- Solo entiende los comandos definidos en `comandos.json`; no interpreta
  instrucciones abiertas en lenguaje natural.
- El reconocimiento depende del micrófono, el ruido y la pronunciación. No hay
  cancelación activa de ruido.
- La lectura depende de la accesibilidad que exponga cada aplicación o página.
- Las funciones de Word, Excel y PowerPoint requieren Office de escritorio.
- La escucha está siempre activa: puede reaccionar a conversaciones cercanas.

## Trabajo futuro

- Palabra de activación o modo "pulsar para hablar".
- Confirmación por voz antes de apagar o bloquear.
- Comando para cerrar aplicaciones.
- Mouse más rápido (cuadrícula numerada o movimiento continuo).
- Dictado continuo y comandos de puntuación.
- Mejorar la lectura en voz alta, incluida la de documentos completos.
- Calibrar e integrar el preprocesamiento y comparar Vosk con Whisper.
- Evaluar Excel y PowerPoint, y probar con más usuarios con discapacidad motriz
  y visual, con un instrumento estandarizado.

## Estructura del proyecto

```
Mouth-ai-live/
├── asr/               reconocimiento de voz (Vosk)
├── interpreter/       interpretación de comandos fijos y dinámicos
├── actions/           ejecución de acciones
├── navigation/        teclado, mouse, lectura, Word, Excel y PowerPoint
├── metrics/           registro y cálculo de métricas
├── gui/               interfaz gráfica (punto de entrada)
├── preprocessing/     detección de voz y normalización (sin integrar)
├── config/comandos.json
├── main_test.py
├── models/            (ignorado por git)
├── requirements.txt
└── .gitignore
```