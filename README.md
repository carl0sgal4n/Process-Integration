# Pinch Lab — Process Integration

Análisis pinch en Python con tablas, curvas de temperatura y entalpía, servicios
externos mínimos, una red ideal de intercambiadores con división de corrientes
y exportación a Excel con fórmulas y gráficos editables.

**[Guía completa en español: LEEME.md](LEEME.md)** — instalación, Google Colab,
explicación de los cálculos, unidades, uso del Excel y límites del modelo.

## Ejecutarlo en Windows

1. Instala [Python 3.11 o superior](https://www.python.org/downloads/).
2. Descarga y descomprime el proyecto; entra en la carpeta `Pinch_Lab`.
3. Abre **`ABRIR_WINDOWS.bat`** con doble clic. La primera ejecución instala
   las dependencias y necesita conexión a Internet.
4. Se abrirá la aplicación en <http://127.0.0.1:8765>. Mantén abierta la consola.
5. Edita las corrientes y el ΔT mínimo, calcula y descarga tu Excel.

También puedes abrir una terminal en la carpeta del proyecto:

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

Para detener el programa, pulsa Ctrl+C en la consola.

## Ejecutarlo en macOS o Linux

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

## Ejecutarlo sin instalar Python

La [guía](LEEME.md#opción-sin-instalar-python-google-colab) incluye las celdas
para [Google Colab](https://colab.research.google.com/). Allí se ejecuta el cálculo
y se descargan los resultados. La interfaz del navegador se ejecuta localmente
con `app.py`.

## Qué incluye

| Archivo | Contenido |
|---|---|
| `pinch.py` | Modelo, tablas, nueve gráficos y ejecución por terminal |
| `excel_export.py` | Excel con fórmulas y seis gráficos nativos |
| `app.py` e `interfaz.html` | Aplicación local para editar datos |
| `requirements.txt` | Dependencias de Python |
| `ABRIR_WINDOWS.bat` | Instalación e inicio en Windows |
| `corrientes_ejemplo.csv` | Datos de las cuatro corrientes del ejercicio |
| `Pinch_Analisis.xlsx` | Excel del caso de referencia |
| `resultados_ejemplo/` | Tablas, gráficos PNG/SVG, Excel y resumen JSON |
| `test_pinch.py` | Pruebas físicas y casos límite |
| `LEEME.md` | Explicación detallada |

## Caso de referencia

Con ΔT mínimo de 10 K:

| Resultado | Valor |
|---|---:|
| Calentamiento mínimo | 70.000 kJ/h |
| Refrigeración mínima | 60.000 kJ/h |
| Calor recuperado | 470.000 kJ/h |
| Temperaturas del pinch | 140 °C caliente / 130 °C fría |

El pinch se calcula a partir de las corrientes y de ΔT mínimo. En Excel puedes
cambiar temperaturas, CP y ΔT mínimo; para cambiar el número de corrientes,
exporta otro libro desde Python o la interfaz.

## Subir este proyecto a GitHub

1. Descomprime `Pinch_Lab_Python.zip` y abre la carpeta `Pinch_Lab`.
2. En tu repositorio, selecciona **Add file → Upload files**. Si está vacío,
   utiliza el enlace para subir un archivo existente.
3. Arrastra el contenido de `Pinch_Lab`, incluida la carpeta de ejemplos, de
   forma que `README.md`, `app.py` y `requirements.txt` queden en la raíz del
   repositorio. El código debe subirse descomprimido.
4. Escribe una descripción del cambio y completa el guardado que ofrece GitHub.

El repositorio guarda el código. Los pasos de ejecución anteriores ponen en
marcha la aplicación; una página de GitHub Pages por sí sola no ejecuta este
servidor Python.

## Validación y alcance

```bash
python -m unittest -v
```

Hay ocho pruebas automatizadas que incluyen el caso de referencia, cambios de
ΔT, sistemas sin recuperación y 100 casos aleatorios. Comprueban balances,
servicios mínimos y diferencias terminales de temperatura.

El modelo considera CP constante y calor sensible. La red usa divisiones y
recombinaciones de corrientes y alcanza los servicios mínimos del modelo; no
optimiza el número de intercambiadores ni el coste. La interfaz se ha comprobado
a nivel de cálculo y exportación, pero sigue pendiente su validación visual
completa en un navegador local. Consulta la guía para el alcance detallado.
