# Pinch Lab — análisis pinch y redes de intercambio en Python

Programa preparado a partir de tus diapositivas 26–30 de la UPC. Los datos de las
cuatro corrientes y ΔT mínimo = 10 K proceden de las diapositivas 27 y 30.
La interfaz, los comentarios del código y esta explicación están en español.

## 1. Qué puedes hacer

- Introducir de 1 a 20 corrientes en la interfaz o en la exportación detallada.
  El cálculo Python también admite más corrientes; el límite del Excel controla
  el tamaño de su tabla de equipos, que crece aproximadamente como N³.
- Cambiar Tin, Tout, CP y ΔT mínimo. El tipo caliente/fría se deduce automáticamente.
- Calcular las cargas de cada corriente, la tabla de intervalos, la cascada,
  todos los nodos pinch y los servicios mínimos.
- Dibujar curvas individuales T–H, compuestas antes y después del ajuste,
  compuestas desplazadas, gran curva compuesta, temperaturas, sensibilidad,
  matriz de intercambios y red por tramos.
- Obtener una asignación de intercambios que alcanza la máxima recuperación,
  con temperaturas de entrada/salida, fracciones de rama, CP de cada rama,
  diferencias térmicas y balances.
- Estimar áreas si introduces U. No se inventa un valor de U.
- Descargar un Excel con fórmulas y **seis gráficos nativos editables**.
  Las tablas, las cargas de los equipos, la sensibilidad y los gráficos se
  recalculan al editar sus entradas. Los PNG/SVG son imágenes de cada ejecución.
- Guardar/abrir casos JSON y reimportar las entradas de un Excel modificado.

## 2. Dónde ejecutarlo

### Opción cómoda en Windows: interfaz en tu navegador

1. Instala **Python 3.11 o superior** desde <https://www.python.org/downloads/>.
2. Descomprime toda la carpeta `Pinch_Lab`. No ejecutes los archivos dentro del ZIP.
3. Abre `ABRIR_WINDOWS.bat` con doble clic.
4. La primera vez crea un entorno `.venv` e instala las dependencias. Necesita
   Internet para esa instalación.
5. Se abre <http://127.0.0.1:8765> en tu navegador. Mantén abierta la consola.
6. Edita la tabla, pulsa **Calcular proceso** y descarga el Excel.
7. Para cerrar el programa, pulsa Ctrl+C en la consola.

La aplicación funciona en tu ordenador. No necesita GitHub, API, cuenta de
Streamlit ni subir tus datos a un servidor.

Si prefieres hacerlo tú desde la terminal de VS Code o PowerShell, abre una
terminal en la carpeta descomprimida y ejecuta:

```powershell
py -3 -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe app.py
```

En macOS/Linux:

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python app.py
```

Si el puerto está ocupado: `python app.py --port 8766`.

### Opción sin instalar Python: Google Colab

Abre <https://colab.research.google.com/>, crea un cuaderno nuevo y ejecuta estas
celdas. Esta opción ejecuta el análisis y exporta los archivos; la interfaz local
del navegador se usa en tu ordenador, no en Colab.

**Celda 1 — subir y descomprimir el ZIP entregado:**

```python
from google.colab import files
import zipfile
uploaded = files.upload()  # Selecciona Pinch_Lab_Python.zip
zip_name = next(name for name in uploaded if name.endswith('.zip'))
with zipfile.ZipFile(zip_name) as z:
    z.extractall('/content/pinch_trabajo')
%cd /content/pinch_trabajo/Pinch_Lab
%pip install -r requirements.txt
```

**Celda 2 — calcular, mostrar tablas y exportar:**

```python
from pinch import Stream, analyse, tables, figures, export_all
from IPython.display import display
import matplotlib.pyplot as plt

# CP en kJ/(h·K); temperaturas en °C. Cambia aquí tus datos.
corrientes = [
    Stream('H1', cp=1000, tin=250, tout=120),
    Stream('H2', cp=4000, tin=200, tout=100),
    Stream('C3', cp=3000, tin=90, tout=150),
    Stream('C4', cp=6000, tin=130, tout=190),
]
resultado = analyse(corrientes, dtmin=10, u=0)
for nombre, tabla in tables(resultado).items():
    print(nombre)
    display(tabla)
for nombre, fig in figures(resultado).items():
    display(fig)
    plt.close(fig)
export_all(resultado, 'mis_resultados')
```

**Celda 3 — descargar:**

```python
files.download('mis_resultados/Pinch_Analisis.xlsx')
```

Para descargar todo, incluidos gráficos y CSV:

```python
import shutil
shutil.make_archive('mis_resultados', 'zip', 'mis_resultados')
files.download('mis_resultados.zip')
```

### Solo código, sin interfaz

```bash
python pinch.py
python pinch.py --csv corrientes_ejemplo.csv --dtmin 20 --out escenario_20K
python pinch.py --excel mi_excel_editado.xlsx --out desde_excel
python pinch.py --dtmin 10 --u 1800 --out con_area
```

En Windows puedes reemplazar `python` por `.venv\Scripts\python.exe`.
En el último ejemplo, U=1800 kJ/(h·m²·K)=0,5 kW/(m²·K) es **solo una hipótesis
ilustrativa**, no un dato del PDF. El valor predeterminado es U=0: área no calculada.

## 3. Resultados del caso de tus apuntes

| Magnitud | kJ/h | kW |
|---|---:|---:|
| Calor de las corrientes calientes | 530.000 | 147,222 |
| Demanda de las corrientes frías | 540.000 | 150,000 |
| Calentamiento externo mínimo QHmin | **70.000** | **19,444** |
| Refrigeración externa mínima QCmin | **60.000** | **16,667** |
| Calor recuperado entre corrientes | **470.000** | **130,556** |

Pinch caliente: **140 °C**. Pinch frío: **130 °C**. Temperatura desplazada:
**135 °C**. Diferencia mínima: **10 K**.

Aunque el déficit neto es solo 10.000 kJ/h, hacen falta 70.000 kJ/h de
calentamiento y 60.000 kJ/h de refrigeración: el nivel de temperatura limita
qué calor puede aprovecharse.

La asignación de la red de ramas calculada para este caso es:

| Donante → receptora | C3 (kJ/h) | C4 (kJ/h) | Refrigeración (kJ/h) |
|---|---:|---:|---:|
| H1 | 32.000 | 98.000 | 0 |
| H2 | 148.000 | 192.000 | 60.000 |
| Calentamiento externo | 0 | 70.000 | — |

Las celdas de esta matriz son **cargas agregadas**, no necesariamente un equipo
por pareja. El método conserva los tramos para explicar la construcción y
produce 22 intercambiadores de ramas en este caso. Algunos se pueden combinar
mediante un diseño posterior que respete temperaturas y puntos de división.
El programa no presenta ese número como un mínimo de equipos.

## 4. Explicación de los cálculos

### Cargas y unidades

`CP = caudal másico × calor específico`, no solo el calor específico.
`Q = CP × |Tin − Tout|`.

El PDF utiliza CP en kJ/(h·K), por lo que Q queda en kJ/h. La conversión a kW
es `Q/3600`. Una diferencia de 10 °C equivale a una diferencia de 10 K.

### Cascada y localización del pinch

Se desplazan las calientes `−ΔTmin/2` y las frías `+ΔTmin/2`. Después se
ordenan todas las temperaturas T* de mayor a menor. En cada intervalo:

```text
ΔH = (ΣCP caliente − ΣCP fría) × (T* superior − T* inferior)
residual siguiente = residual anterior + ΔH
QHmin = max(0, −mínimo residual bruto, incluido el nodo inicial 0)
residual ajustado = residual bruto + QHmin
QCmin = último residual ajustado
```

Los nodos de residual ajustado cero identifican el pinch. Se recuperan sus
temperaturas reales sumando/restando ΔTmin/2. **La temperatura del pinch no es
un dato independiente que se pueda fijar libremente.** Se modifica cambiando
las corrientes o ΔT mínimo. Puede haber varios pinches, una zona pinch o un
pinch de umbral en el extremo de la cascada.

### Curvas

- **Individuales:** cada corriente entre su temperatura inicial y final;
  el origen de entalpía es independiente para cada una.
- **Compuestas originales:** suma de CP activos e integración por temperatura;
  ambas parten de H=0. Esta representación no es todavía un acoplamiento factible.
- **Compuestas ajustadas:** la curva fría se desplaza horizontalmente QCmin.
  El solapamiento representa el calor recuperable, las colas los servicios.
- **Compuestas desplazadas:** mismo eje H, usando las temperaturas T*.
- **Gran compuesta (GCC):** residual ajustado frente a T*. Permite ver niveles
  térmicos de excedente/déficit.
- **Sensibilidad:** se repite la cascada para distintos ΔT mínimos.

En Python se omiten los segmentos de CP=0 en las curvas. Excel puede unir
verticalmente dos temperaturas con la misma entalpía cuando existe un hueco:
ese salto vertical representa **cero carga térmica**, no una corriente activa.

### Intercambios y fracciones de rama

Se unen todos los nodos de entalpía de las compuestas. Para cada tramo se
obtienen sus temperaturas calientes y frías en ambos extremos. Con las
corrientes activas:

```text
fh_i = CP_i / ΣCP calientes
fc_j = CP_j / ΣCP frías
Qij = ΔH_tramo × fh_i × fc_j
CP de la rama caliente i hacia j = CP_i × fc_j
CP de la rama fría j desde i = CP_j × fh_i
```

Las ramas de una misma corriente se vuelven a mezclar a la misma temperatura
al terminar el tramo. No se mezclan fluidos de corrientes distintas. La suma de
las cargas coincide con la capacidad térmica de cada corriente y se verifica
que **ambas diferencias terminales** de cada intercambiador cumplen ΔT mínimo.
En CP constante, comprobar los extremos es suficiente para el tramo lineal.

Es un diseño ideal realizable mediante ramas; alcanza QHmin/QCmin sin imponer
un máximo de equipos. No resuelve un problema de optimización económica ni
de mínima cantidad de intercambiadores. La tabla detallada indica los caudales
térmicos de las ramas para que no se interprete como una red sin divisiones.

### Servicios externos y área

Los tramos con solo curva caliente requieren refrigeración; los que tienen
solo curva fría requieren calentamiento. Sus cargas se desglosan por corriente.
Para seleccionar vapor, agua o refrigerante hace falta especificar sus
temperaturas y verificar la aproximación mínima. No se calculan sus áreas
sin esa información.

Para equipos de proceso, si se aporta U:

```text
ΔT1 = Th_entrada − Tc_salida
ΔT2 = Th_salida − Tc_entrada
ΔTlm = (ΔT1 − ΔT2) / ln(ΔT1/ΔT2)
A = Q / (U × ΔTlm)
```

Si ΔT1=ΔT2, ΔTlm toma ese valor. Si algún extremo es cero, el área ideal es
infinita. ΔTmin=0 es un límite termodinámico, no un diseño de área finita en
el pinch. No se usa la media aritmética general de las diferencias térmicas.

## 5. Cómo se modifica el Excel

1. Abre `Pinch_Analisis.xlsx` en Excel de escritorio.
2. En **Entradas**, edita B4 (ΔT mínimo), B5 (U opcional) y A11:D14 (el caso
   inicial tiene cuatro corrientes). Los valores editables tienen fondo azul.
3. **Resumen** muestra servicios y seis gráficos nativos. **Cascada** muestra
   todos los nodos pinch. **Matriz_Q** contiene la asignación agregada.
4. **Equipos** contiene todas las combinaciones posibles, incluidas las de
   carga cero, para que un cambio de datos pueda activar parejas nuevas.
   Filtra la columna Q por valores mayores que cero. Tras cambiar entradas,
   vuelve a aplicar ese filtro: Excel no reaplica automáticamente los filtros.
5. **Tramos_red** muestra las temperaturas y fracciones que alimentan equipos.
6. **Sensibilidad** permite cambiar los valores de ΔT de los escenarios.
7. No insertes corrientes o filas de cálculo a mano. Para aumentar o disminuir
   su número, vuelve a la interfaz/Python y exporta un libro nuevo.

Las fórmulas y gráficos se actualizan en un programa de hojas de cálculo con
recálculo. Si no ocurre, activa Fórmulas → Opciones de cálculo → Automático,
o pulsa Ctrl+Alt+F9 en Excel. Una vista previa de archivos no equivale a Excel.
Los valores de caché se incluyen para que la primera apertura ya tenga datos.

Puedes leer nuevamente las entradas del libro:

```python
from pinch import read_streams, analyse, export_all
streams, dt, u = read_streams('mi_excel_editado.xlsx')
resultado = analyse(streams, dt, u)
export_all(resultado, 'resultado_actualizado')
```

La reimportación lee solo las entradas y los parámetros; no confía en la
caché de fórmulas. Respeta el nombre `Entradas` y la posición de su tabla.

## 6. Archivos del proyecto

| Archivo | Función |
|---|---|
| `pinch.py` | Cálculo, tablas pandas, gráficos y ejecución por terminal |
| `excel_export.py` | Excel con fórmulas, tablas, sensibilidad y gráficos editables |
| `app.py` | Servidor local en Python para la interfaz |
| `interfaz.html` | Pantalla de edición y visualización |
| `requirements.txt` | Dependencias |
| `ABRIR_WINDOWS.bat` | Inicio para Windows |
| `corrientes_ejemplo.csv` | Datos del PDF |
| `test_pinch.py` | Pruebas de balances, temperaturas y casos límite |
| `resultados_ejemplo/` | Excel, gráficos PNG/SVG, CSV y JSON del caso de referencia |

## 7. Validación y alcance

Ejecuta `python -m unittest -v`. Las pruebas contrastan el caso del PDF,
cambios de ΔT, cambios de tipo de corriente, huecos de temperatura, ausencia
de recuperación, sistemas con solo calientes/frías, zona pinch, ΔT=0, datos
inválidos y 100 problemas aleatorios. Los objetivos se contrastan con un
balance acumulado independiente de la implementación de la cascada.

El modelo supone régimen estacionario, CP constante, calor sensible, ausencia
de pérdidas y posibilidad de dividir y recombinar cada corriente. No incluye
calores latentes, restricciones de contacto entre fluidos, ensuciamiento,
presión, costes, múltiples niveles de utilidad, bombas de calor o CHP.
Las pruebas no convierten el modelo académico en un diseño industrial detallado.

La hoja original se genera con Python/XlsxWriter porque el entregable solicitado
es un programa Python reproducible. Fórmulas y resultados se verifican también
con un motor de hojas de cálculo; una comprobación en LibreOffice no sustituye
una validación visual en tu versión exacta de Microsoft Excel.

La interfaz se comprobó mediante sus funciones de cálculo/exportación y revisión
de JavaScript. El navegador remoto de este entorno bloquea localhost, por lo que
no se pudo hacer una prueba visual completa de la interfaz en un navegador.

## 8. Referencias

- **Fuente del ejercicio:** PDF aportado por el usuario,
  `P2_260921_Energy_Integration_Pinch_Analysis_V1_Lluis_26_27 (1).pdf`,
  diapositivas 26–30, especialmente 27 y 30.
- NPTEL, *Pinch Point Technology*, curvas compuestas y objetivos energéticos:
  <https://archive.nptel.ac.in/content/storage2/courses/112103016/module1/lec11/3.html>
- XlsxWriter, almacenamiento de fórmulas y cachés:
  <https://xlsxwriter.readthedocs.io/working_with_formulas.html>
- XlsxWriter, gráficos nativos:
  <https://xlsxwriter.readthedocs.io/working_with_charts.html>
- Python, instalación:
  <https://www.python.org/downloads/>
