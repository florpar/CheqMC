# cheqMC  
**Chequeo de Z-Score y evaluación de rango para Material de Control**

`cheqMC` es una herramienta diseñada para laboratorios de análisis elemental que trabajan con **materiales de control (MC)**. El programa:

- Extrae valores de **FM Corr** e **Inc Corr** desde varios archivos `_resultados.xlsx`
- Compara los resultados contra un archivo certificado (`.txt`)
- Calcula **Z-Score**, **coincidencia entre intervalos** y **tolerancias ampliadas**
- Marca los valores:  
  - **Rojo** → Z-score > 2  
  - **Amarillo** → Sin intersección entre intervalo medido y certificado  
  - **Verde** → Mejor coincidencia dentro de ±3σ del certificado  
- Genera:
  - Un archivo Excel con resultados consolidados del material
  - Un Excel con formato y colores para análisis visual del cumplimiento del material control

Incluye una **interfaz gráfica en Tkinter**, selección automática de certificados sugeridos, procesamiento de múltiples archivos y normalización inteligente de nucleidos.

---

## 🚀 Características principales

- Lectura automática de múltiples archivos que terminen en `_resultados.xlsx`
- Identificación segura de las columnas “FM Corr / Inc Corr” mediante heurísticas
- Normalización robusta de nombres de nucleidos (`Co60`, `CO-60`, `co60m` → `Co60m`)
- Agrupación inteligente de energías (tolerancia ±5%)
- Consolidación de replicados en una sola fila por nucleido / tipo / energía / detector
- Cálculo de:
  - Intervalos medidos: FM ± Inc
  - Intervalos certificados: C ± δC
  - Intervalo ampliado: C ± 3δC
  - Z-Score según incertidumbres combinadas
- Formateo en Excel:
  - **Rojo:** Z-Score > 2  
  - **Amarillo:** Intervalos sin intersección  
  - **Verde:** Mejor candidato dentro de ±3σ  
- Sugerencia automática del archivo certificado basado en `codificacion.xlsx`
- Organización automática en carpeta `*_control_material`

---

## 📂 Estructura requerida de archivos

### 1. codificacion.xlsx

Debe estar en:

C:\Yaguarete\Standards\codificacion.xlsx

Columnas requeridas:

| sname | cert_file | humedad |
|-------|------------|---------|

Ejemplo:

| sname | cert_file | humedad |
|-------|------------|---------|
| 1633c | Coal-1633C | 5 |
| OTL1  | CTA-OTL-1  | 3 |

---

### 2. Archivos de resultados

Archivos generados por tu pipeline:

*_resultados.xlsx

El programa identifica dentro de la hoja **Mediciones Corregidas** la estructura:

Fila 0: nombres de archivo (A1573, A1574…)  
Fila 1: FM Corr / Inc Corr  
Fila con “Nucleido”: inicio de tabla  

---

### 3. Archivos certificados

Ubicados en:

C:\Yaguarete\Standards\*.txt

Formato:

Nuclido   C_standard   delta_C_standard  
Co60      12.3         0.9  
La140     40.8         1.5  

---

## 🖥️ Uso del programa

### 1. Ejecutar cheqMC

python cheqMC.py

Aparece la ventana principal:

"Chequeo Material Control"

---

### 2. Interfaz gráfica

#### a) Selección de material de control

- Se carga la lista desde codificacion.xlsx
- Al seleccionar:
  - Se actualiza la lista de certificados disponibles
  - Se sugiere el archivo correspondiente

#### b) Selección de archivo certificado

- Aparecen todos los `.txt` en C:\Yaguarete\Standards

#### c) Selección de carpeta con archivos

Debe contener:

A1573_resultados.xlsx  
MC_1633C_resultados.xlsx  
etc.

#### d) Botón “Generar Comparativo”

El programa solicita un **nombre base**, ej.:

Ensayo_Junio

Y genera:

Ensayo_Junio_control_material/

Con:

1. Ensayo_Junio_<material>_control.xlsx  (resultado consolidado)
2. Ensayo_Junio_<material>_rangos.xlsx   (Excel con colores y Z-score)

---

## 🔢 Lógica de procesamiento

### Normalización de nucleidos

Se convierte:

co-60 → Co60  
CO60M → Co60m  
co 60 → Co60  

Usando regex.

---

### Extracción segura de FM Corr / Inc Corr

Se detecta la fila con *Nucleido*  
Se analizan pares **FM Corr / Inc Corr**  
Se toma sólo el **material seleccionado**

---

### Agrupación de energías

Fotopicos del mismo nucleido se agrupan si:

ΔE / E ≤ 5%

---

### Cálculo de intervalos y Z-Score

Intervalo de medición:  
[FM - Inc , FM + Inc]

Intervalo certificado:  
[C – δC , C + δC]

Intervalo extendido (±3σ):  
[C – 3δC , C + 3δC]

Z-Score:  
z = |FM – C| / sqrt(Inc² + δC²)

---

## 🎨 Formato en Excel

Colores aplicados:

Z > 2 → rojo  
Sin solapamiento → amarillo  
Mejor dentro de ±3σ → verde + negrita  

---

## 🧩 Personalización

Todo puede modificarse:

- Textos de GUI  
- Iconos  
- Ruta base  
- Tolerancia de energía  
- Lógica de Z-score  
- Colores  
- Formato de salida  

---

## ⚠️ Limitaciones

- Los archivos resultados deben tener estructura estándar
- La detección de columnas depende del nombre del material en fila 0
- Funciona en Windows (usa os.startfile)
- Archivos certificados deben ser .txt con 3 columnas

---

## 📝 Licencia

Proyecto desarrollado por **Flor** para el control del Material Control en análisis elementales.

Libre para adaptar, modificar y ampliar según las necesidades del laboratorio.
