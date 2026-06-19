# Office Script Visualizer & Mapper (Gold Edition) ✦

Una herramienta web premium, diseñada con estética negra y dorada, para la inspección, análisis, depuración y mapeo de **Office Scripts (TypeScript)** utilizados en procesos contables de Excel y SAP.

## 🌟 Características Principales

1. **Analizador de Scripts ("El Recuadro Grande"):**
   - Carga directamente cualquiera de tus scripts contables preestablecidos.
   - Analiza el código en tiempo real para extraer hojas afectadas, encabezados requeridos, fila de inicio y complejidad.

2. **Disección Visual de Ejecución:**
   - Desglosa paso a paso el código TypeScript en tareas comprensibles (Inicialización, Mapeo, Bucle, Procesamiento).
   - Haz clic en cualquier fase para resaltar la línea de código exacta en el editor.

3. **Mapeador Dinámico de Encabezados:**
   - Si SAP actualiza sus formatos (ej. cambia de `"Nº doc."` a `"N/Doc"`), puedes remapear visualmente los encabezados y actualizar de forma automatizada tu script de TypeScript en un solo clic.

4. **Sugerencias de Optimización Inteligentes:**
   - Detecta bucles con filas de inicio estáticas (hardcodeadas) y ofrece parches automatizados para hacerlas dinámicas en base al escaneo automático del libro.

5. **Simulador de Hojas de Excel:**
   - Visualiza en cámara lenta cómo trabaja el script: busca encabezados, procesa filas, inyecta columnas contables calculadas (`POLIZA`, `FECHA`, `IVA`, `Cliente`) y colorea coincidencias en amarillo en una cuadrícula simulada.

---

## 🛠️ Estructura del Proyecto

```
office-script-visualizer/
├── index.html       # Estructura de la SPA y tabs
├── styles.css       # Estilos CSS premium (Gold & Dark)
├── app.js           # Lógica del simulador, parser y mapeo
└── presets.js       # Archivo de scripts cargados de tu equipo
```

## 🚀 Cómo Ejecutar la Aplicación

Dado que está construida con tecnologías nativas del navegador (HTML5, CSS3, ES6 JavaScript) sin dependencias ni compiladores:

1. **Localmente:** Simplemente haz doble clic en [index.html](index.html) para abrirlo en tu navegador Chrome, Edge o Firefox favorito.
2. **Servidor Local (opcional):** Si deseas ejecutar un servidor ligero en Python, ejecuta:
   ```bash
   python -m http.server 8000
   ```
   Luego ingresa a `http://localhost:8000`.

---
*Desarrollado con pasión para agilizar el trabajo contable de David y Benjamín.*