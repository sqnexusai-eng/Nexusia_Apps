# 🎯 Radar de Talento SQ Nexus IA

Plataforma integral de evaluación curricular algorítmica y comparativa impulsada por IA de última generación (`gemini-3.8-flash`), desarrollada con Python y Streamlit.

---

## 🌟 Características Principales

- **Arquitectura Limpia & Monolítica:** Todo el sistema concentrado en un único archivo ejecutable `app.py`, sin sobreingeniería.
- **Tema Visual SQ Nexus (Dark Mode):** Interfaz corporativa de alto contraste (`#0F172A`, `#1E293B`, acentos `#6366F1` y `#10B981`).
- **Matriz de Requisitos Ponderados:**
  - Tabla interactiva con edición en tiempo real (`st.data_editor`).
  - Validación dinámica de pesos: requiere totalizar exactamente el 100% para habilitar la evaluación.
  - Acciones de matriz: añadir requisito, selector para eliminar requisitos específicos y restaurar valores base.
- **Motor de Evaluación Curricular:**
  - Soporte de extracción de texto para currículums en formato **PDF** (`pypdf`) y **Word DOCX** (`python-docx`).
  - Evaluación analítica estructurada con salida en formato **JSON puro**.
  - Control de cuota automático (pausa preventiva de 12 segundos entre peticiones) para cumplir con el límite de peticiones de la capa gratuita.
- **Dashboard Analítico y Comparativo:**
  - **🕸️ Radar de Competencias (Gráfico de Araña / Telaraña):** Visualización polar interactiva con Plotly (`go.Scatterpolar`) con anillos poligonales concéntricos de 0 a 100.
  - **🏆 Ranking General:** Gráfico de barras horizontales coloreadas según veredicto (*Altamente Recomendado*, *Apto con Observaciones*, *No Recomendado*).
  - **📋 Cuadro de Clasificación (Leaderboard):** Tabla resumen con medallas (🥇, 🥈, 🥉) y estado.
  - **🔍 Fichas Cualitativas:** Pestañas individuales por candidato con desglose por criterio, cálculo de aportación en puntos, fortalezas destacadas y áreas de mejora.
  - **📥 Exportación:** Descarga de informes consolidados en **CSV** y **JSON**.

---

## 🚀 Instalación y Puesta en Marcha

### 1. Clonar el repositorio
```bash
git clone https://github.com/sqnexusai-eng/Nexusia_Apps.git
cd Nexusia_Apps
```

### 2. Crear y activar entorno virtual
```bash
python -m venv .venv

# En Windows:
.\.venv\Scripts\activate

# En Linux/macOS:
source .venv/bin/activate
```

### 3. Instalar dependencias
```bash
pip install streamlit google-generativeai pandas plotly pypdf python-docx
```
*(O con `uv`: `uv sync`)*

### 4. Ejecutar la aplicación
```bash
streamlit run app.py
```

La aplicación se abrirá en `http://localhost:8501`.

---

## 🔑 Configuración de Credenciales
Introduce tu clave de API gratuita de [Google AI Studio](https://aistudio.google.com/) en la barra lateral (`Gemini API Key`) y pulsa **Probar Conexión con IA**.

---

## 📁 Estructura del Proyecto

```text
├── app.py              # Aplicación web completa y motor de evaluación
├── pyproject.toml      # Configuración de dependencias del proyecto
├── sample_cvs/         # Currículums de demostración (PDF y DOCX)
│   ├── CV_Carlos_Mendoza_Senior_Operations.pdf
│   ├── CV_Laura_Gomez_Analista_Datos.docx
│   └── CV_Diego_Torres_Junior_IT.docx
├── README.md           # Documentación del proyecto
└── .gitignore          # Exclusiones de Git
```

---

© SQ Nexus IA | Tecnología de Evaluación Curricular Basada en LLMs
