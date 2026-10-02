import streamlit as st
import google.generativeai as genai
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pypdf import PdfReader
from docx import Document
import json
import re
import io
import os
import time

# ==============================================================================
# 1. CONFIGURACIÓN VISUAL Y TEMA (DARK MODE SQ NEXUS)
# ==============================================================================
st.set_page_config(
    page_title="Radar de Talento SQ Nexus IA",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

CUSTOM_CSS = """
<style>
    /* Paleta SQ Nexus Dark Mode */
    :root {
        --bg-main: #0F172A;
        --bg-card: #1E293B;
        --bg-card-hover: #27354A;
        --border-color: #334155;
        --text-primary: #F8FAFC;
        --text-secondary: #94A3B8;
        --primary-indigo: #4F46E5;
        --primary-indigo-light: #6366F1;
        --accent-emerald: #10B981;
        --accent-amber: #F59E0B;
        --accent-rose: #EF4444;
    }

    .stApp {
        background-color: var(--bg-main);
        color: var(--text-primary);
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
    }

    [data-testid="stSidebar"] {
        background-color: #0b1120;
        border-right: 1px solid var(--border-color);
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] span {
        color: var(--text-primary);
    }

    .sq-card {
        background-color: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 12px;
        padding: 1.25rem;
        margin-bottom: 1.25rem;
        box-shadow: 0 4px 6px -1px rgba(0, 0, 0, 0.2), 0 2px 4px -2px rgba(0, 0, 0, 0.2);
    }

    .sq-header-card {
        background: linear-gradient(135deg, rgba(30, 41, 59, 0.9) 0%, rgba(15, 23, 42, 0.95) 100%);
        border: 1px solid var(--border-color);
        border-left: 4px solid var(--primary-indigo-light);
        border-radius: 12px;
        padding: 1.5rem;
        margin-bottom: 1.5rem;
    }

    .sq-metric-box {
        background-color: var(--bg-card);
        border: 1px solid var(--border-color);
        border-radius: 10px;
        padding: 1rem;
        text-align: center;
    }
    .sq-metric-value {
        font-size: 1.85rem;
        font-weight: 700;
        color: #FFFFFF;
        margin: 0.25rem 0;
    }
    .sq-metric-label {
        font-size: 0.85rem;
        color: var(--text-secondary);
        text-transform: uppercase;
        letter-spacing: 0.05em;
        font-weight: 600;
    }

    .sq-badge {
        display: inline-block;
        padding: 0.35rem 0.8rem;
        font-size: 0.85rem;
        font-weight: 600;
        border-radius: 9999px;
        text-align: center;
    }
    .sq-badge-green {
        background-color: rgba(16, 185, 129, 0.2);
        color: #34D399;
        border: 1px solid rgba(16, 185, 129, 0.4);
    }
    .sq-badge-amber {
        background-color: rgba(245, 158, 11, 0.2);
        color: #FBBF24;
        border: 1px solid rgba(245, 158, 11, 0.4);
    }
    .sq-badge-red {
        background-color: rgba(239, 68, 68, 0.2);
        color: #F87171;
        border: 1px solid rgba(239, 68, 68, 0.4);
    }

    div.stButton > button {
        background: linear-gradient(135deg, #4F46E5 0%, #6366F1 100%) !important;
        color: #FFFFFF !important;
        border: 1px solid #6366F1 !important;
        border-radius: 8px !important;
        padding: 0.55rem 1.25rem !important;
        font-weight: 600 !important;
        box-shadow: 0 4px 14px 0 rgba(79, 70, 229, 0.35) !important;
        transition: all 0.2s ease-in-out !important;
    }
    div.stButton > button:hover {
        background: linear-gradient(135deg, #4338CA 0%, #4F46E5 100%) !important;
        border-color: #818CF8 !important;
        box-shadow: 0 6px 20px 0 rgba(79, 70, 229, 0.5) !important;
        transform: translateY(-1px) !important;
    }
    div.stButton > button:disabled {
        background: #334155 !important;
        border: 1px solid #475569 !important;
        color: #94A3B8 !important;
        cursor: not-allowed !important;
        box-shadow: none !important;
        transform: none !important;
    }

    .stTextInput input, .stTextArea textarea, .stSelectbox [data-baseweb="select"] {
        background-color: var(--bg-card) !important;
        color: var(--text-primary) !important;
        border: 1px solid var(--border-color) !important;
        border-radius: 8px !important;
    }
    .stTextInput input:focus, .stTextArea textarea:focus {
        border-color: var(--primary-indigo-light) !important;
        box-shadow: 0 0 0 2px rgba(99, 102, 241, 0.25) !important;
    }

    .sq-alert-success {
        background-color: rgba(16, 185, 129, 0.12);
        border: 1px solid rgba(16, 185, 129, 0.35);
        color: #A7F3D0;
        padding: 0.85rem 1.15rem;
        border-radius: 8px;
        font-size: 0.95rem;
    }
    .sq-alert-warning {
        background-color: rgba(245, 158, 11, 0.12);
        border: 1px solid rgba(245, 158, 11, 0.35);
        color: #FDE68A;
        padding: 0.85rem 1.15rem;
        border-radius: 8px;
        font-size: 0.95rem;
    }

    hr {
        border-color: var(--border-color) !important;
    }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# ==============================================================================
# 2. FUNCIONES DE EXTRACCIÓN DE TEXTO Y PARSEO DE RESPUESTA
# ==============================================================================
def extract_text_from_pdf(file_bytes: bytes) -> str:
    """Extrae texto completo de todas las páginas de un archivo PDF usando pypdf."""
    try:
        reader = PdfReader(io.BytesIO(file_bytes))
        pages_text = []
        for page in reader.pages:
            t = page.extract_text()
            if t:
                pages_text.append(t.strip())
        return "\n\n".join(pages_text).strip()
    except Exception as e:
        return f"[Error de extracción PDF: {str(e)}]"


def extract_text_from_docx(file_bytes: bytes) -> str:
    """Extrae texto de párrafos y celdas de tablas de un documento DOCX."""
    try:
        doc = Document(io.BytesIO(file_bytes))
        parts = []
        for p in doc.paragraphs:
            if p.text.strip():
                parts.append(p.text.strip())
        for table in doc.tables:
            for row in table.rows:
                cells = [c.text.strip() for c in row.cells if c.text.strip()]
                if cells:
                    parts.append(" | ".join(cells))
        return "\n".join(parts).strip()
    except Exception as e:
        return f"[Error de extracción DOCX: {str(e)}]"


def parse_gemini_json_response(raw_text: str) -> dict:
    """Extrae y parsea el objeto JSON de la respuesta generada por Gemini."""
    clean_text = raw_text.strip()
    if clean_text.startswith("```"):
        clean_text = re.sub(r"^```(?:json)?\s*", "", clean_text)
        clean_text = re.sub(r"\s*```$", "", clean_text)

    match = re.search(r"(\{[\s\S]*\})", clean_text)
    if match:
        clean_text = match.group(1)

    return json.loads(clean_text)


# ==============================================================================
# 3. GESTIÓN DEL ESTADO DE SESIÓN (SESSION STATE)
# ==============================================================================
DEFAULT_REQUIREMENTS = [
    {
        "Requisito Principal": "Python y Pandas",
        "Criterio / Nivel Esperado": "Experiencia sólida en scripts, procesamiento de datos y automatización.",
        "Peso (%)": 35
    },
    {
        "Requisito Principal": "Bases de Datos y SQL",
        "Criterio / Nivel Esperado": "Consultas avanzadas, optimización de queries y diseño relacional.",
        "Peso (%)": 25
    },
    {
        "Requisito Principal": "Optimización de Procesos",
        "Criterio / Nivel Esperado": "Metodologías de mejora continua, resolución de problemas y eficiencia.",
        "Peso (%)": 25
    },
    {
        "Requisito Principal": "Habilidades Blandas y Liderazgo",
        "Criterio / Nivel Esperado": "Trabajo en equipo, comunicación efectiva y orientación al logro.",
        "Peso (%)": 15
    }
]

if "requirements_df" not in st.session_state:
    st.session_state["requirements_df"] = pd.DataFrame(DEFAULT_REQUIREMENTS)

if "req_version" not in st.session_state:
    st.session_state["req_version"] = 0

if "demo_cvs" not in st.session_state:
    st.session_state["demo_cvs"] = []

if "evaluation_results" not in st.session_state:
    st.session_state["evaluation_results"] = None


# ==============================================================================
# 4. BARRA LATERAL (CONFIGURACIÓN SQ NEXUS)
# ==============================================================================
with st.sidebar:
    st.markdown("## Configuración SQ Nexus")
    st.markdown("---")

    # Sección Credenciales IA
    st.markdown("### 🔑 Credenciales IA")
    api_key_default = os.getenv("GEMINI_API_KEY", "")
    api_key = st.text_input(
        "Gemini API Key",
        value=st.session_state.get("gemini_api_key", api_key_default),
        type="password",
        help="Introduce tu clave de API gratuita de Google AI Studio.",
        key="gemini_api_key_input"
    )
    if api_key:
        st.session_state["gemini_api_key"] = api_key.strip()

    st.markdown(
        "[Obtener API Key gratis en Google AI Studio](https://aistudio.google.com/)",
        unsafe_allow_html=True
    )

    # Etiqueta informativa fija
    st.info("⚡ **Motor activo: gemini-3.8-flash (Capa Gratuita Oficial)**")

    # Botón Probar Conexión con IA
    if st.button("Probar Conexión con IA", use_container_width=True):
        current_key = st.session_state.get("gemini_api_key", "").strip()
        if not current_key:
            st.error("❌ Debe ingresar una Gemini API Key antes de probar la conexión.")
        else:
            with st.spinner("Probando conexión con gemini-3.8-flash..."):
                try:
                    genai.configure(api_key=current_key)
                    model = genai.GenerativeModel(
                        "gemini-3.8-flash",
                        generation_config={
                            "temperature": 0.1,
                            "response_mime_type": "application/json"
                        }
                    )
                    test_response = model.generate_content("ping")
                    if test_response and test_response.text:
                        st.success("✅ Conexión exitosa con gemini-3.8-flash (Capa Gratuita Oficial)")
                    else:
                        st.warning("⚠️ El modelo respondió sin contenido.")
                except Exception as e:
                    st.error(f"Error técnico de Google API: {e}")

    st.markdown("---")

    # Sección Carga de CVs
    st.markdown("### 📄 Carga de Currículums")
    uploaded_files = st.file_uploader(
        "Sube archivos CV (.pdf, .docx)",
        accept_multiple_files=True,
        type=["pdf", "docx"],
        help="Arrastra o selecciona currículums en formato PDF o Word DOCX."
    )

    col_demo, col_clear = st.columns(2)
    with col_demo:
        if st.button("🧪 Cargar CVs Demo", use_container_width=True, help="Carga los currículums de prueba en sample_cvs/"):
            demo_folder = "sample_cvs"
            if os.path.exists(demo_folder):
                loaded_demos = []
                for fname in os.listdir(demo_folder):
                    ext = os.path.splitext(fname)[1].lower()
                    if ext in [".pdf", ".docx"]:
                        fpath = os.path.join(demo_folder, fname)
                        with open(fpath, "rb") as f:
                            loaded_demos.append({
                                "name": fname,
                                "bytes": f.read(),
                                "ext": ext
                            })
                if loaded_demos:
                    st.session_state["demo_cvs"] = loaded_demos
                    st.success(f"🧪 {len(loaded_demos)} CVs Demo cargados.")
                else:
                    st.warning("⚠️ No se encontraron archivos .pdf o .docx en sample_cvs/")
            else:
                st.error("❌ No existe la carpeta sample_cvs/ en el directorio.")

    with col_clear:
        if st.button("🗑️ Limpiar", use_container_width=True, help="Elimina los currículums y los resultados actuales"):
            st.session_state["demo_cvs"] = []
            st.session_state["evaluation_results"] = None
            st.rerun()

    # Consolidar lista de CVs (Uploads + Demo)
    available_cvs = []
    seen_names = set()

    for demo in st.session_state.get("demo_cvs", []):
        available_cvs.append(demo)
        seen_names.add(demo["name"])

    if uploaded_files:
        for uf in uploaded_files:
            if uf.name not in seen_names:
                ext = os.path.splitext(uf.name)[1].lower()
                available_cvs.append({
                    "name": uf.name,
                    "bytes": uf.getvalue(),
                    "ext": ext
                })
                seen_names.add(uf.name)

    if available_cvs:
        st.markdown(f"**Currículums listos ({len(available_cvs)}):**")
        for cv in available_cvs:
            icon = "📕" if cv["ext"] == ".pdf" else "📘"
            size_kb = len(cv["bytes"]) / 1024
            st.caption(f"{icon} `{cv['name']}` ({size_kb:.1f} KB)")
    else:
        st.caption("No hay currículums cargados aún.")

    st.markdown("---")
    st.caption("Radar de Talento SQ Nexus IA v1.0 | gemini-3.8-flash (Capa Gratuita Oficial)")


# ==============================================================================
# 5. PANEL PRINCIPAL
# ==============================================================================

# Banner Superior
st.markdown(
    """
    <div class="sq-header-card">
        <h1 style="margin: 0; font-size: 2.2rem; font-weight: 800; color: #FFFFFF; display: flex; align-items: center; gap: 0.6rem;">
            🎯 Radar de Talento SQ Nexus IA
        </h1>
        <p style="margin: 0.5rem 0 0 0; font-size: 1.05rem; color: #94A3B8;">
            Plataforma de evaluación curricular algorítmica y comparativa impulsada por IA de última generación.
        </p>
    </div>
    """,
    unsafe_allow_html=True
)

# Formulario de la Vacante (2 Columnas)
st.markdown("### 📌 Definición de la Vacante")
col_pos, col_desc = st.columns([1, 2])

DEFAULT_JOB_DESCRIPTION = (
    "Buscamos un Analista de Operaciones Senior para liderar la optimización de procesos operativos "
    "y la toma de decisiones basada en datos.\n\n"
    "Responsabilidades Principales:\n"
    "- Analizar grandes volúmenes de datos operativos y de negocio utilizando Python (Pandas) y SQL relacional.\n"
    "- Diseñar e implementar pipelines de automatización y conciliación de información entre sistemas.\n"
    "- Diagnosticar ineficiencias en flujos operativos y aplicar metodologías de mejora continua (Lean / Six Sigma).\n"
    "- Presentar indicadores de gestión (KPIs) a la dirección y liderar iniciativas interdisciplinarias.\n\n"
    "Requisitos Clave:\n"
    "- Experiencia de 4+ años en analítica operativa y modelado de datos.\n"
    "- Dominio comprobable de Python (Pandas), bases de datos SQL y gestión de proyectos de eficiencia."
)

with col_pos:
    job_title = st.text_input(
        "📌 Puesto de Trabajo",
        value="Analista de Operaciones Senior",
        help="Indica el título oficial del cargo a evaluar."
    )

with col_desc:
    job_description = st.text_area(
        "📝 Descripción del Puesto (Responsabilidades y Perfil)",
        value=DEFAULT_JOB_DESCRIPTION,
        height=180,
        help="Detalla las funciones clave, herramientas requeridas y el perfil ideal para la posición."
    )

# Matriz de Requisitos Ponderados
st.markdown("### ⚖️ Matriz de Requisitos Ponderados")
st.caption("Edita los requisitos, criterios esperados y pesos porcentuales. La suma debe totalizar exactamente 100%.")

edited_df = st.data_editor(
    st.session_state["requirements_df"],
    num_rows="dynamic",
    use_container_width=True,
    key=f"requirements_data_editor_{st.session_state['req_version']}",
    column_config={
        "Requisito Principal": st.column_config.TextColumn(
            "Requisito Principal",
            help="Habilidad técnica, competencia o área de evaluación.",
            required=True
        ),
        "Criterio / Nivel Esperado": st.column_config.TextColumn(
            "Criterio / Nivel Esperado",
            help="Profundidad o evidencia requerida en el currículum.",
            required=True
        ),
        "Peso (%)": st.column_config.NumberColumn(
            "Peso (%)",
            help="Ponderación porcentual (0 a 100).",
            min_value=0,
            max_value=100,
            step=5,
            required=True
        )
    }
)

st.session_state["requirements_df"] = edited_df

# Acciones de la Matriz (Agregar, selector para eliminar, restaurar)
st.markdown("#### 🛠️ Acciones de la Matriz:")

current_requirements = st.session_state["requirements_df"]["Requisito Principal"].dropna().tolist()
current_requirements = [r.strip() for r in current_requirements if str(r).strip() != ""]

col_add, col_sel, col_del, col_reset = st.columns([1.5, 2.5, 1.8, 1.3])

with col_add:
    if st.button("➕ Agregar Requisito", use_container_width=True):
        new_row = pd.DataFrame([{
            "Requisito Principal": f"Nuevo Requisito {len(current_requirements) + 1}",
            "Criterio / Nivel Esperado": "Criterio de evaluación específico a medir en el CV",
            "Peso (%)": 10
        }])
        st.session_state["requirements_df"] = pd.concat([st.session_state["requirements_df"], new_row], ignore_index=True)
        st.session_state["req_version"] += 1
        st.rerun()

with col_sel:
    req_to_delete = st.selectbox(
        "Seleccionar requisito a eliminar",
        options=current_requirements if current_requirements else ["(Sin requisitos)"],
        label_visibility="collapsed",
        key="req_delete_selector"
    )

with col_del:
    if st.button("🗑️ Eliminar Requisito", use_container_width=True, disabled=len(current_requirements) <= 1):
        if len(st.session_state["requirements_df"]) > 1:
            st.session_state["requirements_df"] = st.session_state["requirements_df"][
                st.session_state["requirements_df"]["Requisito Principal"] != req_to_delete
            ].reset_index(drop=True)
            st.session_state["req_version"] += 1
            st.rerun()
        else:
            st.warning("Debe haber al menos un requisito en la matriz.")

with col_reset:
    if st.button("🔄 Restaurar", use_container_width=True):
        st.session_state["requirements_df"] = pd.DataFrame(DEFAULT_REQUIREMENTS)
        st.session_state["req_version"] += 1
        st.rerun()

# Validador de suma en tiempo real (métrica + alerta)
try:
    weights_series = pd.to_numeric(
        st.session_state["requirements_df"]["Peso (%)"].astype(str).str.replace("%", "").str.strip(),
        errors="coerce"
    ).fillna(0)
    total_weight = float(weights_series.sum())
except Exception:
    total_weight = 0.0

is_weight_valid = abs(total_weight - 100.0) < 0.01

st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)
col_sum_metric, col_sum_alert = st.columns([1.2, 3.8])

with col_sum_metric:
    if is_weight_valid:
        tag_html = '<div style="font-size: 0.82rem; color: #34D399; font-weight: 600; margin-top: 0.35rem; display: flex; align-items: center; gap: 0.2rem;">↑ Exacto (100%)</div>'
    else:
        tag_html = f'<div style="font-size: 0.82rem; color: #FBBF24; font-weight: 600; margin-top: 0.35rem; display: flex; align-items: center; gap: 0.2rem;">⚠️ Requiere 100% (Actual: {total_weight:.1f}%)</div>'

    st.markdown(
        f"""
        <div style="background-color: #1E293B; border: 1px solid #334155; border-radius: 10px; padding: 0.85rem 1.15rem; min-height: 95px;">
            <div style="font-size: 0.85rem; color: #94A3B8; font-weight: 500;">Suma Total de Pesos</div>
            <div style="font-size: 2.1rem; font-weight: 800; color: #FFFFFF; line-height: 1.1; margin-top: 0.2rem;">{total_weight:.1f}%</div>
            {tag_html}
        </div>
        """,
        unsafe_allow_html=True
    )

with col_sum_alert:
    if is_weight_valid:
        st.markdown(
            """
            <div class="sq-alert-success" style="min-height: 95px; display: flex; align-items: center; margin: 0;">
                <div>
                    ✅ <strong>Ponderación balanceada:</strong> La suma de pesos es exactamente 100%. Todo listo para evaluar.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            f"""
            <div class="sq-alert-warning" style="min-height: 95px; display: flex; align-items: center; margin: 0;">
                <div>
                    ⚠️ <strong>La suma actual es {total_weight:.1f}%.</strong> Debe ser exactamente 100% para poder evaluar. Ajusta los pesos en la matriz superior.
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


# ==============================================================================
# 6. MOTOR DE PROCESAMIENTO Y EVALUACIÓN (CON RATE LIMIT DE SEGURIDAD)
# ==============================================================================
st.markdown("---")

col_eval_btn, col_eval_msg = st.columns([1.8, 3.2])

with col_eval_btn:
    can_evaluate = is_weight_valid and len(available_cvs) > 0
    evaluate_button = st.button(
        "🚀 Evaluar Candidatos",
        disabled=not can_evaluate,
        use_container_width=True,
        help="Ejecuta la evaluación algorítmica con gemini-3.8-flash aplicando control de cuota (12s entre CVs)."
    )

with col_eval_msg:
    if not available_cvs:
        st.caption("ℹ️ *Carga currículums o pulsa '🧪 Cargar CVs Demo' en la barra lateral para evaluar.*")
    elif not is_weight_valid:
        st.caption("ℹ️ *Asegúrate de que la matriz de requisitos sume exactamente el 100% para habilitar la evaluación.*")

if evaluate_button:
    current_key = st.session_state.get("gemini_api_key", "").strip()

    if not current_key:
        st.error("Error técnico de Google API: Se requiere Gemini API Key en la barra lateral para autenticar la petición.")
    elif len(available_cvs) == 0:
        st.error("Error: No hay currículums para evaluar.")
    elif not is_weight_valid:
        st.error(f"Error: La ponderación suma {total_weight:.1f}%. Debe ser exactamente 100%.")
    else:
        # Inicialización del modelo gemini-3.8-flash
        try:
            genai.configure(api_key=current_key)
            model = genai.GenerativeModel(
                "gemini-3.8-flash",
                generation_config={
                    "temperature": 0.1,
                    "response_mime_type": "application/json"
                }
            )
        except Exception as e:
            st.error(f"Error técnico de Google API: {e}")
            st.stop()

        # Preparar especificación de requisitos en texto para el prompt
        req_lines = []
        for _, row in st.session_state["requirements_df"].iterrows():
            req_lines.append(
                f"- Requisito: {row['Requisito Principal']} | Peso: {row['Peso (%)']}% | Criterio Esperado: {row['Criterio / Nivel Esperado']}"
            )
        matriz_text = "\n".join(req_lines)

        progress_bar = st.progress(0.0)
        status_box = st.empty()
        results_list = []
        total_cvs = len(available_cvs)

        for idx, cv_item in enumerate(available_cvs):
            cv_filename = cv_item["name"]
            cv_bytes = cv_item["bytes"]
            cv_ext = cv_item["ext"]

            # Control de Cuota (Rate Limit de seguridad para evitar 429 en Capa Gratuita)
            if idx > 0:
                for remaining in range(12, 0, -1):
                    status_box.info(
                        f"⏳ **Control de Cuota (gemini-3.8-flash):** Esperando {remaining}s antes de evaluar `{cv_filename}` para asegurar respeto a la tasa de peticiones..."
                    )
                    time.sleep(1)

            step_pct = (idx) / total_cvs
            progress_bar.progress(step_pct)
            status_box.markdown(f"🤖 **[{idx+1}/{total_cvs}] Evaluando con gemini-3.8-flash:** `{cv_filename}`...")

            # 1. Extracción de texto
            if cv_ext == ".pdf":
                cv_text = extract_text_from_pdf(cv_bytes)
            else:
                cv_text = extract_text_from_docx(cv_bytes)

            if not cv_text or len(cv_text.strip()) < 20:
                cv_text = f"[Contenido insuficiente o no extraíble para {cv_filename}]"

            # 2. Construcción de prompt estructurado exigiendo JSON
            prompt = f"""Eres un Evaluador Senior de Talento y Selección Técnica altamente riguroso.
Tu tarea es evaluar detalladamente el currículum vitae de un candidato para la siguiente vacante corporativa.

PUESTO VACANTE:
{job_title}

DESCRIPCIÓN DEL PUESTO:
{job_description}

MATRIZ DE REQUISITOS PONDERADOS:
{matriz_text}

CURRÍCULUM DEL CANDIDATO (Archivo: {cv_filename}):
{cv_text}

INSTRUCCIONES CLAVE DE EVALUACIÓN:
1. Extrae el nombre completo del candidato. Si no se puede determinar fehacientemente, usa "{cv_filename}".
2. Para CADA requisito de la matriz, asigna un puntaje de 0 a 100 ('puntaje_asignado') basado estrictamente en las evidencias reales presentes en el currículum.
3. Proporciona una justificación concisa y factual de 1 o 2 frases ('justificacion') fundamentada en los hechos del CV.
4. Calcula 'puntuacion_total' como la suma ponderada exacta: suma de (puntaje_asignado * (peso / 100)).
5. Identifica de 2 a 4 fortalezas clave ('fortalezas') relevantes para la vacante.
6. Identifica de 2 a 4 áreas de mejora o debilidades ('areas_mejora') detectadas en el CV.
7. Determina el 'veredicto' final estrictamente según esta regla:
   - "Altamente Recomendado" (si puntuacion_total >= 80)
   - "Apto con Observaciones" (si 60 <= puntuacion_total < 80)
   - "No Recomendado" (si puntuacion_total < 60)

RESPONDE EXCLUSIVAMENTE CON UN OBJETO JSON VÁLIDO CON ESTA ESTRUCTURA EXACTA (SIN TEXTO EXTRA, SIN EXPLICACIONES FUERA DEL JSON):
{{
  "candidato": "Nombre completo extraído del CV",
  "puntuacion_total": 85.5,
  "desglose": [
    {{
      "requisito": "Nombre exacto del requisito",
      "peso": 35,
      "puntaje_asignado": 90,
      "justificacion": "Justificación concisa basada en hechos del CV"
    }}
  ],
  "fortalezas": ["Punto fuerte 1", "Punto fuerte 2"],
  "areas_mejora": ["Debilidad o vacío 1", "Debilidad o vacío 2"],
  "veredicto": "Altamente Recomendado"
}}"""

            try:
                response = model.generate_content(prompt)
                parsed = parse_gemini_json_response(response.text)

                # Re-validación matemática del puntaje total
                calculated_total = 0.0
                desglose = parsed.get("desglose", [])
                for item in desglose:
                    w = float(item.get("peso", 0))
                    s = float(item.get("puntaje_asignado", 0))
                    calculated_total += s * (w / 100.0)

                parsed["puntuacion_total"] = round(calculated_total, 1)

                sc = parsed["puntuacion_total"]
                if sc >= 80.0:
                    parsed["veredicto"] = "Altamente Recomendado"
                elif sc >= 60.0:
                    parsed["veredicto"] = "Apto con Observaciones"
                else:
                    parsed["veredicto"] = "No Recomendado"

                parsed["archivo"] = cv_filename
                results_list.append(parsed)

            except Exception as e:
                st.error(f"Error técnico de Google API al procesar `{cv_filename}`: {e}")
                error_entry = {
                    "candidato": cv_filename,
                    "archivo": cv_filename,
                    "puntuacion_total": 0.0,
                    "desglose": [],
                    "fortalezas": [],
                    "areas_mejora": [f"Error técnico de Google API: {e}"],
                    "veredicto": "No Recomendado",
                    "error": str(e)
                }
                results_list.append(error_entry)

        progress_bar.progress(1.0)
        status_box.empty()
        st.session_state["evaluation_results"] = results_list
        st.success("🎉 ¡Evaluación completada exitosamente!")
        st.rerun()


# ==============================================================================
# 7. DASHBOARD DE RESULTADOS Y VISUALIZACIONES
# ==============================================================================
eval_results = st.session_state.get("evaluation_results")

if eval_results:
    st.markdown("## 📊 Resultados y Análisis Comparativo")

    # Ordenar candidatos por puntuación descendente
    sorted_results = sorted(eval_results, key=lambda x: x.get("puntuacion_total", 0), reverse=True)

    # 1. Métricas de Alto Nivel
    total_eval = len(sorted_results)
    top_candidate = sorted_results[0] if total_eval > 0 else None
    avg_score = sum(c.get("puntuacion_total", 0) for c in sorted_results) / total_eval if total_eval > 0 else 0
    recommended_count = sum(1 for c in sorted_results if c.get("veredicto") == "Altamente Recomendado")

    mcol1, mcol2, mcol3, mcol4 = st.columns(4)
    with mcol1:
        st.markdown(
            f"""
            <div class="sq-metric-box">
                <div class="sq-metric-label">Candidatos Evaluados</div>
                <div class="sq-metric-value">{total_eval}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with mcol2:
        top_name = top_candidate.get("candidato", "N/A") if top_candidate else "N/A"
        top_sc = top_candidate.get("puntuacion_total", 0) if top_candidate else 0
        st.markdown(
            f"""
            <div class="sq-metric-box">
                <div class="sq-metric-label">Mejor Puntuación</div>
                <div class="sq-metric-value" style="color: #34D399;">{top_sc:.1f} <span style="font-size: 1rem; color: #94A3B8;">/100</span></div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with mcol3:
        st.markdown(
            f"""
            <div class="sq-metric-box">
                <div class="sq-metric-label">Promedio General</div>
                <div class="sq-metric-value" style="color: #818CF8;">{avg_score:.1f}</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    with mcol4:
        pct_rec = (recommended_count / total_eval * 100) if total_eval > 0 else 0
        st.markdown(
            f"""
            <div class="sq-metric-box">
                <div class="sq-metric-label">Altamente Recomendados</div>
                <div class="sq-metric-value" style="color: #60A5FA;">{recommended_count} <span style="font-size: 1rem; color: #94A3B8;">({pct_rec:.0f}%)</span></div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown("<br>", unsafe_allow_html=True)

    # 2. Visualizaciones Gráficas (Radar Multicandidato + Ranking de Barras)
    vcol1, vcol2 = st.columns([1.15, 0.85])

    with vcol1:
        st.markdown("### 🕸️ Radar Comparativo de Competencias")
        radar_fig = go.Figure()

        radar_colors = ["#6366F1", "#10B981", "#F59E0B", "#06B6D4", "#EC4899", "#A855F7"]
        req_names = [row["Requisito Principal"] for _, row in st.session_state["requirements_df"].iterrows()]

        for i, cand in enumerate(sorted_results):
            desglose = cand.get("desglose", [])
            scores_map = {}
            for item in desglose:
                rk = str(item.get("requisito", "")).strip().lower()
                scores_map[rk] = float(item.get("puntaje_asignado", 0))

            r_vals = []
            for rn in req_names:
                rn_clean = rn.strip().lower()
                sc = scores_map.get(rn_clean)
                if sc is None:
                    for k, val in scores_map.items():
                        if k in rn_clean or rn_clean in k:
                            sc = val
                            break
                r_vals.append(sc if sc is not None else 0.0)

            if req_names:
                theta_vals = req_names + [req_names[0]]
                r_vals_closed = r_vals + [r_vals[0]]
            else:
                theta_vals, r_vals_closed = [], []

            color_choice = radar_colors[i % len(radar_colors)]

            radar_fig.add_trace(go.Scatterpolar(
                r=r_vals_closed,
                theta=theta_vals,
                mode='lines+markers',
                fill='toself',
                name=f"{cand.get('candidato')} ({cand.get('puntuacion_total'):.1f} pts)",
                line=dict(color=color_choice, width=2.8),
                marker=dict(size=7, symbol='circle', color=color_choice),
                opacity=0.45,
                hoverinfo="text",
                text=[f"{t}: {r:.0f} pts" for t, r in zip(theta_vals, r_vals_closed)]
            ))

        radar_fig.update_layout(
            polar=dict(
                bgcolor="rgba(30, 41, 59, 0.5)",
                radialaxis=dict(
                    visible=True,
                    range=[0, 100],
                    gridcolor="#334155",
                    gridwidth=1.2,
                    linecolor="#475569",
                    tickfont=dict(color="#94A3B8", size=9),
                    tickvals=[20, 40, 60, 80, 100]
                ),
                angularaxis=dict(
                    gridcolor="#334155",
                    gridwidth=1.2,
                    linecolor="#475569",
                    tickfont=dict(color="#F8FAFC", size=11, family="sans-serif", weight="bold"),
                    direction="clockwise"
                )
            ),
            paper_bgcolor="#0F172A",
            plot_bgcolor="#0F172A",
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.28,
                xanchor="center",
                x=0.5,
                font=dict(color="#F8FAFC", size=10)
            ),
            margin=dict(l=45, r=45, t=20, b=55),
            height=430
        )
        st.plotly_chart(radar_fig, use_container_width=True)

    with vcol2:
        st.markdown("### 🏆 Ranking General de Candidatos")
        df_ranking = pd.DataFrame([
            {
                "Candidato": c.get("candidato"),
                "Puntuación": c.get("puntuacion_total", 0),
                "Veredicto": c.get("veredicto", "No Recomendado")
            }
            for c in sorted_results
        ])

        color_discrete_map = {
            "Altamente Recomendado": "#10B981",
            "Apto con Observaciones": "#F59E0B",
            "No Recomendado": "#EF4444"
        }

        bar_fig = px.bar(
            df_ranking,
            x="Puntuación",
            y="Candidato",
            orientation="h",
            color="Veredicto",
            color_discrete_map=color_discrete_map,
            text="Puntuación",
            range_x=[0, 105]
        )
        bar_fig.update_layout(
            paper_bgcolor="#0F172A",
            plot_bgcolor="#1E293B",
            font=dict(color="#F8FAFC"),
            xaxis=dict(
                gridcolor="#334155",
                title="Puntuación Total (0-100)"
            ),
            yaxis=dict(
                gridcolor="#334155",
                autorange="reversed",
                title=""
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.28,
                xanchor="center",
                x=0.5,
                font=dict(size=10)
            ),
            margin=dict(l=20, r=20, t=20, b=55),
            height=430
        )
        bar_fig.update_traces(textposition='outside')
        st.plotly_chart(bar_fig, use_container_width=True)

    # 3. Tabla Resumen / Leaderboard
    st.markdown("### 📋 Cuadro de Clasificación y Veredictos")
    table_rows = []
    for rank, cand in enumerate(sorted_results, start=1):
        medal = "🥇" if rank == 1 else "🥈" if rank == 2 else "🥉" if rank == 3 else f"#{rank}"
        table_rows.append({
            "Posición": medal,
            "Candidato": cand.get("candidato"),
            "Puntuación Total": f"{cand.get('puntuacion_total'):.1f} / 100",
            "Veredicto": cand.get("veredicto"),
            "Archivo": cand.get("archivo", "-")
        })
    df_leaderboard = pd.DataFrame(table_rows)
    st.dataframe(df_leaderboard, use_container_width=True, hide_index=True)

    # 4. Fichas Detalladas por Candidato
    st.markdown("### 🔍 Fichas Cualitativas y Desglose por Candidato")
    tabs = st.tabs([f"👤 {c.get('candidato')}" for c in sorted_results])

    for tab, cand in zip(tabs, sorted_results):
        with tab:
            verdict = cand.get("veredicto", "No Recomendado")
            badge_class = (
                "sq-badge-green" if verdict == "Altamente Recomendado"
                else "sq-badge-amber" if verdict == "Apto con Observaciones"
                else "sq-badge-red"
            )

            st.markdown(
                f"""
                <div class="sq-card">
                    <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 0.5rem;">
                        <div>
                            <h2 style="margin: 0; font-size: 1.5rem; color: #FFFFFF;">{cand.get('candidato')}</h2>
                            <span style="font-size: 0.85rem; color: #94A3B8;">Archivo origen: <code>{cand.get('archivo', '')}</code></span>
                        </div>
                        <div style="text-align: right;">
                            <span class="sq-badge {badge_class}">{verdict}</span>
                            <div style="font-size: 1.8rem; font-weight: 700; color: #F8FAFC; margin-top: 0.25rem;">
                                {cand.get('puntuacion_total'):.1f} <span style="font-size: 1rem; color: #94A3B8;">/100</span>
                            </div>
                        </div>
                    </div>
                </div>
                """,
                unsafe_allow_html=True
            )

            # Desglose de Requisitos
            st.markdown("#### 📐 Desglose de Evaluación por Criterio")
            desglose = cand.get("desglose", [])
            if desglose:
                d_rows = []
                for d in desglose:
                    peso_val = d.get("peso", 0)
                    punt_val = d.get("puntaje_asignado", 0)
                    aportacion = (punt_val * (peso_val / 100.0))
                    d_rows.append({
                        "Requisito": d.get("requisito"),
                        "Peso (%)": f"{peso_val:.0f}%",
                        "Calificación (0-100)": f"{punt_val:.0f}",
                        "Aportación al Total": f"+{aportacion:.1f} pts",
                        "Justificación Factual del CV": d.get("justificacion")
                    })
                st.dataframe(pd.DataFrame(d_rows), use_container_width=True, hide_index=True)
            else:
                st.caption("No se generó desglose detallado para este candidato.")

            # Cuadrícula de Fortalezas y Áreas de Mejora
            f_col, m_col = st.columns(2)
            with f_col:
                st.markdown(
                    """
                    <div class="sq-card" style="border-left: 4px solid #10B981;">
                        <h4 style="margin: 0 0 0.8rem 0; color: #34D399; display: flex; align-items: center; gap: 0.4rem;">
                            🌟 Fortalezas Clave
                        </h4>
                    """,
                    unsafe_allow_html=True
                )
                fortalezas = cand.get("fortalezas", [])
                if fortalezas:
                    for f in fortalezas:
                        st.markdown(f"• **{f}**")
                else:
                    st.caption("No se especificaron fortalezas destacadas.")
                st.markdown("</div>", unsafe_allow_html=True)

            with m_col:
                st.markdown(
                    """
                    <div class="sq-card" style="border-left: 4px solid #F59E0B;">
                        <h4 style="margin: 0 0 0.8rem 0; color: #FBBF24; display: flex; align-items: center; gap: 0.4rem;">
                            ⚠️ Áreas de Mejora / Brechas
                        </h4>
                    """,
                    unsafe_allow_html=True
                )
                mejoras = cand.get("areas_mejora", [])
                if mejoras:
                    for m in mejoras:
                        st.markdown(f"• {m}")
                else:
                    st.caption("No se detectaron debilidades significativas.")
                st.markdown("</div>", unsafe_allow_html=True)

    # 5. Opciones de Exportación
    st.markdown("---")
    st.markdown("### 📥 Exportar Resultados")
    exp_col1, exp_col2, _ = st.columns([1.5, 1.5, 3])

    with exp_col1:
        csv_data = df_leaderboard.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="📥 Descargar Resultados en CSV",
            data=csv_data,
            file_name="radar_talento_resumen.csv",
            mime="text/csv",
            use_container_width=True
        )

    with exp_col2:
        json_export = json.dumps(sorted_results, indent=2, ensure_ascii=False)
        st.download_button(
            label="📥 Descargar Informe en JSON",
            data=json_export,
            file_name="radar_talento_sq_nexus.json",
            mime="application/json",
            use_container_width=True
        )
