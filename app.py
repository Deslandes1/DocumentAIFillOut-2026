import streamlit as st
import io
import json
from openai import OpenAI
from pypdf import PdfReader, PdfWriter
from docx import Document as DocxDocument
from docxtpl import DocxTemplate
import tempfile
import os
import re

# ---------- PAGE CONFIG ----------
st.set_page_config(page_title="Globalinternet.py Document Filler", layout="wide", initial_sidebar_state="expanded")

# ---------- AUTHENTICATION ----------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False

PASSWORD = "20082010"

if not st.session_state.authenticated:
    st.markdown(
        """
        <style>
        .login-box {
            max-width: 400px;
            margin: 10% auto;
            padding: 2rem;
            background: #6a0dad;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(106, 13, 173, 0.4);
            text-align: center;
        }
        .login-box h2 {
            color: white;
            font-family: 'Segoe UI', sans-serif;
        }
        .login-box input {
            width: 100%;
            padding: 12px;
            margin: 10px 0;
            border-radius: 40px;
            border: none;
            font-size: 1rem;
        }
        .login-box button {
            width: 100%;
            padding: 12px;
            background: #ffd700;
            border: none;
            border-radius: 40px;
            font-weight: bold;
            font-size: 1.1rem;
            cursor: pointer;
        }
        </style>
        <div class="login-box">
            <h2>🔐 Globalinternet.py</h2>
            <p style="color: #ddd;">Enter your password to continue</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
    with st.form("login_form"):
        pwd = st.text_input("Password", type="password", placeholder="Enter password", label_visibility="collapsed")
        submitted = st.form_submit_button("Access")
        if submitted:
            if pwd == PASSWORD:
                st.session_state.authenticated = True
                st.rerun()
            else:
                st.error("Incorrect password")
    st.stop()

# ---------- LANGUAGE DICTIONARY ----------
LANG = {
    "en": {
        "title": "📄 Document Filler",
        "upload_help": "Upload a PDF form with fillable fields or a Word template with placeholders like {{name}}.",
        "upload_label": "Choose a document",
        "upload_btn": "Upload",
        "fields_title": "📋 Fields",
        "auto_fill_hint": "Describe the information you want to fill (e.g., 'John Doe, born 1990, address 123 Main St.')",
        "auto_fill_btn": "✨ Auto-fill with AI",
        "manual_fill_title": "Manual Fill",
        "download_pdf": "⬇️ Download Filled PDF",
        "download_docx": "⬇️ Download Filled Word",
        "ai_suggest": "AI suggestion for this field:",
        "field_label": "Field:",
        "value_label": "Value:",
        "language": "Language",
        "company": "Globalinternet.py Online Software Company",
        "built_by": "Built by Gesner Deslandes, Engineer In Chief",
        "phone": "📞 (509) 4738-5663",
        "email": "✉️ deslandes78@gmail.com",
        "error_upload": "Please upload a PDF or DOCX file.",
        "error_fields": "No fillable fields found in this PDF.",
        "error_api": "OpenAI API key not set. Please add OPENAI_API_KEY in secrets.",
        "success_fill": "Document filled successfully!",
    },
    "fr": {
        "title": "📄 Remplisseur de documents",
        "upload_help": "Téléchargez un formulaire PDF avec champs remplissables ou un modèle Word avec des espaces réservés comme {{name}}.",
        "upload_label": "Choisissez un document",
        "upload_btn": "Télécharger",
        "fields_title": "📋 Champs",
        "auto_fill_hint": "Décrivez les informations à remplir (ex: 'Jean Dupont, né en 1990, adresse 123 rue Principale')",
        "auto_fill_btn": "✨ Remplissage automatique avec IA",
        "manual_fill_title": "Remplissage manuel",
        "download_pdf": "⬇️ Télécharger le PDF rempli",
        "download_docx": "⬇️ Télécharger le Word rempli",
        "ai_suggest": "Suggestion IA pour ce champ :",
        "field_label": "Champ :",
        "value_label": "Valeur :",
        "language": "Langue",
        "company": "Globalinternet.py Société de logiciels en ligne",
        "built_by": "Construit par Gesner Deslandes, Ingénieur en chef",
        "phone": "📞 (509) 4738-5663",
        "email": "✉️ deslandes78@gmail.com",
        "error_upload": "Veuillez télécharger un fichier PDF ou DOCX.",
        "error_fields": "Aucun champ remplissable trouvé dans ce PDF.",
        "error_api": "Clé API OpenAI non définie. Ajoutez OPENAI_API_KEY dans les secrets.",
        "success_fill": "Document rempli avec succès !",
    },
    "es": {
        "title": "📄 Rellenador de documentos",
        "upload_help": "Sube un formulario PDF con campos rellenables o una plantilla de Word con marcadores como {{name}}.",
        "upload_label": "Elige un documento",
        "upload_btn": "Subir",
        "fields_title": "📋 Campos",
        "auto_fill_hint": "Describe la información que deseas rellenar (ej: 'Juan Pérez, nacido en 1990, dirección Calle 123')",
        "auto_fill_btn": "✨ Relleno automático con IA",
        "manual_fill_title": "Relleno manual",
        "download_pdf": "⬇️ Descargar PDF relleno",
        "download_docx": "⬇️ Descargar Word relleno",
        "ai_suggest": "Sugerencia IA para este campo:",
        "field_label": "Campo:",
        "value_label": "Valor:",
        "language": "Idioma",
        "company": "Globalinternet.py Compañía de software en línea",
        "built_by": "Construido por Gesner Deslandes, Ingeniero Jefe",
        "phone": "📞 (509) 4738-5663",
        "email": "✉️ deslandes78@gmail.com",
        "error_upload": "Por favor, sube un archivo PDF o DOCX.",
        "error_fields": "No se encontraron campos rellenables en este PDF.",
        "error_api": "Clave API de OpenAI no configurada. Agrega OPENAI_API_KEY en los secretos.",
        "success_fill": "¡Documento rellenado con éxito!",
    }
}

def t(key):
    return LANG[st.session_state.get("lang", "en")].get(key, key)

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown("""
    <div style="text-align:center; padding:10px;">
        <div style="font-size:4rem;">🌍</div>
        <h3 style="color:#6a0dad;">Globalinternet.py</h3>
        <p style="color:#4a0a6b; font-weight:bold;">Online Software Company</p>
        <hr style="border-color:#6a0dad;">
        <p style="font-size:0.9rem;">💼 Built by <b>Gesner Deslandes</b><br>Engineer In Chief</p>
        <p style="font-size:0.85rem;">📞 (509) 4738-5663<br>✉️ deslandes78@gmail.com</p>
    </div>
    """, unsafe_allow_html=True)

    lang_choice = st.selectbox(
        t("language"),
        options=["en", "fr", "es"],
        format_func=lambda x: {"en": "English", "fr": "Français", "es": "Español"}[x],
        index=0
    )
    st.session_state.lang = lang_choice

# ---------- PURPLE THEME CSS ----------
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(145deg, #f3e8ff 0%, #e6d5f5 100%);
    }
    .stButton>button {
        background-color: #6a0dad;
        color: white;
        border-radius: 40px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #8a2be2;
        transform: scale(1.02);
        box-shadow: 0 8px 20px rgba(106, 13, 173, 0.3);
    }
    .stFileUploader>div>button {
        background-color: #6a0dad !important;
    }
    .stFileUploader>div>button:hover {
        background-color: #8a2be2 !important;
    }
    .stTextInput>div>div>input, .stTextArea>div>textarea {
        border: 2px solid #6a0dad !important;
        border-radius: 30px !important;
        padding: 10px 20px !important;
    }
    .stSelectbox>div>div>select {
        border: 2px solid #6a0dad !important;
        border-radius: 30px !important;
        padding: 8px 15px !important;
    }
    .stAlert {
        border-radius: 20px;
        border-left: 6px solid #6a0dad;
    }
    h1, h2, h3 {
        color: #4a0a6b;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

# ---------- MAIN PAGE ----------
st.title(t("title"))
st.caption(t("company") + " | " + t("built_by"))

# File upload
uploaded_file = st.file_uploader(t("upload_label"), type=["pdf", "docx"], help=t("upload_help"))

# Helper function to get OpenAI client
def get_openai_client():
    api_key = st.secrets.get("OPENAI_API_KEY")
    if not api_key:
        st.error(t("error_api"))
        return None
    return OpenAI(api_key=api_key)

if uploaded_file is not None:
    file_type = uploaded_file.type
    if "pdf" in file_type:
        # Process PDF
        try:
            reader = PdfReader(uploaded_file)
            fields = reader.get_fields()
            if not fields:
                st.warning(t("error_fields"))
                st.stop()
            field_names = list(fields.keys())
            st.success(f"✅ Found {len(field_names)} fillable fields.")
            st.subheader(t("fields_title"))

            # User description for AI auto-fill
            user_description = st.text_area(t("auto_fill_hint"), height=100)
            if st.button(t("auto_fill_btn")):
                client = get_openai_client()
                if client:
                    prompt = f"""
                    You are a helpful assistant that fills PDF form fields.
                    The form has the following fields: {field_names}.
                    The user has described the information as: "{user_description}".
                    Please provide a JSON object mapping each field name to a suitable value.
                    If you don't know a field, suggest a plausible value based on the context.
                    Only output valid JSON.
                    """
                    try:
                        # Use gpt-3.5-turbo as fallback; you can change to "gpt-4" if you have access
                        model = "gpt-3.5-turbo"
                        response = client.chat.completions.create(
                            model=model,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.7
                        )
                        ai_output = response.choices[0].message.content
                        # Extract JSON from response
                        json_match = re.search(r'\{.*\}', ai_output, re.DOTALL)
                        if json_match:
                            ai_values = json.loads(json_match.group())
                        else:
                            ai_values = json.loads(ai_output)
                        for f in field_names:
                            if f in ai_values:
                                st.session_state[f"field_{f}"] = ai_values[f]
                        st.success("✅ AI suggestions applied! Review and adjust below.")
                    except Exception as e:
                        st.error(f"AI error: {e}")

            # Manual fields display
            field_values = {}
            cols = st.columns(2)
            for idx, field in enumerate(field_names):
                with cols[idx % 2]:
                    default_val = st.session_state.get(f"field_{field}", "")
                    st.markdown(f"**{field}**")
                    if default_val:
                        st.caption(f"{t('ai_suggest')} {default_val}")
                    val = st.text_input(t("value_label"), value=default_val, key=f"input_{field}")
                    field_values[field] = val

            # Fill document
            if st.button(t("download_pdf")):
                # Create a filled PDF in memory
                writer = PdfWriter()
                uploaded_file.seek(0)
                reader2 = PdfReader(uploaded_file)
                writer.append(reader2)
                # Update fields
                for field, value in field_values.items():
                    if field in writer.get_fields():
                        writer.update_page_form_field_values(writer.pages[0], {field: value})
                output = io.BytesIO()
                writer.write(output)
                output.seek(0)
                st.download_button(
                    label="📥 Download Filled PDF",
                    data=output,
                    file_name="filled_document.pdf",
                    mime="application/pdf"
                )
                st.success(t("success_fill"))

        except Exception as e:
            st.error(f"Error: {e}")

    elif "word" in file_type or "vnd.openxmlformats-officedocument.wordprocessingml.document" in file_type:
        # Process Word template with placeholders (e.g., {{name}})
        with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
            tmp.write(uploaded_file.getvalue())
            tmp_path = tmp.name

        doc = DocxDocument(tmp_path)
        full_text = " ".join([p.text for p in doc.paragraphs])
        placeholders = re.findall(r'{{(.*?)}}', full_text)
        placeholders = list(set(placeholders))
        if not placeholders:
            st.warning("No placeholders found in this Word document.")
        else:
            st.success(f"✅ Found placeholders: {placeholders}")
            st.subheader(t("fields_title"))

            # User description for AI
            user_description = st.text_area(t("auto_fill_hint"), height=100)
            if st.button(t("auto_fill_btn")):
                client = get_openai_client()
                if client:
                    prompt = f"""
                    You are a helpful assistant that fills Word document placeholders.
                    The placeholders are: {placeholders}.
                    The user described: "{user_description}".
                    Provide a JSON mapping each placeholder to a suitable value.
                    Only output JSON.
                    """
                    try:
                        model = "gpt-3.5-turbo"
                        response = client.chat.completions.create(
                            model=model,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.7
                        )
                        ai_output = response.choices[0].message.content
                        json_match = re.search(r'\{.*\}', ai_output, re.DOTALL)
                        if json_match:
                            ai_values = json.loads(json_match.group())
                        else:
                            ai_values = json.loads(ai_output)
                        for p in placeholders:
                            if p in ai_values:
                                st.session_state[f"field_{p}"] = ai_values[p]
                        st.success("✅ AI suggestions applied!")
                    except Exception as e:
                        st.error(f"AI error: {e}")

            # Manual input for each placeholder
            placeholder_values = {}
            cols = st.columns(2)
            for idx, placeholder in enumerate(placeholders):
                with cols[idx % 2]:
                    default_val = st.session_state.get(f"field_{placeholder}", "")
                    st.markdown(f"**{{{{{placeholder}}}}}**")
                    if default_val:
                        st.caption(f"{t('ai_suggest')} {default_val}")
                    val = st.text_input(t("value_label"), value=default_val, key=f"input_{placeholder}")
                    placeholder_values[placeholder] = val

            # Fill template and download
            if st.button(t("download_docx")):
                context = placeholder_values
                docx_template = DocxTemplate(tmp_path)
                docx_template.render(context)
                output = io.BytesIO()
                docx_template.save(output)
                output.seek(0)
                st.download_button(
                    label="📥 Download Filled Word",
                    data=output,
                    file_name="filled_document.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
                st.success(t("success_fill"))

        os.unlink(tmp_path)

else:
    st.info("📂 " + t("upload_help"))

# ---------- FOOTER ----------
st.markdown("---")
st.caption("© 2026 Globalinternet.py | " + t("company"))
