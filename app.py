import streamlit as st
import io
import json
from groq import Groq
from pypdf import PdfReader, PdfWriter
from docx import Document as DocxDocument
from docxtpl import DocxTemplate
import tempfile
import os
import re
import base64
from gtts import gTTS
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

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
            background: #1a5276;
            border-radius: 20px;
            box-shadow: 0 10px 30px rgba(26, 82, 118, 0.4);
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
        "send_email": "📧 Send to my Google Docs (email)",
        "email_sent": "✅ Document sent to your email! Check your inbox.",
        "email_error": "Email sending failed. Check SMTP settings.",
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
        "error_api": "Groq API key not set. Please add GROQ_API_KEY in secrets.",
        "success_fill": "Document filled successfully!",
        "ai_provider": "Groq (Llama 3.1)",
        "listen_explanation": "🔊 AI Female Voice – How This App Works",
        "recipient_email": "Recipient email",
        "email_subject": "Filled document from Document Filler",
        "email_body": "Please find the filled document attached. You can open it directly in Google Docs from this email.",
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
        "send_email": "📧 Envoyer à mon Google Docs (email)",
        "email_sent": "✅ Document envoyé à votre email ! Vérifiez votre boîte.",
        "email_error": "L'envoi d'email a échoué. Vérifiez les paramètres SMTP.",
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
        "error_api": "Clé API Groq non définie. Ajoutez GROQ_API_KEY dans les secrets.",
        "success_fill": "Document rempli avec succès !",
        "ai_provider": "Groq (Llama 3.1)",
        "listen_explanation": "🔊 Voix IA Féminine – Comment fonctionne cette app",
        "recipient_email": "Email du destinataire",
        "email_subject": "Document rempli depuis Document Filler",
        "email_body": "Veuillez trouver le document rempli en pièce jointe. Vous pouvez l'ouvrir directement dans Google Docs à partir de cet email.",
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
        "send_email": "📧 Enviar a mi Google Docs (email)",
        "email_sent": "✅ ¡Documento enviado a su email! Revise su bandeja.",
        "email_error": "Error al enviar el email. Verifique la configuración SMTP.",
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
        "error_api": "Clave API de Groq no configurada. Agrega GROQ_API_KEY en los secretos.",
        "success_fill": "¡Documento rellenado con éxito!",
        "ai_provider": "Groq (Llama 3.1)",
        "listen_explanation": "🔊 Voz IA Femenina – Cómo funciona esta app",
        "recipient_email": "Email del destinatario",
        "email_subject": "Documento rellenado desde Document Filler",
        "email_body": "Adjunto encontrará el documento rellenado. Puede abrirlo directamente en Google Docs desde este email.",
    }
}

def t(key):
    return LANG[st.session_state.get("lang", "en")].get(key, key)

# ---------- SIDEBAR ----------
with st.sidebar:
    st.markdown(
        """
        <div style="text-align:center; padding:10px;">
            <img src="https://github.com/Deslandes1/DocumentAIFillOut-2026/blob/main/Gesner%20Deslandes%202026.png?raw=true" 
                 style="width:120px; height:120px; border-radius:50%; object-fit:cover; border:3px solid #1a5276; margin-bottom:5px;">
            <h3 style="color:#1a5276; margin:0;">Gesner Deslandes</h3>
            <p style="color:#2c3e50; font-weight:bold; margin:0;">Engineer In Chief</p>
            <hr style="border-color:#1a5276;">
            <p style="font-size:0.9rem;">💼 Built by <b>Gesner Deslandes</b><br>Engineer In Chief</p>
            <p style="font-size:0.85rem;">📞 (509) 4738-5663<br>✉️ deslandes78@gmail.com</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    lang_choice = st.selectbox(
        t("language"),
        options=["en", "fr", "es"],
        format_func=lambda x: {"en": "English", "fr": "Français", "es": "Español"}[x],
        index=0
    )
    st.session_state.lang = lang_choice

    st.info(f"🤖 AI: {t('ai_provider')}")

    st.divider()

    # ---------- VOICE EXPLANATION ----------
    if st.button(t("listen_explanation"), use_container_width=True):
        if st.session_state.lang == "en":
            explanation = (
                "This application helps you fill out any document automatically. "
                "You can upload a PDF form with fillable fields or a Word document with placeholders like double curly braces name. "
                "Describe the information you want to fill, and the AI will suggest values for each field. "
                "You can also manually adjust any field. "
                "Finally, download the filled document as a PDF or Word file, or send it to your email to open in Google Docs. "
                "All data is processed securely, and you can use this for visa applications, contracts, and more."
            )
        elif st.session_state.lang == "fr":
            explanation = (
                "Cette application vous aide à remplir automatiquement tout document. "
                "Vous pouvez télécharger un formulaire PDF avec des champs remplissables ou un document Word avec des espaces réservés comme double accolade nom. "
                "Décrivez les informations à remplir, et l'IA suggérera des valeurs pour chaque champ. "
                "Vous pouvez également ajuster manuellement chaque champ. "
                "Enfin, téléchargez le document rempli au format PDF ou Word, ou envoyez-le à votre email pour l'ouvrir dans Google Docs. "
                "Toutes les données sont traitées de manière sécurisée, et vous pouvez l'utiliser pour des demandes de visa, des contrats, etc."
            )
        else:
            explanation = (
                "Esta aplicación le ayuda a rellenar automáticamente cualquier documento. "
                "Puede subir un formulario PDF con campos rellenables o un documento Word con marcadores como doble llave nombre. "
                "Describa la información que desea rellenar, y la IA sugerirá valores para cada campo. "
                "También puede ajustar manualmente cada campo. "
                "Finalmente, descargue el documento rellenado como PDF o Word, o envíelo a su email para abrirlo en Google Docs. "
                "Todos los datos se procesan de forma segura, y puede usarlo para solicitudes de visa, contratos, etc."
            )
        try:
            tts = gTTS(text=explanation, lang=st.session_state.lang, slow=False)
            with tempfile.NamedTemporaryFile(delete=False, suffix=".mp3") as tmp:
                tts.save(tmp.name)
                audio_path = tmp.name
            with open(audio_path, "rb") as f:
                audio_bytes = f.read()
            st.audio(audio_bytes, format="audio/mp3")
            os.unlink(audio_path)
            st.success("▶️ Audio playing...")
        except Exception as e:
            st.error(f"Could not generate audio: {e}")

# ---------- BLUE THEME CSS ----------
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(145deg, #eaf2f8 0%, #d4e6f1 100%);
    }
    .stButton>button {
        background-color: #1a5276;
        color: white;
        border-radius: 40px;
        border: none;
        padding: 0.5rem 1.5rem;
        font-weight: 600;
        transition: all 0.3s;
    }
    .stButton>button:hover {
        background-color: #2471a3;
        transform: scale(1.02);
        box-shadow: 0 8px 20px rgba(26, 82, 118, 0.3);
    }
    .stFileUploader>div>button {
        background-color: #1a5276 !important;
    }
    .stFileUploader>div>button:hover {
        background-color: #2471a3 !important;
    }
    .stTextInput>div>div>input, .stTextArea>div>textarea {
        border: 2px solid #1a5276 !important;
        border-radius: 30px !important;
        padding: 10px 20px !important;
    }
    .stSelectbox>div>div>select {
        border: 2px solid #1a5276 !important;
        border-radius: 30px !important;
        padding: 8px 15px !important;
    }
    .stAlert {
        border-radius: 20px;
        border-left: 6px solid #1a5276;
    }
    h1, h2, h3 {
        color: #154360;
    }
    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }
    [data-testid="stSidebar"] {
        background: #eaf2f8;
        border-right: 2px solid #1a5276;
    }
</style>
""", unsafe_allow_html=True)

# ---------- EMAIL SENDING FUNCTION ----------
def send_email_attachment(recipient, subject, body, file_bytes, filename):
    try:
        smtp_server = st.secrets.get("SMTP_SERVER", "smtp.gmail.com")
        smtp_port = int(st.secrets.get("SMTP_PORT", 587))
        smtp_user = st.secrets.get("SMTP_USERNAME", "deslandes78@gmail.com")
        smtp_pass = st.secrets.get("SMTP_PASSWORD")
        if not smtp_pass:
            st.error("SMTP password not set in secrets.")
            return False

        msg = MIMEMultipart()
        msg["From"] = smtp_user
        msg["To"] = recipient
        msg["Subject"] = subject
        msg.attach(MIMEText(body, "plain"))

        part = MIMEBase("application", "octet-stream")
        part.set_payload(file_bytes)
        encoders.encode_base64(part)
        part.add_header(
            "Content-Disposition",
            f"attachment; filename={filename}"
        )
        msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)
        server.quit()
        return True
    except Exception as e:
        st.error(f"Email error: {e}")
        return False

# ---------- MAIN PAGE ----------
st.title(t("title"))
st.caption(t("company") + " | " + t("built_by"))

uploaded_file = st.file_uploader(t("upload_label"), type=["pdf", "docx"], help=t("upload_help"))

def get_groq_client():
    api_key = st.secrets.get("GROQ_API_KEY")
    if not api_key:
        st.error(t("error_api"))
        return None
    return Groq(api_key=api_key)

if uploaded_file is not None:
    file_type = uploaded_file.type

    if "pdf" in file_type:
        try:
            reader = PdfReader(uploaded_file)
            fields = reader.get_fields()
            if not fields:
                st.warning(t("error_fields"))
                st.stop()
            field_names = list(fields.keys())
            st.success(f"✅ Found {len(field_names)} fillable fields.")
            st.subheader(t("fields_title"))

            user_description = st.text_area(t("auto_fill_hint"), height=100)
            if st.button(t("auto_fill_btn")):
                client = get_groq_client()
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
                        model = "llama-3.1-8b-instant"
                        response = client.chat.completions.create(
                            model=model,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.7,
                            max_tokens=1024
                        )
                        ai_output = response.choices[0].message.content
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

            col1, col2 = st.columns(2)
            with col1:
                if st.button(t("download_pdf")):
                    writer = PdfWriter()
                    uploaded_file.seek(0)
                    reader2 = PdfReader(uploaded_file)
                    writer.append(reader2)
                    for field, value in field_values.items():
                        if field in writer.get_fields():
                            writer.update_page_form_field_values(writer.pages[0], {field: value})
                    output = io.BytesIO()
                    writer.write(output)
                    output.seek(0)
                    file_bytes = output.getvalue()
                    st.download_button(
                        label="📥 Download Filled PDF",
                        data=file_bytes,
                        file_name="filled_document.pdf",
                        mime="application/pdf"
                    )
                    st.success(t("success_fill"))
                    # Store file for email
                    st.session_state["pdf_bytes"] = file_bytes
                    st.session_state["pdf_filename"] = "filled_document.pdf"

            with col2:
                if st.button(t("send_email")):
                    if "pdf_bytes" not in st.session_state:
                        st.warning("Please generate the PDF first by clicking Download.")
                    else:
                        recipient = st.secrets.get("EMAIL_TO", "deslandes78@gmail.com")
                        subject = t("email_subject")
                        body = t("email_body")
                        if send_email_attachment(
                            recipient,
                            subject,
                            body,
                            st.session_state["pdf_bytes"],
                            st.session_state["pdf_filename"]
                        ):
                            st.success(t("email_sent"))
                        else:
                            st.error(t("email_error"))

        except Exception as e:
            st.error(f"Error: {e}")

    elif "word" in file_type or "vnd.openxmlformats-officedocument.wordprocessingml.document" in file_type:
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

            user_description = st.text_area(t("auto_fill_hint"), height=100)
            if st.button(t("auto_fill_btn")):
                client = get_groq_client()
                if client:
                    prompt = f"""
                    You are a helpful assistant that fills Word document placeholders.
                    The placeholders are: {placeholders}.
                    The user described: "{user_description}".
                    Provide a JSON mapping each placeholder to a suitable value.
                    Only output JSON.
                    """
                    try:
                        model = "llama-3.1-8b-instant"
                        response = client.chat.completions.create(
                            model=model,
                            messages=[{"role": "user", "content": prompt}],
                            temperature=0.7,
                            max_tokens=1024
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

            col1, col2 = st.columns(2)
            with col1:
                if st.button(t("download_docx")):
                    context = placeholder_values
                    docx_template = DocxTemplate(tmp_path)
                    docx_template.render(context)
                    output = io.BytesIO()
                    docx_template.save(output)
                    output.seek(0)
                    file_bytes = output.getvalue()
                    st.download_button(
                        label="📥 Download Filled Word",
                        data=file_bytes,
                        file_name="filled_document.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
                    st.success(t("success_fill"))
                    st.session_state["docx_bytes"] = file_bytes
                    st.session_state["docx_filename"] = "filled_document.docx"

            with col2:
                if st.button(t("send_email")):
                    if "docx_bytes" not in st.session_state:
                        st.warning("Please generate the Word file first by clicking Download.")
                    else:
                        recipient = st.secrets.get("EMAIL_TO", "deslandes78@gmail.com")
                        subject = t("email_subject")
                        body = t("email_body")
                        if send_email_attachment(
                            recipient,
                            subject,
                            body,
                            st.session_state["docx_bytes"],
                            st.session_state["docx_filename"]
                        ):
                            st.success(t("email_sent"))
                        else:
                            st.error(t("email_error"))

        os.unlink(tmp_path)

else:
    st.info("📂 " + t("upload_help"))

# ---------- FOOTER ----------
st.markdown("---")
st.caption("© 2026 Globalinternet.py | " + t("company"))
