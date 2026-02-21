# ============================================
# STEP 1: IMPORTS
# ============================================
import re
import json
import os
import hashlib
from datetime import datetime
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import pdfplumber as pdf
import plotly.express as px
import plotly.graph_objects as go
from docx import Document
from huggingface_hub import InferenceClient
from dotenv import load_dotenv
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.utils import simpleSplit
from io import BytesIO

load_dotenv()

# ============================================
# STEP 2: PAGE CONFIGURATION
# ============================================
st.set_page_config(page_title='GenAI Resume Scanner', page_icon=':file_folder:', layout='wide')
PAGE_COLOR_THEMES = {
    "Default": {
        "main_bg": "#0E1117",
        "sidebar_bg": "#0E1117",
        "text_color": "#F9FAFB",
        "settings_bg": "rgba(14, 17, 23, 0.95)",
        "button_text": "#FFFFFF",
        "chip_text": "#FFFFFF",
        "textbox_border": "rgba(148, 163, 184, 0.35)",
        "textbox_focus": "rgba(59, 130, 246, 0.75)",
    },
    "Midnight Blue": {
        "main_bg": "#0B1B34",
        "sidebar_bg": "#09172C",
        "text_color": "#F9FAFB",
        "settings_bg": "rgba(9, 23, 44, 0.95)",
        "button_text": "#FFFFFF",
        "chip_text": "#FFFFFF",
        "textbox_border": "rgba(148, 163, 184, 0.35)",
        "textbox_focus": "rgba(59, 130, 246, 0.75)",
    },
    "Forest": {
        "main_bg": "#0F241A",
        "sidebar_bg": "#0B1B14",
        "text_color": "#F9FAFB",
        "settings_bg": "rgba(11, 27, 20, 0.95)",
        "button_text": "#FFFFFF",
        "chip_text": "#FFFFFF",
        "textbox_border": "rgba(148, 163, 184, 0.35)",
        "textbox_focus": "rgba(59, 130, 246, 0.75)",
    },
    "Slate": {
        "main_bg": "#1A1F2B",
        "sidebar_bg": "#151A24",
        "text_color": "#F9FAFB",
        "settings_bg": "rgba(21, 26, 36, 0.95)",
        "button_text": "#FFFFFF",
        "chip_text": "#FFFFFF",
        "textbox_border": "rgba(148, 163, 184, 0.35)",
        "textbox_focus": "rgba(59, 130, 246, 0.75)",
    },
    "Light": {
        "main_bg": "#F6F8FC",
        "sidebar_bg": "#EEF2F8",
        "text_color": "#111827",
        "settings_bg": "rgba(238, 242, 248, 0.96)",
        "button_text": "#111827",
        "chip_text": "#111827",
        "textbox_border": "rgba(17, 24, 39, 0.55)",
        "textbox_focus": "rgba(17, 24, 39, 0.9)",
    },
}

# Keep page theme persistent across refresh via URL query params.
query_theme_name = st.query_params.get("theme")
if query_theme_name in PAGE_COLOR_THEMES:
    st.session_state["page_color_theme"] = query_theme_name

selected_theme_name = st.session_state.get("page_color_theme", "Default")
if selected_theme_name not in PAGE_COLOR_THEMES:
    selected_theme_name = "Default"
st.session_state["page_color_theme"] = selected_theme_name
if st.query_params.get("theme") != selected_theme_name:
    st.query_params["theme"] = selected_theme_name


def persist_page_theme_to_query_params():
    chosen_theme = st.session_state.get("page_color_theme", "Default")
    if chosen_theme not in PAGE_COLOR_THEMES:
        chosen_theme = "Default"
        st.session_state["page_color_theme"] = chosen_theme
    st.query_params["theme"] = chosen_theme


selected_theme = PAGE_COLOR_THEMES[selected_theme_name]

theme_css = """
<style>
[data-testid="stAppViewContainer"] {
    background: __MAIN_BG__;
    color: __TEXT_COLOR__;
}
section[data-testid="stSidebar"] {
    background: __SIDEBAR_BG__;
    color: __TEXT_COLOR__;
    border-right: 1px solid rgba(148, 163, 184, 0.28);
}
section[data-testid="stSidebar"] div[data-testid="stSidebarUserContent"] {
    display: flex;
    flex-direction: column;
    height: 100%;
    padding-top: 0.2rem;
}
.previous-scroll {
    overflow-y: auto;
    max-height: calc(100vh - 6rem);
    padding-bottom: 1rem;
    margin-top: -1rem;
}
.previous-scroll h3 {
    margin-top: 0 !important;
    margin-bottom: 0.35rem !important;
}
.settings-bottom {
    margin-top: 0.2rem;
    position: static;
    bottom: auto;
    z-index: auto;
    background: __SETTINGS_BG__;
    padding-top: 0.35rem;
    padding-bottom: 0.25rem;
    border-top: none;
    box-shadow: none;
}
.settings-bottom [data-testid="stExpander"] {
    margin-top: 0 !important;
    margin-bottom: 0 !important;
}
.settings-bottom [data-testid="stExpander"] > details > summary {
    margin-bottom: 0.2rem !important;
}
.settings-bottom div[data-baseweb="select"] {
    margin-top: 0.15rem !important;
    margin-bottom: 0 !important;
}
.settings-bottom [data-testid="stSelectbox"] {
    margin-bottom: 0.1rem !important;
}
.settings-bottom hr {
    margin: 0.08rem 0 0.3rem 0 !important;
}
.settings-divider {
    height: 1px;
    background: rgba(148, 163, 184, 0.35);
    margin: 0.08rem 0 0.28rem 0;
}
button[data-baseweb="tab"] p {
    font-size: 1.08rem;
    font-weight: 600;
}
.keyword-chip {
    display: inline-block;
    margin: 0 0.35rem 0.35rem 0;
    padding: 0.2rem 0.55rem;
    border-radius: 999px;
    border: 1px solid rgba(239, 68, 68, 0.5);
    background: rgba(239, 68, 68, 0.22);
    color: __CHIP_TEXT__;
    font-size: 0.78rem;
    font-weight: 600;
    letter-spacing: 0.02em;
}
section[data-testid="stSidebar"] div[data-testid="stButton"] > button {
    width: 100%;
    height: 2.35rem;
    border-radius: 0.6rem;
    border: 1px solid rgba(59, 130, 246, 0.6);
    background: rgba(59, 130, 246, 0.35);
    color: __BUTTON_TEXT__;
    font-weight: 600;
    font-size: 0.74rem;
    padding: 0.25rem 0.45rem;
    white-space: normal;
    line-height: 1.05;
    margin: 0.72rem 0;
}
section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover {
    background: rgba(37, 99, 235, 0.5);
    border-color: rgba(59, 130, 246, 0.8);
    color: __BUTTON_TEXT__;
}
div[data-baseweb="textarea"] {
    border: 1px solid __TEXTBOX_BORDER__;
    border-radius: 0.5rem;
}
div[data-baseweb="textarea"]:focus-within {
    border: 1px solid __TEXTBOX_FOCUS__;
    box-shadow: 0 0 0 1px __TEXTBOX_FOCUS__;
}
div[data-baseweb="input"] {
    border: 1px solid __TEXTBOX_BORDER__;
    border-radius: 0.5rem;
}
div[data-baseweb="input"]:focus-within {
    border: 1px solid __TEXTBOX_FOCUS__;
    box-shadow: 0 0 0 1px __TEXTBOX_FOCUS__;
}
</style>
"""
theme_css = (
    theme_css
    .replace("__MAIN_BG__", selected_theme["main_bg"])
    .replace("__SIDEBAR_BG__", selected_theme["sidebar_bg"])
    .replace("__TEXT_COLOR__", selected_theme["text_color"])
    .replace("__SETTINGS_BG__", selected_theme["settings_bg"])
    .replace("__BUTTON_TEXT__", selected_theme["button_text"])
    .replace("__CHIP_TEXT__", selected_theme["chip_text"])
    .replace("__TEXTBOX_BORDER__", selected_theme["textbox_border"])
    .replace("__TEXTBOX_FOCUS__", selected_theme["textbox_focus"])
)
st.markdown(theme_css, unsafe_allow_html=True)

if selected_theme_name == "Light":
    st.markdown(
        """
        <style>
        [data-testid="stAppViewContainer"] p,
        [data-testid="stAppViewContainer"] li,
        [data-testid="stAppViewContainer"] label,
        [data-testid="stAppViewContainer"] h1,
        [data-testid="stAppViewContainer"] h2,
        [data-testid="stAppViewContainer"] h3,
        [data-testid="stAppViewContainer"] h4,
        [data-testid="stAppViewContainer"] h5,
        [data-testid="stAppViewContainer"] h6,
        [data-testid="stSidebar"] p,
        [data-testid="stSidebar"] li,
        [data-testid="stSidebar"] label,
        .settings-bottom summary,
        .settings-bottom summary *,
        .settings-bottom [data-testid="stExpander"] summary,
        .settings-bottom [data-testid="stExpander"] summary *,
        .settings-bottom [data-testid="stExpanderDetails"],
        .settings-bottom [data-testid="stExpanderDetails"] *,
        [data-testid="stCaptionContainer"],
        [data-testid="stMetricLabel"],
        [data-testid="stMetricValue"] {
            color: #111827 !important;
        }

        .settings-bottom [data-testid="stExpander"] > details,
        .settings-bottom [data-testid="stExpanderDetails"] {
            background: rgba(255, 255, 255, 0.85) !important;
            border: 1px solid rgba(17, 24, 39, 0.18) !important;
            border-radius: 0.55rem !important;
        }

        .settings-bottom,
        .settings-bottom * {
            color: #111827 !important;
        }

        .settings-bottom [data-testid="stExpander"],
        .settings-bottom [data-testid="stExpander"] > details,
        .settings-bottom [data-testid="stExpander"] > details > summary,
        .settings-bottom [data-testid="stExpanderDetails"] {
            background: #f8fafc !important;
        }

        .settings-bottom [data-testid="stExpander"] svg,
        .settings-bottom [data-testid="stExpander"] summary svg {
            fill: #111827 !important;
            color: #111827 !important;
        }

        .settings-bottom [data-testid="stExpander"] > details > summary,
        .settings-bottom [data-testid="stExpander"] > details > summary p,
        .settings-bottom [data-testid="stExpander"] > details > summary span,
        .settings-bottom [data-testid="stExpander"] > details > summary div {
            color: #111827 !important;
            -webkit-text-fill-color: #111827 !important;
            opacity: 1 !important;
            text-shadow: none !important;
        }

        div[data-baseweb="textarea"],
        div[data-baseweb="input"],
        div[data-baseweb="select"] {
            background: #ffffff !important;
            border: 1px solid rgba(17, 24, 39, 0.62) !important;
            border-radius: 0.5rem !important;
        }

        div[data-baseweb="textarea"]:focus-within,
        div[data-baseweb="input"]:focus-within,
        div[data-baseweb="select"]:focus-within {
            border: 1px solid rgba(17, 24, 39, 0.92) !important;
            box-shadow: 0 0 0 1px rgba(17, 24, 39, 0.92) !important;
        }

        button[data-baseweb="tab"][aria-selected="true"] {
            background: rgba(17, 24, 39, 0.09) !important;
        }

        [data-baseweb="tab-list"] {
            gap: 0.5rem !important;
            margin-bottom: 0.6rem !important;
        }

        button[data-baseweb="tab"] {
            margin: 0.15rem 0 !important;
            padding: 0.5rem 0.8rem !important;
        }

        .compatibility-block {
            margin: 0.5rem 0 0.9rem 0;
            padding: 0;
            border: none;
            background: transparent;
        }

        header[data-testid="stHeader"],
        div[data-testid="stToolbar"],
        div[data-testid="stDecoration"] {
            background: #F6F8FC !important;
        }

        header[data-testid="stHeader"] {
            border-bottom: 1px solid rgba(17, 24, 39, 0.12) !important;
        }

        div[data-testid="stFormSubmitButton"] > button {
            background: #e2e8f0 !important;
            color: #111827 !important;
            border: 1px solid rgba(17, 24, 39, 0.45) !important;
            font-weight: 700 !important;
        }

        div[data-testid="stFormSubmitButton"] > button:hover {
            background: #cbd5e1 !important;
            color: #111827 !important;
            border-color: rgba(17, 24, 39, 0.7) !important;
        }

        div[data-testid="stFormSubmitButton"] > button:focus {
            box-shadow: 0 0 0 2px rgba(17, 24, 39, 0.35) !important;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

HISTORY_FILE = "run_history.json"


def load_history():
    if not os.path.exists(HISTORY_FILE):
        return []
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as file:
            data = json.load(file)
            return data if isinstance(data, list) else []
    except Exception:
        return []


def save_history(history_items):
    with open(HISTORY_FILE, "w", encoding="utf-8") as file:
        json.dump(history_items, file, ensure_ascii=False, indent=2)


def add_history_entry(resume_text, job_description, scores, ai_result=None, edited_resume=None):
    history_items = load_history()
    now = datetime.now().astimezone()
    entry = {
        "saved_at": now.strftime("%-m/%-d/%Y %-I:%M:%S %p"),
        "resume_text": resume_text,
        "job_description": job_description,
        "scores": scores,
        "ai_result": ai_result,
        "edited_resume": edited_resume,
    }
    history_items.insert(0, entry)
    save_history(history_items)


def format_saved_at(saved_at_value):
    if not saved_at_value:
        return "Unknown"
    try:
        parsed_dt = datetime.strptime(saved_at_value, "%Y-%m-%d %H:%M:%S")
        return parsed_dt.strftime("%-m/%-d/%Y %-I:%M:%S %p")
    except Exception:
        return saved_at_value


def resume_text_to_pdf_bytes(resume_text, fit_one_page=False):
    buffer = BytesIO()
    pdf_canvas = canvas.Canvas(buffer, pagesize=LETTER)
    page_width, page_height = LETTER
    left_margin = 46
    right_margin = 46
    top_margin = 42
    bottom_margin = 42
    max_width = page_width - left_margin - right_margin

    lines = (resume_text or "").split("\n")

    if fit_one_page:
        candidate_font_sizes = [11, 10.5, 10, 9.5, 9, 8.5, 8, 7.5, 7, 6.5]
        selected_font_size = candidate_font_sizes[-1]
        selected_wrapped_lines = []

        for font_size in candidate_font_sizes:
            wrapped_lines = []
            for raw_line in lines:
                line_chunks = simpleSplit(raw_line, "Helvetica", font_size, max_width) if raw_line else [""]
                wrapped_lines.extend(line_chunks)

            line_height = max(8, int(font_size * 1.24))
            available_height = page_height - top_margin - bottom_margin
            max_lines_on_one_page = int(available_height // line_height)

            if len(wrapped_lines) <= max_lines_on_one_page:
                selected_font_size = font_size
                selected_wrapped_lines = wrapped_lines
                break

            selected_wrapped_lines = wrapped_lines

        line_height = max(8, int(selected_font_size * 1.24))
        available_height = page_height - top_margin - bottom_margin
        max_lines_on_one_page = int(available_height // line_height)
        rendered_lines = selected_wrapped_lines[:max_lines_on_one_page]

        if len(selected_wrapped_lines) > max_lines_on_one_page and rendered_lines:
            rendered_lines[-1] = (rendered_lines[-1].rstrip() + " …")[:220]

        y_pos = page_height - top_margin
        pdf_canvas.setFont("Helvetica", selected_font_size)
        for line in rendered_lines:
            pdf_canvas.drawString(left_margin, y_pos, line)
            y_pos -= line_height
    else:
        line_height = 14
        y_pos = page_height - top_margin
        pdf_canvas.setFont("Helvetica", 11)
        for raw_line in lines:
            wrapped_lines = simpleSplit(raw_line, "Helvetica", 11, max_width) if raw_line else [""]
            for line in wrapped_lines:
                if y_pos <= bottom_margin:
                    pdf_canvas.showPage()
                    pdf_canvas.setFont("Helvetica", 11)
                    y_pos = page_height - top_margin
                pdf_canvas.drawString(left_margin, y_pos, line)
                y_pos -= line_height

    pdf_canvas.save()
    buffer.seek(0)
    return buffer.getvalue()


def render_pdf_preview(pdf_bytes, height=680):
    if not pdf_bytes:
        st.info("No PDF available to preview yet.")
        return

    try:
        import pypdfium2 as pdfium

        pdf_document = pdfium.PdfDocument(pdf_bytes)
        total_pages = len(pdf_document)
        if total_pages == 0:
            st.info("Unable to render PDF preview.")
            return

        max_preview_pages = min(total_pages, 3)
        preview_images = []
        preview_captions = []

        for page_index in range(max_preview_pages):
            page = pdf_document[page_index]
            bitmap = page.render(scale=1.4)
            preview_images.append(bitmap.to_pil())
            preview_captions.append(f"Page {page_index + 1}")

        st.image(preview_images, caption=preview_captions, use_container_width=True)
        if total_pages > max_preview_pages:
            st.caption(f"Showing first {max_preview_pages} of {total_pages} pages.")
    except Exception:
        st.info(
            "PDF preview is unavailable in this environment. "
            "Install `pypdfium2` to enable in-app preview, or use the download button below."
        )


def summarize_job_description(job_description_text):
    job_description_text = job_description_text or ""
    lines = [line.strip() for line in job_description_text.split("\n") if line.strip()]
    job_title = lines[0] if lines else "Unknown role"

    normalized = re.sub(r"[^a-zA-Z0-9\s]", " ", job_description_text.lower())
    words = normalized.split()
    stop_words = {
        "the", "and", "or", "for", "with", "from", "that", "this", "you", "your", "are", "our",
        "job", "role", "team", "will", "have", "has", "into", "their", "they", "what", "who",
        "where", "when", "why", "how", "all", "any", "but", "not", "can", "use", "using",
    }

    keywords = []
    for word in words:
        if len(word) < 4 or word in stop_words:
            continue
        if word not in keywords:
            keywords.append(word)
        if len(keywords) == 3:
            break

    return job_title, keywords


def extract_company_name(job_description_text):
    job_description_text = job_description_text or ""
    lines = [line.strip() for line in job_description_text.split("\n") if line.strip()]

    for idx, line in enumerate(lines):
        if line.lower() == "company" and idx + 1 < len(lines):
            return lines[idx + 1]

    company_markers = [" inc", " llc", " ltd", " corp", " corporation", " games", " technologies", " labs"]
    for line in lines:
        lower_line = line.lower()
        if any(marker in lower_line for marker in company_markers):
            return line

    return "Unknown Company"

# ============================================
# STEP 3: MAIN TITLE
# ============================================
st.title("GenAI Resume Scanner")




# ============================================
# STEP 4: SIDEBAR SETUP
# ============================================
with st.sidebar:
    st.markdown('<div class="previous-scroll">', unsafe_allow_html=True)
    st.markdown("### Previous")
    history_items = load_history()
    if not history_items:
        st.caption("No previous runs yet.")
    else:
        recent_items = history_items[:3]
        older_items = history_items[3:]

        for idx, item in enumerate(recent_items):
            saved_at = format_saved_at(item.get("saved_at", "Unknown"))
            score = item.get("scores", {}).get("overall_score", 0)
            company_name = extract_company_name(item.get("job_description", ""))
            job_title, keywords = summarize_job_description(item.get("job_description", ""))
            with st.expander(f"{company_name} • {saved_at} • {score}%"):
                btn_col1, btn_col2 = st.columns(2)
                with btn_col1:
                    if st.button("Load into Scanner", key=f"load_recent_{idx}", use_container_width=True):
                        st.session_state["resume_text_input"] = item.get("resume_text", "") or ""
                        st.session_state["job_description_input"] = item.get("job_description", "") or ""
                        st.session_state["submitted_resume_text"] = item.get("resume_text", "") or ""
                        st.session_state["submitted_job_description"] = item.get("job_description", "") or ""
                        st.rerun()
                with btn_col2:
                    if st.button("Preview", key=f"preview_recent_{idx}", use_container_width=True):
                        st.session_state["ai_result"] = item.get("ai_result") or {}
                        st.session_state["edited_resume_text"] = item.get("edited_resume") or ""
                        st.rerun()
                st.caption("Job title")
                st.write(job_title)
                st.caption("Keywords")
                if keywords:
                    st.markdown(
                        "".join([f'<span class="keyword-chip">{keyword}</span>' for keyword in keywords]),
                        unsafe_allow_html=True,
                    )
                else:
                    st.write("No keywords found")

        if older_items:
            with st.expander("Show older previous versions"):
                for idx, item in enumerate(older_items, start=4):
                    saved_at = format_saved_at(item.get("saved_at", "Unknown"))
                    score = item.get("scores", {}).get("overall_score", 0)
                    company_name = extract_company_name(item.get("job_description", ""))
                    job_title, keywords = summarize_job_description(item.get("job_description", ""))
                    with st.expander(f"{idx}. {company_name} • {saved_at} • {score}%"):
                        btn_col1, btn_col2 = st.columns(2)
                        with btn_col1:
                            if st.button("Load into Scanner", key=f"load_older_{idx}", use_container_width=True):
                                st.session_state["resume_text_input"] = item.get("resume_text", "") or ""
                                st.session_state["job_description_input"] = item.get("job_description", "") or ""
                                st.session_state["submitted_resume_text"] = item.get("resume_text", "") or ""
                                st.session_state["submitted_job_description"] = item.get("job_description", "") or ""
                                st.rerun()
                        with btn_col2:
                            if st.button("Preview", key=f"preview_older_{idx}", use_container_width=True):
                                st.session_state["ai_result"] = item.get("ai_result") or {}
                                st.session_state["edited_resume_text"] = item.get("edited_resume") or ""
                                st.rerun()
                        st.caption("Job title")
                        st.write(job_title)
                        st.caption("Keywords")
                        if keywords:
                            st.markdown(
                                "".join([f'<span class="keyword-chip">{keyword}</span>' for keyword in keywords]),
                                unsafe_allow_html=True,
                            )
                        else:
                            st.write("No keywords found")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="settings-bottom">', unsafe_allow_html=True)
    with st.expander("Settings", expanded=False):
        st.selectbox(
            "Page Color",
            options=list(PAGE_COLOR_THEMES.keys()),
            key="page_color_theme",
            on_change=persist_page_theme_to_query_params,
        )
        st.markdown('<div class="settings-divider"></div>', unsafe_allow_html=True)
        default_hf_token = os.getenv("HF_TOKEN", "")
        try:
            default_hf_token = st.secrets.get("HF_TOKEN", default_hf_token)
        except Exception:
            pass
        hf_token = st.text_input("Hugging Face API Token", type="password",
                                 value=default_hf_token,
                                 help="Get a free token at https://huggingface.co/settings/tokens")
    st.markdown("</div>", unsafe_allow_html=True)


#extract the text from the resume
def extract_resume_text(uploaded_file):
    """
    Extract all text content from an uploaded PDF or DOCX resume.
    
    Args:
        uploaded_file: Streamlit UploadedFile object
    
    Returns:
        str: Extracted text, or None if error
    """
    # Get file extension to determine file type
    file_extension = uploaded_file.name.split(".")[-1].lower()
    
    # Initialize empty string to store extracted text
    extracted_text = ""
    
    try:
        # Route to PDF or DOCX logic based on file extension
        if file_extension == "pdf":
            # PDF extraction using pdfplumber
            with pdf.open(uploaded_file) as pdf_document:
                # Iterate through all pages in the PDF
                for page in pdf_document.pages:
                    # Extract text from current page
                    page_text = page.extract_text()
                    # Append extracted text to extracted_text variable
                    if page_text:
                        extracted_text += page_text + "\n"
        
        elif file_extension == "docx":
            # DOCX extraction using python-docx
            # Reset file pointer to beginning (Streamlit files need this)
            uploaded_file.seek(0)
            # Load the document
            doc = Document(uploaded_file)
            # Iterate through all paragraphs in the document
            for paragraph in doc.paragraphs:
                # Extract text from current paragraph
                if paragraph.text:
                    extracted_text += paragraph.text + "\n"
        
        else:
            # Unsupported file type
            st.error(f"Unsupported file type: {file_extension}. Please upload a PDF or DOCX file.")
            return None
        
        # Return the extracted text (cleaned/trimmed)
        return extracted_text.strip() if extracted_text else None
        
    except Exception as e:
        # Handle errors gracefully
        st.error(f"Error extracting text from resume: {e}")
        return None


# --------------------------------------------------
# Parse resume into sections
# --------------------------------------------------
def parse_sections(extracted_text):
    """
    Parse the extracted text into sections.

    Args:
        extracted_text: str - Extracted text from the resume

    Returns:
        dict: Parsed sections, or None if error
    """
    def is_overlapping(new_start, existing_matches):
        if not existing_matches:
            return False
        for match in existing_matches:
            start_index = match["start_index"]
            end_index = match["end_index"]
            if new_start >= start_index and new_start < end_index:
                return True
        return False
    skills = ["skills", "technical skills", "core competencies", "technical skills and competencies", "technical skills and abilities", "technical skills and qualifications", "technical skills and experience", "technical skills and education", "technical skills and training", "technical skills and certification", "technical skills and licenses", "technical skills and awards", "technical skills and achievements", "technical skills and accomplishments", "technical skills and contributions", "technical skills and responsibilities", "technical skills and duties", "technical skills and tasks", "technical skills and projects", "technical skills and responsibilities", "technical skills and duties", "technical skills and tasks", "technical skills and projects"]
    experience = ["experience", "work experience", "professional experience", "professional experience and responsibilities", "professional experience and duties", "professional experience and tasks", "professional experience and projects", "professional experience and responsibilities", "professional experience and duties", "professional experience and tasks", "professional experience and projects"]
    education = ["education", "academic", "academics", "academic background", "academic qualifications", "academic experience", "academic education", "academic training", "academic certification", "academic licenses", "academic awards", "academic achievements", "academic accomplishments", "academic contributions", "academic responsibilities", "academic duties", "academic tasks", "academic projects", "academic responsibilities", "academic duties", "academic tasks", "academic projects"]
    projects = ["projects", "personal projects", "projects and duties", "projects and tasks", "projects and projects", "projects and responsibilities", "projects and duties", "projects and tasks"]
    certifications = ["certifications", "certificates", "certification", "certificate", "certifications and certifications", "certifications and certificates", "certifications and certification", "certifications and certificate"]
    skills_headers = sorted(skills, key=len, reverse=True)
    experience_headers = sorted(experience, key=len, reverse=True)
    education_headers = sorted(education, key=len, reverse=True)
    projects_headers = sorted(projects, key=len, reverse=True)
    certifications_headers = sorted(certifications, key=len, reverse=True)
    original_text = extracted_text
    normalized_text = extracted_text.lower()
    sections_found = []
    for header in skills_headers:
        start_index = normalized_text.find(header)
        if start_index != -1:
            if not is_overlapping(start_index, sections_found):
                sections_found.append({"section_name": header, "start_index": start_index, "end_index": start_index + len(header), "matched_header": header})
    for header in experience_headers:
        start_index = normalized_text.find(header)
        if start_index != -1:
            if not is_overlapping(start_index, sections_found):
                sections_found.append({"section_name": header, "start_index": start_index, "end_index": start_index + len(header), "matched_header": header})
    for header in education_headers:
        start_index = normalized_text.find(header)
        if start_index != -1:
            if not is_overlapping(start_index, sections_found):
                sections_found.append({"section_name": header, "start_index": start_index, "end_index": start_index + len(header), "matched_header": header})
    for header in projects_headers:
        start_index = normalized_text.find(header)
        if start_index != -1:
            if not is_overlapping(start_index, sections_found):
                sections_found.append({"section_name": header, "start_index": start_index, "end_index": start_index + len(header), "matched_header": header})
    for header in certifications_headers:
        start_index = normalized_text.find(header)
        if start_index != -1:
            if not is_overlapping(start_index, sections_found):
                sections_found.append({"section_name": header, "start_index": start_index, "end_index": start_index + len(header), "matched_header": header})
    sections_found_sorted = sorted(sections_found, key=lambda x: x["start_index"])
    sections_dict = {}
    for i in range(len(sections_found_sorted)):
        start = sections_found_sorted[i]["start_index"]
        if i < len(sections_found_sorted) - 1:
            end = sections_found_sorted[i + 1]["start_index"]
        else:
            end = len(original_text)
        section_chunk = original_text[start:end].strip()
        sections_dict[sections_found_sorted[i]["section_name"]] = section_chunk
    return sections_dict


# --------------------------------------------------
# Scoring Engine
# --------------------------------------------------
def compute_scores(parsed_sections, job_description):
    """
    Compare resume (parsed sections) to job description and produce scores.
    Weights: technical/skills 50%, experience 30%, education 20%.
    """
    empty_result = {
        "skills_score": 0,
        "skills_matched": [],
        "skills_missing": [],
        "experience_score": 0,
        "education_score": 0,
        "overall_score": 0,
    }
    if not job_description or not job_description.strip():
        return empty_result
    stop_list = {"the", "and", "or", "a", "of", "to", "in", "for", "is", "on", "with"}
    min_word_len = 2
    job_lower = re.sub(r"[^\w\s]", "", job_description.lower()).strip()
    words = [w for w in job_lower.split() if w not in stop_list and len(w) > min_word_len]
    job_skills = set(words)
    skills_keys = ["skills", "technical skills", "core competencies"]
    resume_skills_text = ""
    for key in parsed_sections:
        if any(sk in key.lower() for sk in skills_keys):
            resume_skills_text = parsed_sections[key]
            break
    resume_lower = resume_skills_text.lower().strip()
    tokens = re.split(r"[,;\n]", resume_lower)
    resume_skills = {s.strip() for s in tokens if s.strip() and len(s.strip()) > 1}
    for js in list(job_skills):
        for rs in resume_skills:
            if js in rs or rs in js:
                resume_skills.add(js)
                break
    exp_keys = ["experience", "work experience", "professional experience"]
    edu_keys = ["education", "academic"]
    other_text = ""
    for key in parsed_sections:
        if any(k in key.lower() for k in exp_keys + edu_keys):
            other_text += " " + parsed_sections[key].lower()
    for js in job_skills:
        if js in other_text:
            resume_skills.add(js)
    matched_skills = job_skills & resume_skills
    skills_missing = job_skills - resume_skills
    skills_score_pct = round(100 * len(matched_skills) / len(job_skills)) if job_skills else 0
    experience_text = ""
    for key in parsed_sections:
        if any(k in key.lower() for k in exp_keys):
            experience_text = parsed_sections[key].lower()
            break
    matched_exp = sum(1 for s in job_skills if s in experience_text)
    experience_score = round(100 * matched_exp / len(job_skills)) if job_skills else 0
    education_text = ""
    for key in parsed_sections:
        if any(k in key.lower() for k in edu_keys):
            education_text = parsed_sections[key].lower()
            break
    degree_markers = ["bachelor", "master", "phd", "bs ", "ms ", "b.s.", "m.s.", "degree", "computer science"]
    education_score = 100 if any(m in education_text for m in degree_markers) else 0
    overall_score = round(0.5 * skills_score_pct + 0.3 * experience_score + 0.2 * education_score)
    return {
        "skills_score": skills_score_pct,
        "skills_matched": sorted(matched_skills),
        "skills_missing": sorted(skills_missing),
        "experience_score": experience_score,
        "education_score": education_score,
        "overall_score": min(100, overall_score),
    }


# --------------------------------------------------
# AI Analysis via Hugging Face Inference API
# --------------------------------------------------
AI_MODEL = "mistralai/Mistral-7B-Instruct-v0.2"

def analyze_with_ai(resume_text, job_description, token):
    if not token:
        return {"error": "Please enter your Hugging Face API token in the sidebar."}

    prompt = f"""You are an expert ATS resume analyst.

Analyze the resume against the job description below.
You MUST return ONLY a valid JSON object — no markdown, no backticks, no explanation, nothing else before or after the JSON.

The JSON must have exactly these keys:
{{
  "summary": "2-3 sentence overall assessment",
  "matched_skills": ["skill1", "skill2"],
  "missing_skills": ["skill1", "skill2"],
  "strengths": ["point1", "point2", "point3"],
  "improvements": ["suggestion1", "suggestion2", "suggestion3"],
  "fit_rating": 7
}}

---RESUME---
{resume_text[:4000]}

---JOB DESCRIPTION---
{job_description[:2000]}

Remember: Return ONLY the JSON object. Start your response with {{ and end with }}"""

    try:
        client = InferenceClient(token=token)
        response = client.chat_completion(
            model=AI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=2048,
            temperature=0.1,  # lower = more deterministic
        )
        raw = response.choices[0].message.content.strip()

        # Attempt 1: try parsing as-is
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            pass

        # Attempt 2: extract content between first { and last }
        start = raw.find("{")
        end = raw.rfind("}") + 1
        if start != -1 and end > start:
            try:
                return json.loads(raw[start:end])
            except json.JSONDecodeError:
                pass

        # Attempt 3: strip markdown code blocks
        cleaned = re.sub(r"```(?:json)?", "", raw).replace("```", "").strip()
        try:
            return json.loads(cleaned)
        except json.JSONDecodeError:
            pass

        return {"error": "The AI returned an unparsable response. Try again.", "raw": raw}

    except Exception as e:
        return {"error": str(e)}


def format_generated_resume_text(resume_text):
    text = (resume_text or "").replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"```(?:text|markdown)?", "", text).replace("```", "")

    raw_lines = [line.rstrip() for line in text.split("\n")]
    formatted_lines = []

    for line in raw_lines:
        stripped = line.strip()

        if not stripped:
            formatted_lines.append("")
            continue

        bullet_match = re.match(r"^[-•*]\s*(.+)$", stripped)
        if bullet_match:
            formatted_lines.append(f"- {bullet_match.group(1).strip()}")
            continue

        heading_like = (
            stripped.endswith(":")
            or (stripped.isupper() and len(stripped.split()) <= 6)
        )
        if heading_like:
            if formatted_lines and formatted_lines[-1] != "":
                formatted_lines.append("")
            formatted_lines.append(stripped)
            formatted_lines.append("")
            continue

        formatted_lines.append(stripped)

    cleaned_text = "\n".join(formatted_lines)
    cleaned_text = re.sub(r"\n{3,}", "\n\n", cleaned_text).strip()
    return cleaned_text


def convert_resume_to_rich_markdown(resume_text):
    text = (resume_text or "").strip()
    if not text:
        return ""

    lines = [line.rstrip() for line in text.split("\n")]
    result_lines = []
    first_content_seen = False

    heading_keywords = {
        "summary", "professional summary", "experience", "work experience", "education",
        "projects", "skills", "technical skills", "certifications", "awards", "leadership",
        "publications", "volunteer", "activities", "interests",
    }

    def linkify(value):
        value = re.sub(r"\b(https?://[^\s)]+)", r"[\1](\1)", value)
        value = re.sub(r"\b(www\.[^\s)]+)", r"[\1](https://\1)", value)
        value = re.sub(r"\b([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,})\b", r"[\1](mailto:\1)", value)
        return value

    for raw_line in lines:
        stripped = raw_line.strip()

        if not stripped:
            if result_lines and result_lines[-1] != "":
                result_lines.append("")
            continue

        normalized = stripped.lower().rstrip(":")
        bullet_match = re.match(r"^[-•*]\s*(.+)$", stripped)

        if not first_content_seen:
            first_content_seen = True
            if len(stripped.split()) <= 6 and "," not in stripped and "|" not in stripped:
                result_lines.append(f"# **{stripped}**")
                continue

        if normalized in heading_keywords or (stripped.isupper() and len(stripped.split()) <= 5):
            if result_lines and result_lines[-1] != "":
                result_lines.append("")
            result_lines.append(f"## **{stripped.rstrip(':')}**")
            result_lines.append("")
            continue

        if bullet_match:
            bullet_text = linkify(bullet_match.group(1).strip())
            result_lines.append(f"- {bullet_text}")
            continue

        if "|" in stripped:
            parts = [part.strip() for part in stripped.split("|") if part.strip()]
            if len(parts) >= 2:
                first = linkify(parts[0])
                second = linkify(parts[1])
                rest = " | ".join(linkify(part) for part in parts[2:])
                formatted_line = f"**{first}** | *{second}*"
                if rest:
                    formatted_line += f" | {rest}"
                result_lines.append(formatted_line)
                continue

        result_lines.append(linkify(stripped))

    markdown_text = "\n".join(result_lines)
    markdown_text = re.sub(r"\n{3,}", "\n\n", markdown_text).strip()
    return markdown_text


def generate_edited_resume(resume_text, job_description, token):
    if not token:
        return {"error": "Please enter your Hugging Face API token in the sidebar."}

    prompt = f"""You are an expert resume writer and ATS optimization specialist.

Rewrite the candidate's resume so it is highly compatible with the job description while remaining truthful to the candidate's real experience and skills.

FORMAT LOCK (must follow):
- Keep the exact same section order.
- Keep the same section headings text unless a heading is clearly incorrect.
- Keep the same bullet style (e.g., '-', '•') and indentation.
- Keep line-break structure as close as possible to the original.
- Keep date/location/company line layout in the same style as the original.
- Keep the same approximate number of bullets per role/project.
- Do NOT convert to a new template.

CONTENT RULES:
- Do not invent fake companies, roles, dates, degrees, certifications, or metrics.
- Keep the resume professional, concise, and ATS-friendly.
- Prioritize keywords from the job description only when they match the candidate's actual background.
- Improve wording for impact, clarity, and relevance while preserving the original structure.

Return ONLY the edited resume text in plain text format.

---ORIGINAL RESUME---
{resume_text[:5000]}

---JOB DESCRIPTION---
{job_description[:3000]}
"""

    try:
        client = InferenceClient(token=token)
        response = client.chat_completion(
            model=AI_MODEL,
            messages=[{"role": "user", "content": prompt}],
            max_tokens=2048,
            temperature=0.1,
        )
        raw = response.choices[0].message.content.strip()
        cleaned = re.sub(r"```(?:text|markdown)?", "", raw).replace("```", "").strip()
        formatted_plain = format_generated_resume_text(cleaned)
        formatted_markdown = convert_resume_to_rich_markdown(formatted_plain)
        return {
            "edited_resume": formatted_plain,
            "edited_resume_markdown": formatted_markdown,
        }
    except Exception as e:
        return {"error": str(e)}


# ============================================
# STEP 5: MAIN AREA LOGIC
# ============================================
tab_original, tab_edited = st.tabs(["Original Resume and Job description", "Edited Version"])

if "submitted_resume_text" not in st.session_state:
    st.session_state["submitted_resume_text"] = ""
if "submitted_job_description" not in st.session_state:
    st.session_state["submitted_job_description"] = ""

with tab_original:
    components.html(
        """
        <script>
        const parentWindow = window.parent;
        const doc = parentWindow.document;
        const listenerVersion = 2;

        if (parentWindow.__doubleEnterSubmitVersion !== listenerVersion) {
            if (parentWindow.__doubleEnterSubmitHandler) {
                doc.removeEventListener("keydown", parentWindow.__doubleEnterSubmitHandler, true);
            }

            parentWindow.__lastEnterPressedAt = 0;

            parentWindow.__doubleEnterSubmitHandler = function (event) {
                if (event.key !== "Enter") return;

                const activeEl = doc.activeElement;
                if (!activeEl || activeEl.tagName.toLowerCase() !== "textarea") return;

                const now = Date.now();
                const prev = parentWindow.__lastEnterPressedAt || 0;
                parentWindow.__lastEnterPressedAt = now;

                if (now - prev <= 1000) {
                    event.preventDefault();

                    const formEl = activeEl.closest("form");
                    if (formEl && typeof formEl.requestSubmit === "function") {
                        formEl.requestSubmit();
                        return;
                    }

                    const formSubmitButton = doc.querySelector('div[data-testid="stFormSubmitButton"] button');
                    if (formSubmitButton) {
                        formSubmitButton.click();
                        return;
                    }

                    const enterButton = Array.from(doc.querySelectorAll("button")).find(
                        (button) => button.innerText && button.innerText.trim() === "Enter"
                    );
                    if (enterButton) enterButton.click();
                }
            };

            doc.addEventListener("keydown", parentWindow.__doubleEnterSubmitHandler, true);
            parentWindow.__doubleEnterSubmitVersion = listenerVersion;
        }
        </script>
        """,
        height=0,
    )

    with st.form("resume_job_form", clear_on_submit=False):
        left_col, right_col = st.columns(2)
        with left_col:
            st.markdown("### Paste Original Resume")
            draft_resume_text = st.text_area(
                "Paste Original Resume",
                height=360,
                placeholder="Copy and paste the full resume text here...",
                key="resume_text_input",
                label_visibility="collapsed",
            )
        with right_col:
            st.markdown("### Paste Job Description")
            draft_job_description = st.text_area(
                "Paste Job Description",
                height=360,
                placeholder="Copy and paste the job description here...",
                key="job_description_input",
                label_visibility="collapsed",
            )

        submitted_inputs = st.form_submit_button("Enter", use_container_width=True)

    if submitted_inputs:
        st.session_state["submitted_resume_text"] = draft_resume_text
        st.session_state["submitted_job_description"] = draft_job_description

        draft_resume_clean = draft_resume_text.strip()
        draft_job_clean = draft_job_description.strip()
        if draft_resume_clean and draft_job_clean:
            submit_scores = compute_scores(parse_sections(draft_resume_clean), draft_job_clean)
            add_history_entry(
                resume_text=draft_resume_text,
                job_description=draft_job_description,
                scores=submit_scores,
                ai_result=st.session_state.get("ai_result"),
                edited_resume=st.session_state.get("edited_resume_text"),
            )
            st.rerun()

    resume_text = st.session_state.get("submitted_resume_text", "")
    job_description = st.session_state.get("submitted_job_description", "")
    resume_text_clean = resume_text.strip()
    job_description_clean = job_description.strip()

    st.caption("Click Enter to submit both fields. Double Enter quickly also submits.")
    st.caption("Pasted formatting (line breaks, bullets, spacing) is preserved.")

    compatibility_score = 0
    if resume_text_clean:
        preview_sections = parse_sections(resume_text_clean)
        compatibility_score = compute_scores(preview_sections, job_description_clean)["overall_score"]

    st.markdown('<div class="compatibility-block">', unsafe_allow_html=True)
    st.subheader("Compatibility with Job Description")
    st.metric("Compatibility", f"{compatibility_score}%")
    st.markdown("</div>", unsafe_allow_html=True)

    if resume_text_clean:
        extracted_text = resume_text_clean
        parsed_sections = parse_sections(extracted_text)
        scores = compute_scores(parsed_sections, job_description_clean)

        st.subheader("Scoring Engine")
        c1, c2, c3, c4 = st.columns(4)
        with c1:
            st.metric("Skills Match", f"{scores['skills_score']}%")
        with c2:
            st.metric("Experience Relevance", f"{scores['experience_score']}%")
        with c3:
            st.metric("Education", f"{scores['education_score']}%")
        with c4:
            st.metric("Overall Fit", f"{scores['overall_score']}%")

        with st.expander("Skills matched"):
            st.write(scores["skills_matched"] if scores["skills_matched"] else "None")
        with st.expander("Skills missing (from job description)"):
            st.write(scores["skills_missing"] if scores["skills_missing"] else "None")

        has_job_desc = bool(job_description_clean)
        if not has_job_desc:
            st.warning("Enter a job description to see match and skill gap.")
        else:
            st.subheader("Match & skill gap")
            overall = scores["overall_score"]
            if overall >= 70:
                gauge_color = "#22c55e"
            elif overall >= 40:
                gauge_color = "#eab308"
            else:
                gauge_color = "#ef4444"
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=overall,
                number={"suffix": "%", "font": {"size": 32}},
                gauge={
                    "axis": {"range": [0, 100], "tickwidth": 1},
                    "bar": {"color": gauge_color},
                    "bgcolor": "white",
                    "borderwidth": 2,
                    "bordercolor": "gray",
                    "steps": [
                        {"range": [0, 40], "color": "#fecaca"},
                        {"range": [40, 70], "color": "#fef08a"},
                        {"range": [70, 100], "color": "#bbf7d0"},
                    ],
                    "threshold": {
                        "line": {"color": gauge_color, "width": 4},
                        "thickness": 0.75,
                        "value": overall,
                    },
                },
                title={"text": "Overall fit", "font": {"size": 18}},
            ))
            fig_gauge.update_layout(
                height=280,
                margin=dict(l=20, r=20, t=50, b=20),
                paper_bgcolor="rgba(0,0,0,0)",
                font=dict(size=12),
            )
            st.plotly_chart(fig_gauge, use_container_width=True)

            matched = (scores["skills_matched"] or [])[:15]
            missing = (scores["skills_missing"] or [])[:15]
            if matched or missing:
                rows = []
                for s in matched:
                    rows.append({"skill": s, "count": 1, "status": "Matched"})
                for s in missing:
                    rows.append({"skill": s, "count": 1, "status": "Missing"})
                df_gap = pd.DataFrame(rows)
                if not df_gap.empty:
                    fig_gap = px.bar(
                        df_gap,
                        x="count",
                        y="skill",
                        orientation="h",
                        color="status",
                        color_discrete_map={"Matched": "#22c55e", "Missing": "#ef4444"},
                        category_orders={"status": ["Matched", "Missing"]},
                    )
                    fig_gap.update_layout(
                        title="Skill gap (job wants vs resume)",
                        xaxis_title="",
                        yaxis_title="",
                        yaxis={"categoryorder": "total ascending"},
                        height=max(300, 24 * len(df_gap)),
                        margin=dict(l=20, r=20, t=40, b=20),
                        showlegend=True,
                        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
                    )
                    fig_gap.update_xaxes(showticklabels=False)
                    st.plotly_chart(fig_gap, use_container_width=True)
            else:
                st.caption("No skill keywords found in job description to compare.")

            sub_names = ["Skills", "Experience", "Education"]
            sub_values = [
                scores["skills_score"],
                scores["experience_score"],
                scores["education_score"],
            ]
            fig_sub = go.Figure(
                go.Bar(
                    x=sub_names,
                    y=sub_values,
                    marker_color=["#3b82f6", "#8b5cf6", "#06b6d4"],
                    text=[f"{v}%" for v in sub_values],
                    textposition="outside",
                )
            )
            fig_sub.update_layout(
                title="Score breakdown (0–100)",
                yaxis=dict(range=[0, 100], title="Score %"),
                height=320,
                margin=dict(l=20, r=20, t=40, b=20),
                showlegend=False,
            )
            st.plotly_chart(fig_sub, use_container_width=True)

        st.write("---")
        st.subheader("AI-Powered Analysis")

        if not has_job_desc:
            st.warning("Enter a job description to enable AI analysis.")
        else:
            if st.button("Run AI Analysis", type="primary", use_container_width=True):
                with st.spinner(f"Analyzing with {AI_MODEL}... this may take a moment."):
                    ai_result = analyze_with_ai(extracted_text, job_description, hf_token)

                if "error" in ai_result:
                    st.error(f"AI analysis failed: {ai_result['error']}")
                    if "raw" in ai_result:
                        with st.expander("Raw AI response"):
                            st.code(ai_result["raw"])
                else:
                    st.session_state["ai_result"] = ai_result

                add_history_entry(
                    resume_text=resume_text,
                    job_description=job_description,
                    scores=scores,
                    ai_result=ai_result,
                    edited_resume=st.session_state.get("edited_resume_text"),
                )

            if "ai_result" in st.session_state:
                ai = st.session_state["ai_result"]

                fit = ai.get("fit_rating", "?")
                if isinstance(fit, int):
                    if fit >= 7:
                        fit_color = "green"
                    elif fit >= 4:
                        fit_color = "orange"
                    else:
                        fit_color = "red"
                    st.markdown(f"### AI Fit Rating: :{fit_color}[{fit} / 10]")

                st.markdown(f"**Summary:** {ai.get('summary', 'N/A')}")

                col_left, col_right = st.columns(2)
                with col_left:
                    st.markdown("**Matched Skills (AI)**")
                    for skill in ai.get("matched_skills", []):
                        st.markdown(f"- :green[{skill}]")
                with col_right:
                    st.markdown("**Missing Skills (AI)**")
                    for skill in ai.get("missing_skills", []):
                        st.markdown(f"- :red[{skill}]")

                st.markdown("---")
                col_s, col_i = st.columns(2)
                with col_s:
                    st.markdown("**Strengths**")
                    for point in ai.get("strengths", []):
                        st.markdown(f"- {point}")
                with col_i:
                    st.markdown("**Suggested Improvements**")
                    for point in ai.get("improvements", []):
                        st.markdown(f"- {point}")

        st.write("---")
        with st.expander("Raw resume text"):
            st.text_area("Raw Resume", value=resume_text, height=220, disabled=True, key="raw_resume_view")
    else:
        st.info("Paste your resume text in the left panel to begin.")

with tab_edited:
    st.subheader("Edited Version")
    has_resume = bool(resume_text_clean)
    has_job_desc = bool(job_description_clean)

    if not has_resume or not has_job_desc:
        st.warning("Paste both the original resume and job description in the first tab to generate an edited version.")
    else:
        source_signature = hashlib.sha256(
            f"{resume_text_clean}\n---\n{job_description_clean}".encode("utf-8")
        ).hexdigest()
        last_generated_signature = st.session_state.get("edited_resume_source_signature")
        last_attempted_signature = st.session_state.get("edited_resume_attempted_signature")

        needs_auto_generation = (
            source_signature != last_generated_signature
            and source_signature != last_attempted_signature
        )

        if needs_auto_generation:
            st.session_state["edited_resume_attempted_signature"] = source_signature
            with st.spinner("Creating an edited resume version tailored to the job description..."):
                edited_result = generate_edited_resume(resume_text_clean, job_description_clean, hf_token)

            if "error" in edited_result:
                st.error(f"Edited resume generation failed: {edited_result['error']}")
            else:
                st.session_state["edited_resume_text"] = edited_result["edited_resume"]
                st.session_state["edited_resume_source_signature"] = source_signature

                history_scores = compute_scores(parse_sections(resume_text_clean), job_description_clean)
                add_history_entry(
                    resume_text=resume_text,
                    job_description=job_description,
                    scores=history_scores,
                    ai_result=st.session_state.get("ai_result"),
                    edited_resume=edited_result.get("edited_resume"),
                )

        if st.button("Regenerate Edited Resume", type="primary", use_container_width=True, key="generate_edited_resume_btn"):
            st.session_state.pop("edited_resume_attempted_signature", None)
            st.session_state.pop("edited_resume_source_signature", None)
            st.rerun()

        if "edited_resume_text" in st.session_state:
            st.markdown("### Edited Resume Output (PDF Preview)")
            edited_pdf_bytes = resume_text_to_pdf_bytes(
                st.session_state["edited_resume_text"],
                fit_one_page=True,
            )
            render_pdf_preview(edited_pdf_bytes)

            st.download_button(
                label="Download Edited Resume (.pdf)",
                data=edited_pdf_bytes,
                file_name="edited_resume.pdf",
                mime="application/pdf",
                use_container_width=True,
            )




