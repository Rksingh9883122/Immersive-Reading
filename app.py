import streamlit as st
from PyPDF2 import PdfReader
from docx import Document
from pptx import Presentation
import tempfile
from pdf2image import convert_from_bytes
import pytesseract
import pytesseract

pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\rsingh6\Downloads\tesseract-ocr-w64-setup-5.5.0.20241111.exe"
)


# -------------------------
# Page Configuration
# -------------------------

st.set_page_config(
    page_title="Immersive Reader",
    layout="wide"
)

# -------------------------
# Extractors
# -------------------------

def ocr_pdf(pdf_file):

    text = ""

    images = convert_from_bytes(
        pdf_file.read()
    )

    for image in images:

        text += pytesseract.image_to_string(
            image
        )

        text += "\n"

    return text


def extract_pdf(file):

    text = ""

    reader = PdfReader(file)

    for page in reader.pages:

        page_text = page.extract_text()

        if page_text:
            text += page_text + "\n"

    return text


def extract_docx(file):

    doc = Document(file)

    return "\n".join(
        para.text
        for para in doc.paragraphs
    )


def extract_pptx(file):

    prs = Presentation(file)

    text = ""

    for slide in prs.slides:

        for shape in slide.shapes:

            if hasattr(shape, "text"):
                text += shape.text + "\n"

    return text


def extract_txt(file):

    return file.read().decode("utf-8")


# -------------------------
# Sidebar
# -------------------------

st.sidebar.title("Reader Settings")

font_size = st.sidebar.slider(
    "Font Size",
    16,
    48,
    28
)

font_color = st.sidebar.color_picker(
    "Font Color",
    "#FFFFFF"
)

background = st.sidebar.selectbox(
    "Theme",
    [
        "Dark",
        "Light",
        "Sepia",
        "Blue"
    ]
)

# Theme Colors

themes = {
    "Dark": "#0F172A",
    "Light": "#FFFFFF",
    "Sepia": "#F5ECD9",
    "Blue": "#E0F2FE"
}

bg_color = themes[background]

# -------------------------
# Title
# -------------------------

st.title("📖 Immersive Reading Assistant")

uploaded_file = st.file_uploader(
    "Upload PDF, DOCX, PPTX, TXT",
    type=["pdf", "docx", "pptx", "txt"]
)

# -------------------------
# File Processing
# -------------------------

if uploaded_file:

    ext = uploaded_file.name.split(".")[-1].lower()

    if ext == "pdf":

        text = extract_pdf(uploaded_file)

        if not text.strip():

            st.info(
                "Scanned PDF detected. Running OCR..."
            )

            uploaded_file.seek(0)

            text = ocr_pdf(uploaded_file)

    elif ext == "docx":
        text = extract_docx(uploaded_file)

    elif ext == "pptx":
        text = extract_pptx(uploaded_file)

    elif ext == "txt":
        text = extract_txt(uploaded_file)

    else:
        text = ""

paragraphs = [
    p.strip()
    for p in text.split("\n")
    if p.strip()
]

if not paragraphs:
    st.error(
        "No text could be extracted from the document."
    )
    st.stop()

if len(paragraphs) == 0:
    st.error(
        "No readable text found in the document. "
        "This PDF may be scanned/image-based."
    )
    st.stop()

    if "current_para" not in st.session_state:
        st.session_state.current_para = 0

    col1, col2, col3 = st.columns(3)

    with col1:

        if st.button("⬅ Previous"):

            if st.session_state.current_para > 0:
                st.session_state.current_para -= 1

    with col2:

        st.write(
            f"Paragraph {st.session_state.current_para + 1} / {len(paragraphs)}"
        )

    with col3:

        if st.button("Next ➡"):

            if st.session_state.current_para < len(paragraphs) - 1:
                st.session_state.current_para += 1

    current_text = paragraphs[
        st.session_state.current_para
    ]

    st.markdown(
        f"""
        <div style="
        background-color:{bg_color};
        padding:40px;
        border-radius:15px;
        min-height:400px;
        ">

        <p style="
        font-size:{font_size}px;
        color:{font_color};
        line-height:2;
        text-align:justify;
        ">

        {current_text}

        </p>

        </div>
        """,
        unsafe_allow_html=True
    )

    # Read Aloud

    tts_html = f"""
    <script>

    function readText() {{

        let msg =
        new SpeechSynthesisUtterance(
            `{current_text}`
        );

        msg.rate = 1.0;
        msg.pitch = 1.0;

        window.speechSynthesis.speak(msg);
    }}

    function stopReading() {{

        window.speechSynthesis.cancel();

    }}

    </script>

    <button onclick="readText()">
    🔊 Read Aloud
    </button>

    <button onclick="stopReading()">
    ⏹ Stop
    </button>

    """

    st.components.v1.html(
        tts_html,
        height=80
    )