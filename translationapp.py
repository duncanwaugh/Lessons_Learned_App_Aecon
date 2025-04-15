import os
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from dotenv import load_dotenv
from openai import OpenAI

# Load environment variables
load_dotenv()

# Set page config and style
st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘", layout="centered")

# Display Aecon logo and style the title
st.image("AECON.png", width=300)
st.markdown("""
    <style>
        .main {
            background-color: #ffffff;
        }
        .stApp {
            background-color: #ffffff;
        }
        h1, h2, h3, h4, h5 {
            color: #c8102e;
        }
        .stButton>button {
            background-color: #c8102e;
            color: white;
            border: None;
        }
        .stDownloadButton>button {
            background-color: #c8102e;
            color: white;
        }
    </style>
""", unsafe_allow_html=True)

# Set up OpenAI client
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

# Extract text and images from PPTX
def extract_text_and_images_from_pptx(file_path):
    prs = Presentation(file_path)
    text = ''
    image_paths = []
    os.makedirs("images", exist_ok=True)

    for i, slide in enumerate(prs.slides):
        for shape in slide.shapes:
            if shape.has_text_frame:
                text += shape.text_frame.text + "\n"
            elif shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                image = shape.image
                ext = image.ext
                image_bytes = image.blob
                image_path = f"images/slide_{i+1}_{len(image_paths)}.{ext}"
                with open(image_path, 'wb') as img_file:
                    img_file.write(image_bytes)
                image_paths.append(image_path)

    return text, image_paths

# Generate summary and extracted sections using OpenAI GPT
def summarize_and_extract(text):
    prompt = f"""
    You are helping prepare a standardized Lessons Learned document from a serious incident.

    Please extract and clearly label the following sections. Each section label should appear on its own line and be followed by its content. Separate sections with one blank line.

    Use these exact labels:
    Title:
    Aecon Business Sector:
    Project/Location:
    Date of Event:
    Event Type:
    Event Summary:
    Contributing Factors:
    Lessons Learned:

    Here is the presentation text:
    {text}
    """

    response = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=1000,
    )

    return response.choices[0].message.content.strip()

# Translate generated content to French (Canadian) using OpenAI
def translate_to_french_openai(text):
    prompt = f"""
    Translate the following safety incident summary into professional French Canadian. Keep the formatting, section headers, and bullet points intact.

    Text to translate:
    {text}
    """

    response = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=1000,
    )

    return response.choices[0].message.content.strip()

# Translate using DeepL API
def translate_to_french_deepl(text):
    api_key = os.getenv("DEEPL_API_KEY")
    if not api_key:
        st.error("DeepL API key not found. Please add it to your .env file.")
        return text

    response = requests.post(
        "https://api-free.deepl.com/v2/translate",
        data={
            "auth_key": api_key,
            "text": text,
            "target_lang": "FR",
            "formality": "more"
        }
    )

    try:
        return response.json()["translations"][0]["text"]
    except Exception as e:
        st.error(f"DeepL translation failed: {e}")
        return text

# Parse structured content into sections
def parse_sections(content):
    import re
    sections = {}
    current_section = None
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.endswith(':') and len(line.split()) < 6:
            current_section = line[:-1].strip()
            sections[current_section] = []
        elif current_section:
            sections[current_section].append(line)
    return sections

# Generate English Word Document
def create_lessons_learned_doc(content, output_path, image_paths=None):
    doc = Document()
    sections = parse_sections(content)

    doc.add_heading(sections.get("Title", ["Lessons Learned Handout"])[0], level=1)

    doc.add_heading("Aecon Business Sector", level=2)
    doc.add_paragraph(' '.join(sections.get("Aecon Business Sector", [])))

    doc.add_heading("Project/Location", level=2)
    doc.add_paragraph(' '.join(sections.get("Project/Location", [])))

    doc.add_heading("Date of Event", level=2)
    doc.add_paragraph(' '.join(sections.get("Date of Event", [])))

    doc.add_heading("Event Type", level=2)
    doc.add_paragraph(' '.join(sections.get("Event Type", [])))

    doc.add_heading("Event Summary", level=2)
    doc.add_paragraph(' '.join(sections.get("Event Summary", [])))

    doc.add_heading("Contributing Factors", level=2)
    for item in sections.get("Contributing Factors", []):
        doc.add_paragraph(item, style='List Bullet')

    doc.add_heading("Lessons Learned", level=2)
    for item in sections.get("Lessons Learned", []):
        doc.add_paragraph(item, style='List Bullet')

    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        table = doc.add_table(rows=0, cols=2)
        table.autofit = True
        for i in range(0, len(image_paths), 2):
            row_cells = table.add_row().cells
            for j in range(2):
                if i + j < len(image_paths):
                    try:
                        paragraph = row_cells[j].paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(image_paths[i + j], width=Inches(2.5))
                    except Exception:
                        row_cells[j].text = "⚠️ Error loading image"

    doc.save(output_path)

# Generate French-specific Word Document
def create_lessons_learned_doc_fr(content, output_path, image_paths=None):
    doc = Document()
    sections = parse_sections(content)

    def get_section(possible_names):
        for name in possible_names:
            if name in sections:
                return sections[name]
        return []

    doc.add_heading(get_section(["Titre", "Title", "EXAMEN D'UN EVENEMENT GRAVE"])[0], level=1)

    doc.add_heading("Secteur d'activité d'Aecon", level=2)
    doc.add_paragraph(' '.join(get_section(["Secteur d'activité d'Aecon", "Secteur d’activité d’Aecon"])))

    doc.add_heading("Projet/Emplacement", level=2)
    doc.add_paragraph(' '.join(get_section(["Projet/Emplacement", "Projet/Lieu"])))

    doc.add_heading("Date de l'événement", level=2)
    doc.add_paragraph(' '.join(get_section(["Date de l'événement"])))

    doc.add_heading("Type d'événement", level=2)
    doc.add_paragraph(' '.join(get_section(["Type d'événement"])))

    doc.add_heading("Résumé de l'événement", level=2)
    doc.add_paragraph(' '.join(get_section(["Résumé de l'événement"])))

    doc.add_heading("Facteurs contributifs", level=2)
    for item in get_section(["Facteurs contributifs"]):
        doc.add_paragraph(item, style='List Bullet')

    doc.add_heading("Leçons apprises", level=2)
    for item in get_section(["Leçons apprises", "Leçons tirées"]):
        doc.add_paragraph(item, style='List Bullet')

    if image_paths:
        doc.add_page_break()
        doc.add_heading("Photos de soutien", level=2)
        table = doc.add_table(rows=0, cols=2)
        table.autofit = True
        for i in range(0, len(image_paths), 2):
            row_cells = table.add_row().cells
            for j in range(2):
                if i + j < len(image_paths):
                    try:
                        paragraph = row_cells[j].paragraphs[0]
                        run = paragraph.add_run()
                        run.add_picture(image_paths[i + j], width=Inches(2.5))
                    except Exception:
                        row_cells[j].text = "⚠️ Erreur de chargement de l'image"

    doc.save(output_path)

# Streamlit UI
st.title('🦺 Serious Event Lessons Learned Generator')

uploaded_file = st.file_uploader("Upload Executive Review PPTX", type=['pptx'])
language = st.selectbox("Choose report language:", ["English", "French (Canadian)"])
translator = None
if language == "French (Canadian)":
    translator = st.radio("Choose translation method:", ["OpenAI (using my account, cost me 50 cents so far and I've probably ran it 200ish times)", "DeepL (using free version, 500 000 characters/month)"])

if uploaded_file:
    if st.button("📄 Generate Lessons Learned Document"):
        input_filepath = f'input/{uploaded_file.name}'
        os.makedirs("input", exist_ok=True)
        with open(input_filepath, 'wb') as f:
            f.write(uploaded_file.getbuffer())

        with st.spinner("Extracting content from presentation..."):
            pptx_text, extracted_images = extract_text_and_images_from_pptx(input_filepath)

        with st.spinner("Generating Lessons Learned summary..."):
            generated_content = summarize_and_extract(pptx_text)

        if language == "French (Canadian)":
            with st.spinner("Translating to French (Canadian)..."):
                if translator == "OpenAI":
                    generated_content = translate_to_french_openai(generated_content)
                else:
                    generated_content = translate_to_french_deepl(generated_content)

        st.success("✅ Generation Complete!")
        st.text_area("📝 Generated Content:", generated_content, height=300)

        output_path = f'generated_lessons_learned_{"fr" if language.startswith("French") else "en"}.docx'
        if language == "French (Canadian)":
            create_lessons_learned_doc_fr(generated_content, output_path, extracted_images)
        else:
            create_lessons_learned_doc(generated_content, output_path, extracted_images)

        with open(output_path, "rb") as file:
            st.download_button(
                label="📥 Download Lessons Learned DOCX",
                data=file,
                file_name=os.path.basename(output_path),
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
