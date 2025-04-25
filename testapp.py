import os
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx.shared import Inches
from docxtpl import DocxTemplate, InlineImage
from dotenv import load_dotenv
from openai import OpenAI
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches

# ─── Load environment variables & secrets ──────────────────────────────────────
load_dotenv()
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
DEEPL_KEY  = os.getenv("DEEPL_API_KEY")
if not OPENAI_KEY:
    st.error("OpenAI API key not found! Please set OPENAI_API_KEY in Streamlit secrets or .env")
client = OpenAI(api_key=OPENAI_KEY)

# ─── Utility Functions ────────────────────────────────────────────────────────
def parse_sections(content: str) -> dict:
    """Split GPT output into labeled sections."""
    sections, current = {}, None
    for line in content.splitlines():
        line = line.strip()
        if not line:
            continue
        if line.endswith(':') and len(line.split()) < 6:
            current = line[:-1].strip()
            sections[current] = []
        elif current:
            sections[current].append(line)
    return sections

def extract_text_and_images_from_pptx(path: str):
    """Pull all text and picture blobs from a PPTX."""
    prs = Presentation(path)
    text, images = "", []
    os.makedirs("images", exist_ok=True)
    for idx, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if shp.has_text_frame:
                text += shp.text_frame.text + "\n"
            elif shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = shp.image.blob
                ext  = shp.image.ext
                fn   = f"images/slide_{idx+1}_{len(images)}.{ext}"
                with open(fn, "wb") as imgf:
                    imgf.write(blob)
                images.append(fn)
    return text, images

def summarize_and_extract(text: str) -> str:
    """Call OpenAI to extract labeled sections from raw incident text."""
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
    resp = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return resp.choices[0].message.content.strip()

def translate_to_french_openai(text: str) -> str:
    """Translate text to French Canadian via OpenAI."""
    prompt = f"""
Translate the following safety incident summary into professional French Canadian. Keep formatting intact.

Text to translate:
{text}
"""
    resp = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role": "user", "content": prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return resp.choices[0].message.content.strip()

def translate_to_french_deepl(text: str) -> str:
    """Translate text to French Canadian via DeepL."""
    if not DEEPL_KEY:
        st.error("DeepL API key not found! Please set DEEPL_API_KEY.")
        return text
    resp = requests.post(
        "https://api-free.deepl.com/v2/translate",
        data={
            "auth_key": DEEPL_KEY,
            "text": text,
            "target_lang": "FR",
            "formality": "more"
        }
    )
    try:
        return resp.json()["translations"][0]["text"]
    except:
        st.error("DeepL translation failed")
        return text

def render_with_docxtpl(sections, template_path, output_path, image_paths):
    # 1) Fill in the Jinja placeholders in your docx template
    tpl = DocxTemplate(template_path)
    context = {
        "TITLE":      sections.get("Title", [""])[0],
        "SECTOR":     " ".join(sections.get("Aecon Business Sector", [])),
        "PROJECT":    " ".join(sections.get("Project/Location", [])),
        "DATE":       " ".join(sections.get("Date of Event", [])),
        "EVENT_TYPE": " ".join(sections.get("Event Type", [])),
        "SUMMARY":    " ".join(sections.get("Event Summary", [])),
        "FACTORS":    sections.get("Contributing Factors", []),
        "LESSONS":    sections.get("Lessons Learned", []),
        # We no longer pass IMAGES into Jinja
    }
    tpl.render(context)
    tpl.save(output_path)

    # 2) Now re-open that file with python-docx and tack on the images
    doc = Document(output_path)
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        table = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            row = table.add_row().cells
            for j in (0,1):
                if i + j < len(image_paths):
                    try:
                        row[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except Exception:
                        row[j].text = "[Image failed to load]"
    doc.save(output_path)


def create_lessons_learned_doc(content: str, output_path: str, image_paths=None):
    """Fallback manual English DOCX generator (python-docx)."""
    sections = parse_sections(content)
    doc = Document()
    doc.add_heading(sections.get("Title", ["Lessons Learned Handout"])[0], level=1)
    doc.add_heading("Aecon Business Sector", level=2)
    doc.add_paragraph(" ".join(sections.get("Aecon Business Sector", [])))
    doc.add_heading("Project/Location", level=2)
    doc.add_paragraph(" ".join(sections.get("Project/Location", [])))
    doc.add_heading("Date of Event", level=2)
    doc.add_paragraph(" ".join(sections.get("Date of Event", [])))
    doc.add_heading("Event Type", level=2)
    doc.add_paragraph(" ".join(sections.get("Event Type", [])))
    doc.add_heading("Event Summary", level=2)
    doc.add_paragraph(" ".join(sections.get("Event Summary", [])))
    doc.add_heading("Contributing Factors", level=2)
    for f in sections.get("Contributing Factors", []):
        doc.add_paragraph(f, style="List Bullet")
    doc.add_heading("Lessons Learned", level=2)
    for l in sections.get("Lessons Learned", []):
        doc.add_paragraph(l, style="List Bullet")
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        table = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            cells = table.add_row().cells
            for j in (0, 1):
                if i + j < len(image_paths):
                    try:
                        cells[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except:
                        cells[j].text = "⚠️ Error loading image"
    doc.save(output_path)

def create_lessons_learned_doc_fr(content: str, output_path: str, image_paths=None):
    """Fallback manual French DOCX generator (python-docx)."""
    sections = parse_sections(content)
    doc = Document()
    def G(k): return sections.get(k, [])
    doc.add_heading(G("Titre")[0] if G("Titre") else "Document d'apprentissage", level=1)
    doc.add_heading("Secteur d'activité d'Aecon", level=2)
    doc.add_paragraph(" ".join(G("Secteur d'activité d'Aecon")))
    doc.add_heading("Projet/Emplacement", level=2)
    doc.add_paragraph(" ".join(G("Projet/Emplacement")))
    doc.add_heading("Date de l'événement", level=2)
    doc.add_paragraph(" ".join(G("Date de l'événement")))
    doc.add_heading("Type d'événement", level=2)
    doc.add_paragraph(" ".join(G("Type d'événement")))
    doc.add_heading("Résumé de l'événement", level=2)
    doc.add_paragraph(" ".join(G("Résumé de l'événement")))
    doc.add_heading("Facteurs contributifs", level=2)
    for f in G("Facteurs contributifs"):
        doc.add_paragraph(f, style="List Bullet")
    doc.add_heading("Leçons apprises", level=2)
    for l in G("Leçons apprises")+G("Leçons tirées"):
        doc.add_paragraph(l, style="List Bullet")
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Photos de soutien", level=2)
        table = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            cells = table.add_row().cells
            for j in (0,1):
                if i + j < len(image_paths):
                    try:
                        cells[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except:
                        cells[j].text = "⚠️ Erreur de chargement de l'image"
    doc.save(output_path)


# ─── Streamlit UI ────────────────────────────────────────────────────────────
st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘")
st.image("AECON.png", width=300)
st.markdown("""
<style>
  .stApp { background:#fff; }
  h1,h2 { color:#c8102e; }
  .stButton>button, .stDownloadButton>button { background:#c8102e; color:#fff; }
  body { font-family:'Segoe UI',sans-serif; }
</style>
""", unsafe_allow_html=True)

st.title("🦺 Serious Event Lessons Learned Generator")
uploaded = st.file_uploader("Upload Executive Review PPTX", type="pptx")
language = st.selectbox("Choose report language:", ["English", "French (Canadian)"])
translator = None
if language == "French (Canadian)":
    translator = st.radio("Translate via:", ["OpenAI", "DeepL"])

use_template = st.checkbox("📑 Use official Aecon Word template", value=True)

if uploaded and st.button("📄 Generate Lessons Learned Document"):
    # 1) Save upload
    in_fp = f"input/{uploaded.name}"
    os.makedirs("input", exist_ok=True)
    with open(in_fp, "wb") as f: f.write(uploaded.getbuffer())

    # 2) Extract text/images
    text, images = extract_text_and_images_from_pptx(in_fp)

    # 3) Summarize
    with st.spinner("Generating summary..."):
        generated = summarize_and_extract(text)

    # 4) Translate if needed
    if language == "French (Canadian)":
        with st.spinner("Translating..."):
            generated = (translate_to_french_openai(generated)
                         if translator == "OpenAI"
                         else translate_to_french_deepl(generated))

    st.success("✅ Generation complete!")
    st.text_area("📝 Generated Content:", generated, height=300)

    # 5) Produce DOCX
    secs = parse_sections(generated)
    lang_code = "fr" if language.startswith("French") else "en"
    out_fp = f"lessons_learned_{lang_code}.docx"
    if use_template:
        render_with_docxtpl(secs, "template.docx", out_fp, images)
    else:
        if language.startswith("French"):
            create_lessons_learned_doc_fr(generated, out_fp, images)
        else:
            create_lessons_learned_doc(generated, out_fp, images)

    # 6) Download
    with open(out_fp, "rb") as f:
        st.download_button(
            "📥 Download DOCX",
            f,
            out_fp,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ─── Footer ──────────────────────────────────────────────────────────────────
st.markdown("""
<hr style="border:none;height:2px;background:#c8102e;"/>
<div style="text-align:center;padding:10px;background:#c8102e;color:#fff;">
  Built by Aecon | For internal use only
</div>
""", unsafe_allow_html=True)
