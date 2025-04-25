import os
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.shared import Inches
from dotenv import load_dotenv
from openai import OpenAI

# ————— Helpers —————

# Load environment variables
load_dotenv()

# Global content parser
def parse_sections(content):
    sections = {}
    current = None
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

# Template-based DOCX generator
def generate_doc_from_template(content, template_path, output_path, image_paths=None):
    sections = parse_sections(content)
    doc = Document(template_path)

    # 1) Simple field replacements
    replacements = {
        "{{TITLE}}": sections.get("Title", [""])[0],
        "{{SECTOR}}": ' '.join(sections.get("Aecon Business Sector", [])),
        "{{PROJECT}}": ' '.join(sections.get("Project/Location", [])),
        "{{DATE}}": ' '.join(sections.get("Date of Event", [])),
        "{{EVENT_TYPE}}": ' '.join(sections.get("Event Type", [])),
    }
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                for run in p.runs:
                    run.text = run.text.replace(key, val)

    # 2) Insert entire generated block at {{CONTENT}}
    for idx, p in enumerate(doc.paragraphs):
        if "{{CONTENT}}" in p.text:
            p.text = ""
            insert_at = idx
            for line in content.split("\n"):
                new_p = doc.add_paragraph(line)
                elem = new_p._p
                # move paragraph into correct position
                doc._body._element.remove(elem)
                doc._body._element.insert(insert_at, elem)
                insert_at += 1
            break

    # 3) Append images
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        tbl = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            cells = tbl.add_row().cells
            for j in range(2):
                if i + j < len(image_paths):
                    try:
                        cells[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except:
                        cells[j].text = "[Image failed to load]"

    doc.save(output_path)

# Manual English DOCX generator
def create_lessons_learned_doc(content, output_path, image_paths=None):
    sections = parse_sections(content)
    doc = Document()

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
    for f in sections.get("Contributing Factors", []):
        doc.add_paragraph(f, style='List Bullet')
    doc.add_heading("Lessons Learned", level=2)
    for l in sections.get("Lessons Learned", []):
        doc.add_paragraph(l, style='List Bullet')

    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        tbl = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            cells = tbl.add_row().cells
            for j in range(2):
                if i + j < len(image_paths):
                    try:
                        cells[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except:
                        cells[j].text = "⚠️ Error loading image"

    doc.save(output_path)

# Manual French DOCX generator
def create_lessons_learned_doc_fr(content, output_path, image_paths=None):
    sections = parse_sections(content)
    doc = Document()

    def get(n): 
        return sections.get(n, [])

    doc.add_heading(get("Titre")[0] if get("Titre") else "Document d'apprentissage", level=1)
    doc.add_heading("Secteur d'activité d'Aecon", level=2)
    doc.add_paragraph(' '.join(get("Secteur d'activité d'Aecon")))
    doc.add_heading("Projet/Emplacement", level=2)
    doc.add_paragraph(' '.join(get("Projet/Emplacement")))
    doc.add_heading("Date de l'événement", level=2)
    doc.add_paragraph(' '.join(get("Date de l'événement")))
    doc.add_heading("Type d'événement", level=2)
    doc.add_paragraph(' '.join(get("Type d'événement")))
    doc.add_heading("Résumé de l'événement", level=2)
    doc.add_paragraph(' '.join(get("Résumé de l'événement")))
    doc.add_heading("Facteurs contributifs", level=2)
    for f in get("Facteurs contributifs"): doc.add_paragraph(f, style='List Bullet')
    doc.add_heading("Leçons apprises", level=2)
    for l in get("Leçons apprises") + get("Leçons tirées"): doc.add_paragraph(l, style='List Bullet')

    if image_paths:
        doc.add_page_break()
        doc.add_heading("Photos de soutien", level=2)
        tbl = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            cells = tbl.add_row().cells
            for j in range(2):
                if i + j < len(image_paths):
                    try:
                        cells[j].paragraphs[0].add_run().add_picture(
                            image_paths[i+j], width=Inches(2.5)
                        )
                    except:
                        cells[j].text = "⚠️ Erreur de chargement de l'image"

    doc.save(output_path)

# ————— OpenAI + DeepL clients —————
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))


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
    res = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return res.choices[0].message.content.strip()

def translate_to_french_openai(text):
    prompt = f"""
Translate the following safety incident summary into professional French Canadian. Keep the formatting, section headers, and bullet points intact.

Text to translate:
{text}
"""
    res = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return res.choices[0].message.content.strip()

def translate_to_french_deepl(text):
    key = st.secrets.get("DEEPL_API_KEY")
    if not key:
        st.error("DeepL API key not found in secrets")
        return text
    r = requests.post(
        "https://api-free.deepl.com/v2/translate",
        data={"auth_key":key,"text":text,"target_lang":"FR","formality":"more"}
    )
    try:
        return r.json()["translations"][0]["text"]
    except:
        st.error("DeepL translation failed")
        return text

# ————— PPTX extract —————

def extract_text_and_images_from_pptx(path):
    prs = Presentation(path)
    txt, imgs = "", []
    os.makedirs("images",exist_ok=True)
    for idx,slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if shp.has_text_frame:
                txt += shp.text_frame.text + "\n"
            elif shp.shape_type==MSO_SHAPE_TYPE.PICTURE:
                blob = shp.image.blob
                ext  = shp.image.ext
                fn   = f"images/slide_{idx+1}_{len(imgs)}.{ext}"
                with open(fn,"wb") as f: f.write(blob)
                imgs.append(fn)
    return txt, imgs

# ————— Streamlit UI —————

st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘", layout="centered")
st.image("AECON.png", width=300)
st.markdown("""
<style>
  .stApp {background:#fff;}
  h1,h2,h3,h4{color:#c8102e;}
  .stButton>button,.stDownloadButton>button{background:#c8102e;color:#fff;}
  body{font-family:'Segoe UI',sans-serif;}
</style>
""",unsafe_allow_html=True)

st.title("🦺 Serious Event Lessons Learned Generator")
uploaded = st.file_uploader("Upload Executive Review PPTX", type="pptx")
lang = st.selectbox("Language:", ["English","French (Canadian)"])
trans = None
if lang=="French (Canadian)": trans = st.radio("Translate with:", ["OpenAI","DeepL"])
use_tpl = st.checkbox("📑 Use official Aecon Word template", True)
if uploaded and st.button("📄 Generate DOCX"):
    # save upload
    inp = f"input/{uploaded.name}"
    os.makedirs("input",exist_ok=True)
    with open(inp,"wb") as f: f.write(uploaded.getbuffer())
    # extract & summarize
    with st.spinner("Extracting PPTX..."): txt,imgs = extract_text_and_images_from_pptx(inp)
    with st.spinner("Summarizing..."): gen = summarize_and_extract(txt)
    # translate if needed
    if lang=="French (Canadian)":
        with st.spinner("Translating..."):
            gen = translate_to_french_openai(gen) if trans=="OpenAI" else translate_to_french_deepl(gen)
    st.success("✅ Done!")
    st.text_area("📝 Generated Content", gen, height=300)
    # generate doc
    out = f"generated_lessons_learned_{'fr' if lang=='French (Canadian)' else 'en'}.docx"
    if use_tpl:
        generate_doc_from_template(gen, "lessons learned template.docx", out, imgs)
    else:
        if lang=="French (Canadian)": create_lessons_learned_doc_fr(gen,out,imgs)
        else:                          create_lessons_learned_doc(gen,out,imgs)
    # download
    with open(out,"rb") as f:
        st.download_button("📥 Download DOCX", f, out,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
st.markdown("""
<hr style="border:none;height:2px;background:#c8102e;"/>
<div style="text-align:center;padding:10px;background:#c8102e;color:#fff;font-size:0.9rem;">
  Built by Aecon | For internal use only
</div>
""",unsafe_allow_html=True)
