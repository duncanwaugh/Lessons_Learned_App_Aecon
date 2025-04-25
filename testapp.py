import os
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from docx.shared import Inches
from dotenv import load_dotenv
from openai import OpenAI

# ————— Load ENV / Secrets —————
load_dotenv()
client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
DEEPL_KEY  = st.secrets.get("DEEPL_API_KEY")  or os.getenv("DEEPL_API_KEY")

# ————— Helpers —————

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

def generate_doc_from_template(content, template_path, output_path, image_paths=None):
    # 1) Parse out the sections
    sections = parse_sections(content)
    doc = Document(template_path)

    # 2) Simple field replacements
    replacements = {
        "{{DATE}}":          ' '.join(sections.get("Date of Event", [])),
        "{{TITLE}}":         sections.get("Title", [""])[0],
        "{{SECTOR}}":        ' '.join(sections.get("Aecon Business Sector", [])),
        "{{PROJECT}}":       ' '.join(sections.get("Project/Location", [])),
        "{{EVENT_TYPE}}":    ' '.join(sections.get("Event Type", [])),
    }
    for p in doc.paragraphs:
        for key, val in replacements.items():
            if key in p.text:
                p.text = p.text.replace(key, val)

    # 3) Inject full generated content at {{CONTENT}}
    for idx, p in enumerate(doc.paragraphs):
        if "{{CONTENT}}" in p.text:
            # clear placeholder
            p.text = ""
            # insert each line as its own paragraph below
            for offset, line in enumerate(content.split("\n")):
                new_p = doc.add_paragraph(line)
                # move it into correct spot
                elem = new_p._p
                body = doc._body._element
                body.remove(elem)
                body.insert(body.index(p._p) + offset + 1, elem)
            break

    # 4) Append images at the end
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Supporting Pictures", level=2)
        tbl = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            row = tbl.add_row().cells
            for j in (0, 1):
                if i + j < len(image_paths):
                    try:
                        row[j].paragraphs[0].add_run().add_picture(
                            image_paths[i + j], width=Inches(2.5)
                        )
                    except:
                        row[j].text = "[Image failed to load]"

    doc.save(output_path)

def create_lessons_learned_doc(content, output_path, image_paths=None):
    # (unchanged manual English DOCX generator)
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
            row = tbl.add_row().cells
            for j in (0, 1):
                if i + j < len(image_paths):
                    try:
                        row[j].paragraphs[0].add_run().add_picture(
                            image_paths[i + j], width=Inches(2.5)
                        )
                    except:
                        row[j].text = "⚠️ Error loading image"
    doc.save(output_path)

def create_lessons_learned_doc_fr(content, output_path, image_paths=None):
    # (unchanged manual French DOCX generator)
    sections = parse_sections(content)
    doc = Document()
    def G(k): return sections.get(k, [])
    doc.add_heading(G("Titre")[0] if G("Titre") else "Document d'apprentissage", level=1)
    doc.add_heading("Secteur d'activité d'Aecon", level=2)
    doc.add_paragraph(' '.join(G("Secteur d'activité d'Aecon")))
    doc.add_heading("Projet/Emplacement", level=2)
    doc.add_paragraph(' '.join(G("Projet/Emplacement")))
    doc.add_heading("Date de l'événement", level=2)
    doc.add_paragraph(' '.join(G("Date de l'événement")))
    doc.add_heading("Type d'événement", level=2)
    doc.add_paragraph(' '.join(G("Type d'événement")))
    doc.add_heading("Résumé de l'événement", level=2)
    doc.add_paragraph(' '.join(G("Résumé de l'événement")))
    doc.add_heading("Facteurs contributifs", level=2)
    for f in G("Facteurs contributifs"): doc.add_paragraph(f, style='List Bullet')
    doc.add_heading("Leçons apprises", level=2)
    for l in G("Leçons apprises") + G("Leçons tirées"): doc.add_paragraph(l, style='List Bullet')
    if image_paths:
        doc.add_page_break()
        doc.add_heading("Photos de soutien", level=2)
        tbl = doc.add_table(rows=0, cols=2)
        for i in range(0, len(image_paths), 2):
            row = tbl.add_row().cells
            for j in (0, 1):
                if i + j < len(image_paths):
                    try:
                        row[j].paragraphs[0].add_run().add_picture(
                            image_paths[i + j], width=Inches(2.5)
                        )
                    except:
                        row[j].text = "⚠️ Erreur de chargement de l'image"
    doc.save(output_path)

# ————— OpenAI + DeepL Calls —————

def summarize_and_extract(text):
    prompt = f"""
You are helping prepare a standardized Lessons Learned document from a serious incident.
...
"""  # same prompt as before
    res = client.chat.completions.create(
        model="gpt-3.5-turbo",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return res.choices[0].message.content.strip()

def translate_to_french_openai(text):
    # your existing OpenAI translation function
    ...

def translate_to_french_deepl(text):
    # your existing DeepL function
    ...

# ————— PPTX Extraction —————

def extract_text_and_images_from_pptx(path):
    prs = Presentation(path)
    txt, imgs = "", []
    os.makedirs("images", exist_ok=True)
    for idx, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if shp.has_text_frame:
                txt += shp.text_frame.text + "\n"
            elif shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = shp.image.blob
                ext = shp.image.ext
                fn = f"images/slide_{idx+1}_{len(imgs)}.{ext}"
                with open(fn, "wb") as f: f.write(blob)
                imgs.append(fn)
    return txt, imgs

# ————— Streamlit UI —————

st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘")
st.image("AECON.png", width=300)
st.markdown("""<style> ... your CSS ... </style>""", unsafe_allow_html=True)

st.title("🦺 Serious Event Lessons Learned Generator")
uploaded = st.file_uploader("Upload PPTX", type="pptx")
lang     = st.selectbox("Language:", ["English","French (Canadian)"])
trans    = None
if lang=="French (Canadian)":
    trans = st.radio("Translate with:", ["OpenAI","DeepL"])
use_tpl  = st.checkbox("📑 Use official Aecon Word template", True)

if uploaded and st.button("📄 Generate DOCX"):
    # 1) save file
    inp = f"input/{uploaded.name}"
    os.makedirs("input", exist_ok=True)
    with open(inp,"wb") as f: f.write(uploaded.getbuffer())

    # 2) extract & summarize
    with st.spinner("Extracting..."): txt, imgs = extract_text_and_images_from_pptx(inp)
    with st.spinner("Summarizing..."): gen = summarize_and_extract(txt)

    # 3) translate
    if lang=="French (Canadian)":
        with st.spinner("Translating..."):
            gen = (translate_to_french_openai(gen)
                   if trans=="OpenAI"
                   else translate_to_french_deepl(gen))

    st.success("✅ Done!")
    st.text_area("📝 Generated Content", gen, height=300)

    # 4) generate DOCX
    out = f"generated_{'fr' if lang.startswith('French') else 'en'}.docx"
    if use_tpl:
        generate_doc_from_template(gen, "lessons learned template.docx", out, imgs)
    else:
        if lang=="French (Canadian)":
            create_lessons_learned_doc_fr(gen, out, imgs)
        else:
            create_lessons_learned_doc(gen, out, imgs)

    # 5) download
    with open(out,"rb") as f:
        st.download_button("📥 Download DOCX", f, out,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("""
<hr style="border:none;height:2px;background:#c8102e;"/>
<div style="text-align:center;padding:10px;background:#c8102e;color:#fff;">
  Built by Aecon | For internal use only
</div>
""", unsafe_allow_html=True)
