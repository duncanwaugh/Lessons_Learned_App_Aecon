import os
import hashlib
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from dotenv import load_dotenv
from openai import OpenAI
from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from PIL import Image as PILImage

# ─── Load API keys ────────────────────────────────────────────────────────────
load_dotenv()
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
DEEPL_KEY  = os.getenv("DEEPL_API_KEY")
if not OPENAI_KEY:
    st.error("Missing OPENAI_API_KEY in .env or Streamlit secrets")
client = OpenAI(api_key=OPENAI_KEY)

# ─── Extract + dedupe text & images from PPTX ─────────────────────────────────
def extract_text_and_images_from_pptx(path: str):
    prs = Presentation(path)
    text, images, seen_hashes = "", [], set()
    os.makedirs("images", exist_ok=True)

    for idx, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if shp.has_text_frame:
                text += shp.text_frame.text + "\n"
            elif shp.has_table:
                for row in shp.table.rows:
                    for cell in row.cells:
                        text += cell.text + "\n"
            elif shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = shp.image.blob
                h = hashlib.sha256(blob).hexdigest()
                if h in seen_hashes:
                    continue
                seen_hashes.add(h)
                ext = shp.image.ext
                fn  = f"images/slide{idx+1}_{len(images)}.{ext}"
                with open(fn, "wb") as imgf:
                    imgf.write(blob)
                images.append(fn)

    return text, images

# ─── OpenAI extraction with detailed summary ──────────────────────────────────
def summarize_and_extract(text: str) -> str:
    system_msg = "You are a concise safety-report writer for Aecon."
    prompt = f"""
You are preparing a formal Lessons Learned report from a serious incident.

**Produce each section clearly labeled.**  
For the **Event Summary**, write a detailed multi-paragraph narrative covering:
  1. Background/context  
  2. Step-by-step sequence of events  
  3. Immediate outcome and injuries/damages  
  4. Broader impacts (delays, reputation, etc.)

For **Contributing Factors** and **Lessons Learned**, use a true bullet list:  
- One factor per line prefixed with a hyphen and a space  
- E.g.:

Contributing Factors:
- Factor one
- Factor two
- Factor three

Lessons Learned:
- Lesson one
- Lesson two

Use these exact labels (and nothing else) so your parser can pick them up:
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
    r = client.chat.completions.create(
        model="gpt-4",
        messages=[
            {"role":"system","content":system_msg},
            {"role":"user","content":prompt}
        ],
        temperature=0.2,
        max_tokens=1500,
    )
    return r.choices[0].message.content.strip()

# ─── Translation helpers ──────────────────────────────────────────────────────
def translate_to_french_openai(text: str) -> str:
    prompt = f"Translate into professional French Canadian, keep formatting:\n\n{text}"
    r = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return r.choices[0].message.content.strip()


def translate_to_spanish_openai(text: str) -> str:
    prompt = f"Translate into professional Spanish, keep formatting, section headers and bullets intact:\n\n{text}"
    r = client.chat.completions.create(
        model="gpt-4",
        messages=[{"role":"user","content":prompt}],
        temperature=0.2,
        max_tokens=1000,
    )
    return r.choices[0].message.content.strip()


def translate_to_deepl(text: str, lang: str) -> str:
    if not DEEPL_KEY:
        st.error("Missing DEEPL_API_KEY in .env or Streamlit secrets")
        return text
    target = "FR" if lang == "French" else "ES"
    r = requests.post(
        "https://api-free.deepl.com/v2/translate",
        data={"auth_key":DEEPL_KEY,
              "text": text,
              "target_lang": target,
              "formality": "more"}
    )
    return r.json()["translations"][0]["text"]

# ─── Parse GPT output into dict ─────────────────────────────────────────────────
def parse_sections(out: str) -> dict:
    """
    Split GPT output into labeled sections, handling both:
      - “Label:” on its own line with subsequent lines
      - “Label: value” on a single line
    """
    import re
    labels = [
        "Title", "Aecon Business Sector", "Project/Location",
        "Date of Event", "Event Type", "Event Summary",
        "Contributing Factors", "Lessons Learned",
    ]
    sections = {}
    current = None
    # matches “Label: rest-of-line”
    pat = re.compile(r'^([^:]+):\s*(.*)$')
    for line in out.splitlines():
        line = line.strip()
        if not line:
            continue
        m = pat.match(line)
        if m and m.group(1) in labels:
            key, rest = m.group(1), m.group(2)
            sections[key] = []
            if rest:
                sections[key].append(rest)
            current = key
        elif current:
            sections[current].append(line)
    return sections

# ─── Render + insert images into template ──────────────────────────────────────
def render_with_docxtpl(secs: dict, tpl_path: str, out_path: str, images: list[str]):
    # verify images via PIL
    valid_exts = {'.png','.jpg','.jpeg','.bmp','.gif'}
    verified = []
    for img in images:
        ext = os.path.splitext(img)[1].lower()
        if ext not in valid_exts:
            continue
        try:
            with PILImage.open(img) as im:
                im.verify()
            verified.append(img)
        except:
            continue
    images = verified

    # Fill template
    tpl = DocxTemplate(tpl_path)
    context = {
        "TITLE":   secs.get("Title", [""])[0],
        "SECTOR":  " ".join(secs.get("Aecon Business Sector", [])),
        "PROJECT":" ".join(secs.get("Project/Location", [])),
        "DATE":   " ".join(secs.get("Date of Event", [])),
        "EVENT_TYPE":" ".join(secs.get("Event Type", [])),
        "SUMMARY":  " \n".join(secs.get("Event Summary", [])),
        "FACTORS":  "\n".join(f"{f}" for f in secs.get("Contributing Factors", [])),
        "LESSONS":  "\n".join(f"{l}" for l in secs.get("Lessons Learned", [])),
    }
    tpl.render(context)
    tpl.save(out_path)

    # Insert images
    doc = Document(out_path)
    img_table = None
    for tbl in doc.tables:
        for cell in tbl._cells:
            if "IMAGE_PLACEHOLDER" in cell.text:
                img_table = tbl
                break
        if img_table:
            break
    if img_table:
        tbl_elm = img_table._tbl
        for row in list(img_table.rows):
            tbl_elm.remove(row._tr)
        for i in range(0, len(images), 2):
            row = img_table.add_row().cells
            for j in (0,1):
                if i+j < len(images):
                    row[j].paragraphs[0].add_run().add_picture(
                        images[i+j], width=Inches(2.5)
                    )
    doc.save(out_path)

# ─── Streamlit UI ────────────────────────────────────────────────────────────
st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘")
st.image("AECON.png", width=300)
st.markdown("""
<style>
  .stApp { background:#fff; }
  h1,h2 { color:#c8102e; }
  .stButton>button,.stDownloadButton>button { background:#c8102e;color:#fff; }
  body { font-family:'Segoe UI',sans-serif; }
</style>
""", unsafe_allow_html=True)

st.title("🦺 Serious Event Lessons Learned Generator")
pptx_file = st.file_uploader("Upload Executive Review PPTX", type="pptx")
lang = st.selectbox("Language:", ["English","French (Canadian)","Spanish"])
translator = None
if lang in ["French (Canadian)", "Spanish"]:
    translator = st.radio("Translate via:", ["OpenAI","DeepL"])

if pptx_file and st.button("📄 Generate DOCX"):
    os.makedirs("input", exist_ok=True)
    in_fp = f"input/{pptx_file.name}"
    with open(in_fp, "wb") as f: f.write(pptx_file.getbuffer())

    raw_text, images = extract_text_and_images_from_pptx(in_fp)
    generated = summarize_and_extract(raw_text)

    # Show raw generated for debugging
    st.text_area("📝 Raw Generated Content", generated, height=300)

    if lang == "French (Canadian)":
        generated = (translate_to_french_openai(generated)
                     if translator=="OpenAI" else translate_to_deepl(generated, "French"))
    elif lang == "Spanish":
        generated = (translate_to_spanish_openai(generated)
                     if translator=="OpenAI" else translate_to_deepl(generated, "Spanish"))

    # Show post-translation for debugging
    if lang != "English":
        st.text_area(f"📝 Translated ({lang})", generated, height=300)

    secs = parse_sections(generated)
    out_fp = f"lessons_learned_{lang[:2].lower()}.docx"
    render_with_docxtpl(secs, "lessons learned template.docx", out_fp, images)

    with open(out_fp, "rb") as f:
        st.download_button("📥 Download DOCX", f, out_fp,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown("""
<hr style="border:none;height:2px;background:#c8102e;"/>
<div style="text-align:center;padding:10px;background:#c8102e;color:#fff;">
  Built by Aecon | For internal use only
</div>
""", unsafe_allow_html=True)
