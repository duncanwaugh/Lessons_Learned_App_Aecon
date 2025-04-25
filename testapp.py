import os
import streamlit as st
import requests
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
from html2docx import html2docx
from dotenv import load_dotenv
from openai import OpenAI

# ─── Load secrets ─────────────────────────────────────────────────────────────
load_dotenv()
OPENAI_KEY = os.getenv("OPENAI_API_KEY")
DEEPL_KEY  = os.getenv("DEEPL_API_KEY")
client     = OpenAI(api_key=OPENAI_KEY)

# ─── Helpers ──────────────────────────────────────────────────────────────────
def parse_sections(content):
    sections, current = {}, None
    for line in content.splitlines():
        line = line.strip()
        if not line: continue
        if line.endswith(':') and len(line.split())<6:
            current = line[:-1].strip()
            sections[current] = []
        elif current:
            sections[current].append(line)
    return sections

def build_html(sections, image_paths):
    # Inline CSS for Aecon styling
    css = """
    body { font-family: 'Segoe UI', sans-serif; margin:20px; }
    h1, h2 { color: #c8102e; }
    ul { margin:0 0 1em 1.5em; }
    .img-row { display:flex; flex-wrap:wrap; gap:10px; }
    .img-row img { width:45%; }
    """
    html = [f"<html><head><meta charset='utf-8'><style>{css}</style></head><body>"]
    html.append(f"<h1>{sections.get('Title',[''])[0]}</h1>")
    html.append(f"<p><strong>Aecon Business Sector:</strong> {' '.join(sections.get('Aecon Business Sector',[]))}</p>")
    html.append(f"<p><strong>Project/Location:</strong> {' '.join(sections.get('Project/Location',[]))}</p>")
    html.append(f"<p><strong>Date of Event:</strong> {' '.join(sections.get('Date of Event',[]))}</p>")
    html.append(f"<p><strong>Event Type:</strong> {' '.join(sections.get('Event Type',[]))}</p>")
    html.append(f"<h2>Event Summary</h2><p>{' '.join(sections.get('Event Summary',[]))}</p>")
    html.append("<h2>Contributing Factors</h2><ul>")
    for f in sections.get('Contributing Factors',[]): html.append(f"<li>{f}</li>")
    html.append("</ul><h2>Lessons Learned</h2><ul>")
    for l in sections.get('Lessons Learned',[]): html.append(f"<li>{l}</li>")
    html.append("</ul>")
    if image_paths:
        html.append("<h2>Supporting Pictures</h2><div class='img-row'>")
        for img in image_paths:
            html.append(f"<img src='{img}'/>")
        html.append("</div>")
    html.append("</body></html>")
    return "".join(html)

def html_to_docx(html: str, output_path: str):
    doc = Document()
    html2docx(html, doc)
    doc.save(output_path)

# ─── Streamlit UI ────────────────────────────────────────────────────────────
st.set_page_config(page_title="Aecon Lessons Learned Generator", page_icon="📘")
st.image("AECON.png", width=300)
st.markdown("""
    <style>
      .stApp { background: #fff; }
      h1,h2 { color: #c8102e; }
      .stButton>button, .stDownloadButton>button { background: #c8102e; color: #fff; }
      body { font-family: 'Segoe UI', sans-serif; }
    </style>
""", unsafe_allow_html=True)

st.title("Serious Event Lessons Learned Generator")

uploaded = st.file_uploader("Upload Executive Review PPTX", type="pptx")
language = st.selectbox("Choose report language:", ["English", "French (Canadian)"])
translator = None
if language == "French (Canadian)":
    translator = st.radio("Translation via:", ["OpenAI", "DeepL"])

if uploaded and st.button("📄 Generate Lessons Learned DOCX"):
    # 1) Save PPTX locally
    os.makedirs("input", exist_ok=True)
    in_path = os.path.join("input", uploaded.name)
    with open(in_path, "wb") as f:
        f.write(uploaded.getbuffer())

    # 2) Extract text & images
    prs = Presentation(in_path)
    full_text, images = "", []
    os.makedirs("images", exist_ok=True)
    for idx, slide in enumerate(prs.slides):
        for shp in slide.shapes:
            if shp.has_text_frame:
                full_text += shp.text_frame.text + "\n"
            elif shp.shape_type == MSO_SHAPE_TYPE.PICTURE:
                blob = shp.image.blob
                ext  = shp.image.ext
                fn   = f"images/slide_{idx+1}_{len(images)}.{ext}"
                with open(fn, "wb") as imgf: imgf.write(blob)
                images.append(fn)

    # 3) Summarize
    with st.spinner("Summarizing..."):
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
{full_text}
"""
        resp = client.chat.completions.create(
            model="gpt-3.5-turbo",
            messages=[{"role":"user","content":prompt}],
            temperature=0.2,
            max_tokens=1000,
        )
        generated = resp.choices[0].message.content.strip()

    # 4) Translate if needed
    if language == "French (Canadian)":
        with st.spinner("Translating..."):
            if translator == "OpenAI":
                trans_prompt = f"Translate into professional French Canadian, keeping formatting:\n\n{generated}"
                t_resp = client.chat.completions.create(
                    model="gpt-4",
                    messages=[{"role":"user","content":trans_prompt}],
                    temperature=0.2,
                    max_tokens=1000,
                )
                generated = t_resp.choices[0].message.content.strip()
            else:
                dl = requests.post(
                    "https://api-free.deepl.com/v2/translate",
                    data={"auth_key":DEEPL_KEY,"text":generated,"target_lang":"FR","formality":"more"},
                )
                generated = dl.json()["translations"][0]["text"]

    st.success("✅ Generation complete!")
    st.text_area("📝 Extracted & Formatted Content:", generated, height=300)

    # 5) Build HTML → DOCX
    secs = parse_sections(generated)
    html = build_html(secs, images)
    out_path = f"lessons_learned_{'fr' if language.startswith('French') else 'en'}.docx"
    html_to_docx(html, out_path)

    # 6) Download
    with open(out_path, "rb") as f:
        st.download_button(
            "📥 Download DOCX", f, out_path,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

# ─── Footer ──────────────────────────────────────────────────────────────────
st.markdown("""
<hr style="border:none;height:2px;background:#c8102e;"/>
<div style="text-align:center;padding:10px;background:#c8102e;color:#fff;">
  Built by Aecon | For internal use only
</div>
""", unsafe_allow_html=True)
