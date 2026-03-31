import streamlit as st
from google import genai
from google.genai import types
import docx
import PyPDF2
from io import BytesIO
import json
import re

# --- Data Extraction ---

def extract_text_from_file(uploaded_file):
    file_name = uploaded_file.name.lower()
    text = ""
    if file_name.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            page_text = page.extract_text()
            if page_text: text += page_text + "\n"
    elif file_name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs:
            if para.text.strip(): text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip(): text += cell.text + " "
    return text

def parse_with_fixed_model(raw_text, api_key):
    # Initializing with the new SDK
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    You are a literal data extraction tool. Extract data into JSON.
    STRICT RULE: Copy professional experience bullets EXACTLY. 
    DO NOT rewrite, summarize, or improve any text. 

    JSON Structure:
    {{
        "FullName": "Name",
        "Education": [{{"School": "Uni", "Degree": "Major"}}],
        "Jobs": [
            {{"Company": "Co", "Title": "Title", "Dates": "MMM YYYY – MMM YYYY", "Bullets": ["Exact Bullet 1"]}}
        ]
    }}

    RESUME TEXT:
    {raw_text}
    """
    
    # FIXED: Using the specific version-locked model string
    response = client.models.generate_content(
        model='gemini-2.0-flash-001', 
        contents=prompt,
        config=types.GenerateContentConfig(response_mime_type='application/json')
    )
    return json.loads(response.text)

def run_level_replace(paragraph, target, replacement):
    """Surgical replacement to preserve Tab Stop alignment."""
    if target.lower() in paragraph.text.lower():
        for run in paragraph.runs:
            if target.lower() in run.text.lower():
                insens_re = re.compile(re.escape(target), re.IGNORECASE)
                run.text = insens_re.sub(str(replacement), run.text)

def generate_docx(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path)

    # 1. Candidate Info
    t0 = doc.tables[0]
    t0.cell(1, 1).text = ai_data.get("FullName", "NAME").title()
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    # 2. Education
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 2 < len(t2.rows):
            t2.cell(i+2, 0).text = edu.get("School", "")
            t2.cell(i+2, 1).text = edu.get("Degree", "")
            t2.cell(i+2, 2).text = "Yes"

    # 3. Work History & Bullet Mapping
    jobs = ai_data.get("Jobs", [])
    for p in doc.paragraphs:
        p_text_low = p.text.lower()
        for i in range(1, 8):
            job = jobs[i-1] if i <= len(jobs) else None
            c_tag, t_tag, b_tag = f"company{i}", f"title{i}", f"job{i}bullets"
            
            if c_tag in p_text_low:
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    for d_tag in ["mmm yyyy – current", "mmm yyyy – mmm yyyy"]:
                        if d_tag in p.text.lower():
                            run_level_replace(p, d_tag, job['Dates'])
                else: p.text = "" 
            elif t_tag in p_text_low:
                if job: run_level_replace(p, t_tag, job['Title'])
                else: p.text = ""
            elif b_tag in p_text_low:
                orig_style = p.style
                p.text = ""
                if job:
                    for b in job['Bullets']:
                        p.insert_paragraph_before(f"• {b}", style=orig_style)

    # 4. Q1-Q5 Interview Logic
    for p in doc.paragraphs:
        for i in range(1, 6):
            q_tag, a_tag = f"Q{i}", f"ANSWER{i}"
            q_val, a_val = manual_inputs.get(f"q{i}"), manual_inputs.get(f"a{i}")
            if q_tag in p.text: p.text = p.text.replace(q_tag, q_val) if q_val else ""
            if a_tag in p.text: p.text = p.text.replace(a_tag, a_val) if a_val else ""

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="Fannie Mae Extractor", layout="wide")
st.title("📄 Fannie Mae Precision Data Extractor")

api_key = st.sidebar.text_input("Gemini API Key", type="password")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Upload Resume", type=["pdf", "docx"])
    location = st.text_input("Current Location (City, ST)")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("LinkedIn Profile Link")

with col2:
    st.subheader("Technical Interview Results")
    qa_data = {}
    for i in range(1, 6):
        qa_data[f"q{i}"] = st.text_input(f"Question {i}", key=f"q_field_{i}")
        qa_data[f"a{i}"] = st.text_area(f"Answer {i}", key=f"a_field_{i}")

if st.button("Generate Formatted Resume") and uploaded_file and api_key:
    try:
        raw_text = extract_text_from_file(uploaded_file)
        data = parse_with_fixed_model(raw_text, api_key)
        manual = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, **qa_data}
        doc_out = generate_docx(data, manual)
        name = data.get("FullName", "Candidate").title()
        st.success(f"Success! Extracted data for {name}")
        st.download_button("Download", doc_out, f"{name}_FannieMae.docx")
    except Exception as e:
        st.error(f"Error: {e}")
