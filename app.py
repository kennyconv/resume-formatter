import streamlit as st
from google import genai
from google.genai import types
import docx
import PyPDF2
from io import BytesIO
import json
import re

# --- Core Functions ---

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

def parse_with_new_sdk(raw_text, api_key):
    # This modern client defaults to the stable v1 endpoint, fixing the 404
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
    You are a data extraction tool. Extract data into JSON.
    STRICT RULE: Copy professional experience bullets EXACTLY. 
    DO NOT rewrite, summarize, or improve any text. 

    JSON Structure:
    {{
        "FullName": "Name",
        "Education": [{{"School": "Uni", "Degree": "Major"}}],
        "Jobs": [
            {{"Company": "Co", "Title": "Title", "Dates": "MMM YYYY – MMM YYYY", "Bullets": ["Bullet 1"]}}
        ]
    }}

    RESUME TEXT:
    {raw_text}
    """
    
    response = client.models.generate_content(
        model='gemini-2.0-flash', 
        contents=prompt,
        config=types.GenerateContentConfig(response_mime_type='application/json')
    )
    return json.loads(response.text)

def run_level_replace(paragraph, target, replacement):
    """Replaces text while preserving the Tab Stops and formatting in your template."""
    if target.lower() in paragraph.text.lower():
        for run in paragraph.runs:
            if target.lower() in run.text.lower():
                insens_re = re.compile(re.escape(target), re.IGNORECASE)
                run.text = insens_re.sub(str(replacement), run.text)

def generate_docx(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path)

    # 1. Candidate Info (Table 0)
    t0 = doc.tables[0]
    t0.cell(1, 1).text = ai_data.get("FullName", "NAME").title()
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    # 2. Education (Table 2)
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
            d_tag1, d_tag2 = "mmm yyyy – current", "mmm yyyy – mmm yyyy"

            if c_tag in p_text_low:
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    # This logic keeps the date on the far right as per your template's Tab Stop
                    if d_tag1 in p.text.lower(): run_level_replace(p, d_tag1, job['Dates'])
                    elif d_tag2 in p.text.lower(): run_level_replace(p, d_tag2, job['Dates'])
                else: p.text = "" # Clears unused roles
            
            elif t_tag in p_text_low:
                if job: run_level_replace(p, t_tag, job['Title'])
                else: p.text = ""

            elif b_tag in p_text_low:
                orig_style = p.style
                p.text = ""
                if job:
                    for b in job['Bullets']:
                        p.insert_paragraph_before(f"• {b}", style=orig_style)

    # 4. Interview Logic (Q1-Q5 and ANSWER1-ANSWER5)
    for p in doc.paragraphs:
        for i in range(1, 6):
            q_tag, a_tag = f"Q{i}", f"ANSWER{i}"
            q_val = manual_inputs.get(f"q{i}")
            a_val = manual_inputs.get(f"a{i}")
            if q_tag in p.text: p.text = p.text.replace(q_tag, q_val) if q_val else ""
            if a_tag in p.text: p.text = p.text.replace(a_tag, a_val) if a_val else ""

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit Interface ---
st.set_page_config(page_title="Fannie Mae Precision Extractor")
st.title("📄 Fannie Mae Precision Data Extractor")

api_key = st.sidebar.text_input("Gemini API Key", type="password")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Upload Resume", type=["pdf", "docx"])
    location = st.text_input("Current Location (City, ST)")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote"])
    former_fm = st.selectbox("Former FM FTE or Contractor?", ["N", "Y"])
    links = st.text_input("LinkedIn Profile Link")

with col2:
    st.subheader("Technical Interview Results")
    qa_data = {}
    for i in range(1, 6):
        qa_data[f"q{i}"] = st.text_input(f"Question {i}", key=f"q_in_{i}")
        qa_data[f"a{i}"] = st.text_area(f"Answer {i}", key=f"a_in_{i}")

if st.button("Generate Formatted Resume") and uploaded_file and api_key:
    try:
        raw_text = extract_text_from_file(uploaded_file)
        data = parse_with_new_sdk(raw_text, api_key)
        manual = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, **qa_data}
        
        output_doc = generate_docx(data, manual)
        name = data.get("FullName", "Candidate").title()
        st.success(f"Success! Data extracted for {name}")
        st.download_button("Download Resume", output_doc, f"{name} Fannie Mae Format.docx")
    except Exception as e:
        st.error(f"Error: {e}")
