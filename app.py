import streamlit as st
import google.generativeai as genai
import docx
import PyPDF2
from io import BytesIO
import json
import os
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
        try:
            for section in doc.sections:
                header = section.header
                if header:
                    for para in header.paragraphs:
                        if para.text.strip(): text += para.text + "\n"
        except: pass
        for para in doc.paragraphs:
            if para.text.strip(): text += para.text + "\n"
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip(): text += cell.text + " "
    return text

def parse_and_generate_with_ai(raw_resume_text, job_description, extra_info, interview_results, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    You are a data extraction tool. You MUST copy professional experience bullets EXACTLY.

    STRICT WORDING RULE: 
    - For the 'Bullets' field in the 'Jobs' section, you MUST copy the text EXACTLY as it appears in the original resume.
    - DO NOT edit, improve, or rephrase. Treat as "read-only" data.

    WORK HISTORY:
    - Extract jobs from most recent to oldest.
    - Format dates as 'Jun 2021 – Nov 2025' or 'Jun 2021 – Current'.

    JSON Structure:
    {{
        "FullName": "Full Name",
        "FirstName": "First Name",
        "Summary": "Summary text...",
        "Skills": [
            {{"Category": "Strategy (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "Uni", "Degree": "Degree", "Status": "Yes"}}
        ],
        "Jobs": [
            {{"Company": "Company", "Title": "Title", "Dates": "Jun 2021 – Nov 2025", "Bullets": ["EXACT BULLET 1"]}}
        ]
    }}

    RESUME: {raw_resume_text}
    JD: {job_description}
    NOTES: {extra_info}
    INTERVIEW: {interview_results}
    """
    
    response = model.generate_content(prompt)
    raw_output = response.text.strip()
    json_string = raw_output.split("```json")[1].split("```")[0].strip() if "```json" in raw_output else raw_output
    return json.loads(json_string)

def run_level_replace(paragraph, target, replacement):
    """Replaces text while keeping all template tabs and paragraph marks intact."""
    if target.lower() in paragraph.text.lower():
        for run in paragraph.runs:
            if target.lower() in run.text.lower():
                # Case-insensitive surgical swap
                insens_re = re.compile(re.escape(target), re.IGNORECASE)
                run.text = insens_re.sub(str(replacement), run.text)

def generate_fm_word_doc(ai_data, manual_inputs, raw_text):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. TABLE COORDINATES (Exact mapping)
    t0 = doc.tables[0]
    t0.cell(1, 1).text = ai_data.get("FullName", "").strip().title()
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 2. EDUCATION & SKILLS TABLES
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 2 < len(t2.rows):
            t2.cell(i+2, 0).text = edu.get("School", "")
            t2.cell(i+2, 1).text = edu.get("Degree", "")
            t2.cell(i+2, 2).text = edu.get("Status", "Yes")

    t3 = doc.tables[3]
    for i, sk in enumerate(ai_data.get("Skills", [])):
        if i + 2 < len(t3.rows):
            t3.cell(i+2, 0).text = sk.get("Category", "")
            t3.cell(i+2, 1).text = sk.get("Exp", "")

    # 3. WORK HISTORY (Surgical Run-Level Mapping)
    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        c_tag, t_tag, b_tag = f"company{i}", f"title{i}", f"Job{i}Bullets"
        d_tag_1, d_tag_2 = "mmm yyyy – Current", "mmm yyyy – mmm yyyy"
        
        for p in doc.paragraphs:
            # Replace Company & Dates without touching the Tab Stop
            if c_tag in p.text.lower():
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    run_level_replace(p, d_tag_1, job['Dates'])
                    run_level_replace(p, d_tag_2, job['Dates'])
                else: p.text = "" # Clears unused role headers
            
            # Replace Title while keeping style
            if t_tag in p.text.lower():
                if job: run_level_replace(p, t_tag, job['Title'])
                else: p.text = ""

            # Inject Bullets using the template's EXACT paragraph style
            if b_tag in p.text:
                orig_style = p.style
                p.text = "" # Clear placeholder but keep the paragraph object
                if job:
                    for b in job['Bullets']:
                        # This creates a new paragraph ABOVE the placeholder with the exact same style
                        doc.paragraphs[doc.paragraphs.index(p)].insert_paragraph_before(f"• {b}", style=orig_style)

    # 4. INTERVIEW RESULTS
    for p in doc.paragraphs:
        if "ANSWER" in p.text:
            p.text = p.text.replace("ANSWER", manual_inputs["interview_results"])

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI Setup ---
st.set_page_config(page_title="Fannie Mae Formatter", layout="wide")
st.title("📄 Fannie Mae Resume Auto-Formatter")

with st.sidebar:
    api_key = st.text_input("Gemini API Key:", type="password")

c1, c2 = st.columns(2)
with c1:
    uploaded_file = st.file_uploader("Upload Resume", type=["pdf", "docx"])
    location = st.text_input("Current Location: (City and State only)", placeholder="Washington, DC")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM FTE or Contractor?", ["N", "Y"])
    links = st.text_input("LinkedIn Profile/GitHub/Portfolio Link")

with c2:
    job_description = st.text_area("Job Description", placeholder="Paste JD here")
    extra_info = st.text_area("Spotlight Call/Other Info", placeholder="Spotlight notes...")
    interview_results = st.text_area("Supplier Technical Interview Results")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
        manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs, raw_text)
        name = ai_data.get("FullName", "Candidate").title()
        st.success(f"Generated for {name}")
        st.download_button(label="Download", data=doc_bytes, file_name=f"{name} Fannie Mae Format.docx")
    except Exception as e:
        st.error(f"Error: {e}")
