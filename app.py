import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt
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
        for para in doc.paragraphs: text += para.text + "\n"
    return text

def get_actual_name(raw_text):
    """Fallback: Grabs the first non-empty line of the resume to prevent 'Alex' hallucinations."""
    lines = [l.strip() for l in raw_text.split('\n') if l.strip()]
    return lines[0] if lines else "CANDIDATE NAME"

def parse_and_generate_with_ai(raw_resume_text, job_description, extra_info, interview_results, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    You are a professional resume formatter. Extract data from the provided resume.
    
    STRICT RULES:
    1. SUMMARY: Start with "[First Name] is a...". Write 4-5 dense, metric-heavy sentences.
    2. SKILLS: Exactly 4 rows. Format: "Functional Category & Strategy (Tool 1, Tool 2, Tool 3)".
    3. JOBS: Standardize dates to "MMM YYYY – Current" or "MMM YYYY – MMM YYYY".

    JSON Structure:
    {{
        "FullName": "Extract correctly from resume",
        "FirstName": "First Name Only",
        "Summary": "Summary text...",
        "Skills": [
            {{"Category": "Functional Category (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "University", "Degree": "Degree Name", "Status": "Yes"}}
        ],
        "Jobs": [
            {{"Company": "Company Name", "Title": "Job Title", "Dates": "MMM YYYY – MMM YYYY", "Bullets": ["b1", "b2"]}}
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

def generate_fm_word_doc(ai_data, manual_inputs, raw_text):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # --- 1. TOP HEADER ---
    for p in doc.paragraphs:
        if "CANDIDATE INFORMATION" in p.text:
            p.text = "CANDIDATE INFORMATION"
            p.runs[0].bold = True

    # --- 2. CANDIDATE INFO TABLE (Table 0) ---
    # FIXED: Using Row indices that preserve your "Name:", "Location:" labels
    t0 = doc.tables[0]
    real_name = get_actual_name(raw_text) if "Alex" in ai_data.get("FullName", "") else ai_data.get("FullName")
    
    t0.cell(0, 1).text = real_name
    t0.cell(1, 1).text = manual_inputs["location"]
    t0.cell(2, 1).text = manual_inputs["remote_onsite"]
    t0.cell(3, 1).text = manual_inputs["former_fm"]
    t0.cell(4, 1).text = manual_inputs["links"]

    # --- 3. SUMMARY (Table 1) ---
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # --- 4. EDUCATION (Table 2) ---
    # FIXED: Skip Header Row 0 (EDUCATION) and Row 1 (School | Degree | Completed)
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        row_idx = i + 2 # Starts filling at Row 2
        if row_idx < len(t2.rows):
            t2.cell(row_idx, 0).text = edu.get("School", "")
            t2.cell(row_idx, 1).text = edu.get("Degree", "")
            t2.cell(row_idx, 2).text = edu.get("Status", "Yes")

    # --- 5. SKILLS (Table 3) ---
    # FIXED: Skip Header Row 0 (SKILL & COMPETENCY) and Row 1 (Skill/Competency | Years)
    t3 = doc.tables[3]
    for i, sk in enumerate(ai_data.get("Skills", [])):
        row_idx = i + 2 # Starts filling at Row 2
        if row_idx < len(t3.rows):
            t3.cell(row_idx, 0).text = sk.get("Category", "")
            t3.cell(row_idx, 1).text = sk.get("Exp", "")

    # --- 6. SURGICAL JOB & INTERVIEW REPLACE ---
    def replace_text_safely(old, new):
        for p in doc.paragraphs:
            if old in p.text:
                p.text = p.text.replace(old, str(new))

    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        if job:
            replace_text_safely(f"COMPANY{i}", job['Company'])
            replace_text_safely(f"TITLE{i}", job['Title'])
            # Replace date placeholder on the same line
            for p in doc.paragraphs:
                if job['Company'] in p.text:
                    p.text = p.text.replace("MMM YYYY – CURRENT", job['Dates']).replace("MMM YYYY – MMM YYYY", job['Dates'])
            
            # Bullets
            bullet_tag = "Bullets" if i == 1 else f"Bullets{i}"
            for p in doc.paragraphs:
                if bullet_tag in p.text:
                    p.text = ""
                    for b in job['Bullets']:
                        new_p = p.insert_paragraph_before(f"• {b}")
        else:
            # Clean unused
            replace_text_safely(f"COMPANY{i}", "")
            replace_text_safely(f"TITLE{i}", "")
            replace_text_safely(f"Bullets{i}", "")

    replace_text_safely("ANSWER", manual_inputs["interview_results"])

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit UI ---
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
    job_description = st.text_area("Job Description", placeholder="Paste the job description here")
    extra_info = st.text_area("Spotlight Call/Other Info", placeholder="Spotlight call notes, transcript, manager feedback, MSP comments, etc.")
    interview_results = st.text_area("Supplier Technical Interview Results")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
        
        manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs, raw_text)
        
        # Pull actual first name for the filename
        fname = get_actual_name(raw_text).split()[0]
        st.success(f"Generated for {get_actual_name(raw_text)}")
        st.download_button("Download Resume", data=doc_bytes, file_name=f"FM_Formatted_{fname}.docx")
    except Exception as e:
        st.error(f"Error: {e}")
