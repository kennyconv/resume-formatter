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
    You are a data extraction tool. You MUST output VALID JSON.
    
    STRICT WORDING RULE: 
    - For the 'Bullets' field, copy the text EXACTLY. 
    - Escape any double quotes inside the text with a backslash.
    - Remove any literal tabs or newlines from inside the bullet strings.

    SKILLS:
    Create 4 strategy rows (e.g., Cybersecurity & SOC Operations...).

    JSON Structure:
    {{
        "FullName": "Full Name",
        "FirstName": "First Name",
        "Summary": "Summary...",
        "Skills": [
            {{"Category": "Strategy", "Exp": "X years"}}
        ],
        "Education": [
            {{"School": "Uni", "Degree": "Degree", "Status": "Yes"}}
        ],
        "Jobs": [
            {{"Company": "Company", "Title": "Title", "Dates": "Jun 2021 – Nov 2025", "Bullets": ["Exact Bullet 1"]}}
        ]
    }}

    RESUME: {raw_resume_text}
    JD: {job_description}
    NOTES: {extra_info}
    INTERVIEW: {interview_results}
    """
    
    response = model.generate_content(prompt)
    res_text = response.text.strip()
    
    # Extract JSON block
    if "```json" in res_text:
        res_text = res_text.split("```json")[1].split("```")[0].strip()
    
    # Clean common JSON errors (trailing commas, etc)
    res_text = re.sub(r',\s*([\]}])', r'\1', res_text)
    
    return json.loads(res_text)

def run_level_replace(paragraph, target, replacement):
    if target.lower() in paragraph.text.lower():
        found = False
        for run in paragraph.runs:
            if target.lower() in run.text.lower():
                insens_re = re.compile(re.escape(target), re.IGNORECASE)
                run.text = insens_re.sub(str(replacement), run.text)
                found = True
        if not found:
            insens_re = re.compile(re.escape(target), re.IGNORECASE)
            new_text = insens_re.sub(str(replacement), paragraph.text)
            for i, run in enumerate(paragraph.runs):
                run.text = new_text if i == 0 else ""
        return True
    return False

def generate_fm_word_doc(ai_data, manual_inputs, raw_text):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. INFO TABLE
    t0 = doc.tables[0]
    t0.cell(1, 1).text = ai_data.get("FullName", "").strip().title()
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 2. TABLES
    for t_idx, data_key in [(2, "Education"), (3, "Skills")]:
        table = doc.tables[t_idx]
        for i, item in enumerate(ai_data.get(data_key, [])):
            if i + 2 < len(table.rows):
                if data_key == "Education":
                    table.cell(i+2, 0).text = item.get("School", "")
                    table.cell(i+2, 1).text = item.get("Degree", "")
                    table.cell(i+2, 2).text = item.get("Status", "Yes")
                else:
                    table.cell(i+2, 0).text = item.get("Category", "")
                    table.cell(i+2, 1).text = item.get("Exp", "")

    # 3. WORK HISTORY
    jobs = ai_data.get("Jobs", [])
    for p in doc.paragraphs:
        p_text_low = p.text.lower()
        for i in range(1, 8):
            job = jobs[i-1] if i <= len(jobs) else None
            c_tag, t_tag, b_tag = f"company{i}", f"title{i}", f"job{i}bullets"
            
            if c_tag in p_text_low:
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    d_tags = ["mmm yyyy – current", "mmm yyyy – mmm yyyy"]
                    for d in d_tags:
                        if d in p.text.lower(): run_level_replace(p, d, job['Dates'])
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

    # 4. INTERVIEW
    for p in doc.paragraphs:
        if "ANSWER" in p.text:
            p.text = p.text.replace("ANSWER", manual_inputs["interview_results"])

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- UI ---
st.set_page_config(page_title="Fannie Mae Formatter", layout="wide")
st.title("📄 Fannie Mae Resume Auto-Formatter")

with st.sidebar:
    api_key = st.text_input("Gemini API Key:", type="password")

c1, c2 = st.columns(2)
with c1:
    uploaded_file = st.file_uploader("Upload Resume", type=["pdf", "docx"])
    location = st.text_input("Current Location (City, ST)")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("LinkedIn/GitHub Link")
with c2:
    job_description = st.text_area("Job Description")
    extra_info = st.text_area("Spotlight Call Info")
    interview_results = st.text_area("Interview Results")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
        manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs, raw_text)
        name = ai_data.get("FullName", "Candidate").title()
        st.success(f"Generated for {name}")
        st.download_button("Download", data=doc_bytes, file_name=f"{name} Fannie Mae Format.docx")
    except Exception as e:
        st.error(f"Error: {e}")
