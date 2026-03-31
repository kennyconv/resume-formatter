import streamlit as st
import google.generativeai as genai
import docx
import PyPDF2
from io import BytesIO
import json
import os

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
    You are a professional resume formatting assistant. 
    STRICT IDENTITY RULE: Extract the candidate's ACTUAL name.
    
    WORK HISTORY RULE:
    - Extract jobs from most recent to oldest.
    - Format dates as 'AbbreviatedMonth YYYY – AbbreviatedMonth YYYY' (e.g., Jun 2021 – Nov 2025).
    - If current, use 'Jun 2021 – Current'.

    JSON Structure:
    {{
        "FullName": "Full Name",
        "FirstName": "First Name",
        "Summary": "Full Summary...",
        "Skills": [
            {{"Category": "Functional Category (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "Uni Name", "Degree": "Degree Name", "Status": "Yes"}}
        ],
        "Jobs": [
            {{"Company": "Company Name", "Title": "Job Title", "Dates": "Jun 2021 – Nov 2025", "Bullets": ["b1", "b2"]}}
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

def surgical_replace(paragraph, old_text, new_text):
    """Replaces text within a paragraph while preserving existing formatting and tab stops."""
    if old_text in paragraph.text:
        # We replace the text in the 'runs' to preserve bold/italic/tabs
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, str(new_text))

def generate_fm_word_doc(ai_data, manual_inputs, raw_text):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. CANDIDATE INFO (Coordinate Mapping)
    t0 = doc.tables[0]
    # Ensure name is Title Case
    clean_name = ai_data.get("FullName", "Candidate").strip().title()
    t0.cell(1, 1).text = clean_name
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    # 2. SUMMARY
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 3. EDUCATION & SKILLS
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

    # 4. WORK HISTORY (Format-Aware Replacement)
    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        
        c_tag = f"company{i}"
        t_tag = f"title{i}"
        d_tag_current = "mmm yyyy – Current"
        d_tag_range = "mmm yyyy – mmm yyyy"
        
        for p in doc.paragraphs:
            if c_tag in p.text:
                if job:
                    surgical_replace(p, c_tag, job['Company'])
                    surgical_replace(p, d_tag_current, job['Dates'])
                    surgical_replace(p, d_tag_range, job['Dates'])
                else:
                    p.text = "" # Remove unused line
            
            if t_tag in p.text:
                if job:
                    surgical_replace(p, t_tag, job['Title'])
                else:
                    p.text = ""

            # Bullets Placeholder
            bullet_tag = "Bullets" if i == 1 else f"Bullets{i}"
            if bullet_tag in p.text:
                p.text = "" # Clear the placeholder
                if job:
                    for bullet in job['Bullets']:
                        new_p = p.insert_paragraph_before(f"• {bullet}")

    # 5. INTERVIEW RESULTS
    for p in doc.paragraphs:
        if "ANSWER" in p.text:
            p.text = p.text.replace("ANSWER", manual_inputs["interview_results"])

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
        
        full_name_clean = ai_data.get("FullName", "Candidate").title()
        st.success(f"Success! Generated for {full_name_clean}")
        st.download_button("Download", data=doc_bytes, file_name=f"{full_name_clean} Fannie Mae Format.docx")
    except Exception as e:
        st.error(f"Error: {e}")
