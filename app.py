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
    You are a professional technical recruiter. 
    
    STRICT IDENTITY RULE: Extract the candidate's ACTUAL name from the resume. 
    
    SKILLS SECTION STYLE (CRITICAL):
    You MUST create exactly 4 rows using this high-level consultant style:
    - Row 1: Cybersecurity & SOC Operations (Threat Detection, Incident Response, Insider Threat Investigations)
    - Row 2: SIEM & Threat Hunting (Splunk, IBM QRadar, Exabeam, Log Analysis, Correlation)
    - Row 3: Network & Security Analysis (TCP/IP, DNS, HTTP/S, Wireshark, Endpoint Security)
    - Row 4: Fraud & Behavioral Risk Analysis (Financial Transactions, Pattern Detection, Root Cause Analysis)
    (Note: If the candidate is a developer or other role, follow this same 'Functional Category & Strategy (Tools, Concepts)' format).

    WORK HISTORY RULE:
    - Extract ALL previous jobs.
    - Format dates exactly as: 'Jun 2021 – Nov 2025' or 'Jun 2021 – Current'.
    - Use 3-letter month abbreviations.

    JSON Structure:
    {{
        "FullName": "Full Name",
        "FirstName": "First Name",
        "Summary": "Full Summary starting with [First Name] is a...",
        "Skills": [
            {{"Category": "Category Name (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "Uni", "Degree": "Degree", "Status": "Yes"}}
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

def run_level_replace(paragraph, target, replacement):
    """Surgical replacement that finds the target even if Word split the runs."""
    if target.lower() in paragraph.text.lower():
        found = False
        for run in paragraph.runs:
            # Case-insensitive replacement within the run
            if target.lower() in run.text.lower():
                import re
                insens_re = re.compile(re.escape(target), re.IGNORECASE)
                run.text = insens_re.sub(str(replacement), run.text)
                found = True
        
        if not found:
            # Fallback if placeholder is split across runs
            import re
            insens_re = re.compile(re.escape(target), re.IGNORECASE)
            new_text = insens_re.sub(str(replacement), paragraph.text)
            for i in range(len(paragraph.runs)):
                paragraph.runs[i].text = ""
            paragraph.runs[0].text = new_text

def generate_fm_word_doc(ai_data, manual_inputs, raw_text):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. TABLE MAPPING (Coordinates)
    t0 = doc.tables[0]
    clean_name = ai_data.get("FullName", "Candidate").strip().title()
    t0.cell(1, 1).text = clean_name
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # Education (Skip headers)
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 2 < len(t2.rows):
            t2.cell(i+2, 0).text = edu.get("School", "")
            t2.cell(i+2, 1).text = edu.get("Degree", "")
            t2.cell(i+2, 2).text = edu.get("Status", "Yes")

    # Skills (Skip headers)
    t3 = doc.tables[3]
    for i, sk in enumerate(ai_data.get("Skills", [])):
        if i + 2 < len(t3.rows):
            t3.cell(i+2, 0).text = sk.get("Category", "")
            t3.cell(i+2, 1).text = sk.get("Exp", "")

    # 2. WORK HISTORY (Format-Locked Replacement)
    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        
        c_tag = f"company{i}"
        t_tag = f"title{i}"
        d_tag_current = "mmm yyyy – Current"
        d_tag_range = "mmm yyyy – mmm yyyy"
        bullet_tag = "Bullets" if i == 1 else f"Bullets{i}"
        
        for p in doc.paragraphs:
            # Match Company and Date line (Preserves Tab)
            if c_tag in p.text.lower():
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    run_level_replace(p, d_tag_current, job['Dates'])
                    run_level_replace(p, d_tag_range, job['Dates'])
                else:
                    p.text = "" # Clears unused roles
            
            # Match Job Title
            if t_tag in p.text.lower():
                if job:
                    run_level_replace(p, t_tag, job['Title'])
                else:
                    p.text = ""

            # Match Bullets
            if bullet_tag in p.text:
                p.text = "" # Clear the tag
                if job:
                    for b in job['Bullets']:
                        # Inserts bullet paragraphs before the placeholder paragraph
                        p.insert_paragraph_before(f"• {b}")

    # 3. INTERVIEW RESULTS
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
    extra_info = st.text_area("Spotlight Call/Other Info", placeholder="Spotlight call notes, feedback, etc.")
    interview_results = st.text_area("Supplier Technical Interview Results")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
        manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
        
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs, raw_text)
        full_name = ai_data.get("FullName", "Candidate").title()
        
        st.success(f"Success! Generated for {full_name}")
        st.download_button(
            label="Download Formatted Resume", 
            data=doc_bytes, 
            file_name=f"{full_name} Fannie Mae Format.docx"
        )
    except Exception as e:
        st.error(f"Error: {e}")
