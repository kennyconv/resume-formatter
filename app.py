import streamlit as st
import google.generativeai as genai
import docx
from docx.shared import Pt
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
        for para in doc.paragraphs: text += para.text + "\n"
    return text

def parse_and_generate_with_ai(raw_resume_text, job_description, extra_info, interview_results, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    You are a professional resume formatter. 
    
    STRICT IDENTITY RULE: 
    - Extract the candidate's ACTUAL name from the top of the resume. 
    - NEVER use "Alex" or "Candidate Name".
    
    SUMMARY RULE:
    - Use [First Name] only. Start with "[First Name] is a..."
    - Write 4-5 dense sentences using metrics.

    JSON Structure:
    {{
        "FullName": "Full Name",
        "FirstName": "First Name",
        "Summary": "Summary text...",
        "Skills": [
            {{"Cat": "Functional Category (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "Uni", "Degree": "Degree", "Status": "Yes"}}
        ],
        "Jobs": [
            {{"Comp": "Company", "Title": "Title", "Dates": "MMM YYYY – MMM YYYY", "Bullets": ["b1", "b2"]}}
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

    # --- 1. COORDINATE-LOCKED TABLE INJECTION ---
    
    # Table 0: CANDIDATE INFORMATION
    # Row 0 is the Header "CANDIDATE INFORMATION" - We Bold it.
    t0 = doc.tables[0]
    t0.cell(0, 0).paragraphs[0].text = "CANDIDATE INFORMATION"
    t0.cell(0, 0).paragraphs[0].runs[0].bold = True
    
    # We put answers in Column 1, starting Row 0 (next to Name label)
    # If AI fails name, we grab first line of raw_text
    raw_name = raw_text.split('\n')[0].strip()
    final_name = ai_data.get("FullName") if "Alex" not in ai_data.get("FullName", "") else raw_name
    
    t0.cell(0, 1).text = final_name
    t0.cell(1, 1).text = manual_inputs["location"]
    t0.cell(2, 1).text = manual_inputs["remote_onsite"]
    t0.cell(3, 1).text = manual_inputs["former_fm"]
    t0.cell(4, 1).text = manual_inputs["links"]

    # Table 1: SUMMARY (Put in Row 1, Column 0)
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # Table 2: EDUCATION (Start at Row 2 to keep headers)
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 2 < len(t2.rows):
            t2.cell(i+2, 0).text = edu.get("School", "")
            t2.cell(i+2, 1).text = edu.get("Degree", "")
            t2.cell(i+2, 2).text = edu.get("Status", "Yes")

    # Table 3: SKILLS (Start at Row 2 to keep headers)
    t3 = doc.tables[3]
    for i, sk in enumerate(ai_data.get("Skills", [])):
        if i + 2 < len(t3.rows):
            t3.cell(i+2, 0).text = sk.get("Cat", "")
            t3.cell(i+2, 1).text = sk.get("Exp", "")

    # --- 2. SURGICAL PLACEHOLDER REPLACEMENT ---
    # We use a loop that replaces the placeholder text directly
    def replace_placeholder(old, new):
        for p in doc.paragraphs:
            if old in p.text:
                p.text = p.text.replace(old, str(new))

    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        if job:
            replace_placeholder(f"COMPANY{i}", job['Comp'])
            replace_placeholder(f"TITLE{i}", job['Title'])
            # Replace the date string on that specific line
            for p in doc.paragraphs:
                if job['Comp'] in p.text:
                    p.text = p.text.replace("MMM YYYY – CURRENT", job['Dates']).replace("MMM YYYY – MMM YYYY", job['Dates'])
            
            # Bullets
            bullet_tag = "Bullets" if i == 1 else f"Bullets{i}"
            for p in doc.paragraphs:
                if bullet_tag in p.text:
                    p.text = ""
                    for b in job['Bullets']:
                        p.insert_paragraph_before(f"• {b}")
        else:
            # Delete unused placeholders
            replace_placeholder(f"COMPANY{i}", "")
            replace_placeholder(f"TITLE{i}", "")
            replace_placeholder(f"Bullets{i}", "")

    # 3. INTERVIEW RESULTS
    replace_placeholder("ANSWER", manual_inputs["interview_results"])

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
        
        # Proper filename extraction
        real_name = raw_text.split('\n')[0].strip()
        fname = real_name.split()[0] if real_name else "Candidate"
        st.success(f"Success! Generated for {real_name}")
        st.download_button("Download Resume", data=doc_bytes, file_name=f"FM_Formatted_{fname}.docx")
    except Exception as e:
        st.error(f"Error: {e}")
