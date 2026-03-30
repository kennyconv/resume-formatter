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
    You are an expert technical recruiter. Your goal is to format a resume for a high-level job submission.
    
    NAME RULES: 
    1. Extract the candidate's FULL NAME correctly from the resume.
    2. In the SUMMARY section, use the candidate's FIRST NAME ONLY to start. 

    SUMMARY GUIDELINES:
    - Write 4-5 sentences.
    - Start with "[First Name] is a..."
    - Analyze the Job Description and Spotlight Notes to identify the "Win Themes".
    - Prove the candidate meets these requirements using specific examples and metrics from their resume and interview.
    - Mention 1-2 previous employers if relevant.
    - Reference a specific technical insight from the interview results to prove depth.

    SKILLS SECTION STYLE GUIDE (CRITICAL):
    - You MUST create exactly 4 rows.
    - Each row must follow this EXACT format: "Functional Competency & Strategy (Tool 1, Tool 2, Concept 1, Tool 3)"
    - DO NOT use simple labels like "Java" or "SIEM". 
    - INSTEAD, use high-level descriptive categories like:
        * "Cybersecurity & SOC Operations (Threat Detection, Incident Response, Insider Threat)"
        * "CI/CD Pipeline Automation & DevOps Testing (Jenkins, Docker, Cloud Environments)"
        * "Cloud Data & API Integration (AWS Redshift, S3, JSON, REST API Validation)"
    - Only include tools/skills the candidate actually possesses.
    - Calculate years of experience and "Year Last Used" based on the resume dates.

    JSON Structure:
    {{
        "Name": "Full Name",
        "FirstName": "First Name",
        "Summary": "Persuasive summary text...",
        "Skills": [
            {{"Skill": "Functional Category (Specific Tools & Concepts)", "Experience": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "University", "Degree": "Level: Concentration", "Completed": "Yes/No"}}
        ],
        "WorkExperience": [
            {{"Company": "Company Name", "Dates": "Month Year – Month Year", "Title": "Job Title", "Bullets": ["bullet 1", "bullet 2"]}}
        ]
    }}

    JOB DESCRIPTION: {job_description}
    SPOTLIGHT NOTES: {extra_info}
    INTERVIEW Q&A: {interview_results}
    CANDIDATE RESUME: {raw_resume_text}
    """
    
    response = model.generate_content(prompt)
    raw_output = response.text.strip()
    
    if "```json" in raw_output:
        json_string = raw_output.split("```json")[1].split("```")[0].strip()
    elif "```" in raw_output:
        json_string = raw_output.split("```")[1].split("```")[0].strip()
    else:
        json_string = raw_output

    return json.loads(json_string)

def generate_fm_word_doc(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. CANDIDATE INFO TABLE (Table 0)
    table = doc.tables[0]
    table.cell(0, 1).text = ai_data.get("Name", "")
    table.cell(1, 1).text = manual_inputs["location"]
    table.cell(2, 1).text = manual_inputs["remote_onsite"]
    table.cell(3, 1).text = manual_inputs["former_fm"]
    table.cell(4, 1).text = manual_inputs["links"]

    # 2. SUMMARY (Table 1)
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 3. EDUCATION (Table 2) - Skip Header Row 0
    edu_table = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        row_idx = i + 1
        if row_idx < len(edu_table.rows):
            cells = edu_table.rows[row_idx].cells
            cells[0].text = edu.get("School", "")
            cells[1].text = edu.get("Degree", "")
            cells[2].text = edu.get("Completed", "Yes")

    # 4. SKILLS (Table 3) - Skip Header Row 0
    skill_table = doc.tables[3]
    for i, skill in enumerate(ai_data.get("Skills", [])):
        row_idx = i + 1
        if row_idx < len(skill_table.rows):
            cells = skill_table.rows[row_idx].cells
            cells[0].text = skill.get("Skill", "")
            cells[1].text = skill.get("Experience", "")

    # 5. WORK EXPERIENCE
    doc.add_paragraph("\nPROFESSIONAL EXPERIENCE").bold = True
    for job in ai_data.get("WorkExperience", []):
        p = doc.add_paragraph()
        run = p.add_run(f"{job.get('Company')}")
        run.bold = True
        p.add_run(f"\t{job.get('Dates')}")
        doc.add_paragraph(job.get("Title", ""))
        for bullet in job.get("Bullets", []):
            doc.add_paragraph(f"• {bullet}")

    # 6. INTERVIEW RESULTS
    doc.add_paragraph("\nSUPPLIER TECHNICAL INTERVIEW RESULTS").bold = True
    doc.add_paragraph(manual_inputs["interview_results"])

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit UI ---
st.set_page_config(page_title="Fannie Mae Formatter", layout="wide")
st.title("📄 Fannie Mae Resume Auto-Formatter")

with st.sidebar:
    api_key = st.text_input("Gemini API Key:", type="password")

col1, col2 = st.columns(2)
with col1:
    uploaded_file = st.file_uploader("Upload Resume", type=["pdf", "docx"])
    location = st.text_input("Current Location (City, ST)", placeholder="e.g., Dallas, TX")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM FTE or Contractor?", ["N", "Y"])
    links = st.text_input("Links (LinkedIn/GitHub)")

with col2:
    job_description = st.text_area("Job Description (Fieldglass)")
    extra_info = st.text_area("Spotlight Call / MSP Notes")
    interview_results = st.text_area("Supplier Technical Interview Q&A")

if st.button("Generate Formatted Resume"):
    if not api_key or not uploaded_file:
        st.error("Missing API Key or Resume File.")
    else:
        try:
            raw_text = extract_text_from_file(uploaded_file)
            ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
            
            manual_inputs = {
                "location": location, "remote_onsite": remote_onsite, 
                "former_fm": former_fm, "links": links, 
                "interview_results": interview_results
            }
            
            doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
            st.success(f"Formatted Resume for {ai_data.get('Name')} is ready!")
            st.download_button("Download Resume", data=doc_bytes, file_name=f"FM_Formatted_{ai_data.get('FirstName')}.docx")
        except Exception as e:
            st.error(f"Error during generation: {e}")
