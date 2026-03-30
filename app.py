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
    You are an expert technical recruiter. 
    
    CRITICAL NAME RULES: 
    1. Extract the candidate's FULL NAME correctly from the resume.
    2. In the SUMMARY section, you MUST use the candidate's FIRST NAME ONLY to start (e.g., "Hassan is a..."). 

    Task 1: Write a highly persuasive 4-5 sentence Summary. 
    - Use this EXACT style: "[First Name] is a [Title] with [X]+ years of experience specializing in [Skills], aligning perfectly with the need for [JD Requirement]. During his tenure at [Company 1] and [Company 2], he developed a deep foundation in [Specific Task]. He conducts end-to-end investigations utilizing [Tools]. As demonstrated in his technical screen, [First Name] leverages [Technical Skill] to [Result]. His unique blend of [Skill A] and [Skill B] ensures he can immediately contribute to [JD Goal]."
    - Include quantitative metrics (e.g., 50+ incidents, 15% efficiency) from the Resume or Interview.
    
    Task 2: Create exactly 4 Skill rows. 
    - Row 1: Cybersecurity & SOC Operations (Threat Detection, Incident Response, Insider Threat Investigations)
    - Row 2: SIEM & Threat Hunting (Splunk, IBM QRadar, Exabeam, Log Analysis, Correlation)
    - Row 3: Network & Security Analysis (TCP/IP, DNS, HTTP/S, Wireshark, Endpoint Security)
    - Row 4: Fraud & Behavioral Risk Analysis (Financial Transactions, Pattern Detection, Root Cause Analysis)
    - Adjust years based on actual resume dates.

    JSON Structure:
    {{
        "Name": "Full Name",
        "FirstName": "First Name",
        "Summary": "Persuasive summary...",
        "Skills": [
            {{"Skill": "Skill Row Name", "Experience": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "University", "Degree": "Degree", "Completed": "Yes/No"}}
        ],
        "WorkExperience": [
            {{"Company": "Company", "Dates": "Dates", "Title": "Title", "Bullets": ["bullet"]}}
        ]
    }}

    JD: {job_description}
    Extra: {extra_info}
    Interview: {interview_results}
    Resume: {raw_resume_text}
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

    # 3. EDUCATION (Table 2)
    edu_table = doc.tables[2]
    # Keep Header (Row 0), Fill Data starting Row 1
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 1 < len(edu_table.rows):
            row = edu_table.rows[i+1].cells
            row[0].text = edu.get("School", "")
            row[1].text = edu.get("Degree", "")
            row[2].text = edu.get("Completed", "Yes")

    # 4. SKILLS (Table 3)
    skill_table = doc.tables[3]
    # Ensure Header "Skill/Competency | Years..." is at Row 0, Data starts Row 1
    for i, skill in enumerate(ai_data.get("Skills", [])):
        if i + 1 < len(skill_table.rows):
            row = skill_table.rows[i+1].cells
            row[0].text = skill.get("Skill", "")
            row[1].text = skill.get("Experience", "")

    # 5. PROFESSIONAL EXPERIENCE (Finding the section or adding at end)
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
    location = st.text_input("Location (City, ST)", placeholder="Washington, DC")
    remote_onsite = st.selectbox("Remote/Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("LinkedIn Link")

with col2:
    job_description = st.text_area("Job Description")
    extra_info = st.text_area("Spotlight/Extra Info")
    interview_results = st.text_area("Interview Q&A")

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
            st.success("Generated!")
            st.download_button("Download Resume", data=doc_bytes, file_name=f"FM_Formatted_{ai_data.get('FirstName')}.docx")
        except Exception as e:
            st.error(f"Error: {e}")
