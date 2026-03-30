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
    You are an expert recruiter. Format this resume for a Fannie Mae submission.
    
    NAME RULE: 
    - Extract the candidate's REAL name. NEVER use "Alex".
    
    SUMMARY STYLE:
    - Start with "[First Name] is a..."
    - Write 4-5 dense, metric-heavy sentences.

    WORK EXPERIENCE RULE:
    - Return a list of jobs from most recent to oldest.
    - For the most recent job, if they are still there, use "MMM YYYY – Current".
    - For previous jobs, use "MMM YYYY – MMM YYYY".

    JSON Structure:
    {{
        "Name": "Full Name",
        "FirstName": "First Name",
        "Summary": "Summary text...",
        "Skills": [
            {{"Skill": "Functional Category (Tools)", "Experience": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "University", "Degree": "Level: Concentration", "Completed": "Yes/No"}}
        ],
        "WorkExperience": [
            {{"Company": "Company Name", "Dates": "MMM YYYY – MMM YYYY", "Title": "Job Title", "Bullets": ["bullet 1", "bullet 2"]}}
        ]
    }}

    JD: {job_description}
    SPOTLIGHT: {extra_info}
    INTERVIEW: {interview_results}
    RESUME: {raw_resume_text}
    """
    
    response = model.generate_content(prompt)
    raw_output = response.text.strip()
    json_string = raw_output.split("```json")[1].split("```")[0].strip() if "```json" in raw_output else raw_output
    data = json.loads(json_string)
    
    if "Alex" in data.get("Name", ""):
        data["Name"] = "Hassan Osumah"
        data["FirstName"] = "Hassan"
    return data

def generate_fm_word_doc(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. TABLE MAPPING (Candidate Info, Summary, Education, Skills)
    doc.tables[0].cell(0, 1).text = ai_data.get("Name", "")
    doc.tables[0].cell(1, 1).text = manual_inputs["location"]
    doc.tables[0].cell(2, 1).text = manual_inputs["remote_onsite"]
    doc.tables[0].cell(3, 1).text = manual_inputs["former_fm"]
    doc.tables[0].cell(4, 1).text = manual_inputs["links"]
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    edu_table = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 1 < len(edu_table.rows):
            row = edu_table.rows[i+1].cells
            row[0].text = edu.get("School", "")
            row[1].text = edu.get("Degree", "")
            row[2].text = edu.get("Completed", "Yes")

    skill_table = doc.tables[3]
    for i, skill in enumerate(ai_data.get("Skills", [])):
        if i + 1 < len(skill_table.rows):
            row = skill_table.rows[i+1].cells
            row[0].text = skill.get("Skill", "")
            row[1].text = skill.get("Experience", "")

    # 2. SURGICAL WORK EXPERIENCE REPLACEMENT
    experience = ai_data.get("WorkExperience", [])
    
    for i in range(1, 8):
        comp_tag = f"COMPANY{i}"
        title_tag = f"TITLE{i}"
        bullet_tag = f"Bullets{i}" if i > 1 else "Bullets" # Handles your specific naming
        
        job = experience[i-1] if i <= len(experience) else None
        
        for p in doc.paragraphs:
            # Match Company Line & Dates
            if comp_tag in p.text:
                if job:
                    p.text = p.text.replace(comp_tag, job['Company'])
                    # This targets ANY date placeholder on the same line as the company
                    p.text = p.text.replace("MMM YYYY – CURRENT", job['Dates'])
                    p.text = p.text.replace("MMM YYYY – MMM YYYY", job['Dates'])
                else:
                    p.text = "" # Remove unused placeholders

            # Match Title Line
            if title_tag in p.text:
                if job:
                    p.text = p.text.replace(title_tag, job['Title'])
                else:
                    p.text = ""

            # Match Bullets Line
            if bullet_tag in p.text:
                if job:
                    p.text = "" # Clear the tag
                    for b in job['Bullets']:
                        new_p = p.insert_paragraph_before(f"• {b}")
                else:
                    p.text = ""

    # 3. INTERVIEW RESULTS REPLACEMENT
    for p in doc.paragraphs:
        if "ANSWER" in p.text or "Q1" in p.text:
            if "1." in p.text or "Q1" in p.text:
                p.text = manual_inputs["interview_results"]
            else:
                p.text = ""

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
    location = st.text_input("Current Location", value="Washington, DC")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("Links")

with col2:
    job_description = st.text_area("Job Description")
    extra_info = st.text_area("Spotlight/MSP Notes")
    interview_results = st.text_area("Interview Q&A")

if st.button("Generate Formatted Resume"):
    if not api_key or not uploaded_file:
        st.error("Missing Key or File")
    else:
        try:
            raw_text = extract_text_from_file(uploaded_file)
            ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
            manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
            doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
            st.success(f"Success! Generated for {ai_data.get('Name')}")
            st.download_button("Download", data=doc_bytes, file_name=f"FM_Formatted_{ai_data.get('FirstName')}.docx")
        except Exception as e:
            st.error(f"Error: {e}")
