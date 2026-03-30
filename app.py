import streamlit as st
import google.generativeai as genai
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
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
    You are an expert technical recruiter and resume formatter.
    
    CRITICAL NAME RULES: 
    1. Correctly extract the candidate's FULL NAME from the resume. Do NOT use placeholders like "Alex Chen".
    2. In the SUMMARY section, you MUST use the candidate's FIRST NAME ONLY to start (e.g., "Hassan is a..."). Never use their last name in the summary paragraph.

    ZERO FABRICATION: Do not invent skills or experience.
    
    Task 1: Write a highly persuasive 4-5 sentence Summary. 
    - Start with [First Name] is a...
    - Focus on selling the manager using keywords from the JD and Spotlight notes. 
    - Include quantitative metrics (e.g., number of incidents handled) if available in the Resume or Interview Results.
    
    Task 2: Create exactly 4 Skill rows. 
    - Use the Category naming convention: "Broad Category (Tool 1, Tool 2)".
    - Ensure the Broad Category uses terminology from the JD (e.g., "SIEM").

    JSON Structure:
    {{
        "Name": "Full Name from Resume",
        "Summary": "Summary text using First Name only...",
        "Skills": [
            {{"Skill": "Category (Tools)", "Experience": "X+ years, current"}}
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
    
    # Clean the JSON output from markdown wrappers
    if "```json" in raw_output:
        json_string = raw_output.split("```json")[1].split("```")[0].strip()
    elif "```" in raw_output:
        json_string = raw_output.split("```")[1].split("```")[0].strip()
    else:
        json_string = raw_output

    try:
        return json.loads(json_string)
    except json.JSONDecodeError:
        # Fallback: Find the actual JSON object within the text
        start_index = json_string.find('{')
        end_index = json_string.rfind('}')
        if start_index != -1 and end_index != -1:
            json_string = json_string[start_index:end_index+1]
            return json.loads(json_string)
        else:
            raise ValueError("AI response formatting error. Please try clicking Generate again.")

def generate_fm_word_doc(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    
    if os.path.exists(template_path):
        doc = docx.Document(template_path)
    else:
        doc = docx.Document()
        st.warning("Template file not found in GitHub. Using default formatting.")

    # 1. FILL CANDIDATE TABLE (Table 0)
    full_name = ai_data.get("Name", "")
    table = doc.tables[0]
    table.cell(0, 1).text = full_name
    table.cell(1, 1).text = manual_inputs["location"]
    table.cell(2, 1).text = manual_inputs["remote_onsite"]
    table.cell(3, 1).text = manual_inputs["former_fm"]
    table.cell(4, 1).text = manual_inputs["links"]

    # 2. FILL SUMMARY (Table 1)
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 3. FILL EDUCATION (Table 2)
    edu_table = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 1 < len(edu_table.rows):
            row = edu_table.rows[i+1].cells
            row[0].text = edu.get("School", "")
            row[1].text = edu.get("Degree", "")
            row[2].text = edu.get("Completed", "Yes")

    # 4. FILL SKILLS (Table 3)
    skill_table = doc.tables[3]
    for i, skill in enumerate(ai_data.get("Skills", [])):
        if i + 1 < len(skill_table.rows):
            row = skill_table.rows[i+1].cells
            row[0].text = skill.get("Skill", "")
            row[1].text = skill.get("Experience", "")

    # 5. WORK EXPERIENCE (Text injection)
    doc.add_paragraph("\nPROFESSIONAL EXPERIENCE").bold = True
    for job in ai_data.get("WorkExperience", []):
        p = doc.add_paragraph()
        p.add_run(f"{job.get('Company', 'Company')}").bold = True
        p.add_run(f"\t{job.get('Dates', '')}")
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
    location = st.text_input("Location (City, ST)")
    remote_onsite = st.selectbox("Remote/Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("LinkedIn Link")

with col2:
    job_description = st.text_area("Job Description")
    extra_info = st.text_area("Spotlight/Extra Info")
    interview_results = st.text_area("Interview Q&A")

if st.button("Generate"):
    if not api_key:
        st.error("Please enter an API Key.")
    elif not uploaded_file:
        st.error("Please upload a resume.")
    else:
        try:
            raw_text = extract_text_from_file(uploaded_file)
            manual_inputs = {
                "location": location, 
                "remote_onsite": remote_onsite, 
                "former_fm": former_fm, 
                "links": links, 
                "interview_results": interview_results
            }
            ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
            doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
            
            st.success("Resume Generated!")
            st.download_button(
                label="Download Formatted Resume", 
                data=doc_bytes, 
                file_name=f"FM_Formatted_{ai_data.get('Name', 'Candidate').replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        except Exception as e:
            st.error(f"An error occurred: {e}")
