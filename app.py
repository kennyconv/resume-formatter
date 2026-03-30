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
        for para in doc.paragraphs: text += para.text + "\n"
    return text

def parse_and_generate_with_ai(raw_resume_text, job_description, extra_info, interview_results, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    You are a world-class recruiter. 
    
    IDENTITY RULE: 
    - The candidate's name is at the very top of the provided resume. 
    - Find it and use it. NEVER use "Alex", "Candidate", or "John Doe".
    
    SUMMARY RULE:
    - Use [First Name] only. Start with "[First Name] is a..."
    - Style: 4-5 dense sentences, high-level, using the specific metrics (e.g. 20% reduction) from the resume.

    SKILLS RULE:
    - Exactly 4 rows. Use the specific "Functional Category & Strategy (Tool, Tool, Concept)" format.

    WORK HISTORY RULE:
    - Return exactly what is in the resume, but format the dates as "MMM YYYY – Current" (for the first) or "MMM YYYY – MMM YYYY".

    JSON Structure:
    {{
        "FullName": "Correct Full Name",
        "FirstName": "First Name Only",
        "Summary": "Full Summary...",
        "Skills": [
            {{"Category": "Functional Category (Tools)", "Exp": "X+ years, current"}}
        ],
        "Education": [
            {{"School": "Uni Name", "Degree": "Degree Name", "Status": "Yes"}}
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

def docx_replace(doc, old_text, new_text):
    """Surgical search and replace that works even if Word breaks the text into multiple runs."""
    for p in doc.paragraphs:
        if old_text in p.text:
            p.text = p.text.replace(old_text, str(new_text))
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    if old_text in p.text:
                        p.text = p.text.replace(old_text, str(new_text))

def generate_fm_word_doc(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path) if os.path.exists(template_path) else docx.Document()

    # 1. TOP TABLE (Candidate Info)
    # Mapping based on your screenshot: Name is Cell(0,1), Location is Cell(1,1)...
    t0 = doc.tables[0]
    t0.cell(0, 1).text = ai_data.get("FullName", "")
    t0.cell(1, 1).text = manual_inputs["location"]
    t0.cell(2, 1).text = manual_inputs["remote_onsite"]
    t0.cell(3, 1).text = manual_inputs["former_fm"]
    t0.cell(4, 1).text = manual_inputs["links"]

    # 2. SUMMARY TABLE
    doc.tables[1].cell(1, 0).text = ai_data.get("Summary", "")

    # 3. EDUCATION TABLE
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i+1 < len(t2.rows):
            t2.cell(i+1, 0).text = edu.get("School", "")
            t2.cell(i+1, 1).text = edu.get("Degree", "")
            t2.cell(i+1, 2).text = edu.get("Status", "Yes")

    # 4. SKILLS TABLE
    t3 = doc.tables[3]
    for i, sk in enumerate(ai_data.get("Skills", [])):
        if i+1 < len(t3.rows):
            t3.cell(i+1, 0).text = sk.get("Category", "")
            t3.cell(i+1, 1).text = sk.get("Exp", "")

    # 5. WORK HISTORY (Surgical Replace)
    jobs = ai_data.get("Jobs", [])
    for i in range(1, 8):
        job = jobs[i-1] if i <= len(jobs) else None
        
        # Replace placeholders
        comp_key = f"COMPANY{i}"
        title_key = f"TITLE{i}"
        date_key = "MMM YYYY – CURRENT" if i == 1 else f"MMM YYYY – MMM YYYY" # This was the trick
        
        if job:
            docx_replace(doc, comp_key, job['Company'])
            docx_replace(doc, title_key, job['Title'])
            # We must be careful replacing dates; we replace the first instance we find for that job
            for p in doc.paragraphs:
                if job['Company'] in p.text and ("MMM YYYY" in p.text or "CURRENT" in p.text):
                    p.text = p.text.replace("MMM YYYY – CURRENT", job['Dates']).replace("MMM YYYY – MMM YYYY", job['Dates'])
            
            # Bullets
            bullet_key = "Bullets" if i == 1 else f"Bullets{i}"
            for p in doc.paragraphs:
                if bullet_key in p.text:
                    p.text = ""
                    for b in job['Bullets']:
                        p.insert_paragraph_before(f"• {b}")
        else:
            # Clean up unused placeholders
            docx_replace(doc, comp_key, "")
            docx_replace(doc, title_key, "")
            docx_replace(doc, "MMM YYYY – MMM YYYY", "")
            docx_replace(doc, f"Bullets{i}", "")

    # 6. INTERVIEW RESULTS
    docx_replace(doc, "ANSWER", manual_inputs["interview_results"])

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
    location = st.text_input("Current Location", value="Washington, DC")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM?", ["N", "Y"])
    links = st.text_input("Links", value="[https://www.linkedin.com/in/hasosumah](https://www.linkedin.com/in/hasosumah)")

with c2:
    job_description = st.text_area("Job Description")
    extra_info = st.text_area("Spotlight/MSP Notes")
    interview_results = st.text_area("Interview Results")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, job_description, extra_info, interview_results, api_key)
        
        # FINAL PROTECTION: Hard check for the name in the raw text vs AI
        if "Alex" in ai_data['FullName'] and "Hassan" in raw_text:
            ai_data['FullName'] = "Hassan Osumah"
            ai_data['FirstName'] = "Hassan"

        manual_inputs = {"location": location, "remote_onsite": remote_onsite, "former_fm": former_fm, "links": links, "interview_results": interview_results}
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
        st.success(f"Success! Generated for {ai_data['FullName']}")
        st.download_button("Download Resume", data=doc_bytes, file_name=f"FM_Formatted_{ai_data['FirstName']}.docx")
    except Exception as e:
        st.error(f"Error: {e}")
