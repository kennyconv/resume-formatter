import streamlit as st
import google.generativeai as genai
import docx
import PyPDF2
from io import BytesIO
import json

# --- Core Functions ---

def extract_text_from_file(uploaded_file):
    """Reads raw text from PDF, DOCX, or TXT files."""
    file_name = uploaded_file.name.lower()
    text = ""
    
    if file_name.endswith('.pdf'):
        pdf_reader = PyPDF2.PdfReader(uploaded_file)
        for page in pdf_reader.pages:
            # Some PDFs return None for blank pages, so we handle that safely
            page_text = page.extract_text()
            if page_text:
                text += page_text + "\n"
                
    elif file_name.endswith('.docx'):
        doc = docx.Document(uploaded_file)
        for para in doc.paragraphs:
            text += para.text + "\n"
            
    elif file_name.endswith('.txt'):
        text = uploaded_file.getvalue().decode("utf-8")
        
    else:
        raise ValueError("Unsupported file format. Please upload a PDF, DOCX, or TXT.")
        
    return text

def parse_and_generate_with_ai(raw_resume_text, job_description, extra_info, interview_results, api_key):
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    prompt = f"""
    You are an expert technical recruiter and resume formatter for a staffing agency supporting Fannie Mae. Your goal is to format a candidate's resume to achieve two outcomes: 
    1) Get "shortlisted" by the VMS/MSP by heavily integrating exact keywords.
    2) Sell the Hiring Manager by writing a highly persuasive narrative that proves the candidate will succeed in this specific role.

    CRITICAL INSTRUCTIONS & GUARDRAILS:
    - ZERO FABRICATION: You must NEVER invent, fabricate, or assume experience, tools, or projects that the candidate does not explicitly have. All claims must be grounded in the provided Resume or Interview Results.
    - SOURCE OF TRUTH: Treat the "Extra Info/Spotlight Notes" as the ultimate source of truth for what the manager actually wants. If the manager emphasizes specific technologies over others, your output must reflect that.
    
    Task 1: Write a highly persuasive "Summary".
    - MUST be exactly 4 to 5 sentences long.
    - The opening sentence MUST strictly follow this format: "[Candidate First Name] is a [Target Job Title] with [X]+ years of experience..."
    - Seamlessly weave in exact keywords, tools, and methodologies requested by the manager.
    - Explicitly mention 1 or 2 of the candidate's past companies. Connect the exact architecture or projects they built there to the specific problems the Hiring Manager needs solved.
    - Use insights from the "Supplier Technical Interview Results" to validate their technical depth and prove they are a perfect fit.
    - Include quantitative metrics (e.g., number of incidents handled, volume of data, team size) if available in the Candidate Resume OR the Supplier Technical Interview Results.
    - NAME-DROP GUARDRAIL: Use metadata like "Cost Center" or "Manager Title" to understand the domain and set the tone, but DO NOT explicitly write those internal names in the summary. ONLY explicitly name-drop a specific team, project, or proprietary system (e.g., "MUSE") in the closing sentence IF it was explicitly mentioned in the Job Description or spoken about in the Spotlight Call transcript.
    - Do not use bullet points. Write one dense, professional paragraph.

    Task 2: Create exactly 4 "Skill & Competency" rows.
    - Map the candidate's actual skills directly to the core requirements prioritized by the manager.
    - Format the Skill name strictly as: "Broad Competency Category (Specific Tool 1, Tool 2, Tool 3)". The 'Broad Competency Category' MUST utilize the exact overarching terminology from the JD (e.g., if the JD asks for "SIEM", use "SIEM" in the category name, do not just list the tools). Include specific tools from the JD/Transcript inside the parentheses ONLY if the candidate actually has experience with them.
    - Format the Experience strictly as: "X+ years, [Year Last Used or 'current']". Calculate years based strictly on the resume dates.

    Task 3: Extract standard Resume Data.
    - Extract Name, Education, and Work Experience strictly as they appear in the resume text.

    Return the information STRICTLY as a JSON object with no additional text or markdown.

    Required JSON Structure:
    {{
        "Name": "Candidate Full Name",
        "Summary": "The 4-5 sentence persuasive summary...",
        "Skills": [
            {{"Skill": "Backend API & Integration (Python, REST APIs)", "Experience": "7+ years, current"}},
            {{"Skill": "Skill Name 2 (Subskills...)", "Experience": "Y+ years, 2022"}},
            {{"Skill": "Skill Name 3 (Subskills...)", "Experience": "Z+ years, current"}},
            {{"Skill": "Skill Name 4 (Subskills...)", "Experience": "W+ years, 2021"}}
        ],
        "Education": [
            {{"School": "University Name", "Degree": "Degree Level and Concentration", "Completed": "Yes or No"}}
        ],
        "WorkExperience": [
            {{
                "Company": "Company Name",
                "Dates": "Month Year - Month Year",
                "Title": "Job Title",
                "Bullets": ["bullet point 1", "bullet point 2"]
            }}
        ]
    }}
    
    --- INPUT DATA ---
    Job Description:
    {job_description}
    
    Spotlight Notes/Extra Info:
    {extra_info}

    Supplier Technical Interview Results (Q&A):
    {interview_results}
    
    Candidate Resume:
    {raw_resume_text}
    """
    
    response = model.generate_content(prompt)
    json_string = response.text.strip()
    if json_string.startswith("```json"):
        json_string = json_string.strip("```json").strip("```")
    elif json_string.startswith("```"):
        json_string = json_string.strip("```")
        
    return json.loads(json_string)

def generate_fm_word_doc(ai_data, manual_inputs):
    doc = docx.Document()

    # 1. CANDIDATE INFORMATION
    table1 = doc.add_table(rows=5, cols=2)
    table1.style = 'Table Grid'
    data_info = [
        ("Name:", ai_data.get("Name", "NAME")),
        ("Current Location: (City and State only)", manual_inputs["location"]),
        ("Remote or Onsite:", manual_inputs["remote_onsite"]),
        ("Former FM FTE or Contractor Y/N: (If yes, add CRC approval)", manual_inputs["former_fm"]),
        ("LinkedIn Profile/GitHub/Portfolio Link", manual_inputs["links"])
    ]
    for i, (label, value) in enumerate(data_info):
        row_cells = table1.rows[i].cells
        row_cells[0].text = label
        row_cells[1].text = value if value else ""
    doc.add_paragraph() 

    # 2. SUMMARY
    table2 = doc.add_table(rows=2, cols=1)
    table2.style = 'Table Grid'
    table2.rows[0].cells[0].text = "SUMMARY (Use this section for a candidate summary, similar to the Fieldglass “Comments” field)"
    table2.rows[1].cells[0].text = ai_data.get("Summary", "SUMMARY MISSING")
    doc.add_paragraph()

    # 3. EDUCATION
    doc.add_paragraph("EDUCATION")
    table3 = doc.add_table(rows=1, cols=3)
    table3.style = 'Table Grid'
    hdr_cells = table3.rows[0].cells
    hdr_cells[0].text = 'School'
    hdr_cells[1].text = 'Degree (Level & Concentration)'
    hdr_cells[2].text = 'Completed (Yes/No)?'

    for edu in ai_data.get("Education", []):
        row_cells = table3.add_row().cells
        row_cells[0].text = edu.get("School", "SCHOOL")
        row_cells[1].text = edu.get("Degree", "DEGREE")
        row_cells[2].text = edu.get("Completed", "Yes")
    doc.add_paragraph()

    # 4. SKILLS & COMPETENCIES
    doc.add_paragraph("SKILL & COMPETENCY REQUIREMENTS (Address all Skill & Competency Requirements from Fieldglass)")
    table4 = doc.add_table(rows=1, cols=2)
    table4.style = 'Table Grid'
    hdr_cells = table4.rows[0].cells
    hdr_cells[0].text = 'Skill/Competency'
    hdr_cells[1].text = 'Years of Experience and Year Last Used'

    for skill_data in ai_data.get("Skills", []):
        row_cells = table4.add_row().cells
        row_cells[0].text = skill_data.get("Skill", "Skill Name")
        row_cells[1].text = skill_data.get("Experience", "X years, current")
    doc.add_paragraph()

    # 5. WORK EXPERIENCE
    for job in ai_data.get("WorkExperience", []):
        p = doc.add_paragraph()
        p.add_run(job.get("Company", "Company Name")).bold = True
        p.add_run(f'\t{job.get("Dates", "mmm yyyy - mmm yyyy")}')
        doc.add_paragraph(job.get("Title", "Job Title"))
        for bullet in job.get("Bullets", []):
            doc.add_paragraph(bullet, style='List Bullet')
        doc.add_paragraph()

    # 6. SUPPLIER TECHNICAL INTERVIEW RESULTS
    doc.add_paragraph("SUPPLIER TECHNICAL INTERVIEW RESULTS").bold = True
    if manual_inputs["interview_results"]:
        for line in manual_inputs["interview_results"].split('\n'):
            if line.strip(): 
                doc.add_paragraph(line.strip())
    else:
        for i in range(1, 6):
            doc.add_paragraph(f"Q{i} ANSWER")

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- Streamlit Web App Interface ---

st.set_page_config(page_title="Fannie Mae Resume Formatter", layout="wide")
st.title("📄 Fannie Mae Resume Auto-Formatter")

with st.sidebar:
    st.header("Settings")
    api_key = st.text_input("Enter Google Gemini API Key:", type="password")
    st.markdown("*Required to extract and generate content.*")

col1, col2 = st.columns(2)

with col1:
    st.subheader("1. Candidate Details")
    uploaded_file = st.file_uploader("Upload Candidate Resume (PDF, DOCX, TXT)", type=["pdf", "docx", "txt"])
    location = st.text_input("Current Location (City, ST):", placeholder="e.g., Dallas, TX")
    remote_onsite = st.selectbox("Remote or Onsite:", ["Onsite", "Remote", "Hybrid"])
    former_fm = st.selectbox("Former FM FTE or contractor:", ["N", "Y"])
    links = st.text_input("LinkedIn/GitHub Link:", placeholder="[https://linkedin.com/in/](https://linkedin.com/in/)...")

with col2:
    st.subheader("2. Role & Interview Context")
    job_description = st.text_area("Job Description:", height=150, placeholder="Paste the JD here...")
    extra_info = st.text_area("Spotlight Call/Extra Info:", height=100, placeholder="Paste notes, manager feedback, etc...")
    interview_results = st.text_area("Supplier Technical Interview Results (Q&A):", height=150, placeholder="Q1: What is DAX?\nA: DAX is...")

st.markdown("---")

if st.button("Generate Formatted Resume", type="primary"):
    if not api_key:
        st.error("Please enter your Gemini API key in the sidebar.")
    elif not uploaded_file:
        st.error("Please upload a resume PDF.")
    else:
        with st.spinner("Extracting text, writing summary, matching skills, and formatting document... This may take 15-30 seconds."):
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
                word_doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
                
                st.success("Resume successfully generated and formatted!")
                st.download_button(
                    label="⬇️ Download Formatted Resume (.docx)",
                    data=word_doc_bytes,
                    file_name=f"FM_Formatted_{ai_data.get('Name', 'Candidate').replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )
            except Exception as e:
                st.error(f"An error occurred during generation: {e}")
