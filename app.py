import streamlit as st
import google.generativeai as genai
import docx
import PyPDF2
from io import BytesIO
import json
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
        # Scan Body
        for para in doc.paragraphs:
            if para.text.strip(): text += para.text + "\n"
        # Scan Tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip(): text += cell.text + " "
    return text

def parse_and_generate_with_ai(raw_resume_text, api_key):
    genai.configure(api_key=api_key)
    # FIXED: Using the stable model identifier to prevent 404
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    You are a data extraction tool. Extract data from the resume into JSON.
    DO NOT rewrite, summarize, or improve the text. Copy bullets EXACTLY.

    1. Extract Full Name.
    2. Extract Education (School, Degree).
    3. Extract Work History (Company, Title, Dates, Bullets).

    JSON Structure:
    {{
        "FullName": "Name",
        "Education": [
            {{"School": "School Name", "Degree": "Degree Name"}}
        ],
        "Jobs": [
            {{"Company": "Company", "Title": "Title", "Dates": "MMM YYYY – MMM YYYY", "Bullets": ["Exact Bullet 1"]}}
        ]
    }}

    RESUME TEXT:
    {raw_resume_text}
    """
    
    response = model.generate_content(prompt)
    res_text = response.text.strip()
    if "```json" in res_text:
        res_text = res_text.split("```json")[1].split("```")[0].strip()
    return json.loads(res_text)

def run_level_replace(paragraph, target, replacement):
    """Surgical replacement to keep Tab Stops and template formatting."""
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

def generate_fm_word_doc(ai_data, manual_inputs):
    template_path = "Fannie Mae Resume Format Template.docx"
    doc = docx.Document(template_path)

    # 1. CANDIDATE INFO (NAME, City, ST, Onsite, Y/N, LINK)
    t0 = doc.tables[0]
    t0.cell(1, 1).text = ai_data.get("FullName", "NAME").title()
    t0.cell(2, 1).text = manual_inputs["location"]
    t0.cell(3, 1).text = manual_inputs["remote_onsite"]
    t0.cell(4, 1).text = manual_inputs["former_fm"]
    t0.cell(5, 1).text = manual_inputs["links"]

    # 2. EDUCATION
    t2 = doc.tables[2]
    for i, edu in enumerate(ai_data.get("Education", [])):
        if i + 2 < len(t2.rows):
            t2.cell(i+2, 0).text = edu.get("School", "SCHOOL")
            t2.cell(i+2, 1).text = edu.get("Degree", "DEGREE")
            t2.cell(i+2, 2).text = "Yes"

    # 3. WORK HISTORY (Mapping Job1Bullets, etc.)
    jobs = ai_data.get("Jobs", [])
    for p in doc.paragraphs:
        p_text_low = p.text.lower()
        for i in range(1, 8):
            job = jobs[i-1] if i <= len(jobs) else None
            c_tag, t_tag, b_tag = f"company{i}", f"title{i}", f"job{i}bullets"
            
            if c_tag in p_text_low:
                if job:
                    run_level_replace(p, c_tag, job['Company'])
                    # Preserves the Right-Tab for dates
                    for d_tag in ["mmm yyyy – current", "mmm yyyy – mmm yyyy"]:
                        if d_tag in p.text.lower():
                            run_level_replace(p, d_tag, job['Dates'])
                else: p.text = "" # Remove unused role blocks
            
            elif t_tag in p_text_low:
                if job: run_level_replace(p, t_tag, job['Title'])
                else: p.text = ""

            elif b_tag in p_text_low:
                orig_style = p.style
                p.text = ""
                if job:
                    for b in job['Bullets']:
                        p.insert_paragraph_before(f"• {b}", style=orig_style)

    # 4. INTERVIEW RESULTS (Q1-Q5 and ANSWER1-ANSWER5)
    for p in doc.paragraphs:
        for i in range(1, 6):
            q_tag, a_tag = f"Q{i}", f"ANSWER{i}"
            val_q = manual_inputs.get(f"q{i}", "")
            val_a = manual_inputs.get(f"a{i}", "")
            
            # If the user provided a question/answer, replace the tag. 
            # If not, delete the line entirely.
            if q_tag in p.text:
                p.text = p.text.replace(q_tag, val_q) if val_q else ""
            if a_tag in p.text:
                p.text = p.text.replace(a_tag, val_a) if val_a else ""

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
    location = st.text_input("Current Location: (City, ST)")
    remote_onsite = st.selectbox("Remote or Onsite", ["Onsite", "Remote"])
    former_fm = st.selectbox("Former FM FTE or Contractor?", ["N", "Y"])
    links = st.text_input("LinkedIn Profile Link")

with c2:
    st.subheader("Supplier Technical Interview Results")
    q1 = st.text_input("Question 1", key="iq1")
    a1 = st.text_area("Answer 1", key="ia1")
    q2 = st.text_input("Question 2", key="iq2")
    a2 = st.text_area("Answer 2", key="ia2")
    q3 = st.text_input("Question 3", key="iq3")
    a3 = st.text_area("Answer 3", key="ia3")
    q4 = st.text_input("Question 4", key="iq4")
    a4 = st.text_area("Answer 4", key="ia4")
    q5 = st.text_input("Question 5", key="iq5")
    a5 = st.text_area("Answer 5", key="ia5")

if st.button("Generate Formatted Resume"):
    try:
        raw_text = extract_text_from_file(uploaded_file)
        ai_data = parse_and_generate_with_ai(raw_text, api_key)
        
        manual_inputs = {
            "location": location, "remote_onsite": remote_onsite, 
            "former_fm": former_fm, "links": links,
            "q1": q1, "a1": a1, "q2": q2, "a2": a2, 
            "q3": q3, "a3": a3, "q4": q4, "a4": a4, "q5": q5, "a5": a5
        }
        
        doc_bytes = generate_fm_word_doc(ai_data, manual_inputs)
        name = ai_data.get("FullName", "Candidate").title()
        st.success(f"Successfully processed {name}")
        st.download_button("Download Resume", data=doc_bytes, file_name=f"{name} Fannie Mae Format.docx")
    except Exception as e:
        st.error(f"Error: {e}")
