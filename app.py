import streamlit as st
import google.generativeai as genai
import docx
from docx.text.paragraph import Paragraph
from docx.shared import Pt
import PyPDF2
import json
import re
import os
import copy

# ====================================================================
# --- STREAMLIT UI & FORM INPUTS ---
# ====================================================================

st.set_page_config(page_title="Fannie Mae Precision Extractor", layout="wide")
st.title("📄 Fannie Mae Precision Extractor & Template Filler")

with st.sidebar:
    st.header("🔑 API Key")
    # SECURITY FIX: Removed hardcoded API key so it is safe for GitHub
    API_KEY = st.text_input("Gemini API Key", type="password")
    st.info("Paste your Gemini API key here to run the tool.")

st.header("📋 Candidate Information")
col1, col2 = st.columns(2)
with col1:
    Current_Location_City_ST = st.text_input("Current Location (City, ST)")
    Remote_or_Onsite = st.selectbox("Remote or Onsite", ["Remote", "Onsite"], index=1)
with col2:
    Former_FM = st.selectbox("Former FM", ["Y - Per CRC, this candidate is eligible for rehire", "N"], index=1)
    LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link")

st.header("🎤 Supplier Technical Interview Results")
qa_col1, qa_col2 = st.columns(2)
with qa_col1:
    Question_1 = st.text_input("Question 1")
    Question_2 = st.text_input("Question 2")
    Question_3 = st.text_input("Question 3")
    Question_4 = st.text_input("Question 4")
    Question_5 = st.text_input("Question 5")
with qa_col2:
    Answer_1 = st.text_area("Answer 1", height=68)
    Answer_2 = st.text_area("Answer 2", height=68)
    Answer_3 = st.text_area("Answer 3", height=68)
    Answer_4 = st.text_area("Answer 4", height=68)
    Answer_5 = st.text_area("Answer 5", height=68)

st.header("📝 Job Description & Notes")
Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200)

st.header("📂 File Uploads")
resume_file = st.file_uploader("📤 Step 1: Upload the resume...", type=['pdf', 'docx', 'doc'])
template_file = st.file_uploader("📄 Step 2: Upload your Word Template...", type=['docx'])

# ====================================================================
# --- DATA EXTRACTION & HELPER FUNCTIONS (100% UNTOUCHED) ---
# ====================================================================

def clean_bullets(bullet_list):
    if isinstance(bullet_list, str):
        bullet_list = bullet_list.split('\n')
    cleaned = []
    for line in bullet_list:
        line = re.sub(r'^[\s\-\•\d\.\*]+', '', line).strip()
        if line:
            cleaned.append(line)
    return "\n".join(cleaned)

def clean_school(name):
    for delimiter in [' - ', ' – ', ' , ', ', ']:
        if delimiter in name:
            name = name.split(delimiter)[0]
    return name.strip()

def standardize_dates(date_str):
    month_map = {
        "01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
        "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec",
        "1": "Jan", "2": "Feb", "3": "Mar", "4": "Apr", "5": "May", "6": "Jun",
        "7": "Jul", "8": "Aug", "9": "Sep",
        "January": "Jan", "February": "Feb", "March": "Mar", "April": "Apr",
        "May": "May", "June": "Jun", "July": "Jul", "August": "Aug",
        "September": "Sep", "Sept": "Sep", "October": "Oct", "November": "Nov", "December": "Dec"
    }
    date_str = re.sub(r"['’](\d{2})", lambda m: ("20" if int(m.group(1)) < 50 else "19") + m.group(1), date_str)
    date_str = re.sub(r'(\d{1,2})/(\d{4})', lambda m: f"{month_map.get(m.group(1).zfill(2), m.group(1))} {m.group(2)}", date_str)
    for full, short in month_map.items():
        if not full.isdigit():
            date_str = re.sub(full, short, date_str, flags=re.IGNORECASE)
    return date_str

def extract_text(file_path):
    text_parts = []
    ext = os.path.splitext(file_path)[1].lower()
    try:
        if ext == '.pdf':
            pdf = PyPDF2.PdfReader(file_path)
            for page in pdf.pages:
                content = page.extract_text()
                if content:
                    text_parts.append(content)
        elif ext in ['.docx', '.doc']:
            doc = docx.Document(file_path)
            for section in doc.sections:
                for p in section.header.paragraphs:
                    if p.text.strip():
                        text_parts.append(p.text)
            for para in doc.paragraphs:
                if para.text.strip():
                    text_parts.append(para.text)
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        if cell.text.strip():
                            text_parts.append(cell.text)
    except:
        with open(file_path, 'rb') as f:
            text_parts.append(f.read().decode('utf-8', errors='ignore'))
    return "\n".join(text_parts)

# ====================================================================
# --- INDESTRUCTIBLE JSON REPAIR (100% UNTOUCHED) ---
# ====================================================================

def repair_and_load_json(raw_text):
    text = raw_text.strip()
    text = re.sub(r'^```[a-zA-Z]*\n|\n```$', '', text, flags=re.MULTILINE).strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        text = re.sub(r'"\s*\n\s*"', '",\n"', text)
        text = re.sub(r'\}\s*\n\s*\{', '},\n{', text)
        text = re.sub(r'\]\s*\n\s*"', '],\n"', text)
        try:
            return json.loads(text)
        except Exception as e:
            st.warning(f"⚠️ Minor Error: AI formatting glitch on this resume. Bypassing crash... ({e})")
            return {}

# ====================================================================
# --- PRECISION TEMPLATE ENGINE (100% UNTOUCHED) ---
# ====================================================================

def replace_tag_safely(p, tag, value, unbold=False):
    if tag.lower() not in p.text.lower():
        return False
    replaced = False
    for run in p.runs:
        if tag.lower() in run.text.lower():
            pattern = re.compile(re.escape(tag), re.IGNORECASE)
            run.text = pattern.sub(str(value), run.text)
            if unbold:
                run.font.bold = False
            replaced = True
    if not replaced:
        full_text = p.text
        pattern = re.compile(re.escape(tag), re.IGNORECASE)
        new_text = pattern.sub(str(value), full_text)
        if p.runs:
            p.runs[0].text = new_text
            if unbold:
                p.runs[0].font.bold = False
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""
    return True

def insert_bullet_line_after(paragraph, text):
    new_p_xml = copy.deepcopy(paragraph._p)
    paragraph._p.addnext(new_p_xml)
    new_p = Paragraph(new_p_xml, paragraph._parent)
    for run in new_p.runs:
        run._element.getparent().remove(run._element)
    if paragraph.runs:
        new_run = new_p.add_run(text)
        ref_font = paragraph.runs[0].font
        new_run.font.bold = ref_font.bold
        new_run.font.italic = ref_font.italic
        if ref_font.color and ref_font.color.rgb:
            new_run.font.color.rgb = ref_font.color.rgb
        new_run.font.size = ref_font.size
        new_run.font.name = ref_font.name
    else:
        new_p.add_run(text)
    return new_p

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    paragraph._p = paragraph._element = None

def process_word_doc(temp_path, mapping, out_path):
    doc = docx.Document(temp_path)
    for table in doc.tables:
        rows_to_delete = []
        for row in table.rows:
            row_text = "".join(cell.text for cell in row.cells).strip()
            for i in range(1, 4):
                tag = f"{{{{School{i}}}}}"
                if tag.lower() in row_text.lower() and not mapping.get(f"School{i}"):
                    rows_to_delete.append(row)
        for row in rows_to_delete:
            table._tbl.remove(row._tr)
    paras_to_remove = []
    kill_keys = ['Company', 'Title', 'Bullets', 'Dates', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'A1', 'A2', 'A3', 'A4', 'A5', 'Summary', 'Skill', 'Years']
    for p in list(doc.paragraphs):
        for key, value in mapping.items():
            tag = f"{{{{{key}}}}}"
            if tag.lower() in p.text.lower():
                if not value or not str(value).strip():
                    if any(k.lower() in key.lower() for k in kill_keys):
                        if p not in paras_to_remove:
                            paras_to_remove.append(p)
                    else:
                        replace_tag_safely(p, tag, "")
                else:
                    if "Bullets" in key:
                        for run in p.runs:
                            if '\n' in run.text:
                                run.text = run.text.replace('\n', '')
                        lines = str(value).split('\n')
                        replace_tag_safely(p, tag, lines[0])
                        curr_p = p
                        for line in lines[1:]:
                            curr_p = insert_bullet_line_after(curr_p, line)
                        spacer = curr_p.insert_paragraph_before("")
                        curr_p._p.addnext(spacer._p)
                        try:
                            spacer.style = doc.styles['Normal']
                        except:
                            pass
                        spacer.paragraph_format.space_after = Pt(0)
                        spacer.paragraph_format.space_before = Pt(0)
                    else:
                        is_answer = key in ['A1', 'A2', 'A3', 'A4', 'A5']
                        replace_tag_safely(p, tag, str(value).strip(), unbold=is_answer)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        tag = f"{{{{{key}}}}}"
                        if tag.lower() in p.text.lower():
                            is_answer = key in ['A1', 'A2', 'A3', 'A4', 'A5']
                            replace_tag_safely(p, tag, str(value) if value else "", unbold=is_answer)
    for p in paras_to_remove:
        try:
            delete_paragraph(p)
        except:
            pass
    doc.save(out_path)
    return out_path

# ====================================================================
# --- MAIN RUNNER (STREAMLIT ADAPTED) ---
# ====================================================================

if st.button("🚀 Generate Submission Document", type="primary"):
    if not API_KEY:
        st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
    elif not resume_file:
        st.error("❌ Error: Please upload a resume.")
    elif not template_file:
        st.error("❌ Error: Please upload a Word template.")
    else:
        with st.spinner("Processing documents and querying AI..."):
            try:
                # Save uploaded files temporarily
                resume_path = f"temp_{resume_file.name}"
                template_path = f"temp_{template_file.name}"
                
                with open(resume_path, "wb") as f:
                    f.write(resume_file.getbuffer())
                with open(template_path, "wb") as f:
                    f.write(template_file.getbuffer())

                raw_text = extract_text(resume_path)
                
                genai.configure(api_key=API_KEY)
                extract_model = genai.GenerativeModel('gemini-2.5-flash')
                summary_model = genai.GenerativeModel('gemini-2.5-flash')

                prompt = f"""
                Return a valid JSON object ONLY.

                RULES:
                1. Name: Pull from top/header (Title Case).
                2. Contractor: Agency is Company, Client is Title.
                3. Education: Extract School and Degree.
                4. Experience: Company, Title, Bullets (LIST), Dates.

                JSON Structure:
                {{
                    "FullName": "",
                    "Education": [{{"School": "", "Degree": ""}}],
                    "Experience": [{{"Company": "", "Title": "", "Bullets": [], "Dates": ""}}]
                }}

                RESUME: {raw_text}
                """

                response = extract_model.generate_content(prompt, generation_config={"response_mime_type": "application/json"})
                data = repair_and_load_json(response.text)

                name = data.get('FullName', '').title()
                mapping = {
                    "FullName": name, "Location": Current_Location_City_ST, "Remote": Remote_or_Onsite,
                    "FormerFM": Former_FM, "Links": LinkedIn_GitHub_Portfolio_Link,
                    "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                    "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                    "SUMMARY": "", "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
                    "SKILL3": "", "YEARS3": "", "SKILL4": "", "YEARS4": ""
                }

                edu = data.get('Education', [])
                for i in range(1, 4):
                    mapping[f"School{i}"] = clean_school(edu[i-1].get('School', '')) if i <= len(edu) else ""
                    mapping[f"Degree{i}"] = edu[i-1].get('Degree', '') if i <= len(edu) else ""

                exp = data.get('Experience', [])
                for i in range(1, 8):
                    if i <= len(exp):
                        mapping[f"Company{i}"] = exp[i-1].get('Company', '')
                        raw_title = exp[i-1].get('Title', '')
                        clean_title = re.sub(r'\s*\(.*$', '', raw_title).strip()
                        mapping[f"Title{i}"] = clean_title
                        mapping[f"Bullets{i}"] = clean_bullets(exp[i-1].get('Bullets', []))
                        mapping[f"Dates{i}"] = standardize_dates(exp[i-1].get('Dates', ''))
                    else:
                        mapping[f"Company{i}"] = ""
                        mapping[f"Title{i}"] = ""
                        mapping[f"Bullets{i}"] = ""
                        mapping[f"Dates{i}"] = ""

                if Job_Description_Notes_etc.strip():
                    try:
                        summary_prompt = f"""
You are a senior technical recruiter writing a Fieldglass submission summary.

Your goal is to clearly and convincingly explain why this candidate will succeed in THIS specific role by connecting their real experience directly to the responsibilities, environment, and day-to-day work of the role.

========================
🔴 CORE REQUIREMENTS
========================
- Write EXACTLY 5 sentences unless 4 sentences are clearly stronger.
- Use ONLY the candidate’s FIRST NAME (never include last name).
- The summary must sound like a real recruiter submission, not a resume summary.
- Tone must be confident, natural, specific, and credible (not generic or overly salesy).
- Every sentence must follow:
  → what they did → where → why it matters for THIS role

========================
🔴 FIRST SENTENCE RULES (STRICT)
========================
- MUST start with:
  "[First Name] is a ..." OR "[First Name] brings ..."
- MUST include:
  • role identity aligned to THIS job
  • total years of experience ("with X+ years of experience")
  • 2–3 role-relevant keywords

DO NOT USE:
- "X+ year engineer"
- "accomplished"
- "seasoned"
- "well-versed in"
- "strong background in"

========================
🔴 EXPERIENCE PRIORITIZATION
========================
- ALWAYS prioritize the MOST RECENT RELEVANT ROLE.
- Default phrasing:
  "In his/her current role at [Company], he/she ..."
- IMPORTANT:
  • Do NOT repeat the candidate’s name again in Sentence 2
  • Use "he/she" after Sentence 1

- Only use older roles if they are clearly more relevant.
- You may reference one prior role if it strengthens the case.

========================
🔴 REQUIRED STRUCTURE
========================

Sentence 1 (Positioning)
- Establish role identity, years, and domain

Sentence 2 (Recent Role + Evidence)
- MUST reference most recent relevant role
- Include real work (systems, models, pipelines, APIs, etc.)
- Be concrete (avoid vague descriptions)

Sentence 3 (TRANSLATION — CRITICAL)
- MUST connect past work → THIS role
- MUST describe what they will BUILD / SUPPORT / OWN
- MUST describe Specific outputs or systems (e.g., pipelins, workflows, APIs, models)
- Avoid vague phrases like "solutions" unless paired with a specific system or function
- Avoid generic nouns such as "applications" or "solutions" unless paired with a specific system (e.g., pipelins, workflows, APIs, models)

Preferred phrasing:
- "This directly translates into building ..."
- "This directly translates into supporting ..."
- "This includes ... which supports ..."
- "This enables ..."

HIGH-QUALITY PATTERN:
- "particularly where [two responsibilities] are tightly coupled"
- Example:
  "particularly where model development and production support are tightly coupled"

DO NOT USE:
- "aligns with"
- "fits well"
- "can apply"
- "perfect fit"
- "role's need for"
- "immediate contribution for"

========================
Sentence 4 (HOW THEY OPERATE — CRITICAL UPGRADE)
========================
- MUST describe how the candidate works in real environments
- MUST start with:
  "He operates within..." OR "She operates within..."

- Include:
  • tools (Python, AWS, etc.)
  • workflows (Git, JIRA, CI/CD, testing)
  • collaboration / delivery style

GOOD PATTERN:
"He operates within structured development workflows, using Python for data processing and model development alongside tools such as Git and JIRA to support testing, version control, and collaborative delivery."

DO NOT:
- Just list tools
- Sound like a resume skills section

========================
Sentence 5 (IMMEDIATE VALUE — NON-GENERIC)
========================
- MUST describe what the candidate will DO in this role
- MUST describe day-to-day responsibilities in concrete terms (e.g., pipelins, workflows, systems, stakeholder interaction)
- MUST avoid vague descriptors like "advanced", "robust", "high-performance" unless tied to specific work
- Avoid using the word "solutions"
- MUST use concrete system language such as:
"workflows", "pipelines", "systems", "outputs"

PREFERRED ACTION WORDS:
- "contribute to developing"
- "supporting"
- "building"
- "working closely with"

GOOD PATTERN:
"He will be able to contribute to developing model-driven solutions, supporting production systems, and working closely with stakeholders to ensure reliability across deployed applications."

DO NOT USE:
- "immediately contribute"
- "seamlessly integrate"
- "high-performance applications"
- "robust systems" (unless specific)

========================
🚫 ANTI-GENERIC LANGUAGE ENFORCEMENT
========================
DO NOT USE generic phrases such as:
- accomplished
- seasoned
- strong candidate
- great fit
- perfect fit
- aligns well
- aligns directly
- well-versed in
- swiftly
- seamlessly
- positions them well
- based on the job description
- based on the manager

If detected → rewrite before output.

========================
🔁 REPETITION CONTROL (CRITICAL)
========================
- Detect repeated domain phrases (e.g., "financial modeling", "data pipelines", "applications")
- If repeated more than twice → replace with equivalent phrasing (e.g., "model-driven workflows", "analytics pipelines", "model outputs")
  
Example replacements:
- "data science applications" → "analytics workflows", "model-driven solutions"
- "building and maintaining" → "developing", "supporting", "operating"
- "financial modeling" → "model-driven analytics", "quantitative workflows", "model outputs"
- "model-driven solutions" → "model-driven workflows", "analytics pipelines", "model outputs"

========================
🟡 CONTEXT PRIORITY
========================
- If manager notes / feedback are provided → PRIORITIZE THEM
- Still include JD keywords for MSP filtering
- Emphasize hands-on execution and real ownership

========================
🟢 SKILLS SECTION RULES
========================
- EXACTLY 4 items
- Must be:
  • specific
  • technical
  • keyword-rich

- Each must include 2–4 tools/technologies

FORMAT:
"Skill Area (Tool1, Tool2, Tool3)"

YEARS FORMAT:
"X+ years, current" OR "X+ years, 2025"

DO NOT USE:
- generic categories like "Software Engineering" or "Data Science"

GOOD EXAMPLES:
- "LLM & GenAI Development (RAG, Fine-tuning, NL2SQL, Prompt Engineering)"
- "Python Data Engineering (PySpark, Pandas, ETL Pipelines)"
- "AWS Data Pipelines (Glue, Lambda, Step Functions, S3)"
- "AWS SageMaker & Cloud ML Engineering (SageMaker, S3, EMR, Endpoint Deployment, Model Monitoring)"
- "MLOps & CI/CD Integration (MLflow, Docker, Kubernetes, Git-based Pipelines, Model Versioning & Retraining)"
- "Python-Based ML & Feature Engineering (Scikit-learn, XGBoost, TensorFlow, PySpark, Large Dataset Processing)"
- "Frontend & Full Stack Collaboration (JavaScript, Angular, React, Microservices Integration)"

========================
🧠 SELF-CHECK VALIDATOR (MANDATORY)
========================
Before outputting, verify:

1. Exactly 4–5 sentences?
2. Sentence 1 includes years + role + domain?
3. Sentence 2 uses most recent relevant role and does NOT repeat name?
4. Sentence 3 clearly translates experience → THIS role?
5. Sentence 4 starts with "He operates within..." and describes real workflow?
6. Sentence 5 is practical and non-generic?
7. No banned phrases used?
8. No repeated phrases?
9. Reads like a recruiter submission (not AI)?

If ANY fail → rewrite internally.

========================
🔵 OUTPUT FORMAT (STRICT)
========================
Return ONLY valid JSON:

{{
  "SUMMARY": "4-5 sentence summary",
  "SKILL1": "Skill Area (tools)",
  "YEARS1": "X+ years, current",
  "SKILL2": "Skill Area (tools)",
  "YEARS2": "X+ years, current",
  "SKILL3": "Skill Area (tools)",
  "YEARS3": "X+ years, current",
  "SKILL4": "Skill Area (tools)",
  "YEARS4": "X+ years, current"
}}

========================
INPUT DATA
========================

Job Description / Notes:
{Job_Description_Notes_etc}

Technical Interview Q&A:
Q1: {Question_1}
A1: {Answer_1}
Q2: {Question_2}
A2: {Answer_2}
Q3: {Question_3}
A3: {Answer_3}
Q4: {Question_4}
A4: {Answer_4}
Q5: {Question_5}
A5: {Answer_5}

Resume:
{raw_text}
"""
                        summary_response = summary_model.generate_content(summary_prompt, generation_config={"response_mime_type": "application/json"})
                        summary_data = repair_and_load_json(summary_response.text)

                        mapping["SUMMARY"] = summary_data.get("SUMMARY", "")
                        mapping["SKILL1"] = summary_data.get("SKILL1", "")
                        mapping["YEARS1"] = summary_data.get("YEARS1", "")
                        mapping["SKILL2"] = summary_data.get("SKILL2", "")
                        mapping["YEARS2"] = summary_data.get("YEARS2", "")
                        mapping["SKILL3"] = summary_data.get("SKILL3", "")
                        mapping["YEARS3"] = summary_data.get("YEARS3", "")
                        mapping["SKILL4"] = summary_data.get("SKILL4", "")
                        mapping["YEARS4"] = summary_data.get("YEARS4", "")
                    except Exception as e:
                        st.warning(f"⚠️ Warning: Summary generation failed. Proceeding without it. ({e})")

                out_file = f"Submission_{name.replace(' ', '_')}.docx"
                process_word_doc(template_path, mapping, out_file)
                
                # Provide download button
                with open(out_file, "rb") as file:
                    btn = st.download_button(
                        label="⬇️ Download Generated Document",
                        data=file,
                        file_name=out_file,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                
                st.success(f"✅ Success! Document is ready for download.")

                # Clean up temp files
                try:
                    os.remove(resume_path)
                    os.remove(template_path)
                except:
                    pass

            except Exception as e:
                st.error(f"❌ Process Failed: {str(e)}")
