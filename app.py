import streamlit as st
from google import genai
from google.genai import types
import docx
from docx.text.paragraph import Paragraph
from docx.shared import Pt
import pypdf
import json
import re
import os
import copy
import time

# ====================================================================
# --- PRELOADED TEMPLATE SETTING ---
# ====================================================================
# The app will automatically look for this file in your GitHub repo
TEMPLATE_FILENAME = "Template.docx"

# ====================================================================
# --- STREAMLIT UI & FORM INPUTS ---
# ====================================================================

st.set_page_config(page_title="Fannie Mae Precision Extractor", layout="wide")

# --- 🔒 PASSWORD PROTECTION LOGIC ---
def check_password():
    """Returns `True` if the user entered the correct password."""
    
    def password_entered():
        """Checks whether a password entered by the user is correct."""
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Delete the password from memory for security
        else:
            st.session_state["password_correct"] = False

    # If the user has already unlocked the app in this session, let them pass
    if st.session_state.get("password_correct", False):
        return True

    # If not, show the login screen
    st.title("🔒 Access Restricted")
    st.text_input(
        "Please enter the team password to use this tool:", 
        type="password", 
        on_change=password_entered, 
        key="password"
    )
    
    # If they tried and failed, show an error
    if "password_correct" in st.session_state:
        st.error("❌ Incorrect password.")
        
    return False

# 🛑 STOP THE APP HERE IF THEY DON'T HAVE THE PASSWORD
if not check_password():
    st.stop()
# ------------------------------------


# 👇 IF THEY PASS THE CHECK, THE REST OF THE APP LOADS NORMALLY 👇
st.title("📄 Fannie Mae Precision Extractor & Template Filler")

with st.sidebar:
    st.header("🔑 API Configuration")
    
    # Logic: Check if the key exists in Streamlit Secrets first
    if "API_KEY" in st.secrets:
        API_KEY = st.secrets["API_KEY"]
        st.success("✅ API Key loaded from Secrets")
    else:
        # Fallback: If no secret is found, show the input box
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
# REMOVED the template uploader. Now it only asks for the resume.
resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'])

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
            pdf = pypdf.PdfReader(file_path)
            for page in pdf.pages:
                content = page.extract_text()
                if content:
                    text_parts.append(content)
        elif ext in ['.docx', '.doc']:
            doc = docx.Document(file_path)
            
            # ==========================================
            # --- NEW: THE TEXT BOX HACK ---
            # Search the raw XML for text boxes and add them to the top
            text_box_found = False
            for node in doc.element.xpath('//w:txbxContent//w:t'):
                if node.text and node.text.strip():
                    text_parts.append(node.text.strip())
                    text_box_found = True
            
            # Add a separator if we actually found text box content
            if text_box_found:
                text_parts.append("--- END OF TEXT BOXES ---")
            # ==========================================

            for section in doc.sections:
                # Group all 3 possible Word headers together
                headers = [section.header, section.first_page_header, section.even_page_header]
                
                for hdr in headers:
                    if hdr: # If the header exists
                        # 1. Check for standard paragraphs
                        for p in hdr.paragraphs:
                            if p.text.strip():
                                text_parts.append(p.text)
                        # 2. Check for hidden tables
                        for table in hdr.tables:
                            for row in table.rows:
                                for cell in row.cells:
                                    if cell.text.strip():
                                        text_parts.append(cell.text)
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

def replace_tag_safely(p, tag, value, unbold=False, force_bold=False):
    if tag.lower() not in p.text.lower():
        return False

    replaced = False
    for run in p.runs:
        if tag.lower() in run.text.lower():
            pattern = re.compile(re.escape(tag), re.IGNORECASE)
            run.text = pattern.sub(str(value), run.text)
            if unbold:
                run.font.bold = False
            if force_bold:
                run.font.bold = True
            replaced = True

    if not replaced:
        full_text = p.text
        pattern = re.compile(re.escape(tag), re.IGNORECASE)
        new_text = pattern.sub(str(value), full_text)

        if p.runs:
            p.runs[0].text = new_text
            if unbold:
                p.runs[0].font.bold = False
            if force_bold:
                p.runs[0].font.bold = True
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
    
    # --- NEW CERTIFICATIONS TABLE LOGIC ---
    # If no certifications, delete the entire table containing the {{Certifications}} tag
    if not mapping.get("Certifications") or not str(mapping.get("Certifications")).strip():
        tables_to_delete = []
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if "{{Certifications}}" in cell.text:
                        if table not in tables_to_delete:
                            tables_to_delete.append(table)
        for table in tables_to_delete:
            table._element.getparent().remove(table._element)
    # --------------------------------------

    # --- NEW EDUCATION TABLE LOGIC ---
    # Check if the AI extracted ANY schools at all
    has_any_education = bool(mapping.get("School1") or mapping.get("School2") or mapping.get("School3"))
    
    for table in doc.tables:
        rows_to_delete = []
        for row in table.rows:
            row_text = "".join(cell.text for cell in row.cells).strip()
            for i in range(1, 4):
                tag = f"{{{{School{i}}}}}"
                if tag.lower() in row_text.lower() and not mapping.get(f"School{i}"):
                    # Only delete the empty row IF they have at least 1 other school listed
                    if has_any_education and row not in rows_to_delete:
                        rows_to_delete.append(row)
        for row in rows_to_delete:
            table._tbl.remove(row._tr)
    # ---------------------------------

    paras_to_remove = []
    kill_keys = ['Company', 'Title', 'Bullets', 'Dates', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'A1', 'A2', 'A3', 'A4', 'A5', 'Summary', 'Skill', 'Years', 'Certifications']

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

                        # --- NEW ENVIRONMENT INJECTION LOGIC ---
                        num_match = re.search(r'\d+', key)
                        if num_match:
                            env_val = mapping.get(f"Environment{num_match.group()}")
                            if env_val and str(env_val).strip():
                                env_p = curr_p.insert_paragraph_before("")
                                curr_p._p.addnext(env_p._p)
                                try:
                                    env_p.style = doc.styles['Normal']
                                except:
                                    pass
                                env_p.paragraph_format.left_indent = Pt(0)
                                env_p.paragraph_format.space_after = Pt(0)
                                
                                clean_env = re.sub(r'^Environment\s*:\s*', '', str(env_val).strip(), flags=re.IGNORECASE)
                                
                                b_run = env_p.add_run("Environment: ")
                                b_run.bold = True
                                b_run.font.name = 'Times New Roman'
                                b_run.font.size = Pt(12)
                                
                                n_run = env_p.add_run(clean_env)
                                n_run.bold = False
                                n_run.font.name = 'Times New Roman'
                                n_run.font.size = Pt(12)
                                
                                curr_p = env_p
                        # ---------------------------------------

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
                        is_question = key in ['Q1', 'Q2', 'Q3', 'Q4', 'Q5']
                        replace_tag_safely(p, tag, str(value).strip(), unbold=is_answer, force_bold=is_question)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        tag = f"{{{{{key}}}}}"
                        if tag.lower() in p.text.lower():
                            is_answer = key in ['A1', 'A2', 'A3', 'A4', 'A5']
                            is_question = key in ['Q1', 'Q2', 'Q3', 'Q4', 'Q5']
                            replace_tag_safely(p, tag, str(value) if value else "", unbold=is_answer, force_bold=is_question)

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
    elif not os.path.exists(TEMPLATE_FILENAME):
        # Checks to make sure the template is actually in your GitHub repo
        st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
    else:
        with st.spinner("Processing document and querying AI..."):
            try:
                # Save uploaded resume temporarily
                resume_path = f"temp_{resume_file.name}"
                
                with open(resume_path, "wb") as f:
                    f.write(resume_file.getbuffer())

                raw_text = extract_text(resume_path)
                
                client = genai.Client(api_key=API_KEY)

                prompt = f"""
                Return a valid JSON object ONLY.

                RULES:
                1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line. (e.g., If the resume says "DUPONT, DELAWARE TCS", you must output exactly "DUPONT").
                3. Education: Extract School and Degree.
                4. Experience: Company, Title, Bullets (LIST), Environment (String, optional), Dates.
                    - For 'Title', clean the string by physically stripping out any employment type modifiers, hyphens, or parentheses at the end of the title (e.g., remove '- Contract', '(Contract)', or '- Consultant').
                    - For 'Environment', you may ONLY extract this if the original resume explicitly uses the word "Environment:" or "Technologies:" at the bottom of the role. If those exact words are not there, you MUST leave it blank "". Do NOT auto-generate or compile an environment list from the bullet points.
                5. Certifications: Extract any certifications listed into a single comma-separated string. If none are found, leave it blank "".

                JSON Structure:
                {{
                    "FullName": "",
                    "Certifications": "",
                    "Education": [{{"School": "", "Degree": ""}}],
                    "Experience": [{{"Company": "", "Title": "", "Bullets": [], "Environment": "", "Dates": ""}}]
                }}

                RESUME: {raw_text}
                """

                # --- SMART RETRY & FALLBACK FOR EXTRACTION ---
                models_to_try = [
                    'gemini-2.5-flash', 
                    'gemini-3-flash-preview', 
                    'gemini-3.1-flash-lite-preview', 
                    'gemini-pro-latest'
                ]
                data = {}
                
                for attempt in range(6):
                    if attempt == 0:
                        current_model = models_to_try[0]
                    elif attempt in [1, 2]:
                        current_model = models_to_try[1]
                    elif attempt in [3, 4]:
                        current_model = models_to_try[2]
                    else:
                        current_model = models_to_try[3]
                    
                    try:
                        response = client.models.generate_content(
                            model=current_model,
                            contents=prompt,
                            config=types.GenerateContentConfig(
                                response_mime_type="application/json"
                            )
                        )
                        data = repair_and_load_json(response.text)
                        break
                    except Exception as e:
                        if "503" in str(e) and attempt < 5:
                            time.sleep(2 ** ((attempt % 2) + 1))
                            continue
                        else:
                            raise e

                name = data.get('FullName', '').title()
                mapping = {
                    "FullName": name, "Location": Current_Location_City_ST, "Remote": Remote_or_Onsite,
                    "Certifications": data.get("Certifications", ""),
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
                        mapping[f"Environment{i}"] = exp[i-1].get('Environment', '')
                        mapping[f"Dates{i}"] = standardize_dates(exp[i-1].get('Dates', ''))
                    else:
                        mapping[f"Company{i}"] = ""
                        mapping[f"Title{i}"] = ""
                        mapping[f"Bullets{i}"] = ""
                        mapping[f"Environment{i}"] = ""
                        mapping[f"Dates{i}"] = ""

                if Job_Description_Notes_etc.strip():
                    try:
                        summary_prompt = f"""
                You are an elite, no-nonsense Senior Technical Recruiter writing an executive submission summary for a Fieldglass portal. 
                The Hiring Manager has 30 seconds to read this. Your goal is to make a punchy, evidence-based business case for why this candidate will succeed.

                ========================
                THE NARRATIVE BLUEPRINT (4 Sentences Max)
                ========================
                Follow this exact structure for the SUMMARY. Every sentence MUST sell the candidate's fit for the role:
                - Sentence 1: The Anchor (Authority). Who are they, what is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description? (Calculate their total overall years. Use their FIRST NAME only. You MUST use the exact job title requested in the JD if the candidate's history support it. Avoid generic fluff; use specific technical keywords from the JD).
                - Sentence 2: The Alignment (The Hook). You MUST explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job? (Do not skip their current job just to find a better keyword match 5 years ago. Frame it as a direct 1:1 match for the manager's current challenge).
                - Sentence 3: The Execution & Impact (The Proof). How did they build it, and why does it matter? (Weave the specific tools/methodologies into an "Execution Statement" that highlights the complexity, scale, or business impact. DO NOT write a comma-separated list of tools. E.g., "By engineering PySpark pipelines across Glue and EMR, she orchestrated TB-scale workflows that eliminated bottlenecks and ensured 99% reliability." DO NOT repeat verbs from Sentence 2).
                - Sentence 4: The Closer (The ROI). Based on their past execution, what specific value will they deliver on Day 1 in THIS new role? (Do NOT start with "Because" or "With." Use a strong, direct structure like: "[First Name]'s success in [X] makes them an immediate asset for [Y]." Note: Refer to the work as "this team," "this project," or "the application". Do NOT mention the physical location/city).

                ========================
                STYLE & TONE RULES (STRICT)
                ========================
                - Write like a human pitching to a colleague. Confident, direct, and factual.
                - TECH MATCHING: Strictly align the tools you highlight with the JD. If the JD asks for AWS, highlight AWS. Do not highlight competing tech (like Azure or GCP) just because it's prominent in the resume, unless it is their only experience.
                - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded') instead of passive verbs like 'Assisted', 'Collaborated', or 'Helped'.
                - SHOW, DON'T TELL. 
                  🔴 Bad: "John's background in AWS makes him a great fit for this role."
                  🟢 Good: "Because John spent the last three years building highly available data lakes in AWS, he can immediately step in to optimize your current infrastructure."
                - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely," "exceptionally well-prepared," "fits well," "aligns with," "enterprise-grade platforms."
                - Do NOT use transition crutches: "Additionally," "Furthermore," "Moreover."
                - Do NOT repeat or restate the job description.

                ========================
                EXAMPLE OF A PERFECT SUMMARY
                ========================
                "Sarah is a senior data engineer with 8+ years of experience architecting cloud-native data migrations within heavily regulated financial environments. Most recently at Capital One, she led the end-to-end migration of a legacy on-prem data warehouse to AWS, directly mirroring the scale and compliance rigor required for this team's current cloud initiative. By engineering automated ETL pipelines with Python, PySpark, and Apache Airflow, she processed 5TB of daily transaction data and reduced reporting latency by 40%. Sarah's success in navigating complex data governance structures makes her an immediate asset for driving this AWS migration."

                ========================
                🔴 Q&A USAGE RULES (CRITICAL)
                ========================
                - The resume is the PRIMARY source of truth for experience.
                - The technical interview Q&A is SECONDARY and should only be used to:
                  • clarify depth
                  • add supporting detail
                  • reinforce experience already shown in the resume

                DO NOT:
                - Base the summary primarily on Q&A responses.
                - Introduce tools, systems, or concepts that are only mentioned in Q&A but not supported by the resume.
                - Overweight theoretical or idealized answers from Q&A.

                If there is any conflict: 👉 prioritize the resume over Q&A.
                The summary should reflect what the candidate has DONE, not just what they can describe.

                ========================
                SKILLS SECTION
                ========================
                - EXACTLY 4 items
                - Prioritize the specific "Must-Have" technologies AND methodologies (e.g., Agile, Data Governance) requested in the JD.
                - Highly relevant + keyword-rich
                - Use only tools explicitly mentioned in resume/Q&A/notes

                Format:
                "Skill Area (Tool1, Tool2, Tool3, Tool4)"

                Years format:
                "X+ years, current" OR "X+ years, 2026"

                ========================
                🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                ========================
                1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                2. DO NOT INFER OR ASSUME. If the JD asks for "Redshift" and the resume only says "AWS", you MUST NOT write "Redshift" under any circumstances. 
                3. DO NOT INFLATE TO MATCH THE JD. If the candidate lacks a requested skill, omit it completely. It is better to have an incomplete match than a fabricated one.
                4. FACT AUDIT: Before outputting the final JSON, you must verify every single tool mentioned in your SUMMARY and SKILLS. If a tool exists in the JD but not in the Resume/Q&A, you must delete it from your output.

                ========================
                YEARS ACCURACY RULES (REALISTIC RECRUITER MODE)
                ========================
                - Use April 2026 as the current date.
                - Calculate years based on total progressive experience in that SKILL DOMAIN, but DO NOT blindly apply their maximum total years to every single bucket.
                - Foundational skills (e.g., Python, SQL, general engineering) should get their maximum total years (e.g., 7+).
                - Advanced/Specialized tools (e.g., SageMaker, Kubernetes, Cloud Architecture) should realistically be calculated at 1-2 years less than their maximum total experience (e.g., 5+ or 6+) unless the resume proves otherwise. 
                - Default to "current" if they are still working in this general technical field. Do NOT use past years (e.g., "2022"). Always bridge past experience to their current tenure.

                ========================
                🧠 FINAL POLISH CHECKLIST (ONE PASS ONLY)
                ========================
                Before outputting the JSON, evaluate your drafted SUMMARY and SKILLS against these 7 checks:
                1. SIMPLIFIED: No run-on sentences or unnecessary filler.
                2. HUMAN TONE: No generic claims like "highly experienced."
                3. TOOL DENSITY: Maximum 3 tools per sentence.
                4. NO REPETITION: Do not repeat verbs or concepts across sentences.
                5. THE PITCH: Ensure a natural, confident recruiter tone.
                6. CURRENT JOB: Did you explicitly name their most recent employer in Sentence 2?

                Output only the final, polished JSON.

                ========================
                OUTPUT FORMAT (STRICT)
                ========================
                Return ONLY valid JSON:

                {{
                  "SUMMARY": "4 sentence summary following the blueprint",
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
                        # We give the server a 2-second breather between the Extraction and Summary calls
                        time.sleep(2) 
                        
                        # --- SMART RETRY & FALLBACK FOR SUMMARY ---
                        summary_data = {}
                        
                        for attempt in range(6):
                            if attempt == 0:
                                current_model = models_to_try[0]
                            elif attempt in [1, 2]:
                                current_model = models_to_try[1]
                            elif attempt in [3, 4]:
                                current_model = models_to_try[2]
                            else:
                                current_model = models_to_try[3]
                            
                            try:
                                summary_response = client.models.generate_content(
                                    model=current_model,
                                    contents=summary_prompt,
                                    config=types.GenerateContentConfig(
                                        response_mime_type="application/json"
                                    )
                                )
                                summary_data = repair_and_load_json(summary_response.text)
                                break
                            except Exception as api_e:
                                if "503" in str(api_e) and attempt < 5:
                                    time.sleep(2 ** ((attempt % 2) + 1))
                                    continue
                                else:
                                    raise api_e

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
                
                # INJECT DATA INTO THE PRELOADED TEMPLATE
                process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                
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

                # Clean up the temporary resume file
                try:
                    os.remove(resume_path)
                except:
                    pass

            except Exception as e:
                st.error(f"❌ Process Failed: {str(e)}")
