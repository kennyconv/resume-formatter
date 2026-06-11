import streamlit as st
import pandas as pd
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
import subprocess
from pdf2docx import Converter
from docx.shared import Pt

# ====================================================================
# --- STREAMLIT UI & PASSWORD LOGIC ---
# ====================================================================

st.set_page_config(page_title="Resume Formatter", layout="wide")

# --- CUSTOM CSS FOR INPUT BOXES ---
st.markdown(
    """
    <style>
    /* Add a visible outline and subtle background shading to all inputs, text areas, and dropdowns */
    div[data-baseweb="input"], 
    div[data-baseweb="base-input"], 
    div[data-baseweb="select"] {
        border: 1px solid #7c7c8c !important; /* Gray border */
        background-color: rgba(128, 128, 128, 0.1) !important; /* Slight transparent shading */
        border-radius: 6px !important; /* Smooth rounded corners */
    }
    </style>
    """, 
    unsafe_allow_html=True
)
# ----------------------------------

def check_password():
    """Returns `True` if the user entered the correct password."""
    def password_entered():
        if st.session_state["password"] == st.secrets["APP_PASSWORD"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    st.title("🔒 Access Restricted")
    st.text_input(
        "Please enter the team password to use this tool:", 
        type="password", 
        on_change=password_entered, 
        key="password"
    )
    
    if "password_correct" in st.session_state:
        st.error("❌ Incorrect password.")
        
    return False

# 🛑 STOP THE APP HERE IF THEY DON'T HAVE THE PASSWORD
if not check_password():
    st.stop()


# ====================================================================
# --- UNIVERSAL HELPER FUNCTIONS (Shared by all clients) ---
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

def format_peraton_dates(date_str):
    """Enforces MM/YYYY to MM/YYYY or Present for Peraton"""
    if not date_str:
        return ""
    # Swap common hyphens/dashes to the required " to "
    date_str = re.sub(r'\s*[-–—]\s*', ' to ', date_str)
    # Ensure "Present" or "Current" is properly formatted
    date_str = re.sub(r'(?i)present', 'Present', date_str)
    date_str = re.sub(r'(?i)current', 'Present', date_str)
    return date_str.strip()

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
            text_box_found = False
            for node in doc.element.xpath('//w:txbxContent//w:t'):
                if node.text and node.text.strip():
                    text_parts.append(node.text.strip())
                    text_box_found = True
            if text_box_found:
                text_parts.append("--- END OF TEXT BOXES ---")
            for section in doc.sections:
                headers = [section.header, section.first_page_header, section.even_page_header]
                for hdr in headers:
                    if hdr: 
                        for p in hdr.paragraphs:
                            if p.text.strip():
                                text_parts.append(p.text)
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

def replace_tag_safely(p, tag, value, unbold=False, force_bold=False, font_name=None):
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
            if font_name:
                run.font.name = font_name
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
            if font_name:
                p.runs[0].font.name = font_name
            for i in range(1, len(p.runs)):
                p.runs[i].text = ""
    return True

def insert_bullet_line_after(paragraph, text, font_name=None):
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
        if font_name:
            new_run.font.name = font_name
        else:
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
    
    # Check if this is the Peraton tool
    is_peraton = "CERTIFICATION1" in mapping
    target_font = 'Calibri' if is_peraton else None
    env_font = 'Calibri' if is_peraton else 'Times New Roman'

    # --- TWEAK 1: Remove "Certifications" Table & Header (All Clients) ---
    has_certifications = True
    if is_peraton:
        if not mapping.get("CERTIFICATION1"):
            has_certifications = False
    else:
        if not mapping.get("Certifications") or not str(mapping.get("Certifications")).strip():
            has_certifications = False

    if not has_certifications:
        tables_to_delete = []
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Made case-insensitive to ensure it catches the tag even if typed differently in template
                    if "{{certifications}}" in cell.text.lower() or "{{certification1}}" in cell.text.lower():
                        if table not in tables_to_delete:
                            tables_to_delete.append(table)
        for table in tables_to_delete:
            table._element.getparent().remove(table._element)
            
        paras = list(doc.paragraphs)
        for i, p in enumerate(paras):
            # Made case-insensitive to ensure the header text is caught
            if p.text.strip().lower() == "certifications":
                delete_paragraph(p)
                # Prevent a double-space gap by deleting the empty line above it if it exists
                if i > 0 and not paras[i-1].text.strip():
                    try:
                        delete_paragraph(paras[i-1])
                    except:
                        pass

    # --- TWEAK 2: Remove Q&A Section for ALL clients if empty (Including Fannie Mae) ---
    has_qa = any(mapping.get(k) and str(mapping.get(k)).strip() for k in ["Q1", "A1", "Q2", "A2", "Q3", "A3", "Q4", "A4", "Q5", "A5"])

    if not has_qa:
        paras = list(doc.paragraphs)
        for i, p in enumerate(paras):
            if "SUPPLIER TECHNICAL INTERVIEW RESULTS" in p.text.upper():
                delete_paragraph(p)
                # Prevent a double-space gap by deleting the empty line above it if it exists
                if i > 0 and not paras[i-1].text.strip():
                    try:
                        delete_paragraph(paras[i-1])
                    except:
                        pass
        
        tables_to_delete = []
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    # Made case-insensitive for safety
                    if "{{q1}}" in cell.text.lower():
                        if table not in tables_to_delete:
                            tables_to_delete.append(table)
        for table in tables_to_delete:
            table._element.getparent().remove(table._element)
                
    # --- PERATON TWEAK 2: Space before Relevant Professional Experience ---
    if is_peraton:
        for p in list(doc.paragraphs):
            if "Relevant Professional Experience" in p.text:
                p.insert_paragraph_before("")
                break # Only do it once
            
    has_any_education = bool(mapping.get("School1") or mapping.get("School2") or mapping.get("School3"))
    for table in doc.tables:
        rows_to_delete = []
        for row in table.rows:
            row_text = "".join(cell.text for cell in row.cells).strip()
            for i in range(1, 4):
                tag = f"{{{{School{i}}}}}"
                if tag.lower() in row_text.lower() and not mapping.get(f"School{i}"):
                    if has_any_education and row not in rows_to_delete:
                        rows_to_delete.append(row)
        for row in rows_to_delete:
            table._tbl.remove(row._tr)
            
    paras_to_remove = []
    # Updated kill_keys with 'Certification', 'School', and 'Responsible'
    kill_keys = ['Company', 'Title', 'Bullets', 'Dates', 'Q1', 'Q2', 'Q3', 'Q4', 'Q5', 'A1', 'A2', 'A3', 'A4', 'A5', 'Summary', 'Skill', 'Years', 'Certification', 'School', 'Responsible', 'BillRate', 'WorkStatus']
    for p in list(doc.paragraphs):
        for key, value in mapping.items():
            tag = f"{{{{{key}}}}}"
            if tag.lower() in p.text.lower():
                if not value or not str(value).strip():
                    if any(k.lower() in key.lower() for k in kill_keys):
                        if p not in paras_to_remove:
                            paras_to_remove.append(p)
                    else:
                        replace_tag_safely(p, tag, "", font_name=target_font)
                else:
                    if "Bullets" in key:
                        for run in p.runs:
                            if '\n' in run.text:
                                run.text = run.text.replace('\n', '')
                        lines = str(value).split('\n')
                        replace_tag_safely(p, tag, lines[0], font_name=target_font)
                        curr_p = p
                        for line in lines[1:]:
                            curr_p = insert_bullet_line_after(curr_p, line, font_name=target_font)
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
                                b_run.font.name = env_font
                                b_run.font.size = Pt(12)
                                n_run = env_p.add_run(clean_env)
                                n_run.bold = False
                                n_run.font.name = env_font
                                n_run.font.size = Pt(12)
                                curr_p = env_p
                        
                        # --- PERATON TWEAK 3: Disable auto-spacer for Peraton ---
                        if not is_peraton:
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
                        replace_tag_safely(p, tag, str(value).strip(), unbold=is_answer, force_bold=is_question, font_name=target_font)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        tag = f"{{{{{key}}}}}"
                        if tag.lower() in p.text.lower():
                            is_answer = key in ['A1', 'A2', 'A3', 'A4', 'A5']
                            is_question = key in ['Q1', 'Q2', 'Q3', 'Q4', 'Q5']
                            replace_tag_safely(p, tag, str(value) if value else "", unbold=is_answer, force_bold=is_question, font_name=target_font)
    for p in paras_to_remove:
        try:
            delete_paragraph(p)
        except:
            pass
    doc.save(out_path)
    return out_path


# ====================================================================
# --- 🔴 CLIENT APP 1: FANNIE MAE 🔴 ---
# ====================================================================
def fannie_mae_app():
    TEMPLATE_FILENAME = "Template.docx"
    st.title("Fannie Mae Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="fm_loc")
        Remote_or_Onsite = st.selectbox("Remote or Onsite", ["Remote", "Onsite"], index=1, key="fm_rem")
    with col2:
        Former_FM = st.selectbox("Former FM", ["Y - Per CRC, this candidate is eligible for rehire", "N"], index=1, key="fm_form")
        LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link", key="fm_link")

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="fm_q1")
        Question_2 = st.text_input("Question 2", key="fm_q2")
        Question_3 = st.text_input("Question 3", key="fm_q3")
        Question_4 = st.text_input("Question 4", key="fm_q4")
        Question_5 = st.text_input("Question 5", key="fm_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="fm_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="fm_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="fm_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="fm_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="fm_a5")

    st.header("📝 Job Description & Notes")

    # --- NEW GOOGLE SHEET LOGIC ---
    # IMPORTANT: Replace this placeholder with your actual Google Sheet link. 
    # The sheet's share settings MUST be set to "Anyone with the link can view".
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1swKaDhGdlmO0ILv7tngkRCbsK-7jF89cdCpf5bypzwM/edit?usp=sharing"

    @st.cache_data(ttl=600)
    def load_reqs(url):
        try:
            # Converts standard Google Sheet URL to a direct CSV download link
            if "/edit" in url:
                csv_url = url.split("/edit")[0] + "/export?format=csv"
            else:
                csv_url = url
            df = pd.read_csv(csv_url)
            return df.fillna("")
        except Exception:
            return pd.DataFrame()

    df_reqs = load_reqs(SHEET_URL)
    options = ["None"]
    req_mapping = {}

    if not df_reqs.empty:
        for idx, row in df_reqs.iterrows():
            req_id = str(row.get("ID", "")).replace(".0", "")
            vms_id = str(row.get("VMS ID", "")).replace(".0", "")
            title = str(row.get("Job Title", ""))
            manager = str(row.get("Manager", "")) # NEW: Extract Manager
            desc = str(row.get("Description", ""))
            notes = str(row.get("FG Notes", ""))
            transcript = str(row.get("Spotlight Transcript", "")) # NEW: Extract Transcript

            # Build the dropdown string safely, ignoring empty values
            option_parts = [req_id, vms_id, title, manager]
            option_str = " - ".join([p.strip() for p in option_parts if p.strip()])
            
            if option_str and option_str != "-":
                options.append(option_str)
                # Added transcript to the mapping dictionary
                req_mapping[option_str] = {"desc": desc, "notes": notes, "transcript": transcript}

    if "fm_jd" not in st.session_state:
        st.session_state.fm_jd = ""
    if "fm_notes" not in st.session_state:
        st.session_state.fm_notes = ""
    if "fm_trans" not in st.session_state:
        st.session_state.fm_trans = ""

    def update_jd_text():
        if "fm_req_selector" not in st.session_state:
            return

        selected = st.session_state.fm_req_selector
        if selected != "None" and selected in req_mapping:
            st.session_state.fm_jd = req_mapping[selected].get("desc", "").strip()
            st.session_state.fm_notes = req_mapping[selected].get("notes", "").strip()
            st.session_state.fm_trans = req_mapping[selected].get("transcript", "").strip()
        else:
            # Clear the text boxes if "None" is selected
            st.session_state.fm_jd = ""
            st.session_state.fm_notes = ""
            st.session_state.fm_trans = ""
            
    selected_req = st.selectbox(
        "🔍 Select a Fannie Mae Requisition (Type to search):",
        options=options,
        key="fm_req_selector",
        on_change=update_jd_text
    )

    Job_Description = st.text_area(
        "1. Official Job Description (Paste generic JD here):", 
        height=150, 
        key="fm_jd"
    )

    Chat_Notes = st.text_area(
        "2. Fieldglass Chat Notes (Optional - Overrides JD):", 
        height=100, 
        key="fm_notes"
    )

    Call_Transcript = st.text_area(
        "3. Spotlight Call Transcript (Optional - Highest Priority):", 
        height=100, 
        key="fm_trans"
    )

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="fm_res")

    if st.button("🚀 Generate Fannie Mae Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
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

                    if Job_Description.strip():
                        try:
                            summary_prompt = f"""
                            You are an elite, no-nonsense Senior Technical Recruiter writing an executive submission summary for a Fieldglass portal. 
                            The Hiring Manager has 30 seconds to read this. Your goal is to make a punchy, evidence-based business case for why this candidate will succeed, optimized for both human reading and ATS exact-match parsing.

                            ========================
                            🚨 HIERARCHY OF TRUTH: JD vs. MANAGER NOTES (CRITICAL)
                            ========================
                            You are receiving up to three sources of information for this role:
                            1. Spotlight Call Transcript (HIGHEST PRIORITY)
                            2. Fieldglass Chat Notes (HIGH PRIORITY)
                            3. Official Job Description (BASELINE)

                            - OVERRIDE RULE: The Call Transcript and Chat Notes represent the Hiring Manager's actual, immediate needs. They are the ABSOLUTE SOURCE OF TRUTH. If they contradict or update the official Job Description (e.g., lowering required years of experience, changing the core tech stack, or negating responsibilities), you MUST strictly follow the Transcript and Notes.
                            - ATS COMPROMISE: To satisfy the Fieldglass parsing algorithm, you must still weave in high-level, generalized keywords from the baseline Job Description where applicable, but NEVER highlight technical skills or responsibilities that the manager explicitly negated in their notes or calls.

                            ========================
                            THE NARRATIVE BLUEPRINT (4 Sentences Max)
                            ========================
                            Follow this exact structure for the SUMMARY. Every sentence MUST sell the candidate's fit for the role:
                            - Sentence 1: The Anchor (Authority). Who are they, what is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description? (Calculate their total overall years strictly rounding DOWN to the nearest whole year formatted as 'X+ years'. Use their FIRST NAME only. You MUST use the exact job title requested in the JD if the candidate's history supports it. Avoid generic fluff).
                            - Sentence 2: The Alignment (The Hook). You MUST explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job? (Frame it as a direct 1:1 match for the manager's current challenge).
                            - Sentence 3: The Execution & Impact (The Proof). How did they build it, and why does it matter? (Weave the specific tools/methodologies into an "Execution Statement" that highlights the complexity, scale, or business impact. DO NOT write a comma-separated list of tools. DO NOT repeat verbs from Sentence 2).
                            - Sentence 4: The Closer (The ROI). Based on their past execution, what specific value will they deliver on Day 1 in THIS new role? (Use a strong, direct structure like: "[First Name]'s success in [X] makes them an immediate asset for driving this team's [specific technical goal/initiative]." Do NOT mention the physical location/city).

                            CRITICAL ATS HACK: Across these 4 sentences, you MUST seamlessly embed 2-3 exact phrases from the Job Description (including soft skills like "changing priorities" or "system analysis") to maximize the Fieldglass match score. Do not force them if they ruin the sentence flow, but prioritize exact phrase matching where supported by the resume.

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
                            "Sarah is a Senior Data Engineer with 7+ years of experience architecting cloud-native data migrations within heavily regulated financial environments. Most recently at Capital One, she led the end-to-end migration of a legacy on-prem data warehouse to AWS, directly mirroring the scale and compliance rigor required for this team's current cloud initiative. By engineering automated ETL pipelines with Python, PySpark, and Apache Airflow, she processed 5TB of daily transaction data and reduced reporting latency by 40%. Sarah's success in navigating complex data governance structures makes her an immediate asset for driving this team's AWS migration goals."

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
                            - Use June 2026 as the current date.
                            - STRICT MATH (DATE-DRIVEN ONLY): Calculate years strictly based on the earliest chronological date provided in the 'Professional Experience' or 'Work History' section. You MUST completely IGNORE any self-reported years of experience in the candidate's summary blurb (e.g., if their summary claims "12+ years" but their listed jobs only go back to 2019, you must calculate from 2019). Round DOWN to the nearest whole year and use the exact format "X+ years". Do not use phrases like "nearly X years". (e.g., If the job history calculates to 7 years and 10 months, output "7+ years". NEVER round up to 8+). 
                            - Foundational skills (e.g., Python, SQL, general engineering) should get their maximum calculated years.
                            - Advanced/Specialized tools (e.g., SageMaker, Kubernetes, Cloud Architecture) should realistically be calculated at 1-2 years less than their maximum total experience unless the resume explicitly proves Day 1 usage. 
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
                            7. ACCURATE MATH: Did you strictly round down their years of experience and use the 'X+ years' format to prevent ATS parser flags?

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
                            Official Job Description:
                            {Job_Description}

                            Fieldglass Chat Notes:
                            {Chat_Notes}

                            Spotlight Call Transcript:
                            {Call_Transcript}

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
                            time.sleep(2) 
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

                    out_file = f"Submission_Fannie_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 🔵 CLIENT APP 2: PERATON 🔵 ---
# ====================================================================
def peraton_app():
    TEMPLATE_FILENAME = "Peraton_Template.docx"
    st.title("Peraton Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="per_q1")
        Question_2 = st.text_input("Question 2", key="per_q2")
        Question_3 = st.text_input("Question 3", key="per_q3")
        Question_4 = st.text_input("Question 4", key="per_q4")
        Question_5 = st.text_input("Question 5", key="per_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="per_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="per_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="per_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="per_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="per_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="per_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="per_res")

    if st.button("🚀 Generate Peraton Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
                    3. Education: Extract School and Degree.
                    4. Experience: Company, Title, Responsible, Bullets (LIST), Environment (String, optional), Dates.
                        - For 'Title', clean the string by physically stripping out any employment type modifiers, hyphens, or parentheses at the end of the title.
                        - For 'Responsible', write exactly 1 sentence summarizing what the candidate was responsible for at this specific job, based on their bullets. Do NOT use the candidate's name or pronouns (he/she). Start the sentence with the exact words "Responsible for ".
                        - For 'Environment', you may ONLY extract this if the original resume explicitly uses the word "Environment:" or "Technologies:" at the bottom of the role.
                        - For 'Dates', you MUST format the dates strictly as "MM/YYYY to MM/YYYY" (e.g., "10/2019 to 07/2024") or "MM/YYYY to Present" if it is their current position. Convert all months to their 2-digit number.
                    5. Certifications: Extract active certifications into an ARRAY of strings. Do not include classes or courses taken. Keep it strictly to the certification name.

                    JSON Structure:
                    {{
                        "FullName": "",
                        "Certifications": [],
                        "Education": [{{"School": "", "Degree": ""}}],
                        "Experience": [{{"Company": "", "Title": "", "Responsible": "", "Bullets": [], "Environment": "", "Dates": ""}}]
                    }}

                    RESUME: {raw_text}
                    """

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

                    name = data.get('FullName', '').title().strip()
                    name_parts = name.split(' ', 1)
                    first_name = name_parts[0] if len(name_parts) > 0 else ""
                    last_name = name_parts[1] if len(name_parts) > 1 else ""

                    mapping = {
                        "Firstname": first_name, "Lastname": last_name, 
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY": "", 
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", "SUMMARY5": "",
                        "CERTIFICATION1": "", "CERTIFICATION2": "", "CERTIFICATION3": "", "CERTIFICATION4": "",
                        "SKILL1": "", "SKILL2": "", "SKILL3": "", "SKILL4": ""
                    }

                    # Handle Certifications Array
                    certs = data.get("Certifications", [])
                    if isinstance(certs, str): 
                        certs = [c.strip() for c in certs.split(',') if c.strip()]
                    for i in range(1, 5):
                        mapping[f"CERTIFICATION{i}"] = certs[i-1] if i <= len(certs) else ""

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
                            mapping[f"RESPONSIBLE{i}"] = exp[i-1].get('Responsible', '')
                            mapping[f"Bullets{i}"] = clean_bullets(exp[i-1].get('Bullets', []))
                            mapping[f"Environment{i}"] = exp[i-1].get('Environment', '')
                            
                            # USE NEW PERATON DATE FORMATTER HERE
                            mapping[f"Dates{i}"] = format_peraton_dates(exp[i-1].get('Dates', ''))
                        else:
                            mapping[f"Company{i}"] = ""
                            mapping[f"Title{i}"] = ""
                            mapping[f"RESPONSIBLE{i}"] = ""
                            mapping[f"Bullets{i}"] = ""
                            mapping[f"Environment{i}"] = ""
                            mapping[f"Dates{i}"] = ""

                    if Job_Description_Notes_etc.strip():
                        try:
                            summary_prompt = f"""
                            You are an elite, no-nonsense Senior Technical Recruiter writing an executive submission summary for a Peraton Fieldglass portal. 
                            Your goal is to make an evidence-based business case for why this candidate will succeed, strictly following Peraton's specific submission guidelines.

                            ========================
                            OVERALL SUMMARY RULES (STRICT)
                            ========================
                            You must output the "SUMMARY" field as a single paragraph of 150 words or less.
                            - FIRST PERSON: You MUST write this summary entirely in the 1st person perspective (e.g., "I am a Data Analyst with...", "My background includes..."). Do NOT use the candidate's name or 3rd person pronouns (he/she/they).
                            - It must highlight exactly why they are the best candidate for this specific position.
                            - You MUST explicitly state their total number of years of professional experience.
                            - Be sure the paragraph seamlessly highlights the specific skills required for the position.

                            ========================
                            SUMMARY OF QUALIFICATIONS RULES (STRICT)
                            ========================
                            You must output EXACTLY 5 distinct strings mapped to "SUMMARY1" through "SUMMARY5".
                            - TAILOR TO THE JD: Only include qualifications specifically tailored to the position being applied for. Omit irrelevant history.
                            - ACTION VERBS: Start EVERY point with a strong action verb (e.g., Engineered, Managed, Architected).
                            - QUANTIFY EVERYTHING: You MUST substantiate/quantify each point with a number derived from the resume/Q&A (e.g., number of systems, applications, users, bandwidth, endpoints, data scale, or team size). 
                            - PRIORITIZATION: Place the most relevant points first.
                            - You MUST provide all 5 bullets. Do not leave any blank. 
                            - DO NOT include the physical bullet character (•) in the text. The Word document template will apply the bullet automatically.

                            ========================
                            TECHNICAL SKILLS RULES (STRICT)
                            ========================
                            - ONLY include technical skills and tools directly relevant to the position being applied for.
                            - DO NOT include basic/universal skills (e.g., MS Word, Excel, PowerPoint, Outlook).
                            - Divide the relevant skills into exactly 4 logical buckets mapped to SKILL1 through SKILL4. 

                            Format:
                            "Skill Area (Tool1, Tool2, Tool3, Tool4)"

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth.
                            - The technical interview Q&A is SECONDARY and should only be used to clarify depth or extract quantifiable metrics (numbers/scale) to support the Summary bullets.
                            - If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME.
                            3. FACT AUDIT: Before outputting the final JSON, verify every tool and NUMBER you used. If a metric/number doesn't exist in the resume/Q&A, do not invent one. 

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY": "150 word paragraph highlighting fit and total years of experience...",
                              "SUMMARY1": "Engineered [X] resulting in [Y]...",
                              "SUMMARY2": "Architected [X] for [Y] users...",
                              "SUMMARY3": "Managed [X] endpoints...",
                              "SUMMARY4": "Developed [X] applications...",
                              "SUMMARY5": "Optimized [X] systems...",
                              "SKILL1": "Skill Area (tools)",
                              "SKILL2": "Skill Area (tools)",
                              "SKILL3": "Skill Area (tools)",
                              "SKILL4": "Skill Area (tools)"
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
                            time.sleep(2) 
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
                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
                            mapping["SUMMARY5"] = summary_data.get("SUMMARY5", "")
                            
                            mapping["SKILL1"] = summary_data.get("SKILL1", "")
                            mapping["SKILL2"] = summary_data.get("SKILL2", "")
                            mapping["SKILL3"] = summary_data.get("SKILL3", "")
                            mapping["SKILL4"] = summary_data.get("SKILL4", "")

                        except Exception as e:
                            st.warning(f"⚠️ Warning: Summary generation failed. Proceeding without it. ({e})")

                    out_file = f"Submission_Peraton_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 🟢 CLIENT APP 3: CAPITAL ONE 🟢 ---
# ====================================================================
def capital_one_app():
    TEMPLATE_FILENAME = "CapitalOne_Template.docx"
    st.title("Capital One Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="co_loc")
        
        # --- NEW REMOTE/HYBRID LOGIC ---
        Remote_or_Onsite = st.selectbox("Remote or Onsite/Hybrid", ["Remote", "Onsite/Hybrid"], index=0, key="co_rem")
        co_location_name = ""
        if Remote_or_Onsite == "Onsite/Hybrid":
            co_location_name = st.text_input("Location Name (e.g., HQ2)", key="co_loc_name")
            
        # --- NEW START DATE LOGIC ---
        co_start = st.text_input("Availability to Start", key="co_start")

    with col2:
        # --- NEW FORMER EMPLOYEE LOGIC ---
        Former_CO = st.selectbox("Former Capital One", ["Y", "N"], index=1, key="co_form")
        co_mgr = ""
        co_dates = ""
        if Former_CO == "Y":
            co_mgr = st.text_input("Manager Name", key="co_mgr")
            co_dates = st.text_input("Dates Worked", key="co_dates")
            
        LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link", key="co_link")

    # Mapping logic variables based on above inputs
    if Remote_or_Onsite == "Onsite/Hybrid" and co_location_name.strip():
        remote_val = f"Onsite/Hybrid - {co_location_name.strip()}"
    else:
        remote_val = Remote_or_Onsite
        
    if Former_CO == "Y":
        former_co_val = f"Y - {co_mgr.strip()}, {co_dates.strip()}"
    else:
        former_co_val = "N"

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="co_q1")
        Question_2 = st.text_input("Question 2", key="co_q2")
        Question_3 = st.text_input("Question 3", key="co_q3")
        Question_4 = st.text_input("Question 4", key="co_q4")
        Question_5 = st.text_input("Question 5", key="co_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="co_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="co_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="co_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="co_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="co_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="co_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="co_res")

    if st.button("🚀 Generate Capital One Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
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
                        "FullName": name, "Location": Current_Location_City_ST, "Remote": remote_val,
                        "Start": co_start,
                        "Certifications": data.get("Certifications", ""),
                        "FormerCapitalOne": former_co_val, "Links": LinkedIn_GitHub_Portfolio_Link,
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", 
                        "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
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
                            THE NARRATIVE BLUEPRINT (4 Bullet Points)
                            ========================
                            Generate exactly 4 distinct bullet points mapped to SUMMARY1 through SUMMARY4. 
                            CRITICAL RULE: DO NOT use the candidate's name or 3rd person pronouns (he/she/they) in ANY of the bullet points. Start each bullet with a strong action verb or direct statement.
                            
                            - SUMMARY1 (The Anchor): What is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description?
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job?
                            - SUMMARY3 (The Execution & Impact): How did they build it, and why does it matter? Weave specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (The Closer): Based on their past execution, what specific value will they deliver on Day 1 in THIS new role?

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth for experience.
                            - The technical interview Q&A is SECONDARY.
                            If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            SKILLS SECTION
                            ========================
                            - EXACTLY 4 items
                            - Prioritize the specific "Must-Have" technologies.
                            - Use only tools explicitly mentioned in resume/Q&A/notes
                            Format: "Skill Area (Tool1, Tool2, Tool3, Tool4)"
                            Years format: "X+ years, current" OR "X+ years, 2026"

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. 
                            4. FACT AUDIT: Before outputting the final JSON, verify every single tool.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY1": "Bullet 1 following the blueprint",
                              "SUMMARY2": "Bullet 2 following the blueprint",
                              "SUMMARY3": "Bullet 3 following the blueprint",
                              "SUMMARY4": "Bullet 4 following the blueprint",
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
                            time.sleep(2) 
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

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
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

                    # 1. Generate the .docx file first
                    docx_file = f"Submission_CapitalOne_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, docx_file)
                    
                    # 2. Convert the .docx to .pdf using LibreOffice (Linux compatible)
                    pdf_file = f"Submission_CapitalOne_{name.replace(' ', '_')}.pdf"
                    
                    try:
                        subprocess.run([
                            'libreoffice', '--headless', '--convert-to', 'pdf', 
                            docx_file, '--outdir', os.getcwd()
                        ], check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
                    except Exception as e:
                        st.error(f"⚠️ PDF Conversion failed: {e}. Ensure 'libreoffice' is in packages.txt.")
                    
                    # 3. Provide the PDF for download (fallback to docx if PDF fails to generate)
                    if os.path.exists(pdf_file):
                        with open(pdf_file, "rb") as file:
                            btn = st.download_button(
                                label="⬇️ Download Generated PDF",
                                data=file,
                                file_name=pdf_file,
                                mime="application/pdf",
                                type="primary"
                            )
                        st.success(f"✅ Success! PDF Document is ready for download.")
                    else:
                        with open(docx_file, "rb") as file:
                            btn = st.download_button(
                                label="⬇️ Download Generated Word Doc (PDF Failed)",
                                data=file,
                                file_name=docx_file,
                                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                type="primary"
                            )
                        st.warning("⚠️ Success, but PDF conversion failed. Provided Word Document instead.")

                    # 4. Cleanup the temporary files
                    try:
                        os.remove(resume_path)
                        if os.path.exists(docx_file):
                            os.remove(docx_file)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 🟣 CLIENT APP 4: ADUSA 🟣 ---
# ====================================================================
def adusa_app():
    TEMPLATE_FILENAME = "ADUSA_Template.docx"
    st.title("ADUSA Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="ad_loc")
        
        Remote_or_Onsite = st.selectbox("Remote or Onsite/Hybrid", ["Remote", "Onsite/Hybrid"], index=0, key="ad_rem")
        ad_location_name = ""
        if Remote_or_Onsite == "Onsite/Hybrid":
            ad_location_name = st.text_input("Location Name (e.g., HQ2)", key="ad_loc_name")
            
        ad_start = st.text_input("Availability to Start", key="ad_start")

    with col2:
        Former_ADUSA = st.selectbox("Former ADUSA", ["Y", "N"], index=1, key="ad_form")
        ad_mgr = ""
        ad_dates = ""
        if Former_ADUSA == "Y":
            ad_mgr = st.text_input("Manager Name", key="ad_mgr")
            ad_dates = st.text_input("Dates Worked", key="ad_dates")
        # LinkedIn Removed for ADUSA

    if Remote_or_Onsite == "Onsite/Hybrid" and ad_location_name.strip():
        remote_val = f"Onsite/Hybrid - {ad_location_name.strip()}"
    else:
        remote_val = Remote_or_Onsite
        
    if Former_ADUSA == "Y":
        former_adusa_val = f"Y - {ad_mgr.strip()}, {ad_dates.strip()}"
    else:
        former_adusa_val = "N"

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="ad_q1")
        Question_2 = st.text_input("Question 2", key="ad_q2")
        Question_3 = st.text_input("Question 3", key="ad_q3")
        Question_4 = st.text_input("Question 4", key="ad_q4")
        Question_5 = st.text_input("Question 5", key="ad_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="ad_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="ad_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="ad_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="ad_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="ad_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="ad_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="ad_res")

    if st.button("🚀 Generate ADUSA Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
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
                        "FullName": name, "Location": Current_Location_City_ST, "Remote": remote_val,
                        "Start": ad_start,
                        "Certifications": data.get("Certifications", ""),
                        "FormerADUSA": former_adusa_val, 
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", 
                        "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
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
                            THE NARRATIVE BLUEPRINT (4 Bullet Points)
                            ========================
                            Generate exactly 4 distinct bullet points mapped to SUMMARY1 through SUMMARY4. 
                            CRITICAL RULE: DO NOT use the candidate's name or 3rd person pronouns (he/she/they) in ANY of the bullet points. Start each bullet with a strong action verb or direct statement.
                            
                            - SUMMARY1 (The Anchor): What is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description?
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job?
                            - SUMMARY3 (The Execution & Impact): How did they build it, and why does it matter? Weave specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (The Closer): Based on their past execution, what specific value will they deliver on Day 1 in THIS new role?

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth for experience.
                            - The technical interview Q&A is SECONDARY.
                            If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            SKILLS SECTION
                            ========================
                            - EXACTLY 4 items
                            - Prioritize the specific "Must-Have" technologies.
                            - Use only tools explicitly mentioned in resume/Q&A/notes
                            Format: "Skill Area (Tool1, Tool2, Tool3, Tool4)"
                            Years format: "X+ years, current" OR "X+ years, 2026"

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. 
                            4. FACT AUDIT: Before outputting the final JSON, verify every single tool.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY1": "Bullet 1 following the blueprint",
                              "SUMMARY2": "Bullet 2 following the blueprint",
                              "SUMMARY3": "Bullet 3 following the blueprint",
                              "SUMMARY4": "Bullet 4 following the blueprint",
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
                            time.sleep(2) 
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

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
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

                    out_file = f"Submission_ADUSA_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 🟡 CLIENT APP 5: CBRE 🟡 ---
# ====================================================================
def cbre_app():
    TEMPLATE_FILENAME = "CBRE_Template.docx"
    st.title("CBRE Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="cb_loc")
        
        Remote_or_Onsite = st.selectbox("Remote or Onsite/Hybrid", ["Remote", "Onsite/Hybrid"], index=0, key="cb_rem")
        cb_location_name = ""
        if Remote_or_Onsite == "Onsite/Hybrid":
            cb_location_name = st.text_input("Location Name (e.g., HQ2)", key="cb_loc_name")
            
        cb_start = st.text_input("Availability to Start", key="cb_start")

    with col2:
        Former_CBRE = st.selectbox("Former CBRE", ["Y", "N"], index=1, key="cb_form")
        cb_mgr = ""
        cb_dates = ""
        if Former_CBRE == "Y":
            cb_mgr = st.text_input("Manager Name", key="cb_mgr")
            cb_dates = st.text_input("Dates Worked", key="cb_dates")
            
        LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link", key="cb_link")

    if Remote_or_Onsite == "Onsite/Hybrid" and cb_location_name.strip():
        remote_val = f"Onsite/Hybrid - {cb_location_name.strip()}"
    else:
        remote_val = Remote_or_Onsite
        
    if Former_CBRE == "Y":
        former_cbre_val = f"Y - {cb_mgr.strip()}, {cb_dates.strip()}"
    else:
        former_cbre_val = "N"

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="cb_q1")
        Question_2 = st.text_input("Question 2", key="cb_q2")
        Question_3 = st.text_input("Question 3", key="cb_q3")
        Question_4 = st.text_input("Question 4", key="cb_q4")
        Question_5 = st.text_input("Question 5", key="cb_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="cb_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="cb_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="cb_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="cb_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="cb_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="cb_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="cb_res")

    if st.button("🚀 Generate CBRE Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
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
                        "FullName": name, "Location": Current_Location_City_ST, "Remote": remote_val,
                        "Start": cb_start,
                        "Certifications": data.get("Certifications", ""),
                        "FormerCBRE": former_cbre_val, "Links": LinkedIn_GitHub_Portfolio_Link,
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", 
                        "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
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
                            THE NARRATIVE BLUEPRINT (4 Bullet Points)
                            ========================
                            Generate exactly 4 distinct bullet points mapped to SUMMARY1 through SUMMARY4. 
                            CRITICAL RULE: DO NOT use the candidate's name or 3rd person pronouns (he/she/they) in ANY of the bullet points. Start each bullet with a strong action verb or direct statement.
                            
                            - SUMMARY1 (The Anchor): What is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description?
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job?
                            - SUMMARY3 (The Execution & Impact): How did they build it, and why does it matter? Weave specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (The Closer): Based on their past execution, what specific value will they deliver on Day 1 in THIS new role?

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth for experience.
                            - The technical interview Q&A is SECONDARY.
                            If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            SKILLS SECTION
                            ========================
                            - EXACTLY 4 items
                            - Prioritize the specific "Must-Have" technologies.
                            - Use only tools explicitly mentioned in resume/Q&A/notes
                            Format: "Skill Area (Tool1, Tool2, Tool3, Tool4)"
                            Years format: "X+ years, current" OR "X+ years, 2026"

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. 
                            4. FACT AUDIT: Before outputting the final JSON, verify every single tool.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY1": "Bullet 1 following the blueprint",
                              "SUMMARY2": "Bullet 2 following the blueprint",
                              "SUMMARY3": "Bullet 3 following the blueprint",
                              "SUMMARY4": "Bullet 4 following the blueprint",
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
                            time.sleep(2) 
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

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
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

                    out_file = f"Submission_CBRE_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 🟠 CLIENT APP 6: BNSF 🟠 ---
# ====================================================================
def bnsf_app():
    TEMPLATE_FILENAME = "BNSF_Template.docx"
    st.title("BNSF Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="bn_loc")
        
        Remote_or_Onsite = st.selectbox("Remote or Onsite/Hybrid", ["Remote", "Onsite/Hybrid"], index=0, key="bn_rem")
        bn_location_name = ""
        if Remote_or_Onsite == "Onsite/Hybrid":
            bn_location_name = st.text_input("Location Name (e.g., HQ2)", key="bn_loc_name")
            
        bn_start = st.text_input("Availability to Start", key="bn_start")
        
        # --- NEW BNSF SPECIFIC FIELDS ---
        Bill_Rate = st.text_input("Bill Rate", key="bn_rate")
        Work_Status = st.selectbox("Work Status", ["USC", "GC", "Visa"], index=0, key="bn_status")

    with col2:
        Former_BNSF = st.selectbox("Former BNSF", ["Y", "N"], index=1, key="bn_form")
        bn_mgr = ""
        bn_dates = ""
        if Former_BNSF == "Y":
            bn_mgr = st.text_input("Manager Name", key="bn_mgr")
            bn_dates = st.text_input("Dates Worked", key="bn_dates")
            
        LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link", key="bn_link")

    if Remote_or_Onsite == "Onsite/Hybrid" and bn_location_name.strip():
        remote_val = f"Onsite/Hybrid - {bn_location_name.strip()}"
    else:
        remote_val = Remote_or_Onsite
        
    if Former_BNSF == "Y":
        former_bnsf_val = f"Y - {bn_mgr.strip()}, {bn_dates.strip()}"
    else:
        former_bnsf_val = "N"

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="bn_q1")
        Question_2 = st.text_input("Question 2", key="bn_q2")
        Question_3 = st.text_input("Question 3", key="bn_q3")
        Question_4 = st.text_input("Question 4", key="bn_q4")
        Question_5 = st.text_input("Question 5", key="bn_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="bn_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="bn_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="bn_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="bn_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="bn_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="bn_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="bn_res")

    if st.button("🚀 Generate BNSF Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
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
                        "FullName": name, "Location": Current_Location_City_ST, "Remote": remote_val,
                        "Start": bn_start,
                        "Certifications": data.get("Certifications", ""),
                        "FormerBNSF": former_bnsf_val, "Links": LinkedIn_GitHub_Portfolio_Link,
                        "BillRate": Bill_Rate, "WorkStatus": Work_Status,
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", 
                        "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
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
                            THE NARRATIVE BLUEPRINT (4 Bullet Points)
                            ========================
                            Generate exactly 4 distinct bullet points mapped to SUMMARY1 through SUMMARY4. 
                            CRITICAL RULE: DO NOT use the candidate's name or 3rd person pronouns (he/she/they) in ANY of the bullet points. Start each bullet with a strong action verb or direct statement.
                            
                            - SUMMARY1 (The Anchor): What is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description?
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job?
                            - SUMMARY3 (The Execution & Impact): How did they build it, and why does it matter? Weave specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (The Closer): Based on their past execution, what specific value will they deliver on Day 1 in THIS new role?

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth for experience.
                            - The technical interview Q&A is SECONDARY.
                            If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            SKILLS SECTION
                            ========================
                            - EXACTLY 4 items
                            - Prioritize the specific "Must-Have" technologies.
                            - Use only tools explicitly mentioned in resume/Q&A/notes
                            Format: "Skill Area (Tool1, Tool2, Tool3, Tool4)"
                            Years format: "X+ years, current" OR "X+ years, 2026"

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. 
                            4. FACT AUDIT: Before outputting the final JSON, verify every single tool.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY1": "Bullet 1 following the blueprint",
                              "SUMMARY2": "Bullet 2 following the blueprint",
                              "SUMMARY3": "Bullet 3 following the blueprint",
                              "SUMMARY4": "Bullet 4 following the blueprint",
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
                            time.sleep(2) 
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

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
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

                    out_file = f"Submission_BNSF_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- ⬜ CLIENT APP 7: DALLAS GENERIC ⬜ ---
# ====================================================================
def dallas_generic_app():
    TEMPLATE_FILENAME = "DallasGeneric_Template.docx"
    st.title("Dallas Generic Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="dg_loc")
        
        Remote_or_Onsite = st.selectbox("Remote or Onsite/Hybrid", ["Remote", "Onsite/Hybrid"], index=0, key="dg_rem")
        dg_location_name = ""
        if Remote_or_Onsite == "Onsite/Hybrid":
            dg_location_name = st.text_input("Location Name (e.g., HQ2)", key="dg_loc_name")
            
        dg_start = st.text_input("Availability to Start", key="dg_start")

    with col2:
        LinkedIn_GitHub_Portfolio_Link = st.text_input("LinkedIn/GitHub/Portfolio Link", key="dg_link")

    if Remote_or_Onsite == "Onsite/Hybrid" and dg_location_name.strip():
        remote_val = f"Onsite/Hybrid - {dg_location_name.strip()}"
    else:
        remote_val = Remote_or_Onsite

    st.header("🎤 Supplier Technical Interview Results")
    qa_col1, qa_col2 = st.columns(2)
    with qa_col1:
        Question_1 = st.text_input("Question 1", key="dg_q1")
        Question_2 = st.text_input("Question 2", key="dg_q2")
        Question_3 = st.text_input("Question 3", key="dg_q3")
        Question_4 = st.text_input("Question 4", key="dg_q4")
        Question_5 = st.text_input("Question 5", key="dg_q5")
    with qa_col2:
        Answer_1 = st.text_area("Answer 1", height=68, key="dg_a1")
        Answer_2 = st.text_area("Answer 2", height=68, key="dg_a2")
        Answer_3 = st.text_area("Answer 3", height=68, key="dg_a3")
        Answer_4 = st.text_area("Answer 4", height=68, key="dg_a4")
        Answer_5 = st.text_area("Answer 5", height=68, key="dg_a5")

    st.header("📝 Job Description & Notes")
    Job_Description_Notes_etc = st.text_area("Paste JD, Manager Notes, etc. here:", height=200, key="dg_jd")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="dg_res")

    if st.button("🚀 Generate Dallas Generic Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found. Please make sure it is uploaded to your GitHub repository.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors (e.g., "TCS", "Cognizant") that share the same line.
                    3. Education: Extract School, Degree, and DegreeStatus.
                        - DegreeStatus: Evaluate if the degree is completed or in progress.
                        - Return "Yes" if the resume indicates the degree is finished.
                        - Return "Pursuing" if the resume indicates ongoing study, contains the word "Pursuing", "In-progress", or lists a graduation date in the future (relative to June 2026).
                    4. Experience: Company, Title, Bullets (LIST), Environment (String, optional), Dates.
                        - For 'Title', clean the string by physically stripping out any employment type modifiers, hyphens, or parentheses at the end of the title (e.g., remove '- Contract', '(Contract)', or '- Consultant').
                        - For 'Environment', you may ONLY extract this if the original resume explicitly uses the word "Environment:" or "Technologies:" at the bottom of the role. If those exact words are not there, you MUST leave it blank "". Do NOT auto-generate or compile an environment list from the bullet points.
                    5. Certifications: Extract any certifications listed into a single comma-separated string. If none are found, leave it blank "".

                    JSON Structure:
                    {{
                        "FullName": "",
                        "Certifications": "",
                        "Education": [{{"School": "", "Degree": "", "DegreeStatus": "Yes"}}],
                        "Experience": [{{"Company": "", "Title": "", "Bullets": [], "Environment": "", "Dates": ""}}]
                    }}

                    RESUME: {raw_text}
                    """

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
                        "FullName": name, "Location": Current_Location_City_ST, "Remote": remote_val,
                        "Start": dg_start,
                        "Certifications": data.get("Certifications", ""),
                        "Links": LinkedIn_GitHub_Portfolio_Link, 
                        "Q1": Question_1, "A1": Answer_1, "Q2": Question_2, "A2": Answer_2,
                        "Q3": Question_3, "A3": Answer_3, "Q4": Question_4, "A4": Answer_4, "Q5": Question_5, "A5": Answer_5,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", 
                        "SKILL1": "", "YEARS1": "", "SKILL2": "", "YEARS2": "",
                        "SKILL3": "", "YEARS3": "", "SKILL4": "", "YEARS4": ""
                    }

                    edu = data.get('Education', [])
                    for i in range(1, 4):
                        mapping[f"School{i}"] = clean_school(edu[i-1].get('School', '')) if i <= len(edu) else ""
                        mapping[f"Degree{i}"] = edu[i-1].get('Degree', '') if i <= len(edu) else ""
                        mapping[f"DegreeStatus{i}"] = edu[i-1].get('DegreeStatus', 'Yes') if i <= len(edu) else ""

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
                            THE NARRATIVE BLUEPRINT (4 Bullet Points)
                            ========================
                            Generate exactly 4 distinct bullet points mapped to SUMMARY1 through SUMMARY4. 
                            CRITICAL RULE: DO NOT use the candidate's name or 3rd person pronouns (he/she/they) in ANY of the bullet points. Start each bullet with a strong action verb or direct statement.
                            
                            - SUMMARY1 (The Anchor): What is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description?
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. What is the most impressive, relevant project they recently delivered that proves they can handle this specific job?
                            - SUMMARY3 (The Execution & Impact): How did they build it, and why does it matter? Weave specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (The Closer): Based on their past execution, what specific value will they deliver on Day 1 in THIS new role?

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected', 'Spearheaded').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            🔴 Q&A USAGE RULES (CRITICAL)
                            ========================
                            - The resume is the PRIMARY source of truth for experience.
                            - The technical interview Q&A is SECONDARY.
                            If there is any conflict: 👉 prioritize the resume over Q&A.

                            ========================
                            SKILLS SECTION
                            ========================
                            - EXACTLY 4 items
                            - Prioritize the specific "Must-Have" technologies.
                            - Use only tools explicitly mentioned in resume/Q&A/notes
                            Format: "Skill Area (Tool1, Tool2, Tool3, Tool4)"
                            Years format: "X+ years, current" OR "X+ years, 2026"

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. 
                            4. FACT AUDIT: Before outputting the final JSON, verify every single tool.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON:

                            {{
                              "SUMMARY1": "Bullet 1 following the blueprint",
                              "SUMMARY2": "Bullet 2 following the blueprint",
                              "SUMMARY3": "Bullet 3 following the blueprint",
                              "SUMMARY4": "Bullet 4 following the blueprint",
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
                            time.sleep(2) 
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

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
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

                    out_file = f"Submission_DallasGeneric_{name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- 📄 NEW APP: PDF TO WORD CONVERTER 📄 ---
# ====================================================================
def pdf_to_word_app():
    st.title("PDF to Word Converter")
    st.markdown("Upload a PDF resume below to convert it to a Word Document (.docx) without any AI formatting or content changes.")

    st.header("📂 File Upload")
    resume_file = st.file_uploader("📤 Upload PDF Resume...", type=['pdf'], key="pdf2word_res")

    if st.button("🔄 Convert to Word (Adobe High Fidelity)", type="primary"):
        if not resume_file:
            st.error("❌ Error: Please upload a PDF file.")
        else:
            with st.spinner("Converting via Adobe PDF Services... this may take 15-30 seconds."):
                try:
                    # Pull Adobe credentials from Streamlit secrets
                    client_id = st.secrets["PDF_SERVICES_CLIENT_ID"]
                    client_secret = st.secrets["PDF_SERVICES_CLIENT_SECRET"]

                    # Read PDF bytes directly from uploader (no temp file needed)
                    pdf_bytes = resume_file.read()

                    # Call Adobe converter
                    from adobe_converter import convert_pdf_to_docx
                    docx_bytes = convert_pdf_to_docx(pdf_bytes, client_id, client_secret)

                    docx_filename = f"Converted_{resume_file.name.replace('.pdf', '')}.docx"

                    st.download_button(
                        label="⬇️ Download Word Document",
                        data=docx_bytes,
                        file_name=docx_filename,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary"
                    )
                    st.success("✅ Success! Document converted with full formatting preserved.")

                except KeyError:
                    st.error("❌ Adobe API credentials not found in Streamlit Secrets. Please add PDF_SERVICES_CLIENT_ID and PDF_SERVICES_CLIENT_SECRET.")
                except RuntimeError as e:
                    st.error(f"❌ Adobe Conversion Failed: {str(e)}")
                except Exception as e:
                    st.error(f"❌ Unexpected Error: {str(e)}")


# ====================================================================
# --- 🔵 CLIENT APP 8: DELOITTE 🔵 ---
# ====================================================================
def deloitte_app():
    TEMPLATE_FILENAME = "Deloitte_Template.docx"
    st.title("Deloitte Precision Extractor")

    with st.sidebar:
        st.header("🔑 API Configuration")
        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input("Gemini API Key", type="password")
            st.info("Paste your Gemini API key here to run the tool.")

    st.header("📋 Candidate Information")
    col1, col2 = st.columns(2)
    with col1:
        Legal_Name = st.text_input("Candidate Legal Name", key="del_legal")
        Current_Location_City_ST = st.text_input("Current Location (City, ST)", key="del_loc")
    with col2:
        Preferred_Name = st.text_input("Preferred Name", key="del_pref")
        Time_Off = st.text_input("Upcoming Scheduled Time off", key="del_toff")

    st.header("📝 Job Description & Notes")

    # --- GOOGLE SHEET LOGIC FOR DELOITTE ---
    SHEET_URL = "https://docs.google.com/spreadsheets/d/1swKaDhGdlmO0ILv7tngkRCbsK-7jF89cdCpf5bypzwM/edit?gid=2029588481#gid=2029588481"

    @st.cache_data(ttl=600)
    def load_deloitte_reqs(url):
        try:
            # Use the exact GID for the Deloitte sheet
            base_url = url.split("/edit")[0]
            csv_url = base_url + "/export?format=csv&gid=2029588481"
            df = pd.read_csv(csv_url)
            return df.fillna("")
        except Exception:
            return pd.DataFrame()

    df_reqs = load_deloitte_reqs(SHEET_URL)
    options = ["None"]
    req_mapping = {}

    if not df_reqs.empty:
        for idx, row in df_reqs.iterrows():
            # Column A (index 0) = ID, Column B (index 1) = VMS ID, Column C (index 2) = Job Title
            req_id = str(row.iloc[0]).replace(".0", "") if len(row) > 0 else ""
            vms_id = str(row.iloc[1]).replace(".0", "") if len(row) > 1 else ""
            title = str(row.iloc[2]) if len(row) > 2 else ""
            
            # Column I (index 8) = Job Description
            desc = str(row.iloc[8]) if len(row) > 8 else ""

            # Skills starting from Column J (index 9) to the end of the row
            skills = []
            if len(row) > 9:
                for val in row.iloc[9:]:
                    val_str = str(val).strip()
                    if val_str:
                        skills.append(val_str)
            skills_matrix_text = "\n".join(skills)

            option_parts = [req_id, vms_id, title]
            option_str = " - ".join([p.strip() for p in option_parts if p.strip()])
            
            if option_str and option_str != "-":
                options.append(option_str)
                req_mapping[option_str] = {"desc": desc, "skills": skills_matrix_text}

    if "del_jd" not in st.session_state:
        st.session_state.del_jd = ""
    if "del_notes" not in st.session_state:
        st.session_state.del_notes = ""
    if "del_skills" not in st.session_state:
        st.session_state.del_skills = ""

    def update_jd_text():
        if "del_req_selector" not in st.session_state:
            return
        selected = st.session_state.del_req_selector
        if selected != "None" and selected in req_mapping:
            st.session_state.del_jd = req_mapping[selected].get("desc", "").strip()
            st.session_state.del_skills = req_mapping[selected].get("skills", "").strip()
            st.session_state.del_notes = ""
        else:
            st.session_state.del_jd = ""
            st.session_state.del_skills = ""
            st.session_state.del_notes = ""
            
    selected_req = st.selectbox(
        "🔍 Select a Deloitte Requisition (Type to search):",
        options=options,
        key="del_req_selector",
        on_change=update_jd_text
    )

    Job_Description = st.text_area("1. Official Job Description (Paste generic JD here):", height=150, key="del_jd")
    Chat_Notes = st.text_area("2. Notes (Optional - Overrides JD):", height=100, key="del_notes")
    Skills_Matrix = st.text_area("3. Skills Matrix:", height=100, key="del_skills")

    st.header("📂 File Uploads")
    resume_file = st.file_uploader("📤 Upload the candidate's resume...", type=['pdf', 'docx', 'doc'], key="del_res")

    if st.button("🚀 Generate Deloitte Submission", type="primary"):
        if not API_KEY:
            st.error("❌ Error: Please enter your Gemini API Key in the sidebar.")
        elif not resume_file:
            st.error("❌ Error: Please upload a resume.")
        elif not os.path.exists(TEMPLATE_FILENAME):
            st.error(f"❌ Error: The preloaded template '{TEMPLATE_FILENAME}' was not found.")
        else:
            with st.spinner("Processing document and querying AI..."):
                try:
                    resume_path = f"temp_{resume_file.name}"
                    with open(resume_path, "wb") as f:
                        f.write(resume_file.getbuffer())

                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors.
                    3. Education: Extract School and Degree.
                    4. Experience: Company, Title, Bullets (LIST), Environment (String, optional), Dates.
                        - For 'Title', clean the string by physically stripping out any employment type modifiers, hyphens, or parentheses at the end of the title.
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

                    models_to_try = ['gemini-2.5-flash', 'gemini-3-flash-preview', 'gemini-3.1-flash-lite-preview', 'gemini-pro-latest']
                    data = {}
                    
                    for attempt in range(6):
                        current_model = models_to_try[0] if attempt == 0 else (models_to_try[1] if attempt in [1, 2] else (models_to_try[2] if attempt in [3, 4] else models_to_try[3]))
                        try:
                            response = client.models.generate_content(
                                model=current_model,
                                contents=prompt,
                                config=types.GenerateContentConfig(response_mime_type="application/json")
                            )
                            data = repair_and_load_json(response.text)
                            break
                        except Exception as e:
                            if "503" in str(e) and attempt < 5:
                                time.sleep(2 ** ((attempt % 2) + 1))
                                continue
                            else:
                                raise e

                    ai_name = data.get('FullName', '').title()
                    final_name = Legal_Name.strip() if Legal_Name.strip() else ai_name

                    # --- ADD THIS LOGIC HERE TO DEFINE resume_body ---
                    # We split the text to remove the Recruiter's Summary and keep the rest
                    if "Technical Skills" in raw_text:
                        resume_body = "Technical Skills" + raw_text.split("Technical Skills", 1)[1]
                    elif "Links:" in raw_text:
                        resume_body = "Links:" + raw_text.split("Links:", 1)[1]
                    else:
                        resume_body = raw_text 
                    # --------------------------------------------------

                    # --- UPDATED MAPPING TO SUPPORT 8 SKILLS + COMPANIES + 5 SUMMARIES ---
                    mapping = {
                        "FullName": final_name, "PreferredName": Preferred_Name, 
                        "TimeOff": Time_Off, "Location": Current_Location_City_ST,
                        "RESUME_BODY": resume_body,
                        "SUMMARY1": "", "SUMMARY2": "", "SUMMARY3": "", "SUMMARY4": "", "SUMMARY5": "",
                        "SKILL1": "", "YEARS1": "", "SKILL1COMPANIES": "",
                        "SKILL2": "", "YEARS2": "", "SKILL2COMPANIES": "",
                        "SKILL3": "", "YEARS3": "", "SKILL3COMPANIES": "",
                        "SKILL4": "", "YEARS4": "", "SKILL4COMPANIES": "",
                        "SKILL5": "", "YEARS5": "", "SKILL5COMPANIES": "",
                        "SKILL6": "", "YEARS6": "", "SKILL6COMPANIES": "",
                        "SKILL7": "", "YEARS7": "", "SKILL7COMPANIES": "",
                        "SKILL8": "", "YEARS8": "", "SKILL8COMPANIES": ""
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

                    if Job_Description.strip():
                        try:
                            # --- UPDATED SUMMARY PROMPT FOR DELOITTE ---
                            summary_prompt = f"""
                            You are an elite, no-nonsense Senior Technical Recruiter writing an executive submission summary for a Fieldglass portal. 
                            The Hiring Manager has 30 seconds to read this. Your goal is to make a punchy, evidence-based business case for why this candidate will succeed, optimized for both human reading and ATS exact-match parsing.

                            ========================
                            🚨 HIERARCHY OF TRUTH: JD vs. MANAGER NOTES (CRITICAL)
                            ========================
                            You are receiving two primary sources of information for this role:
                            1. Notes (HIGH PRIORITY)
                            2. Official Job Description (BASELINE)

                            - OVERRIDE RULE: The Notes represent the Hiring Manager's actual, immediate needs. They are the ABSOLUTE SOURCE OF TRUTH. If they contradict or update the official Job Description, you MUST strictly follow the Notes.
                            - ATS COMPROMISE: To satisfy the Fieldglass parsing algorithm, you must still weave in high-level, generalized keywords from the baseline Job Description where applicable, but NEVER highlight technical skills or responsibilities that the manager explicitly negated in their notes.

                            ========================
                            THE NARRATIVE BLUEPRINT (5 Bullet Points)
                            ========================
                            Follow this exact structure for the SUMMARY. Generate exactly 5 distinct bullet points mapped to SUMMARY1 through SUMMARY5. Every point MUST sell the candidate's fit for the role:
                            - SUMMARY1 (The Anchor): The candidate's TOTAL progressive professional experience and dominant expertise that solves the PRIMARY technical "must-have" of the Job Description. (Use their FIRST NAME only. Calculate total overall years strictly rounding DOWN to the nearest whole year formatted as 'X+ years').
                            - SUMMARY2 (The Alignment): Explicitly name their CURRENT or most recent employer. Frame their most impressive, relevant project as a direct 1:1 match for the manager's current challenge.
                            - SUMMARY3 (The Execution): How did they build it? Weave the specific tools/methodologies into an "Execution Statement" that highlights complexity, scale, or business impact.
                            - SUMMARY4 (Secondary Expertise): Highlight a secondary skill, architecture, or methodology requested in the JD that the candidate also possesses.
                            - SUMMARY5 (The Closer): Based on their past execution, state what specific value they will deliver on Day 1 in THIS new role.

                            CRITICAL ATS HACK: Across these 5 bullets, you MUST seamlessly embed exact phrases from the Job Description to maximize the Fieldglass match score.

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - TECH MATCHING: Strictly align the tools you highlight with the JD. 
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            SKILLS SECTION (8 SKILLS + COMPANIES)
                            ========================
                            - EXTRACT UP TO 8 SKILLS MAXIMUM.
                            - Prioritize the specific "Must-Have" technologies AND methodologies requested in the JD and the provided Skills Matrix.
                            - You MUST identify all companies from the candidate's Resume Experience section where they utilized each skill.

                            Format: "Skill Area (Tool1, Tool2)"
                            Years format: "X+ years, current" OR "X+ years"
                            Companies format: "Company A, Company B" (Extract exact company names from the resume where the skill was used)

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume.
                            2. DO NOT INFER OR ASSUME. 
                            3. DO NOT INFLATE TO MATCH THE JD. If the candidate lacks a requested skill, omit it completely.

                            ========================
                            YEARS ACCURACY RULES (REALISTIC RECRUITER MODE)
                            ========================
                            - Use June 2026 as the current date.
                            - STRICT MATH (DATE-DRIVEN ONLY): Calculate years strictly based on the earliest chronological date provided in the 'Professional Experience'. Round DOWN to the nearest whole year.

                            ========================
                            OUTPUT FORMAT (STRICT)
                            ========================
                            Return ONLY valid JSON. If less than 8 skills are found, leave the remaining skill strings empty ("").

                            {{
                              "SUMMARY1": "Bullet 1...",
                              "SUMMARY2": "Bullet 2...",
                              "SUMMARY3": "Bullet 3...",
                              "SUMMARY4": "Bullet 4...",
                              "SUMMARY5": "Bullet 5...",
                              "SKILL1": "Skill Area",
                              "YEARS1": "X+ years",
                              "SKILL1COMPANIES": "Company A, Company B",
                              "SKILL2": "Skill Area",
                              "YEARS2": "X+ years",
                              "SKILL2COMPANIES": "Company C",
                              "SKILL3": "",
                              "YEARS3": "",
                              "SKILL3COMPANIES": "",
                              "SKILL4": "",
                              "YEARS4": "",
                              "SKILL4COMPANIES": "",
                              "SKILL5": "",
                              "YEARS5": "",
                              "SKILL5COMPANIES": "",
                              "SKILL6": "",
                              "YEARS6": "",
                              "SKILL6COMPANIES": "",
                              "SKILL7": "",
                              "YEARS7": "",
                              "SKILL7COMPANIES": "",
                              "SKILL8": "",
                              "YEARS8": "",
                              "SKILL8COMPANIES": ""
                            }}

                            ========================
                            INPUT DATA
                            ========================
                            Official Job Description:
                            {Job_Description}

                            Notes:
                            {Chat_Notes}

                            Skills Matrix:
                            {Skills_Matrix}

                            Resume:
                            {raw_text}
                            """
                            time.sleep(2) 
                            summary_data = {}
                            
                            for attempt in range(6):
                                current_model = models_to_try[0] if attempt == 0 else (models_to_try[1] if attempt in [1, 2] else (models_to_try[2] if attempt in [3, 4] else models_to_try[3]))
                                try:
                                    summary_response = client.models.generate_content(
                                        model=current_model,
                                        contents=summary_prompt,
                                        config=types.GenerateContentConfig(response_mime_type="application/json")
                                    )
                                    summary_data = repair_and_load_json(summary_response.text)
                                    break
                                except Exception as api_e:
                                    if "503" in str(api_e) and attempt < 5:
                                        time.sleep(2 ** ((attempt % 2) + 1))
                                        continue
                                    else:
                                        raise api_e

                            mapping["SUMMARY1"] = summary_data.get("SUMMARY1", "")
                            mapping["SUMMARY2"] = summary_data.get("SUMMARY2", "")
                            mapping["SUMMARY3"] = summary_data.get("SUMMARY3", "")
                            mapping["SUMMARY4"] = summary_data.get("SUMMARY4", "")
                            mapping["SUMMARY5"] = summary_data.get("SUMMARY5", "")
                            
                            for i in range(1, 9):
                                mapping[f"SKILL{i}"] = summary_data.get(f"SKILL{i}", "")
                                mapping[f"YEARS{i}"] = summary_data.get(f"YEARS{i}", "")
                                mapping[f"SKILL{i}COMPANIES"] = summary_data.get(f"SKILL{i}COMPANIES", "")
                                
                        except Exception as e:
                            st.warning(f"⚠️ Warning: Summary generation failed. Proceeding without it. ({e})")

                    out_file = f"Submission_Deloitte_{final_name.replace(' ', '_')}.docx"
                    process_word_doc(TEMPLATE_FILENAME, mapping, out_file)
                    
                    with open(out_file, "rb") as file:
                        btn = st.download_button(
                            label="⬇️ Download Generated Document",
                            data=file,
                            file_name=out_file,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            type="primary"
                        )
                    
                    st.success(f"✅ Success! Document is ready for download.")

                    try:
                        os.remove(resume_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- MAIN ROUTING LOGIC (The Dropdown Page) ---
# ====================================================================

st.title("Resume Formatter")
st.markdown("Please select your client account below to access the customized formatter.")

client_selection = st.selectbox(
    "Select Client Account or Tool:", 
    ["Fannie Mae", "Deloitte", "Peraton", "Capital One", "ADUSA", "CBRE", "BNSF", "Dallas Generic", "PDF to Word"]
)

st.divider()

if client_selection == "Fannie Mae":
    fannie_mae_app()
elif client_selection == "Deloitte":
    deloitte_app()
elif client_selection == "Peraton":
    peraton_app()
elif client_selection == "Capital One":
    capital_one_app()
elif client_selection == "ADUSA":
    adusa_app()
elif client_selection == "CBRE":
    cbre_app()
elif client_selection == "BNSF":
    bnsf_app()
elif client_selection == "Dallas Generic":
    dallas_generic_app()
elif client_selection == "PDF to Word":
    pdf_to_word_app()
