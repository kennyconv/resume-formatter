import streamlit as st
import google.generativeai as genai
import docx
from docx.text.paragraph import Paragraph
from docx.shared import Pt
import pypdf
import json
import re
import os
import copy

# ====================================================================
# --- PRELOADED TEMPLATE SETTING ---
# ====================================================================
# The app will automatically look for this file in your GitHub repo
TEMPLATE_FILENAME = "Template.docx"

# ====================================================================
# --- STREAMLIT UI & FORM INPUTS ---
# ====================================================================

st.set_page_config(page_title="Fannie Mae Precision Extractor", layout="wide")
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

Your goal is to clearly and convincingly explain why this candidate will succeed in THIS specific role by making a strong, evidence-based case using their real experience.

========================
🔴 CORE OBJECTIVE (CRITICAL)
========================
This is NOT a resume summary.

This should read like a recruiter explaining:
👉 why this candidate will succeed in THIS role
👉 what they’ve already done that makes this job easy for them

Every sentence must be grounded in:
- real experience
- real systems
- real ownership

========================
🔴 STYLE & TONE RULES
========================
- Write 4–5 sentences total
- Use ONLY the candidate’s FIRST NAME
- Tone must be natural, confident, and credible (not robotic or templated)
- Vary sentence structure (DO NOT follow a fixed blueprint)
- Write like a human making a case, not AI following rules

========================
🔴 WHAT GOOD LOOKS LIKE
========================
Strong summaries:
- Highlight specific work (pipelines, APIs, models, systems, architecture)
- Show ownership (built, designed, implemented, led)
- Subtly connect experience to the role (NOT forced mapping)
- Feel tailored to THIS candidate and THIS role

Weak summaries:
- Repeat job description language
- Say “experience with X” without context
- Use generic phrases like:
  "aligns with", "fits well", "can contribute", "strong background"
- Follow identical structure every time

========================
🚫 HARD RULES (DO NOT BREAK)
========================
DO NOT:
- Use “this directly translates”
- Parrot or restate the job description
- Copy phrasing from the resume without adding insight
- Use generic filler language
- Sound templated or repetitive

========================
🔴 REQUIRED CONTENT
========================

Sentence 1:
- Clearly establish who they are + years of experience + core domain

Sentence 2–3:
- Focus on MOST RELEVANT EXPERIENCE
- Prefer MOST RECENT ROLE (if relevant)
- Include specific systems, tools, or work they’ve done
- Show depth (not just listing technologies)

Sentence 3–4:
- Explain WHY that experience matters for THIS role
- This should feel natural, not forced
- Think: “they’ve already done something very similar”

Final Sentence:
- Close with a confident, grounded statement of impact
- Focus on what they’ll be able to step in and do

========================
🧠 THINK LIKE THIS (IMPORTANT)
========================
Before writing, ask yourself:

"What has this person already done that makes this job easy for them?"

Then write the summary around THAT.

========================
🟡 CONTEXT PRIORITY
========================
- If manager notes / feedback exist → PRIORITIZE THEM
- Still include key technical keywords from the JD (for MSP filtering)
- Emphasize hands-on experience and real ownership

========================
🟢 SKILLS SECTION RULES
========================
- EXACTLY 4 items
- Each must be:
  • highly relevant to the role
  • keyword-rich (for MSP search)
  • specific (not generic categories)

FORMAT:
"Skill Area (Tool1, Tool2, Tool3, Tool4)"

YEARS FORMAT:
"X+ years, current" OR "X+ years, 2025"

GOOD EXAMPLES:
- "LLM & GenAI Engineering (RAG, Fine-tuning, NL2SQL, Prompt Engineering)"
- "AWS Data Pipelines (Glue, Lambda, Step Functions, S3)"
- "Python Data Engineering (PySpark, Pandas, ETL Pipelines)"
- "MLOps & Model Lifecycle (MLflow, Kubeflow, Docker, CI/CD)"

DO NOT:
- Use generic labels like “Software Engineering”
- Include irrelevant tools just to fill space

========================
🟢 SKILLS ACCURACY RULES
========================
- Only include skills, tools, and platforms explicitly supported by the resume, interview answers, or provided notes.
- Do NOT infer tools that are merely adjacent or commonly used together.
- If a tool is not directly mentioned, do not include it.
- Prefer narrower, fully supported skills over broader, inferred ones.

========================
🟢 YEARS ACCURACY RULES
========================
- Use April 2026 as the current date when calculating years.
- Calculate years conservatively based on the actual timeline in the resume.
- Do NOT assign years based on general career length or related experience.
- Distinguish between broad experience and exact tool/platform experience.
- If exact years for a specific tool are unclear, use the lowest clearly supported number.
- Use "current" only if the candidate is using that skill/tool in their current or most recent role.
- Otherwise use the latest supported year from the resume.

========================
🧠 FINAL SELF-CHECK (MANDATORY)
========================
Before output:

- Does this sound like a real recruiter wrote it?
- Does it avoid repeating the job description?
- Does it highlight REAL work (not vague claims)?
- Does it clearly show WHY this candidate will succeed?
- Does it feel tailored (not reusable)?
- Are all tools in the skills section explicitly supported by the resume or interview answers?
- Are years based on the exact tool/skill rather than general domain experience?
- Were years calculated using April 2026 as the current date?

If not → rewrite internally before output.

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
