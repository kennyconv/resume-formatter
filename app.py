import streamlit as st
import pandas as pd
from streamlit_gsheets import GSheetsConnection
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
import smtplib
import hashlib
from email.message import EmailMessage
from pdf2docx import Converter
from copy import deepcopy
from datetime import date

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
    """
    Normalize only date detail that is actually present in the source.

    IMPORTANT:
    - Never invent a month when the source provides only a year.
    - "2018 - 2020" must remain year-only.
    - "2018 - Present" must remain year-only on the start side.
    - Month/year inputs may still be standardized normally.
    """
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
    """
    Format Peraton dates without inventing missing month detail.

    Examples:
    - "10/2019 - 07/2024" -> "10/2019 to 07/2024"
    - "2019 - 2024"       -> "2019 to 2024"
    - "2019 - Present"    -> "2019 to Present"

    If the original resume gives only a year, keep only the year.
    """
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

def collapse_extra_blank_paragraphs(doc):
    """
    Collapse consecutive empty BODY paragraphs to a single blank paragraph.

    This is intentionally conservative:
    - Only operates on top-level document-body paragraphs.
    - Does NOT touch paragraphs inside tables.
    - Does NOT remove page-break or section-break paragraphs.
    - Preserves one intentional blank paragraph between sections.
    """

    body = doc.element.body
    previous_was_blank = False

    for element in list(body):

        # A table or other non-paragraph element breaks the blank-paragraph run.
        if not element.tag.endswith("}p"):
            previous_was_blank = False
            continue

        # ------------------------------------------------------------
        # DO NOT TOUCH STRUCTURAL PARAGRAPHS
        # ------------------------------------------------------------

        # Preserve section breaks.
        if element.xpath(".//w:sectPr"):
            previous_was_blank = False
            continue

        # Preserve explicit page/column/line breaks.
        if element.xpath(".//w:br"):
            previous_was_blank = False
            continue

        # Preserve drawings/images/shapes.
        if (
            element.xpath(".//w:drawing")
            or element.xpath(".//w:pict")
            or element.xpath(".//w:object")
        ):
            previous_was_blank = False
            continue

        # ------------------------------------------------------------
        # DETERMINE WHETHER PARAGRAPH IS VISUALLY EMPTY
        # ------------------------------------------------------------

        text_nodes = element.xpath(".//w:t")
        text = "".join(
            node.text or ""
            for node in text_nodes
        ).strip()

        is_blank = not text

        # ------------------------------------------------------------
        # COLLAPSE ONLY CONSECUTIVE BLANKS
        # ------------------------------------------------------------

        if is_blank:
            if previous_was_blank:
                body.remove(element)
            else:
                previous_was_blank = True
        else:
            previous_was_blank = False

def ensure_experience_spacing(doc, mapping):
    """
    Ensure one visible blank paragraph between consecutive work-history entries.

    This specifically protects dynamically generated jobs beyond the original
    seven template slots, including abbreviated older roles that have
    Company/Title/Dates but no bullets. Existing spacing is preserved; a new
    blank paragraph is added only when one is missing.
    """
    paragraphs = list(doc.paragraphs)
    company_pattern = re.compile(
        r"\{\{company(\d+)\}\}",
        flags=re.IGNORECASE,
    )

    role_locations = []

    for idx, paragraph in enumerate(paragraphs):
        match = company_pattern.search(paragraph.text or "")
        if not match:
            continue

        role_num = int(match.group(1))
        company_value = str(
            mapping.get(f"Company{role_num}", "") or ""
        ).strip()

        if company_value:
            role_locations.append((role_num, idx))

    if len(role_locations) < 2:
        return

    # Bottom-up so inserting a paragraph does not invalidate earlier indexes.
    for _, idx in reversed(role_locations[1:]):
        paragraphs = list(doc.paragraphs)

        if idx <= 0 or idx >= len(paragraphs):
            continue

        current_p = paragraphs[idx]
        previous_p = paragraphs[idx - 1]

        # Do nothing when the template/output already has a separator.
        if not previous_p.text.strip():
            continue

        blank = current_p.insert_paragraph_before("")

        try:
            blank.style = doc.styles["Normal"]
        except Exception:
            pass

        blank.paragraph_format.space_before = Pt(0)
        blank.paragraph_format.space_after = Pt(0)


def expand_experience_placeholders(doc, mapping):
    """
    Dynamically extend a template's work-history placeholder block when the
    candidate has more roles than the template was originally built to hold.

    Existing templates can stay exactly as they are (for example, through
    {{Company7}} / {{Dates7}} / {{Title7}} / {{Bullets7}}). If mapping contains
    Company8, Company9, etc., this function clones the final existing experience
    block and renumbers the placeholders before the normal replacement logic runs.

    This is deliberately template-driven rather than capped at an arbitrary
    number such as 20.
    """

    # Determine how many experience entries the mapping actually contains.
    mapped_indices = []
    for key in mapping.keys():
        match = re.fullmatch(
            r"(?i)(?:company|dates|title|bullets|environment|responsible)(\d+)",
            str(key),
        )
        if match:
            mapped_indices.append(int(match.group(1)))

    if not mapped_indices:
        return

    required_max = max(mapped_indices)

    # Locate the highest-numbered complete experience block already present
    # in the template body. Current client templates use paragraph-based
    # Company/Dates -> Title -> Bullets blocks.
    paragraphs = list(doc.paragraphs)
    placeholder_pattern = re.compile(
        r"\{\{(company|dates|title|bullets|environment|responsible)(\d+)\}\}",
        flags=re.IGNORECASE,
    )

    locations = {}
    for idx, paragraph in enumerate(paragraphs):
        for kind, number in placeholder_pattern.findall(paragraph.text or ""):
            number = int(number)
            locations.setdefault(number, {}).setdefault(kind.lower(), []).append(idx)

    complete_indices = [
        number
        for number, parts in locations.items()
        if "company" in parts and "title" in parts and "bullets" in parts
    ]

    if not complete_indices:
        return

    template_max = max(complete_indices)

    if required_max <= template_max:
        return

    parts = locations[template_max]
    block_start = min(
        parts["company"] + parts.get("dates", []) + parts["title"] + parts["bullets"]
    )
    block_end = max(
        parts["company"] + parts.get("dates", []) + parts["title"] + parts["bullets"]
    )

    # Include any Environment placeholder that belongs inside the same final block.
    if parts.get("environment"):
        block_end = max(block_end, max(parts["environment"]))

    # Preserve the template's inter-role spacing. Peraton, for example,
    # uses one blank Normal paragraph between experience blocks, while
    # the Fannie-style templates do not.
    if (
        block_start > 0
        and not paragraphs[block_start - 1].text.strip()
    ):
        block_start -= 1

    source_block = paragraphs[block_start:block_end + 1]
    if not source_block:
        return

    insert_after = source_block[-1]._p

    # Clone the final template block once for each additional candidate role.
    for new_index in range(template_max + 1, required_max + 1):
        cloned_elements = []

        for source_paragraph in source_block:
            new_p = copy.deepcopy(source_paragraph._p)

            # Renumber every placeholder in the cloned block that belongs to
            # the original final template index.
            text_nodes = new_p.xpath(".//w:t")
            combined_text = "".join(
                node.text or ""
                for node in text_nodes
            )

            # Word frequently splits one placeholder across multiple runs
            # (for example "{{" + "company" + "7" + "}}").  Once we confirm
            # this cloned paragraph belongs to the final template role, safely
            # renumber the placeholder fragments across those runs.
            if re.search(
                rf"(?i)\{{\{{(?:company|dates|title|bullets|environment|responsible){template_max}\}}\}}",
                combined_text,
            ):
                for text_node in text_nodes:
                    if not text_node.text:
                        continue

                    text_node.text = re.sub(
                        rf"(?i)(company|dates|title|bullets|environment|responsible){template_max}",
                        lambda m: f"{m.group(1)}{new_index}",
                        text_node.text,
                    )

                    if text_node.text.strip() == str(template_max):
                        text_node.text = text_node.text.replace(
                            str(template_max),
                            str(new_index),
                        )

            insert_after.addnext(new_p)
            insert_after = new_p
            cloned_elements.append(new_p)


def process_word_doc(temp_path, mapping, out_path):
    doc = docx.Document(temp_path)
    expand_experience_placeholders(doc, mapping)
    ensure_experience_spacing(doc, mapping)
    
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

    # Collapse extra blank lines created when optional sections/tables were removed.
    collapse_extra_blank_paragraphs(doc)
    
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
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="fm_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="fm_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="fm_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="fm_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="fm_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="fm_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="fm_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="fm_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="fm_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="fm_a5")

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
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
                            The summary should read like an experienced recruiter verbally explaining to the hiring manager why this candidate deserves an interview. Your goal is to make a punchy, evidence-based business case for why this candidate will succeed, optimized for both human reading and ATS exact-match parsing.

                            ========================
                            JD vs. MANAGER INPUT (CRITICAL)
                            ========================

                            You are receiving up to three sources of information:

                            1. Official Job Description
                            2. Spotlight Call Transcript
                            3. Fieldglass Chat Notes

                            OFFICIAL JOB DESCRIPTION:
                            Primary source for ATS keyword coverage and phrase matching.

                            SPOTLIGHT CALL + CHAT NOTES:
                            Primary source for hiring manager priorities and emphasis.

                            If manager input contradicts or removes a requirement from the Job Description, follow the manager.

                            Preserve strong coverage of the Official Job Description because automated shortlisting is assumed to rely primarily on Job Description terminology.

                            Never highlight technical skills or responsibilities that the manager explicitly negated.

                            ========================
                            ATS COVERAGE RULE (CRITICAL)
                            ========================

                            Assume the Fieldglass AI shortlisting engine evaluates candidates primarily against the Official Job Description.

                            Therefore:
                            - The Official Job Description is the primary source for keyword coverage and phrase matching.
                            - Spotlight Call transcripts and Chat Notes should influence emphasis and prioritization, but should not reduce coverage of important Job Description concepts.
                            - Preserve exact terminology from the Job Description whenever supported by the resume.
                            - Prefer exact Job Description terminology over synonyms when both are supported by the resume.
                            - Exact phrase overlap is preferred whenever it does not distort the truth.
                            - When multiple resume experiences are equally relevant, prefer the one that produces the greatest overlap with Job Description terminology and concepts.

                            ========================
                            THE NARRATIVE BLUEPRINT (4 Sentences Max)
                            ========================
                            Follow this exact structure for the SUMMARY. Every sentence MUST sell the candidate's fit for the role:
                            - Sentence 1: The Anchor (Authority). Who are they, what is their TOTAL progressive professional experience, and what is their dominant expertise that solves the PRIMARY technical "must-have" of the Job Description? (Calculate their total overall years strictly rounding DOWN to the nearest whole year formatted as 'X+ years'. Use their FIRST NAME only. You MUST use the exact job title requested in the JD if the candidate's history supports it. Avoid generic fluff).
                            - Sentence 2: The Alignment (The Hook). You MUST explicitly name the current employer unless an earlier role provides substantially stronger evidence for the target role. Highlight the most relevant accomplishment or responsibility that demonstrates success performing the core duties of this role. Show why this experience provides strong evidence they can perform the responsibilities of the target position.
                            - Sentence 3: The Execution & Impact (The Proof). Explain how they executed the work and why it mattered. Highlight the scale, complexity, operational ownership, or business impact. Weave tools and methodologies naturally into the sentence. DO NOT write a comma-separated list of tools. DO NOT repeat verbs from Sentence 2.
                            - Sentence 4: The Closer (The ROI). Based on their past execution, what specific value will they deliver on Day 1 in THIS new role? (Use a strong, direct structure like: "[First Name]'s success in [X] makes them an immediate asset for driving this team's [specific technical goal/initiative]." Do NOT mention the physical location/city).

                            CRITICAL ATS HACK: Across these 4 sentences, you MUST seamlessly embed 2-3 exact phrases from the Job Description (including soft skills like "changing priorities" or "system analysis") to maximize the Fieldglass match score. Do not force them if they ruin the sentence flow, but prioritize exact phrase matching where supported by the resume.

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - TECH MATCHING: Strictly align the tools you highlight with the JD. If the JD asks for AWS, highlight AWS. Do not highlight competing tech (like Azure or GCP) just because it's prominent in the resume, unless it is their only experience.
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - SHOW, DON'T TELL. 
                              🔴 Bad: "John's background in AWS makes him a great fit for this role."
                              🟢 Good: "Because John spent the last three years building highly available data lakes in AWS, he can immediately step in to optimize your current infrastructure."
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely," "exceptionally well-prepared," "fits well," "aligns with," "enterprise-grade platforms."
                            - Do NOT use transition crutches: "Additionally," "Furthermore," "Moreover."
                            - Do NOT repeat or restate the job description.
                            - Prefer clear business language over consultant buzzwords.
                            - Favor evidence and ownership over dramatic wording.
                            - Use verbs that naturally fit the work performed. Do not force "leadership verbs."

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
                            EXPERIENCE PRIORITY RULE (CRITICAL)
                            ========================

                            Prefer accomplishments explicitly demonstrated in experience bullets over technologies listed only in the skills section.

                            Favor things the candidate built, created, supported, deployed, maintained, optimized, or owned rather than technologies merely listed in a skills inventory.

                            Prefer recent and recurring responsibilities over isolated projects whenever both are supported by the resume.

                            When multiple experiences are relevant, prefer the experience that overlaps with the largest number of Job Description requirements and responsibilities.

                            ========================
                            OPERATIONAL OWNERSHIP RULE
                            ========================

                            Prefer recurring responsibilities and direct ownership over one-time projects or transformation initiatives.

                            Unless the role is explicitly implementation-focused, prioritize evidence of day-to-day execution, production support, issue resolution, audits, compliance, platform ownership, and measurable business outcomes over project participation or consulting activities.

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
                            "X+ years, current" OR "X+ years, "

                            ========================
                            🚨 ZERO-TOLERANCE HALLUCINATION RULES (CRITICAL)
                            ========================
                            1. THE RESUME IS THE ONLY SOURCE OF TRUTH. You are strictly forbidden from copying a skill, tool, or technology from the Job Description and assigning it to the candidate unless it physically appears in their Resume or Q&A.
                            2. DO NOT INFER OR ASSUME. If the JD asks for "Redshift" and the resume only says "AWS", you MUST NOT write "Redshift" under any circumstances. 
                            3. DO NOT INFLATE TO MATCH THE JD. If the candidate lacks a requested skill, omit it completely. It is better to have an incomplete match than a fabricated one.
                            4. FACT AUDIT: Before outputting the final JSON, you must verify every single tool mentioned in your SUMMARY and SKILLS. If a tool exists in the JD but not in the Resume/Q&A, you must delete it from your output.
                            5. DO NOT ELEVATE MANAGER LANGUAGE INTO CANDIDATE EXPERIENCE. If the Job Description, Chat Notes, or Spotlight Call mention technologies, responsibilities, or business goals, you may reference them as objectives of the target role, but you MUST NOT imply that the candidate performed those activities unless supported by the resume or Q&A.

                            ========================
                            YEARS ACCURACY RULES (REALISTIC RECRUITER MODE)
                            ========================
                            - Use October 2026 as the current date.
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
                            8. OWNERSHIP TEST: Did I emphasize what the candidate repeatedly owned and executed, rather than simply the most impressive-sounding project?
                            9. ATS COVERAGE TEST: Did I naturally incorporate the most important technologies, responsibilities, and terminology from the Official Job Description that are supported by the resume?
                            10. EVIDENCE TEST: Would every skill, tool, accomplishment, and responsibility in the summary survive a side-by-side comparison against the actual resume?
                            11. JD OVERLAP TEST: Did I emphasize the experiences that provide the strongest overlap with the Official Job Description rather than simply the most impressive accomplishments?

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
# --- 🟣 FREDDIE MAC HELPER FUNCTIONS 🟣 ---
# ====================================================================

FREDDIE_VNDLY_SKILLS_FILENAME = "freddie_vndly_skills_curated.json"


@st.cache_data(show_spinner=False)
def fred_load_vndly_skill_catalog(path=FREDDIE_VNDLY_SKILLS_FILENAME):
    """
    Load the curated VNDLY skill catalogue bundled with the Streamlit app.
    """
    if not os.path.exists(path):
        return []

    try:
        with open(path, "r", encoding="utf-8") as f:
            payload = json.load(f)

        raw_skills = payload.get("skills", payload)

        if not isinstance(raw_skills, list):
            return []

        cleaned = []

        for item in raw_skills:
            if not isinstance(item, dict):
                continue

            skill_name = str(
                item.get("skill_name")
                or item.get("value")
                or item.get("label")
                or ""
            ).strip()

            if not skill_name:
                continue

            cleaned.append(
                {
                    "id": item.get("id"),
                    "skill_name": skill_name,
                    "catalog_tier": str(
                        item.get("catalog_tier", "canonical_preferred")
                    ).strip() or "canonical_preferred",
                }
            )

        return cleaned

    except Exception:
        return []


def fred_cap_skill_details(value, max_items=6):
    """
    Freddie-only safety cap for the right-hand Skills-table column.

    Gemini is instructed to return at most six comma-separated details, but this
    deterministic pass guarantees the Word output never exceeds that limit.
    """
    raw = str(value or "").strip()
    if not raw:
        return ""

    items = []
    seen = set()

    for part in raw.split(","):
        clean = re.sub(r"\s+", " ", str(part or "")).strip()
        key = clean.casefold()

        if not clean or key in seen:
            continue

        items.append(clean)
        seen.add(key)

        if len(items) >= max_items:
            break

    return ", ".join(items)


def fred_normalize_skill_name(value):
    """Conservative normalization used only for exact catalogue lookups."""
    return re.sub(
        r"\s+",
        " ",
        str(value or "").replace("\u00a0", " "),
    ).strip().casefold()


def fred_normalize_match_surface(value):
    """Normalize Freddie text for safe token/phrase matching."""
    value = str(value or "")
    value = (
        value.replace("\u00a0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    value = re.sub(r"\s+", " ", value).strip().casefold()
    return value


def fred_text_contains_term(text_value, term):
    """
    Boundary-aware Freddie matcher.

    This deliberately prevents short catalogue values from matching inside
    unrelated words (for example `Git` inside `digital`). It also protects
    short terms such as AI, R, C, QA, BI, and UI from substring false positives.
    """
    text = fred_normalize_match_surface(text_value)
    term_text = fred_normalize_match_surface(term).strip()

    if not text or not term_text:
        return False

    # Permit flexible whitespace without changing the actual characters in the
    # catalogue/source term. Boundaries are based on alphanumerics rather than
    # Python's \b so values such as C++, C#, .NET, J2EE, and CI/CD remain usable.
    pieces = [re.escape(piece) for piece in term_text.split() if piece]
    if not pieces:
        return False

    body = r"\s+".join(pieces)
    pattern = rf"(?<![a-z0-9]){body}(?![a-z0-9])"
    return bool(re.search(pattern, text, flags=re.IGNORECASE))


def fred_text_contains_any(text_value, terms):
    return any(
        fred_text_contains_term(text_value, term)
        for term in (terms or [])
        if str(term or "").strip()
    )


def fred_normalize_quote_match_text(value):
    """Normalize harmless punctuation/whitespace for source-quote verification."""
    value = str(value or "")
    value = (
        value.replace("\u00a0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    value = value.casefold()
    value = re.sub(r"[^a-z0-9+#./'-]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def fred_source_quote_found(source_text, quote):
    source = fred_normalize_quote_match_text(source_text)
    quote_norm = fred_normalize_quote_match_text(quote)
    if not quote_norm or len(quote_norm) < 5:
        return False
    return quote_norm in source


def fred_evidence_segments(source_text):
    """Return real source excerpts used only to verify AI evidence citations."""
    text = str(source_text or "").replace("\r\n", "\n").replace("\r", "\n")
    if not text.strip():
        return []

    segments = []
    seen = set()

    def add(value):
        value = re.sub(r"\s+", " ", str(value or "")).strip()
        value = re.sub(r"^[•·▪◦*-]+\s*", "", value).strip()
        if len(value) < 10 or len(value) > 650:
            return
        key = fred_normalize_quote_match_text(value)
        if key and key not in seen:
            seen.add(key)
            segments.append(value)

    lines = [line.strip() for line in text.split("\n") if line.strip()]
    for i, line in enumerate(lines):
        add(line)
        if i + 1 < len(lines):
            add(line + " " + lines[i + 1])
        if i + 2 < len(lines):
            add(line + " " + lines[i + 1] + " " + lines[i + 2])

    return segments


def fred_candidate_evidence_quote_found(candidate_source_text, quote):
    """
    Verify AI evidence metadata without requiring identical punctuation or an
    identical excerpt boundary. This never creates candidate evidence; it only
    confirms that the same substantive concepts co-occur in a real source excerpt.
    """
    if fred_source_quote_found(candidate_source_text, quote):
        return True

    quote_text = str(quote or "").replace("...", " ").strip()
    quote_tokens = fred_signal_tokens(quote_text)
    if not quote_tokens:
        return False

    quote_set = set(quote_tokens)
    identifiers = [
        fred_normalize_quote_match_text(value)
        for value in fred_explicit_identifiers(quote_text)
        if fred_normalize_quote_match_text(value)
    ]

    for segment in fred_evidence_segments(candidate_source_text):
        segment_tokens = set(fred_signal_tokens(segment))
        if not segment_tokens:
            continue
        segment_norm = fred_normalize_quote_match_text(segment)
        if identifiers and any(identifier not in segment_norm for identifier in identifiers):
            continue
        matched = len(quote_set.intersection(segment_tokens))
        if len(quote_set) <= 3:
            required = len(quote_set)
        else:
            required = max(3, int(len(quote_set) * 0.70 + 0.9999))
        if matched >= required:
            return True

    return False


def fred_vndly_summary_evidence_found(candidate_source_text, evidence):
    """
    Validate VNDLY-comment evidence. Resume/clarification citations may use a
    harmlessly shortened excerpt. UI-entered logistics may be rendered naturally,
    but this flexibility is limited to the CANDIDATE SUBMISSION FACTS block.
    """
    if fred_candidate_evidence_quote_found(candidate_source_text, evidence):
        return True

    marker = "CANDIDATE SUBMISSION FACTS:"
    source_text = str(candidate_source_text or "")
    if marker not in source_text:
        return False

    facts_text = source_text.split(marker, 1)[1]
    evidence_tokens = fred_signal_tokens(evidence)
    fact_tokens = set(fred_signal_tokens(facts_text))
    if not evidence_tokens or not fact_tokens:
        return False
    return all(token in fact_tokens for token in evidence_tokens)


def fred_numeric_years(value):
    match = re.search(r"(\d+(?:\.\d+)?)", str(value or ""))
    if not match:
        return None
    try:
        return float(match.group(1))
    except Exception:
        return None


def fred_should_suppress_total_years(calculated_total_experience, vndly_context):
    """Presentation-only guard against unnecessarily advertising overqualification."""
    vndly_context = vndly_context or {}
    total_years = fred_numeric_years(calculated_total_experience)
    max_years = fred_numeric_years(vndly_context.get("max_experience", ""))
    if total_years is None or max_years is None or max_years <= 0:
        return False
    return total_years >= max_years + 8 or (max_years <= 10 and total_years >= max_years * 2)


def fred_suppress_total_career_years(value, calculated_total_experience, vndly_context):
    """
    Remove only the deterministic total-career number when it is strategically
    counterproductive. The replacement consumes modifiers such as "over" so it
    cannot create broken phrases like "over extensive experience".
    """
    text = str(value or "")
    if not text or not fred_should_suppress_total_years(calculated_total_experience, vndly_context):
        return text

    total_years = fred_numeric_years(calculated_total_experience)
    if total_years is None:
        return text
    year_token = str(int(total_years)) if float(total_years).is_integer() else str(total_years).rstrip("0").rstrip(".")
    modifier = r"(?:(?:over|more\s+than|approximately|about|nearly)\s+)?"

    # e.g. "over 30+ years of professional experience" -> "extensive professional experience"
    text = re.sub(
        rf"\b{modifier}{re.escape(year_token)}\s*\+?\s*years?\s+of\s+",
        "extensive ",
        text,
        flags=re.IGNORECASE,
    )
    # e.g. "30+ years professional experience" -> "extensive professional experience"
    text = re.sub(
        rf"\b{modifier}{re.escape(year_token)}\s*\+?\s*years?\s+(?=(?:professional|overall|total)\s+experience\b)",
        "extensive ",
        text,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s+", " ", text).strip()


def fred_grounding_error_is_hard(error_text):
    """Only genuine fabrication/scope/exclusion issues are blocking safety issues."""
    value = str(error_text or "").casefold()
    hard_markers = [
        "unsupported high-impact claim",
        "unsupported named identifier",
        "unsupported identifier",
        "manager-explicitly-excluded",
        "uses manager-explicitly-excluded",
        "emphasizes manager-explicitly-excluded",
    ]
    return any(marker in value for marker in hard_markers)


def fred_drop_hard_grounding_items(data, hard_errors):
    """
    Last-resort safety action that drops complete unsafe sentences/details/rows.
    It never splices resume fragments into generated prose.
    """
    result = copy.deepcopy(data) if isinstance(data, dict) else {}
    sentences = result.get("SUMMARY_SENTENCES", [])
    skills = result.get("SKILLS", [])
    sentences = sentences if isinstance(sentences, list) else []
    skills = skills if isinstance(skills, list) else []

    summary_drop = set()
    row_drop = set()
    detail_drop = {}

    for error in hard_errors or []:
        sm = re.search(r"SUMMARY sentence\s+(\d+)", str(error), flags=re.IGNORECASE)
        if sm:
            summary_drop.add(int(sm.group(1)) - 1)
            continue

        dm = re.search(r"SKILL row\s+(\d+)\s+detail\s+'([^']+)'", str(error), flags=re.IGNORECASE)
        if dm:
            row_idx = int(dm.group(1)) - 1
            detail_drop.setdefault(row_idx, set()).add(dm.group(2).casefold())
            continue

        rm = re.search(r"SKILL row\s+(\d+)", str(error), flags=re.IGNORECASE)
        if rm:
            row_drop.add(int(rm.group(1)) - 1)

    result["SUMMARY_SENTENCES"] = [
        entry for idx, entry in enumerate(sentences) if idx not in summary_drop
    ]

    safe_skills = []
    for idx, row in enumerate(skills):
        if idx in row_drop or not isinstance(row, dict):
            continue
        row = copy.deepcopy(row)
        drops = detail_drop.get(idx, set())
        if drops:
            details = fred_normalize_audited_detail_records(row.get("details", []))
            row["details"] = [d for d in details if str(d.get("item", "")).casefold() not in drops]
        if row.get("details", []):
            safe_skills.append(row)
    result["SKILLS"] = safe_skills
    return result


def fred_string_list(value):
    if isinstance(value, list):
        return [
            re.sub(r"\s+", " ", str(item or "")).strip()
            for item in value
            if str(item or "").strip()
        ]
    if value is None:
        return []
    cleaned = re.sub(r"\s+", " ", str(value)).strip()
    return [cleaned] if cleaned else []


def fred_clean_candidate_clarifications(value):
    """
    Keep candidate-confirmed facts while removing obvious recruiter question text.

    The optional Freddie clarification box is a candidate-evidence source. Users may
    paste a Q&A exchange, so question wording must never become evidence merely
    because it names a tool, responsibility, or scope the candidate was asked about.
    """
    raw = str(value or "").replace("\r\n", "\n").replace("\r", "\n")
    if not raw.strip():
        return ""

    cleaned_lines = []
    question_start = re.compile(
        r"^\s*(?:q(?:uestion)?\s*\d*\s*[:.)-]?\s*|\d+\s*[.)-]\s*)?"
        r"(?:have|has|do|does|did|can|could|are|is|was|were|will|would|what|when|where|why|how)\b",
        flags=re.IGNORECASE,
    )

    for raw_line in raw.split("\n"):
        line = raw_line.strip()
        if not line:
            if cleaned_lines and cleaned_lines[-1] != "":
                cleaned_lines.append("")
            continue

        # Drop obvious interviewer/recruiter questions. Candidate answers such as
        # "Yes. I shared ownership..." remain untouched.
        if "?" in line and question_start.search(line):
            continue
        if re.match(r"^\s*(?:q(?:uestion)?\s*\d*\s*[:.)-])", line, flags=re.IGNORECASE):
            continue

        cleaned_lines.append(line)

    # Collapse excessive blank lines while preserving readable item separation.
    result = "\n".join(cleaned_lines).strip()
    result = re.sub(r"\n{3,}", "\n\n", result)
    return result


def fred_candidate_evidence_text(resume_text, candidate_clarifications=""):
    """
    Build the Freddie candidate-truth surface used by grounding checks.

    Resume remains the primary source. Optional candidate-confirmed follow-up
    answers may add facts the original resume omitted, but role/JD sources never
    become candidate evidence.
    """
    resume = str(resume_text or "").strip()
    clarifications = fred_clean_candidate_clarifications(candidate_clarifications)
    if not clarifications:
        return resume
    return (
        resume
        + "\n\nCANDIDATE-CONFIRMED CLARIFICATIONS / FOLLOW-UP ANSWERS:\n"
        + clarifications
    ).strip()


def fred_clean_generated_text(value):
    """Freddie output style cleanup, including the user's no-em-dash preference."""
    cleaned = re.sub(r"\s+", " ", str(value or "")).strip()
    # Never emit em dashes in generated recruiter-facing writing.
    cleaned = cleaned.replace("\u2014", ", ")
    cleaned = re.sub(r"\s+,\s+", ", ", cleaned)
    return cleaned.strip()


def fred_signal_tokens(value):
    """Conservative lexical tokens for Freddie candidate-evidence checks."""
    stop_words = {
        "a", "an", "and", "as", "at", "by", "for", "from", "in", "into",
        "of", "on", "or", "the", "to", "using", "with", "within", "via",
        "support", "supports", "supported", "experience", "experienced",
        "skill", "skills", "work", "working", "professional", "professionals",
        "provide", "provides", "provided", "providing", "function", "functions",
        "ability", "able", "help", "helps", "helping",
        # Generic framing nouns should not require literal resume wording when the
        # substantive domain/competency itself is proven. For example,
        # "Multifamily Business Context" is grounded by Multifamily + business
        # experience; the word "context" is only a label wrapper.
        "context", "industry", "domain",
    }

    derivational_roots = {
        "management": "manag",
        "manager": "manag",
        "managed": "manag",
        "managing": "manag",
        "analysis": "analy",
        "analyst": "analy",
        "analyze": "analy",
        "analyzed": "analy",
        "analytical": "analy",
        "requirements": "require",
        "requirement": "require",
        "reporting": "report",
        "reports": "report",
        "reported": "report",
        "communications": "communicat",
        "communication": "communicat",
        "communicating": "communicat",
        "presentation": "present",
        "presentations": "present",
        "presenting": "present",
        "automation": "automat",
        "automated": "automat",
        "automating": "automat",
        "governance": "govern",
        "validation": "valid",
        "validate": "valid",
        "validated": "valid",
        "testing": "test",
        "tested": "test",
        "planning": "plan",
        "prioritization": "priorit",
        "prioritize": "priorit",
        "prioritized": "priorit",
        "stakeholders": "stakeholder",
        "independent": "independ",
        "independently": "independ",
        "identification": "identif",
        "identify": "identif",
        "identifies": "identif",
        "identified": "identif",
        "identifying": "identif",
        "strategic": "strateg",
        "strategy": "strateg",
        "alignment": "align",
        "align": "align",
        "aligned": "align",
        "aligning": "align",
        "technical": "tech",
        "technology": "tech",
        "technologies": "tech",
        "translate": "translat",
        "translated": "translat",
        "translating": "translat",
        "translation": "translat",
        "transformation": "transform",
        "transform": "transform",
        "transformed": "transform",
        "transforming": "transform",
        # Common competency-label noun/verb pairs. These normalize natural
        # recruiter labels to the same concept as resume action verbs without
        # granting support for a different skill or named technology.
        "facilitation": "facilitat",
        "facilitate": "facilitat",
        "facilitated": "facilitat",
        "facilitating": "facilitat",
        "integration": "integrat",
        "integrations": "integrat",
        "integrate": "integrat",
        "integrated": "integrat",
        "integrating": "integrat",
        # Coordination, collaboration, and partnership are equivalent umbrella
        # concepts for competency-label grounding. Concrete role nouns/tools still
        # require their own evidence, and authority claims remain separately strict.
        "coordination": "collab",
        "coordinator": "collab",
        "coordinate": "collab",
        "coordinated": "collab",
        "coordinating": "collab",
        "collaboration": "collab",
        "collaborate": "collab",
        "collaborated": "collab",
        "collaborating": "collab",
        "partnership": "collab",
        "partnerships": "collab",
        "partner": "collab",
        "partners": "collab",
        "partnered": "collab",
        "partnering": "collab",
        "implementation": "implement",
        "implement": "implement",
        "implemented": "implement",
        "implementing": "implement",
        "delivery": "deliver",
        "deliver": "deliver",
        "delivered": "deliver",
        "delivering": "deliver",
        "documentation": "document",
        "document": "document",
        "documented": "document",
        "documenting": "document",
        "recommendation": "recommend",
        "recommendations": "recommend",
        "recommend": "recommend",
        "recommended": "recommend",
        "recommending": "recommend",
        "escalation": "escalat",
        "escalations": "escalat",
        "escalate": "escalat",
        "escalated": "escalat",
        "escalating": "escalat",
    }

    # Hyphens are treated as word separators here so natural umbrella labels
    # such as "Business-to-Technology Translation" can be compared against
    # resume wording like "translating business needs ... technical solutions."
    # Slash/plus/hash/dot remain available for concrete tokens such as AI/ML,
    # C++, C#, and .NET-style identifiers.
    raw_tokens = re.findall(
        r"[A-Za-z0-9]+(?:[./+#][A-Za-z0-9]+)*",
        str(value or ""),
    )
    result = []

    for token in raw_tokens:
        lowered = token.casefold()
        if lowered.endswith("-level") and len(lowered) > len("-level"):
            lowered = lowered[:-len("-level")]
        if lowered in stop_words:
            continue
        normalized = derivational_roots.get(lowered, lowered)
        if normalized not in derivational_roots.values() and len(normalized) > 4:
            if normalized.endswith("ies") and len(normalized) > 5:
                normalized = normalized[:-3] + "y"
                normalized = derivational_roots.get(normalized, normalized)
            else:
                for suffix in ("ing", "ed", "es", "s"):
                    if normalized.endswith(suffix) and len(normalized) - len(suffix) >= 4:
                        normalized = normalized[:-len(suffix)]
                        # Re-apply the semantic root map after plural/tense stemming
                        # (for example, automations -> automation -> automat).
                        normalized = derivational_roots.get(normalized, normalized)
                        break
        if normalized and normalized not in result:
            result.append(normalized)

    return result


def fred_claim_supported_by_evidence(claim, evidence):
    """Require a clear majority of substantive claim tokens in cited evidence."""
    claim_tokens = fred_signal_tokens(claim)
    evidence_tokens = set(fred_signal_tokens(evidence))
    if not claim_tokens:
        return False
    matched = sum(1 for token in claim_tokens if token in evidence_tokens)
    if len(claim_tokens) == 1:
        return matched == 1
    required = max(1, int((len(claim_tokens) * 0.60) + 0.9999))
    return matched >= required


def fred_claim_supported_by_candidate_context(
    claim,
    cited_evidence,
    full_candidate_evidence,
):
    """
    Validate natural skill labels/details without demanding every umbrella word
    appear inside one short citation.

    The cited excerpt must still anchor the claim. The broader candidate evidence
    may then supply closely related context from another verified resume/clarification
    excerpt. This is intentionally used for generic competency labels/details only;
    concrete identifiers and high-impact authority claims remain separately strict.
    """
    if fred_claim_supported_by_evidence(claim, cited_evidence):
        return True

    claim_tokens = fred_signal_tokens(claim)
    cited_tokens = set(fred_signal_tokens(cited_evidence))
    full_tokens = set(fred_signal_tokens(full_candidate_evidence))

    if not claim_tokens or not cited_tokens or not full_tokens:
        return False

    # At least one substantive claim concept must be present in the exact cited
    # evidence so an unrelated quote cannot validate a row/detail.
    anchored = [token for token in claim_tokens if token in cited_tokens]
    if not anchored:
        return False

    matched_full = sum(1 for token in claim_tokens if token in full_tokens)
    if len(claim_tokens) == 1:
        return matched_full == 1

    # Umbrella labels may synthesize closely related verified facts across the
    # candidate record, but still require a strong majority of their concepts.
    required = max(1, int((len(claim_tokens) * 0.67) + 0.9999))
    return matched_full >= required


def fred_explicit_identifiers(value):
    """Concrete acronyms/products/framework identifiers that should not be invented."""
    text = str(value or "")
    identifiers = []

    for token in re.findall(r"\b[A-Z][A-Z0-9./+#-]{1,}\b", text):
        if token not in identifiers:
            identifiers.append(token)

    for token in re.findall(r"\b[A-Z][a-z]+[A-Z][A-Za-z0-9]*\b", text):
        if token not in identifiers:
            identifiers.append(token)

    return identifiers


def fred_explicit_identifiers_supported(value, evidence_text, resume_text):
    evidence_norm = fred_normalize_quote_match_text(evidence_text)
    resume_norm = fred_normalize_quote_match_text(resume_text)
    unsupported = []

    for identifier in fred_explicit_identifiers(value):
        ident = fred_normalize_quote_match_text(identifier)
        if ident and ident not in evidence_norm and ident not in resume_norm:
            unsupported.append(identifier)

    return unsupported


def fred_unsupported_high_impact_claims(output_text, evidence_text):
    """
    Protect claim strength/authority across Freddie job families.

    A supported noun phrase does not make a stronger modifier true. This is
    intentionally generic across technical, business, operations, design,
    finance, legal, HR, and program-management requisitions.
    """
    output = str(output_text or "")
    evidence = str(evidence_text or "")

    rules = [
        ("independent/independently", r"\bindependent(?:ly)?\b", r"\bindependent(?:ly)?\b"),
        ("lead/led/leading", r"\b(?:lead|leads|leading|led)\b", r"\b(?:lead|leads|leading|led)\b"),
        ("ownership/owned", r"\b(?:own|owns|owned|owning|ownership)\b", r"\b(?:own|owns|owned|owning|ownership)\b"),
        ("manage/managed/managing/manager", r"\b(?:manage|manages|managed|managing|manager)\b", r"\b(?:manage|manages|managed|managing|manager)\b"),
        ("oversight/oversee", r"\b(?:oversee|oversees|oversaw|overseeing|oversight)\b", r"\b(?:oversee|oversees|oversaw|overseeing|oversight)\b"),
        ("architect/architected", r"\b(?:architect|architects|architected|architecting)\b", r"\b(?:architect|architects|architected|architecting)\b"),
        ("expert/expertise", r"\b(?:expert|expertise)\b", r"\b(?:expert|expertise)\b"),
        ("advanced", r"\badvanced\b", r"\badvanced\b"),
        ("hands-on", r"\bhands?[\s-]+on\b", r"\bhands?[\s-]+on\b"),
        ("enterprise-wide", r"\benterprise[\s-]+wide\b", r"\benterprise[\s-]+wide\b"),
        ("organization-wide", r"\b(?:organization|organisation)[\s-]+wide\b", r"\b(?:organization|organisation)[\s-]+wide\b"),
        ("global/globally", r"\bglobal(?:ly)?\b", r"\bglobal(?:ly)?\b"),
        ("end-to-end", r"\bend[\s-]+to[\s-]+end\b", r"\bend[\s-]+to[\s-]+end\b"),
        ("high-volume", r"\bhigh[\s-]+volume\b", r"\bhigh[\s-]+volume\b"),
        ("large-scale", r"\blarge[\s-]+scale\b", r"\blarge[\s-]+scale\b"),
        ("production", r"\bproduction\b", r"\bproduction\b"),
    ]

    unsupported = []
    for label, output_pattern, evidence_pattern in rules:
        if re.search(output_pattern, output, flags=re.IGNORECASE) and not re.search(
            evidence_pattern, evidence, flags=re.IGNORECASE
        ):
            unsupported.append(label)

    return unsupported


def fred_requirement_phrase_hits(output_text, requirement_items):
    """Return meaningful requirement phrases reused in an output string."""
    output_norm = fred_normalize_quote_match_text(output_text)
    hits = []
    if not output_norm:
        return hits

    output_tokens = fred_signal_tokens(output_text)
    for item in requirement_items or []:
        item_text = str(item or "").strip()
        if not fred_normalize_quote_match_text(item_text):
            continue
        tokens = [token for token in fred_signal_tokens(item_text) if len(token) >= 3]
        for size in (4, 3, 2):
            if len(tokens) < size:
                continue
            found = False
            for idx in range(len(tokens) - size + 1):
                phrase_tokens = tokens[idx:idx + size]
                phrase = " ".join(phrase_tokens)
                for out_idx in range(len(output_tokens) - size + 1):
                    if output_tokens[out_idx:out_idx + size] == phrase_tokens:
                        if phrase not in hits:
                            hits.append(phrase)
                        found = True
                        break
                if found:
                    break
            if found:
                break
    return hits[:8]


def fred_output_hits_excluded_requirement(output_text, structured_req):
    """Block only true negative/disqualifying manager guidance."""
    structured_req = fred_coerce_structured_req(structured_req)
    excluded = structured_req.get("explicitly_excluded_requirements", []) or []
    return fred_requirement_phrase_hits(output_text, excluded)


def fred_output_hits_deemphasized_requirement(output_text, structured_req):
    """
    Identify lower-priority concepts without treating them as prohibited.

    v11 intentionally separates 'the team does not use X / X is not central'
    from 'do not submit candidates with X / X is disqualifying'.
    """
    structured_req = fred_coerce_structured_req(structured_req)
    items = structured_req.get("deemphasized_requirements", []) or []
    # Backward compatibility for cached v10 analysis objects.
    if not items:
        items = structured_req.get("deemphasized_or_negated_requirements", []) or []
    return fred_requirement_phrase_hits(output_text, items)


def fred_split_sentences(text):
    """Lightweight sentence splitter for Freddie audit checks."""
    text = re.sub(r"\s+", " ", str(text or "")).strip()
    if not text:
        return []
    parts = re.split(r"(?<=[.!?])\s+(?=[A-Z0-9])", text)
    return [part.strip() for part in parts if part.strip()]


def fred_role_only_phrase_hits(output_text, role_text, resume_text, evidence_text=""):
    """
    Find meaningful multi-word phrases copied from the role sources that are not
    present in the candidate resume and are not independently supported by the
    cited candidate evidence.
    """
    output_words = fred_normalize_quote_match_text(output_text).split()
    role_norm = fred_normalize_quote_match_text(role_text)
    resume_norm = fred_normalize_quote_match_text(resume_text)
    evidence = str(evidence_text or "")

    stop_words = {
        "a", "an", "and", "as", "at", "be", "by", "for", "from", "in",
        "is", "of", "on", "or", "the", "to", "into", "with", "this", "that",
        "their", "his", "her", "candidate", "role", "team", "freddie", "mac",
        "supplier", "call", "hiring", "manager", "emphasized", "highlighted",
        "discussed", "noted", "across", "multiple",
    }

    hits = []
    seen = set()

    for size in (5, 4, 3, 2):
        if len(output_words) < size:
            continue
        for idx in range(len(output_words) - size + 1):
            words = output_words[idx:idx + size]
            phrase = " ".join(words)
            if phrase in seen:
                continue
            signals = [w for w in words if w not in stop_words and len(w) >= 4]
            if len(signals) < 2:
                continue
            if words[0] in stop_words or words[-1] in stop_words:
                continue
            if phrase in role_norm and phrase not in resume_norm:
                if evidence and fred_claim_supported_by_candidate_context(
                    phrase, evidence, resume_text
                ):
                    continue
                if any(phrase in existing for existing in hits):
                    continue
                hits.append(phrase)
                seen.add(phrase)

    return hits[:10]


def fred_role_only_signal_tokens(output_text, role_text, resume_text, evidence_text=""):
    """
    Return distinctive candidate-facing words that appear in Freddie role sources
    but not in the resume/evidence. Useful for catching subtle requirement leakage
    such as adding "executive" to otherwise supported reporting experience.
    """
    output_tokens = fred_signal_tokens(output_text)
    role_tokens = set(fred_signal_tokens(role_text))
    resume_tokens = set(fred_signal_tokens(resume_text))
    evidence_tokens = set(fred_signal_tokens(evidence_text))

    generic = {
        "freddie", "candidate", "team", "role", "current", "manager",
        "require", "preferred", "priority", "business", "technical",
        "professional", "experience", "work", "support", "project",
        "supplier", "hiring", "emphasized", "highlighted", "discussed", "noted",
    }

    hits = []
    for token in output_tokens:
        if token in generic or len(token) < 6:
            continue
        if token in role_tokens and token not in resume_tokens and token not in evidence_tokens:
            if token not in hits:
                hits.append(token)

    return hits[:10]


def fred_catalog_lookup(catalog):
    """Build a normalized-name -> exact VNDLY record lookup."""
    lookup = {}

    for item in catalog or []:
        if not isinstance(item, dict):
            continue

        name = str(item.get("skill_name", "")).strip()
        if not name:
            continue

        lookup[fred_normalize_skill_name(name)] = item

    return lookup


def fred_sanitize_recommended_skills(raw_items, catalog):
    """
    Force Gemini recommendations back onto exact catalogue values.
    """
    if not isinstance(raw_items, list):
        return []

    lookup = fred_catalog_lookup(catalog)
    cleaned = []
    seen = set()

    for item in raw_items:
        if isinstance(item, dict):
            proposed = (
                item.get("skill_name")
                or item.get("value")
                or item.get("label")
                or ""
            )
            evidence = str(item.get("evidence", "") or "").strip()
            reason = str(item.get("reason", "") or "").strip()
        else:
            proposed = str(item)
            evidence = ""
            reason = ""

        exact = lookup.get(fred_normalize_skill_name(proposed))

        if not exact:
            continue

        key = fred_normalize_skill_name(exact["skill_name"])
        if key in seen:
            continue

        seen.add(key)

        cleaned.append(
            {
                "id": exact.get("id"),
                "skill_name": exact["skill_name"],
                "catalog_tier": exact.get(
                    "catalog_tier",
                    "canonical_preferred",
                ),
                "evidence": evidence,
                "reason": reason,
            }
        )

    return cleaned


def fred_dedupe_skill_groups(*groups):
    """Remove duplicates while preserving priority order."""
    seen = set()
    output = []

    for group in groups:
        clean_group = []

        for item in group or []:
            name = str(item.get("skill_name", "")).strip()
            key = fred_normalize_skill_name(name)

            if not name or key in seen:
                continue

            seen.add(key)
            clean_group.append(item)

        output.append(clean_group)

    return output


def fred_resolve_structured_vndly_skills(raw_field, catalog):
    """
    Resolve a raw Google-Sheet Must/Nice field back to the exact VNDLY skill
    values that produced it.

    The sheet stores multiple selected skills as one comma-separated string, but
    some individual VNDLY skills also contain commas.  This uses the known exact
    catalogue and a longest-match dynamic-programming pass so we do not mistake
    a JD keyword such as "Java" for an exact requisition-selected skill.
    """
    raw = re.sub(r"\s+", " ", str(raw_field or "").replace("\u00a0", " ")).strip()

    if not raw:
        return []

    candidates = []

    for item in catalog or []:
        if not isinstance(item, dict):
            continue

        name = re.sub(
            r"\s+",
            " ",
            str(item.get("skill_name", "") or "").replace("\u00a0", " "),
        ).strip()

        if name:
            candidates.append((name, item))

    # Prefer the longest exact catalogue value at any ambiguous position.
    candidates.sort(key=lambda pair: len(pair[0]), reverse=True)

    raw_fold = raw.casefold()
    memo = {}

    def skip_separator(pos):
        while pos < len(raw) and raw[pos].isspace():
            pos += 1

        if pos < len(raw) and raw[pos] in {",", ";", "|"}:
            pos += 1
            while pos < len(raw) and raw[pos].isspace():
                pos += 1

        return pos

    def solve(pos):
        while pos < len(raw) and raw[pos].isspace():
            pos += 1

        if pos >= len(raw):
            return []

        if pos in memo:
            return memo[pos]

        for name, item in candidates:
            name_fold = name.casefold()

            if not raw_fold.startswith(name_fold, pos):
                continue

            end = pos + len(name)

            # Exact skill must end at the raw-field end or at a separator.
            check = end
            while check < len(raw) and raw[check].isspace():
                check += 1

            if check < len(raw) and raw[check] not in {",", ";", "|"}:
                continue

            next_pos = skip_separator(check)

            remainder = solve(next_pos)

            if remainder is not None:
                result = [item] + remainder
                memo[pos] = result
                return result

        memo[pos] = None
        return None

    resolved = solve(0)

    if resolved is not None:
        return resolved

    # Conservative fallback: only accept a whole-field exact catalogue match.
    whole = fred_catalog_lookup(catalog).get(
        fred_normalize_skill_name(raw)
    )

    return [whole] if whole else []


def fred_filter_exact_req_recommendations(raw_items, allowed_items, catalog):
    """
    Validate Gemini output against the deterministic set of skills actually
    selected on this requisition's structured Must/Nice field.
    """
    sanitized = fred_sanitize_recommended_skills(raw_items, catalog)

    allowed = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in (allowed_items or [])
        if str(item.get("skill_name", "")).strip()
    }

    if not allowed:
        return []

    return [
        item
        for item in sanitized
        if fred_normalize_skill_name(item.get("skill_name", "")) in allowed
    ]


FREDDIE_VNDLY_SKILL_POOL_LIMIT = 40
# Freddie/VNDLY permits selecting more than eight manual skills.
# v11 therefore has NO final hard cap; the pool limit only prevents runaway AI output.


def fred_extract_work_history_text(raw_resume_text):
    """
    Return the portion of the resume most likely to contain dated work history.

    Heading detection is line-anchored so a summary sentence such as
    "20+ years of professional experience" cannot accidentally make the entire
    resume (including a top-level skills inventory) count as dated work-history
    evidence.
    """
    text = str(raw_resume_text or "")
    if not text:
        return ""

    heading_pattern = re.compile(
        r"(?im)^\s*(?:PROFESSIONAL EXPERIENCE|WORK EXPERIENCE|WORK HISTORY|EMPLOYMENT HISTORY|EXPERIENCE)\s*:?[ \t]*$"
    )
    match = heading_pattern.search(text)

    if not match:
        # Full-resume fallback remains available separately at evidence strength
        # 1. Do not inflate a skills-list mention to dated-history strength 2.
        return ""

    work = text[match.end():]

    end_pattern = re.compile(
        r"(?im)^\s*(?:EDUCATION|ACADEMIC(?: BACKGROUND)?|CERTIFICATIONS?|TRAINING)\s*:?[ \t]*$"
    )
    end_match = end_pattern.search(work)
    if end_match:
        work = work[:end_match.start()]

    return work


def fred_skill_work_history_evidence(skill_name, raw_resume_text):
    """
    Conservative deterministic check for whether a VNDLY skill is evidenced in
    the dated work-history portion of the resume rather than only in a summary
    or top-level skills inventory.

    Returns:
      2 = clear dated work-history evidence
      1 = supported elsewhere in the resume / broader semantic domain evidence
      0 = no meaningful candidate evidence found
    """
    full_resume = str(raw_resume_text or "")
    full_fold = full_resume.casefold()
    work = fred_extract_work_history_text(full_resume)
    work_fold = work.casefold()
    skill = str(skill_name or "").strip()
    key = fred_normalize_skill_name(skill)

    if not skill:
        return 0

    alias_map = {
        "java": ["java"],
        "j2ee": ["j2ee", "java ee", "jee"],
        "java database connectivity (jdbc)": ["jdbc", "java database connectivity"],
        "java servlets": ["servlet", "servlets"],
        "javaserver pages (jsp)": ["jsp", "jsps", "javaserver pages"],
        "spring framework": ["spring framework", "spring "],
        "spring boot": ["spring boot"],
        "relational database": [
            "relational database", "oracle", "db2", "sql server", "mysql",
            "postgresql", "sybase", "jdbc", "sql queries", "sql statements"
        ],
        "relational databases": [
            "relational database", "oracle", "db2", "sql server", "mysql",
            "postgresql", "sybase", "jdbc", "sql queries", "sql statements"
        ],
        "junit testing": ["junit"],
        "mockito": ["mockito"],
        "git": ["git"],
        "jenkins (software)": ["jenkins"],
        "docker (software)": ["docker"],
        "maven": ["maven"],
        "gradle": ["gradle"],
        "eclipse ide": ["eclipse"],
        "python": ["python"],
        "sql": ["sql"],
        "snowflake": ["snowflake"],
        "json": ["json"],
        "opencv": ["opencv", "open cv"],
        "pillow": ["pillow"],
        "pytorch": ["pytorch", "py torch"],
        "tensorflow": ["tensorflow", "tensor flow"],
        "data engineering": [
            "data engineer", "data engineering", "data pipeline", "etl", "elt",
            "pyspark", "spark", "aws glue"
        ],
        "data quality": [
            "data quality", "quality rule", "quality check", "reconciliation",
            "data validation"
        ],
        "data validation": [
            "data validation", "validate data", "validated data", "validation",
            "reconciliation"
        ],
        "automation": ["automation", "automated", "automate"],
        "automation testing": [
            "automation testing", "automated testing", "test automation",
            "automation test", "automated test"
        ],
        "test automation": [
            "test automation", "automation testing", "automated testing",
            "automation test", "automated test"
        ],
        "test automation tools": [
            "test automation", "automation testing", "automated testing",
            "automation framework", "testing framework"
        ],
        "data modeling": [
            "data modeling", "data model", "logical model", "physical model",
            "dimensional model", "entity relationship", "erwin", "er/studio",
            "er studio", "powerdesigner"
        ],
        "ai/ml": [
            "ai/ml", "artificial intelligence", "machine learning",
            "microsoft copilot", "github copilot", "chatgpt", "generative ai",
            "genai"
        ],
        "jira": ["jira"],
        "sharepoint": ["sharepoint"],
        "power bi": ["power bi"],
        "power apps": ["power apps"],
        "power automate": ["power automate"],
        "power platform": ["power platform", "power apps", "power automate", "power bi"],
        "agile delivery": ["agile", "scrum", "sprint planning", "backlog", "pi planning"],
        "agile": ["agile", "scrum", "sprint planning"],
        "product management": ["product manager", "product management", "product owner", "product roadmap", "product roadmaps"],
        "program management": ["program manager", "program management", "program delivery", "program planning", "portfolio management", "program portfolio"],
        "enterprise program management": ["program manager", "program management", "program delivery", "program planning", "portfolio management", "program portfolio", "multiple projects", "concurrent projects", "concurrent initiatives"],
        "project management": ["project manager", "project management", "project delivery", "project planning", "managing specific projects", "multiple projects"],
        "business analysis": ["business analyst", "business analysis", "business requirements", "requirements elicitation", "requirements analysis", "user requirements", "business needs", "translating complex business needs", "translate business needs"],
        "requirements analysis": ["requirements analysis", "requirements elicitation", "business requirements", "technical requirements", "requirements session", "requirements sessions"],
        "stakeholder management": ["stakeholder management", "stakeholders", "business stakeholders"],
        "executive reporting": ["executive reporting", "executive-level reporting", "executive and program-level reporting", "program-level reporting", "executive dashboard", "executive dashboards", "executive status", "status reports", "leadership reporting", "leadership presentation", "leadership presentations"],
        "risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks"],
        "business risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks", "risks and issues"],
        "user story": ["user story", "user stories", "epics", "story writing"],
        "user stories": ["user story", "user stories", "epics", "story writing"],
        "devops": ["devops", "devsecops"],
        "quality assurance": ["quality assurance", "qa engineer", "qa testing"],
        "software testing": [
            "software testing", "functional testing", "regression testing",
            "integration testing", "system testing", "uat"
        ],
        "finance": [
            "bank", "banking", "financial", "investment", "portfolio",
            "securities", "trading", "trade", "retirement",
            "asset", "capital markets"
        ],
        "financial services": [
            "bank", "banking", "financial", "investment", "portfolio",
            "securities", "trading", "trade", "retirement",
            "asset", "capital markets"
        ],
        "mortgage industry": ["mortgage", "appraisal", "loan", "collateral", "uad", "underwriting"],
    }

    aliases = alias_map.get(key)

    if aliases:
        if fred_text_contains_any(work, aliases):
            return 2
        if fred_text_contains_any(full_resume, aliases):
            # Supported elsewhere in the resume (for example a top-level
            # technical-skills section), but not proven as strongly as dated
            # work-history evidence.
            return 1
        return 0

    # Generic fallback: strip catalogue parentheticals and test the meaningful
    # core phrase. Dated work-history evidence is stronger than a top-level
    # skills-list mention, but both are valid candidate evidence.
    core = re.sub(r"\([^)]*\)", "", skill).strip().casefold()

    if core and len(core) >= 2 and fred_text_contains_term(work, core):
        return 2

    if core and len(core) >= 2 and fred_text_contains_term(full_resume, core):
        return 1

    # Domain-like catalogue values can be supported semantically even if the
    # exact catalogue phrase is not present.
    domain_terms = {
        "mortgage": ["mortgage", "appraisal", "loan", "uad"],
        "capital markets": ["capital markets", "securities", "trading", "trade "],
        "banking": ["bank", "banking"],
        "insurance": ["insurance"],
    }

    for domain_key, terms in domain_terms.items():
        if domain_key in key and fred_text_contains_any(work, terms):
            return 1

    return 0


def fred_find_skill_evidence_excerpt(skill_name, raw_resume_text):
    """
    Return a real candidate-resume line that supports a VNDLY catalogue skill.

    This is used for deterministic auditing/ranking; it never fabricates an
    evidence sentence. Work-history lines are preferred over top-level skills.
    """
    full_resume = str(raw_resume_text or "")
    work = fred_extract_work_history_text(full_resume)
    aliases = fred_skill_aliases_for_requirement(skill_name)

    for source in (work, full_resume):
        for line in str(source).splitlines():
            clean = re.sub(r"\s+", " ", line).strip()
            if clean and fred_text_contains_any(clean, aliases):
                return clean

    # Generic whole-label fallback for catalogue values not covered by an alias.
    core = re.sub(r"\([^)]*\)", "", str(skill_name or "")).strip()
    if core:
        for source in (work, full_resume):
            for line in str(source).splitlines():
                clean = re.sub(r"\s+", " ", line).strip()
                if clean and fred_text_contains_term(clean, core):
                    return clean

    return ""


def fred_skill_priority_category(skill_name, structured_req):
    """Classify a supported VNDLY skill without confusing priority with prohibition."""
    structured_req = fred_coerce_structured_req(structured_req)

    excluded_text = " ".join(
        str(x)
        for x in (structured_req.get("explicitly_excluded_requirements", []) or [])
    )
    if excluded_text and fred_skill_matches_text(skill_name, excluded_text):
        return ""

    deemphasized_text = " ".join(
        str(x)
        for x in (structured_req.get("deemphasized_requirements", []) or [])
    )
    if deemphasized_text and fred_skill_matches_text(skill_name, deemphasized_text):
        return "jd_supported_coverage"

    explicit_must_text = " ".join(
        str(x)
        for x in (structured_req.get("explicit_must_have_requirements", []) or [])
    )
    preferred_text = " ".join(
        str(x)
        for x in (structured_req.get("preferred_requirements", []) or [])
    )

    # Explicit "nice to have / preferred / helpful" language wins over a broad
    # responsibility-section mention unless the same skill is also explicitly
    # labeled Must Have.
    if (
        fred_skill_matches_text(skill_name, preferred_text)
        and not fred_skill_matches_text(skill_name, explicit_must_text)
    ):
        return "preferred_job_intelligence"

    if fred_skill_matches_text(skill_name, explicit_must_text):
        return "required_job_intelligence"

    manager_optional_lines = [
        str(x)
        for x in (structured_req.get("role_clarifications", []) or [])
        if fred_skill_matches_text(skill_name, str(x))
        and re.search(
            r"(?i)\b(?:nice to have|helpful|a plus|preferred|optional|not required|helps?)\b",
            str(x),
        )
    ]
    if manager_optional_lines:
        return "preferred_job_intelligence"

    manager_text = " ".join(
        str(x)
        for key in [
            "current_manager_priorities",
            "selection_feedback",
            "role_clarifications",
        ]
        for x in (structured_req.get(key, []) or [])
    )

    if fred_skill_matches_text(skill_name, manager_text):
        return "manager_priority"

    if fred_skill_is_formal_must_have(skill_name, structured_req):
        return "required_job_intelligence"

    if fred_skill_matches_text(skill_name, preferred_text):
        return "preferred_job_intelligence"

    required_text = " ".join(
        str(x)
        for x in (structured_req.get("required_requirements", []) or [])
    )
    if required_text and fred_skill_matches_text(skill_name, required_text):
        return "required_job_intelligence"

    return "additional_high_value"


def fred_supplement_priority_skills(
    ranked_skills,
    structured_req,
    catalog,
    raw_resume_text,
    selected_must_items=None,
    selected_nice_items=None,
    vndly_context=None,
):
    """
    Deterministically supplement Gemini's VNDLY pool with candidate-supported
    requisition skills it may have omitted.

    This protects exact structured Must/Nice values, formal Required skills,
    current manager priorities, and explicit Preferred skills from being crowded
    out by merely interesting candidate strengths.
    """
    output = [dict(item) for item in (ranked_skills or [])]
    seen = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in output
    }

    exact_must = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in (selected_must_items or [])
        if str(item.get("skill_name", "") or "").strip()
    }
    exact_nice = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in (selected_nice_items or [])
        if str(item.get("skill_name", "") or "").strip()
    }

    for cat_item in catalog or []:
        if not isinstance(cat_item, dict):
            continue

        name = str(cat_item.get("skill_name", "") or "").strip()
        if not name:
            continue

        key = fred_normalize_skill_name(name)
        if key in seen:
            continue

        tier = str(cat_item.get("catalog_tier", "canonical_preferred") or "")
        if tier not in {"canonical_preferred", "compound_review", "legacy_free_text"}:
            continue

        evidence_strength = fred_skill_work_history_evidence(name, raw_resume_text)
        if evidence_strength <= 0:
            continue

        if key in exact_must:
            category = "exact_structured_must"
        elif key in exact_nice:
            category = "exact_structured_nice"
        else:
            if tier not in {"canonical_preferred", "compound_review"}:
                continue
            category = fred_skill_priority_category(name, structured_req)

        # Rescue an Additional skill only when it is genuinely supported AND
        # materially relevant to this requisition. This keeps no-cap behavior
        # useful without turning the list into a generic candidate inventory.
        if category == "additional_high_value" and not fred_additional_skill_role_relevance(
            name, structured_req, vndly_context or {}
        ):
            continue

        evidence = fred_find_skill_evidence_excerpt(name, raw_resume_text)
        if not evidence:
            continue

        output.append(
            {
                "id": cat_item.get("id"),
                "skill_name": name,
                "catalog_tier": tier,
                "source_category": category,
                "evidence": evidence,
                "reason": "Deterministically supported current Freddie requirement/priority",
                "work_history_evidence_strength": evidence_strength,
            }
        )
        seen.add(key)

    return output


def fred_vndly_skill_family(skill_name):
    """High-confidence skill families used for Freddie redundancy control."""
    key = fred_normalize_skill_name(skill_name)

    families = {
        "spring": {
            "spring",
            "java spring",
            "spring framework",
            "spring boot",
            "spring mvc",
            "spring web",
            "spring cloud",
        },
        "relational_database": {
            "relational database",
            "relational databases",
            "oracle",
            "db2",
            "ibm db2",
            "microsoft sql server",
            "mysql",
            "postgresql",
            "sybase",
        },
        "testing": {
            "junit testing",
            "mockito",
            "unit testing",
        },
        "source_control": {
            "git",
            "github",
            "gitlab",
        },
        "test_automation": {
            "automation",
            "automation testing",
            "test automation",
            "test automation tools",
            "automated testing",
            "testing automation",
        },
        "finance_generic": {
            "finance",
            "financial services",
            "financial service",
        },
        "data_quality_validation": {
            "data quality",
            "data validation",
            "data quality validation",
            "validation",
        },
        "program_management_label": {
            "program management",
            "enterprise program management",
        },
        "risk_management_label": {
            "risk management",
            "business risk management",
        },
    }

    for family_name, members in families.items():
        if key in members:
            return family_name

    return ""


def fred_skill_is_explicitly_named_in_job(skill_name, structured_req):
    """
    Conservative check for whether a catalogue skill is explicitly named in
    current formal requisition intelligence.
    """
    structured_req = fred_coerce_structured_req(structured_req)
    req_text = json.dumps(
        structured_req,
        ensure_ascii=False,
    ).casefold()

    key = fred_normalize_skill_name(skill_name)

    aliases = {
        "java database connectivity (jdbc)": [
            "jdbc",
            "java database connectivity",
        ],
        "javaserver pages (jsp)": [
            "jsp",
            "javaserver pages",
            "java server pages",
        ],
        "java servlets": ["servlet", "servlets"],
        "junit testing": ["junit"],
        "mockito": ["mockito"],
        "j2ee": ["j2ee", "j2ee technologies", "java ee"],
        "java": ["java"],
        "git": ["git"],
        "eclipse ide": ["eclipse"],
        "jenkins (software)": ["jenkins"],
        "docker (software)": ["docker"],
        "gradle": ["gradle"],
        "db2": ["db2"],
        "ibm db2": ["ibm db2", "db2"],
        "microsoft sql server": ["sql server"],
        "spring framework": ["spring framework"],
        "relational database": [
            "relational database",
            "relational databases",
        ],
        "financial services": [
            "finance background",
            "financial services",
            "financial domain",
        ],
        "finance": [
            "finance background",
            "finance",
            "financial",
        ],
        "python": ["python"],
        "sql": ["sql"],
        "snowflake": ["snowflake"],
        "json": ["json"],
        "opencv": ["opencv", "open cv"],
        "pillow": ["pillow"],
        "pytorch": ["pytorch", "py torch"],
        "tensorflow": ["tensorflow", "tensor flow"],
        "data engineering": ["data engineering"],
        "data quality": ["data quality"],
        "data validation": ["data validation", "validation"],
        "automation": ["automation", "automate", "automated"],
        "automation testing": ["automation testing", "automated testing", "test automation"],
        "test automation": ["test automation", "automation testing", "automated testing"],
        "test automation tools": ["test automation", "automation testing", "testing framework"],
        "data modeling": ["data modeling", "logical data modeling", "logical model", "physical model"],
        "ai/ml": ["ai/ml", "artificial intelligence", "machine learning", "ai enabled", "ai-enabled"],
        "jira": ["jira"],
        "sharepoint": ["sharepoint"],
        "power bi": ["power bi"],
        "power apps": ["power apps"],
        "power automate": ["power automate"],
        "power platform": ["power platform", "power apps", "power automate", "power bi"],
        "agile delivery": ["agile", "scrum", "sprint planning", "backlog", "pi planning"],
        "agile": ["agile", "scrum", "sprint planning"],
        "product management": ["product management", "product manager", "product owner", "product roadmap"],
        "program management": ["program management", "program manager", "program planning", "program delivery", "portfolio management"],
        "enterprise program management": ["program management", "program manager", "program planning", "program delivery", "portfolio management", "multiple projects", "concurrent projects", "concurrent initiatives"],
        "project management": ["project management", "project manager", "project planning", "managing specific projects", "multiple projects"],
        "business analysis": ["business analysis", "business analyst", "business requirements", "requirements elicitation", "user requirements", "requirements sessions", "business needs", "translate business"],
        "requirements analysis": ["requirements analysis", "requirements elicitation", "business requirements", "technical requirements", "user requirements", "requirements sessions", "business needs", "translate business"],
        "stakeholder management": ["stakeholder management", "stakeholders", "business stakeholders"],
        "executive reporting": ["executive reporting", "executive-level reporting", "executive and program-level reporting", "program-level reporting", "executive dashboard", "executive dashboards", "executive status", "status reports", "leadership reporting", "leadership presentation", "leadership presentations"],
        "risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks"],
        "business risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks", "risks and issues"],
        "user story": ["user story", "user stories", "epics", "story writing"],
        "user stories": ["user story", "user stories", "epics", "story writing"],
        "devops": ["devops", "devsecops"],
        "quality assurance": ["quality assurance", "qa"],
        "software testing": ["software testing", "functional testing", "regression testing", "system testing"],
        "mortgage industry": ["mortgage", "appraisal", "collateral", "uad", "underwriting"],
    }

    terms = aliases.get(key, [key])
    return fred_text_contains_any(req_text, terms)


def fred_coerce_structured_req(value):
    """
    Return Freddie structured requisition intelligence as a dict.

    The tailoring prompt historically stores this object as a JSON string for
    interpolation into Gemini prompts. VNDLY ranking helpers, however, need the
    actual dict. Accept either form so a future refactor cannot recreate the
    "'str' object has no attribute 'get'" package failure.
    """
    if isinstance(value, dict):
        return value

    if isinstance(value, str):
        value = value.strip()
        if not value:
            return {}
        try:
            parsed = json.loads(value)
        except Exception:
            return {}
        return parsed if isinstance(parsed, dict) else {}

    return {}


def fred_skill_aliases_for_requirement(skill_name):
    """Aliases used to map catalogue skills back to Freddie's requirement text."""
    key = fred_normalize_skill_name(skill_name)
    aliases = {
        "j2ee": ["j2ee", "j2ee technologies", "java ee"],
        "java database connectivity (jdbc)": ["jdbc", "java database connectivity"],
        "javaserver pages (jsp)": ["jsp", "javaserver pages", "java server pages"],
        "java servlets": ["servlet", "servlets"],
        "junit testing": ["junit"],
        "mockito": ["mockito"],
        "spring framework": ["spring framework"],
        "spring boot": ["spring boot"],
        "relational database": ["relational database", "relational databases"],
        "financial services": ["finance background", "financial services", "financial domain"],
        "eclipse ide": ["eclipse"],
        "jenkins (software)": ["jenkins"],
        "docker (software)": ["docker"],
        "microsoft sql server": ["sql server"],
        "ibm db2": ["ibm db2", "db2"],
        "db2": ["db2"],
        "git": ["git"],
        "gradle": ["gradle"],
        "maven": ["maven"],
        "java": ["java"],
        "python": ["python"],
        "sql": ["sql"],
        "snowflake": ["snowflake"],
        "json": ["json"],
        "opencv": ["opencv", "open cv"],
        "pillow": ["pillow"],
        "pytorch": ["pytorch", "py torch"],
        "tensorflow": ["tensorflow", "tensor flow"],
        "data engineering": ["data engineering"],
        "data quality": ["data quality"],
        "data validation": ["data validation", "validation"],
        "automation": ["automation", "automate", "automated"],
        "automation testing": ["automation testing", "automated testing", "test automation"],
        "test automation": ["test automation", "automation testing", "automated testing"],
        "test automation tools": ["test automation", "automation testing", "testing framework"],
        "data modeling": ["data modeling", "logical data modeling", "logical model", "physical model"],
        "ai/ml": ["ai/ml", "artificial intelligence", "machine learning", "ai enabled", "ai-enabled"],
        "jira": ["jira"],
        "sharepoint": ["sharepoint"],
        "power bi": ["power bi"],
        "power apps": ["power apps"],
        "power automate": ["power automate"],
        "power platform": ["power platform", "power apps", "power automate", "power bi"],
        "agile delivery": ["agile", "scrum", "sprint planning", "backlog", "pi planning"],
        "agile": ["agile", "scrum", "sprint planning"],
        "product management": ["product management", "product manager", "product owner", "product roadmap", "product analyst", "product team"],
        "program management": ["program management", "program manager", "program planning", "program delivery", "portfolio management"],
        "enterprise program management": ["program management", "program manager", "program planning", "program delivery", "portfolio management", "multiple projects", "concurrent projects", "concurrent initiatives"],
        "project management": ["project management", "project manager", "project planning", "managing specific projects", "multiple projects"],
        "business analysis": ["business analysis", "business analyst", "business requirements", "requirements elicitation", "user requirements", "requirements sessions", "business needs", "translate business"],
        "requirements analysis": ["requirements analysis", "requirements elicitation", "business requirements", "technical requirements", "user requirements", "requirements sessions", "business needs", "translate business"],
        "stakeholder management": ["stakeholder management", "stakeholders", "business stakeholders"],
        "executive reporting": ["executive reporting", "executive-level reporting", "executive and program-level reporting", "program-level reporting", "executive dashboard", "executive dashboards", "executive status", "status reports", "leadership reporting", "leadership presentation", "leadership presentations"],
        "risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks"],
        "business risk management": ["risk management", "risk identification", "risk assessment", "risk register", "issue tracking", "program risks", "project risks", "risks and issues"],
        "user story": ["user story", "user stories", "epics", "story writing"],
        "user stories": ["user story", "user stories", "epics", "story writing"],
        "devops": ["devops", "devsecops"],
        "quality assurance": ["quality assurance", "qa"],
        "software testing": ["software testing", "functional testing", "regression testing", "system testing"],
        "finance": ["finance background", "finance", "financial"],
        "mortgage industry": ["mortgage", "appraisal", "collateral", "uad", "underwriting"],
    }
    return aliases.get(key, [key])


def fred_skill_matches_text(skill_name, text_value):
    return fred_text_contains_any(
        text_value,
        fred_skill_aliases_for_requirement(skill_name),
    )


def fred_skill_requirement_group(skill_name, structured_req):
    """Return the 0-based must-have competency group that names this skill."""
    structured_req = fred_coerce_structured_req(structured_req)
    groups = structured_req.get("must_have_competency_groups", []) or []
    for group_index, group in enumerate(groups):
        if not isinstance(group, dict):
            continue
        terms = group.get("terms", []) or []
        if not isinstance(terms, list):
            terms = []
        group_text = " ".join(
            [
                str(group.get("name", "") or ""),
                " ".join(str(x) for x in terms),
                str(group.get("reason", "") or ""),
            ]
        )
        if fred_skill_matches_text(skill_name, group_text):
            return group_index
    return 99


def fred_skill_requirement_term_position(skill_name, structured_req):
    """Position of the skill inside its highest-priority must-have group."""
    structured_req = fred_coerce_structured_req(structured_req)
    groups = structured_req.get("must_have_competency_groups", []) or []
    for group_index, group in enumerate(groups):
        if not isinstance(group, dict):
            continue
        terms = group.get("terms", []) or []
        if not isinstance(terms, list):
            terms = []
        for term_index, term in enumerate(terms):
            if fred_skill_matches_text(skill_name, str(term)):
                return group_index, term_index
        if fred_skill_matches_text(skill_name, str(group.get("name", "") or "")):
            return group_index, 99
    return 99, 99


def fred_skill_clause_strength(skill_name, structured_req):
    """
    Prefer skills named in focused/dedicated mandatory clauses over tools buried
    in a long multi-tool list. Higher is stronger.
    """
    structured_req = fred_coerce_structured_req(structured_req)
    clauses = []
    for key in ["explicit_must_have_requirements", "required_requirements", "domain_requirements"]:
        value = structured_req.get(key, []) or []
        if isinstance(value, list):
            clauses.extend(value)
        elif isinstance(value, str) and value.strip():
            clauses.append(value)

    best = 0
    aliases = fred_skill_aliases_for_requirement(skill_name)

    for clause in clauses:
        clause_text = str(clause or "")
        fold = clause_text.casefold()
        if not fred_text_contains_any(clause_text, aliases):
            continue

        # Dedicated requirements score higher than long tool bundles.
        separators = clause_text.count(",") + len(re.findall(r"\band\b", fold)) + clause_text.count(";")
        strength = max(1, 12 - separators)

        # Repeated mention across formal clauses is additional signal.
        occurrences = sum(
            1 for alias in aliases
            if alias and fred_text_contains_term(clause_text, alias)
        )
        strength += min(3, occurrences)
        best = max(best, strength)

    return best


def fred_supplement_explicit_required_skills(
    ranked_skills,
    structured_req,
    catalog,
    raw_resume_text,
):
    """
    Deterministically add candidate-supported catalogue skills that are explicitly
    required by the current requisition, even when Gemini omitted them from its
    candidate pool. This protects high-signal requirements such as J2EE/JDBC.
    """
    output = [dict(item) for item in (ranked_skills or [])]
    seen = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in output
    }

    for cat_item in catalog or []:
        name = str(cat_item.get("skill_name", "") or "").strip()
        if not name:
            continue

        key = fred_normalize_skill_name(name)
        if key in seen:
            continue

        if cat_item.get("catalog_tier") not in {"canonical_preferred", "compound_review"}:
            continue

        if not fred_skill_is_explicitly_named_in_job(name, structured_req):
            continue

        evidence_strength = fred_skill_work_history_evidence(name, raw_resume_text)
        if evidence_strength <= 0:
            continue

        output.append(
            {
                "id": cat_item.get("id"),
                "skill_name": name,
                "catalog_tier": cat_item.get("catalog_tier", "canonical_preferred"),
                "source_category": "required_job_intelligence",
                "evidence": "Dated work-history evidence",
                "reason": "Explicit current Freddie requirement",
                "work_history_evidence_strength": evidence_strength,
            }
        )
        seen.add(key)

    return output


def fred_skill_is_formal_must_have(skill_name, structured_req):
    """
    True only when the catalogue skill is explicitly named in Freddie's formal
    mandatory requirement text. Manager/Spotlight weighting is intentionally
    unchanged; this helper protects concrete Must-Have technologies from being
    diluted by broad generic labels.
    """
    structured_req = fred_coerce_structured_req(structured_req)

    formal_parts = []
    for key in [
        "explicit_must_have_requirements",
        "required_requirements",
        "domain_requirements",
    ]:
        value = structured_req.get(key, []) or []
        if isinstance(value, list):
            formal_parts.extend(str(x) for x in value)
        elif isinstance(value, str) and value.strip():
            formal_parts.append(value)

    formal_text = " ".join(formal_parts)
    return fred_skill_matches_text(skill_name, formal_text)


def fred_vndly_is_broad_generic_skill(skill_name):
    """
    Broad concepts remain useful, but they should not outrank a candidate-supported
    concrete technology explicitly required by Freddie.
    """
    key = fred_normalize_skill_name(skill_name)

    return key in {
        "automation",
        "automation testing",
        "test automation tools",
        "data engineering",
        "data quality",
        "data validation",
        "finance",
        "financial services",
        "software development",
        "software engineering",
        "programming",
    }


def fred_vndly_redundancy_preference(skill_name):
    """
    Tie-break only among known redundant/synonymous catalogue concepts.
    Lower is better. This is not requisition weighting; it simply prefers the
    clearer VNDLY label when two selections communicate essentially the same
    screening signal.
    """
    key = fred_normalize_skill_name(skill_name)

    preferences = {
        # "Financial Services" is more useful/specific than generic "Finance"
        # when only one of those redundant domain labels should consume a slot.
        "enterprise program management": 0,
        "program management": 1,
        "business risk management": 0,
        "risk management": 1,
        "financial services": 0,
        "finance": 1,

        # "Data Quality" is the cleaner umbrella when "data validation" would
        # otherwise duplicate the same screening dimension.
        "data quality": 0,
        "data validation": 1,

        # When the requisition is about automated testing, prefer the most
        # precise testing label over generic Automation or tool-bucket labels.
        "test automation": 0,
        "automation testing": 1,
        "test automation tools": 2,
        "automated testing": 2,
        "automation": 3,
    }

    return preferences.get(key, 0)


def fred_required_skill_sort_key(item, structured_req, original_index=0):
    category_priority = {
        "exact_structured_must": 0,
        "required_job_intelligence": 1,
        "manager_priority": 2,
        "exact_structured_nice": 3,
        "preferred_job_intelligence": 4,
        "jd_supported_coverage": 5,
        "additional_high_value": 6,
    }
    category = str(item.get("source_category", "") or "").strip().lower()
    group_index = fred_skill_requirement_group(item.get("skill_name", ""), structured_req)
    clause_strength = fred_skill_clause_strength(item.get("skill_name", ""), structured_req)
    evidence = int(item.get("work_history_evidence_strength", 0) or 0)
    return (
        category_priority.get(category, 99),
        fred_vndly_redundancy_preference(
            item.get("skill_name", "")
        ),
        -clause_strength,
        group_index,
        -evidence,
        original_index,
    )


def fred_apply_vndly_redundancy_control(
    ranked_skills,
    structured_req,
    cap=None,
):
    """
    Final Freddie VNDLY selection with NO artificial hard cap.

    Keep every materially relevant, candidate-supported skill while removing
    duplicates and obvious synonym/family noise. This mirrors the real VNDLY
    workflow, where recruiters may select more than eight skills when they all
    add useful screening signal.
    """
    candidates = []
    seen = set()
    structured_req = fred_coerce_structured_req(structured_req)

    for idx, item in enumerate(ranked_skills or []):
        name = str(item.get("skill_name", "") or "").strip()
        key = fred_normalize_skill_name(name)
        if not name or key in seen:
            continue
        copy_item = dict(item)
        copy_item["_sort_key"] = fred_required_skill_sort_key(copy_item, structured_req, idx)
        candidates.append(copy_item)
        seen.add(key)

    if not candidates:
        return []

    selected = []
    selected_keys = set()
    family_counts = {}

    def can_add(item):
        name = str(item.get("skill_name", "") or "").strip()
        key = fred_normalize_skill_name(name)
        if not name or key in selected_keys:
            return False

        family = fred_vndly_skill_family(name)
        selected_names = {
            fred_normalize_skill_name(existing.get("skill_name", ""))
            for existing in selected
        }

        # Remove only true duplicate/synonym noise. Distinct business screening
        # dimensions are allowed to coexist because VNDLY has no eight-skill cap.
        if "j2ee" in selected_names and key in {
            "java servlets",
            "javaserver pages (jsp)",
        }:
            return False

        family_caps = {
            "spring": 1,
            "relational_database": 1,
            "source_control": 1,
            "test_automation": 1,
            "finance_generic": 1,
            "data_quality_validation": 1,
            "program_management_label": 1,
            "risk_management_label": 1,
        }
        if family and family_counts.get(family, 0) >= family_caps.get(family, 99):
            return False
        return True

    for item in sorted(candidates, key=lambda x: x["_sort_key"]):
        if can_add(item):
            selected.append(item)
            key = fred_normalize_skill_name(item.get("skill_name", ""))
            selected_keys.add(key)
            family = fred_vndly_skill_family(item.get("skill_name", ""))
            if family:
                family_counts[family] = family_counts.get(family, 0) + 1

    for item in selected:
        item.pop("_sort_key", None)

    # Optional internal safety only; normal callers pass cap=None.
    if isinstance(cap, int) and cap > 0:
        return selected[:cap]
    return selected


def fred_additional_skill_role_relevance(skill_name, structured_req, vndly_context):
    """
    Conservative relevance gate for `additional_high_value` manual skills.

    No-cap does not mean no discipline. Candidate strengths that are materially
    unrelated to the current role should not appear just because they exist.
    """
    structured_req = fred_coerce_structured_req(structured_req)
    vndly_context = vndly_context or {}
    role_text = "\n".join(
        [
            str(vndly_context.get("desc", "") or ""),
            str(vndly_context.get("notes", "") or ""),
            str(vndly_context.get("transcript", "") or ""),
            json.dumps(structured_req, ensure_ascii=False, default=str),
        ]
    )

    # Direct requirement/role-language match is always enough.
    if fred_skill_matches_text(skill_name, role_text):
        return True

    key = fred_normalize_skill_name(skill_name)
    adjacency = {
        "jira": ["agile", "scrum", "sprint", "backlog", "project management", "product management"],
        "confluence": ["agile", "scrum", "requirements", "documentation", "project management"],
        "product management": ["product analyst", "product owner", "product team", "roadmap", "product"],
        "project management": ["project", "projects", "project delivery", "project planning", "managing specific projects", "manage specific projects"],
        "enterprise program management": ["program management", "program specialist", "portfolio", "multiple projects", "concurrent projects"],
        "business risk management": ["risk", "issue tracking", "escalation", "program health"],
        "snowflake": ["snowflake", "data warehouse", "data engineering", "data platform", "sql database"],
        "git": ["source control", "version control", "software development", "developer", "engineering"],
    }
    terms = adjacency.get(key, [])
    return fred_text_contains_any(role_text, terms) if terms else False


def fred_rerank_vndly_skills_by_evidence(
    ranked_skills,
    raw_resume_text,
    structured_req=None,
    vndly_context=None,
    cap=None,
):
    """
    Requirement-aware evidence reranking for the VNDLY candidate pool.

    Every recommended catalogue skill must survive BOTH:
    1. candidate evidence validation against the ORIGINAL resume; and
    2. source-category normalization so manager emphasis is not mislabeled as a
       formal Required skill.
    """
    rescored = []
    structured_req = fred_coerce_structured_req(structured_req)
    vndly_context = vndly_context or {}

    official_jd = str(vndly_context.get("desc", "") or "")
    msp_notes = str(vndly_context.get("notes", "") or "")
    spotlight = str(vndly_context.get("transcript", "") or "")
    authoritative_role_text = "\n".join([official_jd, msp_notes, spotlight])

    for original_index, item in enumerate(ranked_skills or []):
        name = str(item.get("skill_name", "") or "").strip()
        if not name:
            continue

        evidence_strength = fred_skill_work_history_evidence(
            name,
            raw_resume_text,
        )

        # Absolute candidate-evidence gate.
        if evidence_strength <= 0:
            continue

        copy_item = dict(item)

        # Gemini is asked for a verbatim resume excerpt. If it returns a
        # paraphrase, replace it with a deterministic real resume line rather
        # than trusting invented evidence text.
        candidate_evidence = str(copy_item.get("evidence", "") or "").strip()
        if (
            not candidate_evidence
            or not fred_source_quote_found(raw_resume_text, candidate_evidence)
            or not fred_claim_supported_by_evidence(name, candidate_evidence)
        ):
            candidate_evidence = fred_find_skill_evidence_excerpt(
                name,
                raw_resume_text,
            )

        if not candidate_evidence:
            continue

        copy_item["evidence"] = candidate_evidence
        copy_item["work_history_evidence_strength"] = evidence_strength

        category = str(copy_item.get("source_category", "") or "").strip().lower()

        # Exact structured categories are already validated against the exact
        # requisition fields. Every inferred category is reclassified from the
        # structured requisition so Gemini cannot call a manager priority
        # "Required" or promote a nice-to-have.
        if category not in {"exact_structured_must", "exact_structured_nice"}:
            deterministic_category = fred_skill_priority_category(
                name,
                structured_req,
            )
            if not deterministic_category:
                continue
            category = deterministic_category
            copy_item["source_category"] = category

        if category == "additional_high_value" and not fred_additional_skill_role_relevance(
            name, structured_req, vndly_context
        ):
            continue

        # Validate the role-side citation when one is required. This evidence
        # proves relevance only; it never proves candidate capability.
        role_evidence = str(copy_item.get("role_evidence", "") or "").strip()

        if category == "manager_priority":
            if not (
                role_evidence
                and (
                    fred_source_quote_found(spotlight, role_evidence)
                    or fred_source_quote_found(msp_notes, role_evidence)
                )
            ):
                # The structured analysis already established the manager signal,
                # so retain the category but replace untraceable AI prose with a
                # transparent deterministic reason rather than a fake quote.
                role_evidence = ""

        elif category in {"required_job_intelligence", "preferred_job_intelligence"}:
            if role_evidence and not fred_source_quote_found(
                authoritative_role_text,
                role_evidence,
            ):
                role_evidence = ""

        copy_item["role_evidence"] = role_evidence

        rescored.append(
            (
                fred_required_skill_sort_key(
                    copy_item,
                    structured_req,
                    original_index,
                ),
                copy_item,
            )
        )

    rescored.sort(key=lambda row: row[0])
    results = [row[1] for row in rescored]
    if isinstance(cap, int) and cap > 0:
        return results[:cap]
    return results


def fred_build_final_ranked_vndly_skills(
    raw_ranked_items,
    selected_must_items,
    selected_nice_items,
    catalog,
    cap=None,
):
    """
    Validate Gemini's ranked recommendation list before evidence reranking.

    Priority is decided in the AI ranking prompt, then this function enforces:
    - exact catalogue values only;
    - structured Must/Nice claims must truly be attached to this requisition;
    - legacy/free-text values are allowed only for exact structured req skills;
    - no duplicates.

    `cap` is an optional internal pool-safety limit, not a VNDLY business cap.
    """
    if not isinstance(raw_ranked_items, list):
        return []

    catalog_lookup = fred_catalog_lookup(catalog)

    allowed_must = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in (selected_must_items or [])
        if str(item.get("skill_name", "") or "").strip()
    }
    allowed_nice = {
        fred_normalize_skill_name(item.get("skill_name", ""))
        for item in (selected_nice_items or [])
        if str(item.get("skill_name", "") or "").strip()
    }

    valid_categories = {
        "exact_structured_must",
        "exact_structured_nice",
        "required_job_intelligence",
        "manager_priority",
        "preferred_job_intelligence",
        "jd_supported_coverage",
        "additional_high_value",
    }

    output = []
    seen = set()

    for raw_item in raw_ranked_items:
        # Validate the candidate pool here. Final selection later removes
        # redundancy but does not impose an artificial eight-skill cap.

        if not isinstance(raw_item, dict):
            continue

        proposed = (
            raw_item.get("skill_name")
            or raw_item.get("value")
            or raw_item.get("label")
            or ""
        )
        category = str(
            raw_item.get("source_category", "")
        ).strip().lower()

        if category not in valid_categories:
            continue

        exact = catalog_lookup.get(
            fred_normalize_skill_name(proposed)
        )

        if not exact:
            continue

        name = str(exact.get("skill_name", "") or "").strip()
        key = fred_normalize_skill_name(name)

        if not name or key in seen:
            continue

        tier = exact.get(
            "catalog_tier",
            "canonical_preferred",
        )

        # Exact structured categories must actually come from the exact req field.
        if category == "exact_structured_must":
            if key not in allowed_must:
                continue
        elif category == "exact_structured_nice":
            if key not in allowed_nice:
                continue
        else:
            # Legacy/free-text is never used for inferred JD/manager/additional
            # recommendations. It is reserved for exact req-attached tags.
            if tier not in {
                "canonical_preferred",
                "compound_review",
            }:
                continue

        output.append(
            {
                "id": exact.get("id"),
                "skill_name": name,
                "catalog_tier": tier,
                "source_category": category,
                "evidence": str(
                    raw_item.get("evidence", "") or ""
                ).strip(),
                "role_evidence": str(
                    raw_item.get("role_evidence", "") or ""
                ).strip(),
                "reason": str(
                    raw_item.get("reason", "") or ""
                ).strip(),
            }
        )
        seen.add(key)

    return output


def fred_split_ranked_vndly_skills(ranked_skills):
    """Split the globally-ranked list into UI/email groups without changing rank."""
    groups = {
        "exact_must_have_skills": [],
        "exact_nice_to_have_skills": [],
        "required_vndly_skills_from_job_intelligence": [],
        "manager_priority_vndly_skills": [],
        "preferred_vndly_skills_from_job_intelligence": [],
        "jd_coverage_vndly_skills": [],
        "additional_vndly_skills": [],
    }

    category_to_key = {
        "exact_structured_must": "exact_must_have_skills",
        "exact_structured_nice": "exact_nice_to_have_skills",
        "required_job_intelligence": "required_vndly_skills_from_job_intelligence",
        "manager_priority": "manager_priority_vndly_skills",
        "preferred_job_intelligence": "preferred_vndly_skills_from_job_intelligence",
        "jd_supported_coverage": "jd_coverage_vndly_skills",
        "additional_high_value": "additional_vndly_skills",
    }

    for item in ranked_skills or []:
        key = category_to_key.get(
            str(item.get("source_category", "")).strip().lower()
        )
        if key:
            groups[key].append(item)

    return groups


def fred_extract_spotlight_hiring_manager_name(spotlight_transcript):
    """
    Freddie-only helper that returns a hiring-manager name only when the
    transcript explicitly labels ONE unique person as Hiring Manager.

    When multiple hiring managers are labeled, return blank so recruiter-facing
    comments refer to "the hiring team" / "supplier call" rather than
    accidentally attributing another manager's statement to the first speaker.
    """
    transcript = str(spotlight_transcript or "").strip()
    if not transcript:
        return ""

    # Supplier-call transcripts are often Markdown-bolded (for example
    # **Bruce (Hiring Manager):**). Strip formatting before parsing names.
    transcript = transcript.replace("**", "")

    patterns = [
        r"(?im)^\s*([A-Z][A-Za-z'’.\-]+(?:\s+[A-Z][A-Za-z'’.\-]+){0,2})\s*\(\s*Hiring\s+Manager\s*\)\s*:",
        r"(?im)^\s*([A-Z][A-Za-z'’.\-]+(?:\s+[A-Z][A-Za-z'’.\-]+){0,2})\s*[-–\u2014]\s*Hiring\s+Manager\s*:",
    ]

    names = []
    for pattern in patterns:
        for match in re.finditer(pattern, transcript):
            name = re.sub(r"\s+", " ", match.group(1)).strip()
            if name and name.casefold() not in {n.casefold() for n in names}:
                names.append(name)

    if not names:
        return ""
    first_names = [name.split()[0] for name in names]
    if len(first_names) == 1:
        return first_names[0]
    if len(first_names) == 2:
        return f"{first_names[0]} and {first_names[1]}"
    return ", ".join(first_names[:-1]) + f", and {first_names[-1]}"


def fred_clean_vndly_submission_summary(summary, candidate_name):
    """Deterministic cleanup for the recruiter-facing VNDLY Comments field."""
    if not summary:
        return ""

    cleaned = fred_clean_generated_text(summary)
    full_name = str(candidate_name or "").strip()
    first_name = full_name.split()[0] if full_name else ""
    last_name = full_name.split()[-1] if full_name else ""

    if full_name and first_name:
        cleaned = re.sub(
            r"^\s*" + re.escape(full_name) + r"\b",
            first_name,
            cleaned,
            count=1,
            flags=re.IGNORECASE,
        )

    # Remove honorific constructions the model should never use in this field.
    if last_name and first_name:
        cleaned = re.sub(
            rf"\b(?:Mr|Ms|Mrs|Dr)\.?\s+{re.escape(last_name)}\b",
            first_name,
            cleaned,
            flags=re.IGNORECASE,
        )
        cleaned = re.sub(
            rf"\b(?:Mr|Ms|Mrs|Dr)\.?\s+{re.escape(candidate_name)}\b",
            first_name,
            cleaned,
            flags=re.IGNORECASE,
        )

    cleaned = re.sub(r"\b(?:Mr|Ms|Mrs|Dr)\.\s+", "", cleaned, flags=re.IGNORECASE)

    # Strip a few recurring editorial tails without rewriting factual content.
    cleaned = re.sub(
        r",?\s+(?:which is|making (?:him|her|them))\s+(?:crucial|critical|ideal)\s+for\s+(?:this|the)\s+[^.]+\.",
        ".",
        cleaned,
        flags=re.IGNORECASE,
    )

    return cleaned.strip()


def fred_validate_vndly_summary(
    summary,
    evidence_list,
    raw_resume_text,
    structured_req,
    vndly_context,
):
    """Deterministic safety checks for recruiter-facing Freddie VNDLY comments."""
    errors = []
    summary = str(summary or "").strip()
    evidence_list = fred_string_list(evidence_list)
    vndly_context = vndly_context or {}

    if not summary:
        return ["VNDLY summary is blank."]

    if not evidence_list:
        errors.append("VNDLY summary has no candidate evidence.")

    valid_evidence = []
    for evidence in evidence_list:
        if fred_vndly_summary_evidence_found(raw_resume_text, evidence):
            valid_evidence.append(evidence)
        else:
            errors.append(
                "VNDLY summary cites candidate evidence not found in the supplied candidate sources: "
                + evidence[:120]
            )

    joined_evidence = " ".join(valid_evidence)

    unsupported_qualifiers = fred_unsupported_high_impact_claims(
        summary,
        joined_evidence,
    )
    if unsupported_qualifiers:
        errors.append(
            "VNDLY summary contains unsupported high-impact claim(s): "
            + ", ".join(unsupported_qualifiers)
        )

    unsupported_identifiers = fred_explicit_identifiers_supported(
        summary,
        joined_evidence,
        raw_resume_text,
    )
    if unsupported_identifiers:
        errors.append(
            "VNDLY summary contains unsupported named identifier(s): "
            + ", ".join(unsupported_identifiers)
        )

    excluded_hits = fred_output_hits_excluded_requirement(
        summary,
        structured_req,
    )
    if excluded_hits:
        errors.append(
            "VNDLY summary uses a manager-explicitly-excluded concept: "
            + ", ".join(excluded_hits)
        )

    role_text = "\n".join(
        [
            str(vndly_context.get("desc", "") or ""),
            str(vndly_context.get("notes", "") or ""),
            str(vndly_context.get("transcript", "") or ""),
        ]
    )

    # The final sentence may explicitly explain transferability to the target
    # role. Earlier sentences should not quietly convert Freddie wording into
    # candidate history.
    sentences = fred_split_sentences(summary)
    for sentence in sentences[:-1]:
        hits = fred_role_only_phrase_hits(
            sentence,
            role_text,
            raw_resume_text,
            joined_evidence,
        )
        token_hits = fred_role_only_signal_tokens(
            sentence,
            role_text,
            raw_resume_text,
            joined_evidence,
        )
        if hits:
            errors.append(
                "VNDLY summary contains role-source wording not independently "
                "supported as candidate history: " + "; ".join(hits)
            )
        if token_hits:
            errors.append(
                "VNDLY summary contains distinctive role-source term(s) absent "
                "from candidate evidence: " + ", ".join(token_hits)
            )

    return errors


def fred_repair_vndly_summary(
    api_key,
    candidate_name,
    draft_summary,
    validator_errors,
    raw_resume_text,
    structured_req,
    vndly_context,
    calculated_total_experience,
    candidate_positioning_title,
):
    """One evidence-grounded repair pass for Freddie VNDLY Comments."""
    prompt = f"""
Return valid JSON ONLY.

You are repairing Freddie Mac VNDLY recruiter Comments after deterministic
candidate-grounding checks found problems.

CANDIDATE: {candidate_name}
CANDIDATE POSITIONING TITLE: {candidate_positioning_title or "BLANK"}
PYTHON-CALCULATED TOTAL EXPERIENCE: {calculated_total_experience or "BLANK"}

VALIDATOR ERRORS:
- """ + "\n- ".join(validator_errors) + f"""

DRAFT COMMENTS:
{draft_summary}

RULES:
- 3-5 compact sentences, one paragraph, about 105-155 words and never over 175.
- The candidate evidence source below contains the original resume and may also
  contain candidate-confirmed follow-up facts. Use only facts supported there.
- Official JD, Spotlight transcript, and MSP notes define role priorities only.
- Do not turn a role requirement into candidate history.
- Do not use manager-explicitly-excluded items.
- Merely de-emphasized JD skills may remain as secondary MSP/JD coverage when
  candidate evidence supports them and they do not displace current priorities.
- If a claim is not directly supported, remove or soften it rather than inventing
  evidence.
- Final sentence may explain truthful transferability to the Freddie need.
- NEVER use an em dash.
- Never mention "VNDLY requirements" or similar process-oriented matching language.
- Return 2-8 short candidate-evidence excerpts supporting every claim. Resume
  and candidate-clarification evidence must be VERBATIM. UI-entered logistics
  facts may be cited using a concise natural rendering of the structured fact.

STRUCTURED REQUISITION:
{json.dumps(fred_coerce_structured_req(structured_req), indent=2, ensure_ascii=False)}

OFFICIAL JD:
{str((vndly_context or {}).get("desc", "") or "")}

MSP/VNDLY NOTES:
{str((vndly_context or {}).get("notes", "") or "")}

SPOTLIGHT TRANSCRIPT:
{str((vndly_context or {}).get("transcript", "") or "")}

CANDIDATE EVIDENCE SOURCE:
{raw_resume_text}

OUTPUT:
{{
  "VNDLY_SUMMARY": "",
  "VNDLY_SUMMARY_EVIDENCE": [""]
}}
"""
    return fred_generate_json(api_key, prompt)


def fred_build_vndly_candidate_package(
    api_key,
    candidate_name,
    raw_resume_text,
    structured_req,
    vndly_context,
    vetting_for_ai,
    resume_summary,
    calculated_total_experience,
    candidate_positioning_title,
    candidate_submission_facts=None,
    candidate_clarifications="",
):
    """
    Create the Freddie-only VNDLY Comments text and skill selections.
    """
    structured_req = fred_coerce_structured_req(structured_req)
    candidate_submission_facts = candidate_submission_facts or {}
    candidate_evidence_text = fred_candidate_evidence_text(
        raw_resume_text, candidate_clarifications
    )
    submission_fact_lines = [
        f"{key}: {value}"
        for key, value in candidate_submission_facts.items()
        if str(value or "").strip()
    ]
    vndly_summary_candidate_evidence = candidate_evidence_text
    if submission_fact_lines:
        vndly_summary_candidate_evidence += (
            "\n\nCANDIDATE SUBMISSION FACTS:\n" + "\n".join(submission_fact_lines)
        )

    catalog = fred_load_vndly_skill_catalog()

    if not catalog:
        raise RuntimeError(
            f"VNDLY skill catalogue '{FREDDIE_VNDLY_SKILLS_FILENAME}' "
            "was not found or could not be read."
        )

    catalog_for_ai = [
        {
            "id": item.get("id"),
            "skill_name": item.get("skill_name", ""),
            "catalog_tier": item.get(
                "catalog_tier",
                "canonical_preferred",
            ),
        }
        for item in catalog
    ]

    catalog_json = json.dumps(
        catalog_for_ai,
        ensure_ascii=False,
        separators=(",", ":"),
    )

    vndly_context = vndly_context or {}

    spotlight_hiring_manager_name = fred_extract_spotlight_hiring_manager_name(
        vndly_context.get("transcript", "")
    )

    raw_must = str(vndly_context.get("must_have_skills", "") or "").strip()
    raw_nice = str(vndly_context.get("nice_to_have_skills", "") or "").strip()

    selected_must_items = fred_resolve_structured_vndly_skills(
        raw_must,
        catalog,
    )
    selected_nice_items = fred_resolve_structured_vndly_skills(
        raw_nice,
        catalog,
    )

    selected_must_json = json.dumps(
        [
            {
                "id": item.get("id"),
                "skill_name": item.get("skill_name", ""),
                "catalog_tier": item.get(
                    "catalog_tier",
                    "canonical_preferred",
                ),
            }
            for item in selected_must_items
        ],
        ensure_ascii=False,
        indent=2,
    )

    selected_nice_json = json.dumps(
        [
            {
                "id": item.get("id"),
                "skill_name": item.get("skill_name", ""),
                "catalog_tier": item.get(
                    "catalog_tier",
                    "canonical_preferred",
                ),
            }
            for item in selected_nice_items
        ],
        ensure_ascii=False,
        indent=2,
    )

    package_prompt = f"""
Return a valid JSON object ONLY.

You are preparing the VNDLY submission metadata for a Freddie Mac contingent
worker candidate. This is SEPARATE from the resume Summary.

CANDIDATE:
{candidate_name}

CANDIDATE SUBMISSION FACTS (UI-entered logistics; use only when they materially
help the MSP shortlist decision):
{json.dumps(candidate_submission_facts, indent=2, ensure_ascii=False)}

CANDIDATE-CONFIRMED CLARIFICATIONS / FOLLOW-UP ANSWERS:
{candidate_clarifications or "NONE PROVIDED"}

ALREADY-APPROVED RESUME SUMMARY:
{resume_summary}

RECOMMENDED CANDIDATE POSITIONING TITLE:
{candidate_positioning_title or "BLANK"}

SPOTLIGHT HIRING MANAGER NAME - EXPLICIT TRANSCRIPT LABEL ONLY:
{spotlight_hiring_manager_name or "BLANK"}

CANDIDATE TITLE RULE:
- The title above was selected from the candidate's evidence + actual role
  function using a natural external-market title.
- If it is populated and you identify the candidate by title in the VNDLY
  Submission Summary, use that title.
- Do NOT substitute the structured VNDLY Job Title or rate-card title.
- The VNDLY Job Title may still describe the requisition in job metadata, but it
  is not automatically the candidate's professional identity.
- If the positioning title is blank, use a concise functional identity supported
  by the resume; do not default to an awkward VMS/rate-card title.

PYTHON-CALCULATED TOTAL PROFESSIONAL EXPERIENCE:
{calculated_total_experience}

CRITICAL EXPERIENCE RULE:
- If the Python-calculated total above is populated, that is the ONLY total
  career-experience figure you may use.
- Never use, repeat, paraphrase, or round a candidate-written estimate such as
  "close to 20 years" or "nearly 20 years" when it differs from the calculated
  value.
- Do not claim a requirement-specific minimum number of years unless the dated
  work history itself supports that minimum for the relevant technology/domain.
- A technology named only in the candidate's Professional Summary or skills
  inventory can establish familiarity, but by itself cannot prove a multi-year
  threshold.

STRUCTURED REQUISITION INTELLIGENCE:
{json.dumps(structured_req, indent=2, ensure_ascii=False)}

IMPORTANT JOB-INTELLIGENCE RULE:
- Formal requirements are NOT limited to structured VNDLY Must/Nice skill tags.
- The Official JD and authoritative manager/MSP/Spotlight guidance can establish
  Required or Preferred requirements even when the structured skill fields are blank.
- Use the existing structured requisition intelligence hierarchy exactly as supplied.
- A requirement that is formally mandatory in the JD or authoritative manager guidance
  belongs under REQUIRED_VNDLY_SKILLS_FROM_JOB_INTELLIGENCE when candidate evidence supports it.
- A requirement that is explicitly preferred/nice-to-have in the JD or authoritative
  manager guidance belongs under PREFERRED_VNDLY_SKILLS_FROM_JOB_INTELLIGENCE when supported.

STRUCTURED VNDLY MUST HAVE SKILLS FIELD - RAW TEXT:
{raw_must or "BLANK"}

DETERMINISTIC EXACT MUST-HAVE SKILLS SELECTED ON THIS REQUISITION:
{selected_must_json}

STRUCTURED VNDLY NICE TO HAVE SKILLS FIELD - RAW TEXT:
{raw_nice or "BLANK"}

DETERMINISTIC EXACT NICE-TO-HAVE SKILLS SELECTED ON THIS REQUISITION:
{selected_nice_json}

SUPPLIER VETTING RESPONSES:
{vetting_for_ai}

OFFICIAL VNDLY JOB DESCRIPTION - ROLE EVIDENCE SOURCE:
{vndly_context.get("desc", "")}

VNDLY / MSP NOTES - ROLE EVIDENCE SOURCE:
{vndly_context.get("notes", "")}

SPOTLIGHT CALL TRANSCRIPT - HIGHEST-PRIORITY CURRENT MANAGER SOURCE:
{vndly_context.get("transcript", "")}

FULL ORIGINAL CANDIDATE RESUME:
{raw_resume_text}

EXACT VNDLY SKILL CATALOGUE:
{catalog_json}

======================================================================
TASK 1 - VNDLY COMMENTS / MSP SHORTLIST JUSTIFICATION
======================================================================

Write the recruiter Comments that will be pasted into VNDLY. This is NOT merely
a shortened resume Summary. Treat it as a supplier-to-MSP explanation of WHY
this candidate deserves manager review.

PRIMARY AUDIENCE:
An MSP recruiter who knows the requisition and attended/reviewed the supplier
call, but may not be deeply technical.

PRIMARY GOALS - IN THIS ORDER:
1. Make the shortlist decision easy.
2. Show that the supplier searched against the hiring manager's actual current
   priorities, not just generic JD keywords.
3. Give the MSP concrete candidate evidence they can confidently pass to the
   hiring manager.

MENTAL MODEL:
The reader should be able to answer:
- Who is this candidate?
- Which 2-4 things the manager cares about do they genuinely demonstrate?
- Where did they prove those things?
- Why does that proven background transfer to this Freddie role?

Use plain business language first and technical detail only where it strengthens
the shortlist case.

LENGTH / STRUCTURE:
- 3-5 compact sentences as ONE paragraph.
- Target roughly 105-155 words. Do not exceed 175 words.
- The comments are allowed to be more explicit than the resume Summary because
  their first job is to help the MSP make a fast shortlist decision.
- When candidate submission facts contain a materially helpful logistics signal,
  such as former Freddie/rehire eligibility or confirmed availability for a
  required onsite schedule, a short opening clause may state it. Do not mention
  compensation or unnecessary immigration details.
- Do NOT lead with a very large total-years number when it could create an
  unnecessary overqualification concern. Use a natural authority phrase instead.
- Also return VNDLY_SUMMARY_EVIDENCE containing 2-8 short evidence excerpts
  that collectively support every candidate claim. Experience/skill/ownership
  evidence must come from the original resume or candidate-confirmed
  clarifications. UI-entered logistics such as former-Freddie status, rehire
  eligibility, current location, and onsite availability may be cited from the
  structured CANDIDATE SUBMISSION FACTS. Role-source wording may explain
  transferability, but may not substitute for candidate evidence.
- Sentence 1: natural candidate positioning title + the strongest genuine role
  anchors. Use the deterministic total-career figure only when it strengthens
  the shortlist case and does not unnecessarily advertise overqualification.
- Sentence 2: strongest concrete recent/repeated proof. Name the employer when
  it improves credibility.
- Sentence 3: the next most important supported proof cluster, ideally including
  relevant technical/domain depth.
- Sentence 4: explain why the proven experience transfers to THIS role and makes
  the candidate worth manager review.

HIRING-MANAGER NAME RULE:
- If SPOTLIGHT HIRING MANAGER NAME above is populated, it may contain one or
  multiple first names (for example "Bruce and Colton"). Use that name string
  ONCE when it makes the comments more useful to the MSP.
- Best use: connect proven candidate evidence to a priority the hiring manager
  explicitly emphasized on the supplier/Spotlight call.
- Natural pattern:
  "This combination gives Arunabh the hands-on testing and automation background
  Mark emphasized on the supplier call, with the technical depth to support
  structured testing of Freddie Mac's data-modeling platform."
- Do NOT use the name merely for decoration or claim the manager emphasized
  something unless STRUCTURED REQUISITION INTELLIGENCE actually supports it.
- If the name is BLANK, do not invent or infer a hiring-manager name.
- Do NOT substitute the VNDLY Resource Manager for the hiring manager unless the
  transcript itself explicitly labels that person as Hiring Manager.

REQUIREMENT PRIORITY:
- Start from the CURRENT substantive requirement hierarchy in STRUCTURED
  REQUISITION INTELLIGENCE, not from whichever technologies are most prominent
  on the candidate's resume.
- Highest priority: current manager/MSP/Spotlight selection guidance and why
  prior candidates failed, when available.
- Then formal Must Haves / required competencies / domain requirements.
- Then explicit preferred skills only when they materially strengthen this
  particular candidate.
- Do not spend scarce summary space on adjacent candidate strengths that are not
  important to the current requisition.
- Prefer recurring responsibilities and direct hands-on execution over oversight,
  strategy, or one-time projects unless the role specifically asks for those.

EVIDENCE / TRANSFERABILITY BOUNDARY - ABSOLUTE:
- Candidate facts must come from the resume, official candidate vetting, or
  candidate-confirmed clarifications supplied above.
- Freddie's JD, manager comments, MSP notes, and target environment define WHY
  an experience matters; they do NOT become candidate history.
- Do not write that the candidate tested data-modeling workflows, built an AI
  platform, used a named database, worked in mortgage, etc. unless candidate
  evidence actually supports that statement.
- It IS appropriate to say supported prior experience "provides a foundation
  for", "transfers to", or "can support" a Freddie responsibility.
- When evidence is adjacent rather than direct, explain the transfer honestly
  instead of upgrading it into direct experience.

STYLE - VERY IMPORTANT:
- Write like an experienced recruiter sending a concise internal candidate note.
- Use the candidate's FIRST NAME naturally after the first sentence.
- NEVER use honorifics or formal references such as "Mr.", "Ms.", "Mrs.",
  "Dr.", "Mr. Zaman", "Ms. Fisher", "the candidate", etc.
- Preserve employer names exactly as they appear in the resume. Do not silently
  correct/rebrand company names from outside knowledge.
- Use clear business English, not consultant language or inflated adjectives.
- Prefer factual evidence over claims about fit.
- Keep the paragraph skimmable; avoid long comma-heavy inventory sentences.
- Do not simply paraphrase or duplicate the formatted resume Summary.
- The final sentence may be more explicit than the resume Summary about WHY the
  candidate should advance, as long as it is phrased as evidence-based
  transferability rather than unsupported praise.

GOOD MODEL - communications example:
"Kristin is a senior communications specialist with 18+ years of experience
 developing business communications, change communications, and audience-focused
 content across technology and government environments. At CGI, she managed
 enterprise communication channels and repositories including SharePoint for a
 7,000+ employee technology organization and supported high-volume communications
 workstreams. She has independently created newsletters, presentations, web and
 digital content, implementation communications, and executive materials while
 translating complex business and technical information into clear,
 audience-specific messaging. Her hands-on content production, SharePoint
 experience, operational coordination, and technology-change communications
 support the need for a seasoned individual contributor who can independently
 execute communications deliverables."

The example demonstrates the desired DENSITY and recruiter tone. Do not copy its
wording or force its structure onto technical roles.

BANNED / AVOID:
- "strongly matching", "aligning perfectly", "perfectly aligns", "ideal fit",
  "great fit", "strong fit", "well suited", "well-suited"
- "positions him/her/them", "meets all requirements", "directly matches"
- "possesses extensive", "possesses solid expertise", "brings a wealth of"
- "crucial for this role", "critical for this role", "data-intensive role"
- "immediate asset", "Day 1", "hit the ground running"
- Honorifics: "Mr.", "Ms.", "Mrs.", "Dr."
- Generic closing judgments when the evidence itself is stronger

OTHER RULES:
- Never invent or infer candidate experience.
- Do not merge evidence from separate employers into one sentence in a way that
  falsely implies all tools/responsibilities occurred together.
- Do not discuss weaknesses, missing skills, compensation, citizenship,
  sponsorship, submission mechanics, shortlisting strategy, or AI matching
  strategy. A concise former-Freddie/rehire or required-onsite availability fact
  may be used when it materially reduces MSP uncertainty.
- Never write phrases such as "aligns with VNDLY requirements", "matches the
  VNDLY requirements", or otherwise tell the reader that the candidate was
  optimized against VNDLY/JD requirements. Demonstrate the match through facts.
- NEVER use an em dash in the comments.
- AI/GenAI may be mentioned only when it is genuinely a job technology and the
  candidate has supported experience.

======================================================================
======================================================================
TASK 2 - FINAL VNDLY MANUAL SKILLS (NO FIXED HARD CAP)
======================================================================

Produce ONE ranked shortlist of the BEST VNDLY skills to select for this
candidate submission.

CRITICAL:
- There is NO fixed manual-skill cap in VNDLY for this workflow.
- Build a HIGH-SIGNAL list, not a comprehensive inventory.
- Return every candidate-supported skill that materially improves requirement
  coverage, MSP search/review value, or hiring-manager relevance.
- Stop when additional skills would be redundant, weak, or tangential.
- It is completely acceptable to return a short list for a narrow role or a
  longer list for a broad role.
- Rank every item from highest to lowest submission value.
- IMPORTANT: A technical skill appearing only in the candidate's Professional
  Summary or top-level skills inventory MUST NOT outrank a current Required skill
  that is explicitly demonstrated inside dated work-history roles.
- For example, if Docker appears only in a summary but J2EE/JDBC are explicitly
  shown in dated roles and are current JD requirements, J2EE/JDBC must rank above
  Docker.

PRIORITY HIERARCHY:
1. Exact structured VNDLY Must-Have skill attached to this requisition,
   when the candidate genuinely satisfies the COMPLETE requirement.
2. Explicit formal MUST-HAVE / REQUIRED skill from the Official JD.
3. CURRENT MANAGER PRIORITY from Spotlight/MSP selection guidance when it
   materially affects interview selection, even when it is not formally labeled
   Required. Keep this category distinct so a manager emphasis is not falsely
   presented to the recruiter as a formal Must Have.
4. Exact structured VNDLY Nice-To-Have skill attached to this requisition,
   when genuinely satisfied.
5. Explicit PREFERRED / NICE-TO-HAVE skill from the JD or manager guidance.
6. Additional candidate-supported skill only when it adds unusually useful
   signal for THIS requisition.

CLASSIFICATION FIREWALL:
- `required_job_intelligence` means FORMALLY required by the JD/current formal
  Freddie requirement language. Do not use it merely because a manager discussed
  or emphasized the topic.
- `jd_supported_coverage` means the candidate genuinely has a skill/concept that
  appears in the official JD but current manager guidance has made it secondary
  rather than central. Keep it for MSP/JD coverage when useful, but rank it
  below current manager priorities and explicit preferred items.
- `manager_priority` means the Spotlight/MSP/hiring team clearly emphasized the
  competency as important to selection, but it is not established as formally
  mandatory.
- If Freddie says a skill is "nice to have", "helpful", "a plus", or "not
  required", it MUST be Preferred even if the role uses that skill operationally.
- If a manager says the team does not use an artifact/process (for example User
  Stories), DE-EMPHASIZE it rather than automatically deleting it. When the
  candidate genuinely has that experience and the official JD still contains it,
  it may be retained as `jd_supported_coverage` for MSP/JD recognition.
- Only an item in `explicitly_excluded_requirements` is automatically prohibited.

SELECTION PRINCIPLES:
- Optimize for SIGNAL, not coverage.
- CANDIDATE-EVIDENCE GATE: never recommend a catalogue skill solely because it
  appears in Freddie's JD, structured Must/Nice fields, manager comments, or
  Spotlight transcript. The ORIGINAL RESUME or supplied candidate-confirmed
  clarification must independently support every recommended VNDLY skill.
  Candidate vetting may also clarify factual depth when it describes performed
  experience rather than a theoretical answer.
- If the role requires Data Modeling but the resume does not demonstrate data
  modeling concepts/work, do NOT recommend `Data Modeling`.
- If the role requires SQL but the resume does not show SQL, do NOT recommend
  `SQL`.
- A skill that directly represents the core requirement is more valuable than
  several narrower synonyms/components of the same concept.
- Prefer the exact terminology Freddie is screening for.
- Prefer skills supported by dated, substantive work history.
- Prefer recent and recurring evidence over a skills-list mention.
- Prefer a broad exact requirement such as "Relational Database" over selecting
  Oracle + DB2 + SQL Server + MySQL individually, unless a specific database is
  itself a critical screening requirement.
- Prefer "Spring Framework" over also consuming another slot with "Spring Boot"
  when the formal requirement is Spring Framework, unless Spring Boot itself is
  separately important to the role.
- Do not spend limited attention on generic soft skills (Problem Solving, Analytical
  Thinking, Critical Thinking, Writing, Communication, etc.) unless the role or
  manager makes that specific competency unusually central AND it is more
  valuable than a technical/domain requirement.
- Do not spend limited attention on adjacent candidate strengths that are not central
  to the requisition.
- Do not use one skill as a substitute for a different required skill
  (for example Maven is not proof of Gradle).
- Avoid redundant parent/child/synonym selections. Because every selected skill should add distinct value,
  coverage of distinct explicit requirements is usually more valuable than a
  second closely related variant.
- TEST-AUTOMATION REDUNDANCY RULE: never recommend more than ONE of these when
  they represent the same candidate capability:
  `Automation`, `Automation Testing`, `Test Automation`,
  `Test Automation Tools`, `Automated Testing`.
  Prefer the most role-specific exact value the candidate truly supports.
  For a QA automation role, normally prefer `Test Automation` over generic
  `Automation`.
- DATA-QUALITY REDUNDANCY RULE: normally choose one of `Data Quality` and
  `data validation` unless the requisition clearly treats them as materially
  different screening dimensions.
- FINANCE REDUNDANCY RULE: normally choose one of `Finance` and
  `Financial Services`.
- SOURCE-CONTROL REDUNDANCY RULE: do not consume multiple slots with Git/GitHub/
  GitLab unless a specific platform is independently central to the role.
- Example: when Spring Framework is already selected and JDBC/J2EE are explicit
  candidate-proven requirements, do not spend another slot on Spring Boot merely
  because it is a candidate strength.
- Candidate evidence must genuinely support every selection.

STRUCTURED-SKILL RULE:
- If source_category is `exact_structured_must`, skill_name MUST come from:
  {selected_must_json}
- If source_category is `exact_structured_nice`, skill_name MUST come from:
  {selected_nice_json}
- Long/legacy/free-text catalogue values are allowed ONLY in those two exact
  structured categories.
- For all other categories, use clean `canonical_preferred` values first and
  `compound_review` only when materially more precise. Never use
  `legacy_free_text`.

SOURCE CATEGORY must be exactly one of:
- exact_structured_must
- required_job_intelligence
- manager_priority
- exact_structured_nice
- preferred_job_intelligence
- jd_supported_coverage
- additional_high_value

EVIDENCE FIELDS - REQUIRED FOR EVERY RECOMMENDATION:
- `evidence` = ONE short VERBATIM candidate-evidence excerpt proving the skill.
  This may come from the original resume or candidate-confirmed clarification.
  Do not paraphrase it and do not cite the JD/Spotlight here.
- `role_evidence` = ONE short VERBATIM excerpt from the Official JD, Spotlight
  transcript, or MSP/VNDLY notes showing why the skill matters to THIS role.
  `additional_high_value` may leave role_evidence blank when the value is
  candidate-specific rather than an explicit requisition requirement.
- A role requirement is never candidate evidence and candidate evidence is never
  proof that Freddie requires the skill. Keep those two functions separate.

BEFORE RETURNING:
Ask yourself: "Which candidate-supported skills materially improve this
submission and are worth selecting manually in VNDLY?"

There is NO fixed maximum. Return every genuinely supported, role-relevant skill
that adds meaningful shortlist/search value, then stop before the list becomes
noise. Do not add weak or tangential skills merely because the candidate has them.

Then perform this audit:
1. Remove every skill that is supported only by the JOB and not by the candidate.
2. Collapse synonym/redundancy families to the single strongest label.
3. Confirm that concrete candidate-supported Must Haves (for example Python,
   SQL, JDBC, Snowflake, JSON, Cucumber, etc. when applicable) are not being
   crowded out by broad labels such as Automation, DevOps, Data Engineering, or
   Data Validation.
4. Prefer dated work-history evidence over a top-level skills-list mention.
5. Do not omit a clearly work-history-evidenced Required skill merely to include
   a weaker summary-only Required skill.

======================================================================
OUTPUT
======================================================================

Return ONLY:

{{
  "VNDLY_SUMMARY": "",
  "VNDLY_SUMMARY_EVIDENCE": [""],
  "FINAL_RECOMMENDED_VNDLY_SKILLS": [
    {{
      "skill_name": "",
      "source_category": "",
      "evidence": "",
      "role_evidence": "",
      "reason": ""
    }}
  ]
}}
"""

    raw_package = fred_generate_json(
        api_key,
        package_prompt,
    )

    if not isinstance(raw_package, dict):
        raise RuntimeError(
            "Gemini returned an invalid VNDLY package payload."
        )

    vndly_summary = fred_clean_vndly_submission_summary(
        raw_package.get("VNDLY_SUMMARY", ""),
        candidate_name,
    )
    vndly_summary = fred_suppress_total_career_years(
        vndly_summary,
        calculated_total_experience,
        vndly_context,
    )
    vndly_summary_evidence = fred_string_list(
        raw_package.get("VNDLY_SUMMARY_EVIDENCE", [])
    )

    if not vndly_summary:
        raise RuntimeError(
            "Gemini returned no VNDLY submission summary."
        )

    vndly_summary_errors = fred_validate_vndly_summary(
        vndly_summary,
        vndly_summary_evidence,
        vndly_summary_candidate_evidence,
        structured_req,
        vndly_context,
    )

    if vndly_summary_errors:
        repaired_summary = fred_repair_vndly_summary(
            api_key,
            candidate_name,
            vndly_summary,
            vndly_summary_errors,
            vndly_summary_candidate_evidence,
            structured_req,
            vndly_context,
            calculated_total_experience,
            candidate_positioning_title,
        )

        if isinstance(repaired_summary, dict):
            vndly_summary = fred_clean_vndly_submission_summary(
                repaired_summary.get("VNDLY_SUMMARY", ""),
                candidate_name,
            )
            vndly_summary = fred_suppress_total_career_years(
                vndly_summary,
                calculated_total_experience,
                vndly_context,
            )
            vndly_summary_evidence = fred_string_list(
                repaired_summary.get("VNDLY_SUMMARY_EVIDENCE", [])
            )

        vndly_summary_errors = fred_validate_vndly_summary(
            vndly_summary,
            vndly_summary_evidence,
            vndly_summary_candidate_evidence,
            structured_req,
            vndly_context,
        )

    if vndly_summary_errors:
        hard_vndly_errors = [
            error for error in vndly_summary_errors
            if fred_grounding_error_is_hard(error)
        ]
        if hard_vndly_errors:
            second_repair = fred_repair_vndly_summary(
                api_key,
                candidate_name,
                vndly_summary,
                hard_vndly_errors,
                vndly_summary_candidate_evidence,
                structured_req,
                vndly_context,
                calculated_total_experience,
                candidate_positioning_title,
            )
            if isinstance(second_repair, dict):
                candidate = fred_clean_vndly_submission_summary(
                    second_repair.get("VNDLY_SUMMARY", ""),
                    candidate_name,
                )
                candidate = fred_suppress_total_career_years(
                    candidate, calculated_total_experience, vndly_context
                )
                candidate_evidence = fred_string_list(
                    second_repair.get("VNDLY_SUMMARY_EVIDENCE", [])
                )
                candidate_errors = fred_validate_vndly_summary(
                    candidate,
                    candidate_evidence,
                    vndly_summary_candidate_evidence,
                    structured_req,
                    vndly_context,
                )
                candidate_hard = [
                    error for error in candidate_errors
                    if fred_grounding_error_is_hard(error)
                ]
                if not candidate_hard and candidate:
                    vndly_summary = candidate
                    vndly_summary_evidence = candidate_evidence
                    hard_vndly_errors = []

        if hard_vndly_errors:
            # Use the already-grounded coherent resume Summary rather than
            # splicing raw resume fragments. This preserves readable English.
            vndly_summary = fred_clean_vndly_submission_summary(
                resume_summary, candidate_name
            )
            vndly_summary = fred_suppress_total_career_years(
                vndly_summary, calculated_total_experience, vndly_context
            )
            print(
                "Freddie VNDLY hard safety fallback used coherent resume summary: "
                + " | ".join(hard_vndly_errors[:6])
            )
        else:
            print(
                "Freddie VNDLY soft grounding diagnostics accepted after repair: "
                + " | ".join(vndly_summary_errors[:6])
            )

    ranked_skill_pool = fred_build_final_ranked_vndly_skills(
        raw_package.get(
            "FINAL_RECOMMENDED_VNDLY_SKILLS",
            [],
        ),
        selected_must_items,
        selected_nice_items,
        catalog,
        cap=FREDDIE_VNDLY_SKILL_POOL_LIMIT,
    )

    ranked_skill_pool = fred_supplement_priority_skills(
        ranked_skill_pool,
        structured_req,
        catalog,
        candidate_evidence_text,
        selected_must_items=selected_must_items,
        selected_nice_items=selected_nice_items,
        vndly_context=vndly_context,
    )

    evidence_ranked_skills = fred_rerank_vndly_skills_by_evidence(
        ranked_skill_pool,
        candidate_evidence_text,
        structured_req=structured_req,
        vndly_context=vndly_context,
        cap=FREDDIE_VNDLY_SKILL_POOL_LIMIT,
    )

    ranked_skills = fred_apply_vndly_redundancy_control(
        evidence_ranked_skills,
        structured_req,
        cap=None,
    )

    grouped_skills = fred_split_ranked_vndly_skills(
        ranked_skills
    )

    return {
        "candidate_positioning_title": str(
            candidate_positioning_title or ""
        ).strip(),
        "vndly_summary": vndly_summary,
        "final_recommended_vndly_skills": ranked_skills,
        **grouped_skills,
    }



def fred_skill_names(items):
    names = []
    for item in (items or []):
        if not isinstance(item, dict):
            continue
        name = str(item.get("skill_name", "") or "").strip()
        if name:
            names.append(name)
    return names


def fred_build_submission_email_body(
    candidate_name,
    vndly_context,
    package,
):
    """Plain-text ready-to-use Freddie submission email."""
    vndly_context = vndly_context or {}
    package = package or {}

    job_id = str(vndly_context.get("job_id", "") or "").strip()
    job_title = str(vndly_context.get("job_title", "") or "").strip()
    manager = str(vndly_context.get("resource_manager", "") or "").strip()

    ranked_skills = package.get(
        "final_recommended_vndly_skills",
        [],
    ) or []

    source_labels = {
        "exact_structured_must": "Exact Structured Must",
        "exact_structured_nice": "Exact Structured Nice",
        "required_job_intelligence": "Required",
        "manager_priority": "Manager Priority",
        "preferred_job_intelligence": "Preferred",
        "jd_supported_coverage": "JD Coverage",
        "additional_high_value": "Additional",
    }

    if ranked_skills:
        skill_lines = []
        for index, item in enumerate(ranked_skills, start=1):
            name = str(item.get("skill_name", "") or "").strip()
            category = str(
                item.get("source_category", "") or ""
            ).strip().lower()
            label = source_labels.get(category, "Recommended")
            if name:
                skill_lines.append(
                    f"{index}. {name} [{label}]"
                )
        skills_text = "\n".join(skill_lines)
    else:
        skills_text = "None recommended"

    lines = [
        f"Candidate: {candidate_name}",
        f"Freddie Job: {job_id}" + (f" - {job_title}" if job_title else ""),
        f"Resource Manager: {manager or 'N/A'}",
        "",
        "VNDLY COMMENTS",
        str(package.get("vndly_summary", "") or "").strip(),
        "",
        "RECOMMENDED VNDLY SKILLS",
        skills_text,
    ]

    return "\n".join(lines).strip()


def fred_send_submission_email(
    candidate_name,
    vndly_context,
    package,
):
    """
    Send the Freddie submission metadata through Gmail SMTP.
    Credentials live only in Streamlit Secrets.
    """
    sender = str(
        st.secrets.get(
            "FREDDIE_EMAIL_SENDER",
            "kennyconv@gmail.com",
        )
    ).strip()
    recipient = str(
        st.secrets.get(
            "FREDDIE_EMAIL_RECIPIENT",
            "kkerrigan@conv.com",
        )
    ).strip()
    app_password = str(
        st.secrets.get(
            "FREDDIE_EMAIL_APP_PASSWORD",
            "",
        )
    ).replace(" ", "").strip()

    if not sender or not recipient or not app_password:
        raise RuntimeError(
            "Freddie email Secrets are incomplete. Expected "
            "FREDDIE_EMAIL_SENDER, FREDDIE_EMAIL_RECIPIENT, and "
            "FREDDIE_EMAIL_APP_PASSWORD."
        )

    job_id = str((vndly_context or {}).get("job_id", "") or "").strip()
    job_title = str((vndly_context or {}).get("job_title", "") or "").strip()

    subject_parts = [
        "Freddie VNDLY Submission",
        job_id,
        candidate_name,
    ]
    subject = " - ".join(
        part for part in subject_parts if str(part).strip()
    )

    if job_title:
        subject += f" ({job_title})"

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(
        fred_build_submission_email_body(
            candidate_name,
            vndly_context,
            package,
        )
    )

    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
        timeout=30,
    ) as smtp:
        smtp.login(sender, app_password)
        smtp.send_message(msg)


def fred_generation_email_key(
    candidate_name,
    job_id,
    raw_resume_text,
    vndly_summary,
):
    """
    Stable session key preventing duplicate sends for the same generated package.
    """
    payload = "||".join(
        [
            str(candidate_name or ""),
            str(job_id or ""),
            str(raw_resume_text or ""),
            str(vndly_summary or ""),
        ]
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def fred_repair_and_load_json(raw_text):
    """Freddie-only JSON repair without global Streamlit warning side effects."""
    text = str(raw_text or "").strip()
    text = re.sub(r'^```[a-zA-Z]*\n|\n```$', '', text, flags=re.MULTILINE).strip()
    match = re.search(r'\{.*\}', text, re.DOTALL)
    if match:
        text = match.group()

    candidates = [text]

    repaired = text
    repaired = re.sub(r'"\s*\n\s*"', '",\n"', repaired)
    repaired = re.sub(r'\}\s*\n\s*\{', '},\n{', repaired)
    repaired = re.sub(r'\]\s*\n\s*"', '],\n"', repaired)
    repaired = re.sub(r',\s*([}\]])', r'\1', repaired)
    # Gemini occasionally emits an otherwise-valid object with an unquoted
    # structural key. Quote only key-shaped tokens immediately after { or ,.
    repaired = re.sub(
        r'([\{,]\s*)([A-Za-z_][A-Za-z0-9_ ]*?)(\s*:)',
        lambda m: m.group(1) + json.dumps(m.group(2).strip()) + m.group(3),
        repaired,
    )
    if repaired != text:
        candidates.append(repaired)

    for candidate in candidates:
        try:
            data = json.loads(candidate, strict=False)
            if isinstance(data, dict):
                return data
        except Exception:
            continue

    return {}


def fred_generate_json(api_key, prompt, max_attempts=6):
    """
    Freddie-specific Gemini JSON helper with the same retry philosophy
    already used throughout the app.
    """
    client = genai.Client(api_key=api_key)

    models_to_try = [
        "gemini-2.5-flash",
        "gemini-3-flash-preview",
        "gemini-3.1-flash-lite-preview",
        "gemini-pro-latest",
    ]

    last_error = None

    for attempt in range(max_attempts):
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
                ),
            )

            data = fred_repair_and_load_json(response.text)

            if data:
                return data

            last_error = ValueError(
                "Gemini returned malformed or unusable Freddie Mac JSON."
            )

        except Exception as e:
            last_error = e

            if "503" in str(e) and attempt < max_attempts - 1:
                time.sleep(2 ** ((attempt % 2) + 1))
                continue

            if attempt < max_attempts - 1:
                time.sleep(1)
                continue

    raise RuntimeError(
        f"Gemini could not return usable Freddie Mac JSON after retries. "
        f"Last error: {last_error}"
    )


def fred_extract_vetting_block(job_description):
    """
    Deterministic first-pass locator for Freddie's Supplier Vetting Questions.
    This does NOT decide what constitutes a question. It only gives Gemini
    a high-signal excerpt while Gemini also receives the complete JD.
    """
    if not job_description:
        return ""

    text = str(job_description).replace("\r\n", "\n").replace("\r", "\n")

    heading_pattern = re.compile(
        r"(?im)^[ \t]*"
        r"(?:required[ \t]+)?"
        r"supplier[ \t]+vetting[ \t]+questions?"
        r"[ \t]*(?:[:\-–\u2014])[ \t]*(.*)?$"
    )

    match = heading_pattern.search(text)

    if not match:
        return ""

    start = match.start()

    # Common Freddie headings that usually terminate the vetting block.
    stop_pattern = re.compile(
        r"(?im)^[ \t]*("
        r"job description|"
        r"position overview|"
        r"position description|"
        r"responsibilities|"
        r"key responsibilities|"
        r"required qualifications|"
        r"preferred qualifications|"
        r"qualifications|"
        r"must have|"
        r"must-have|"
        r"nice to have|"
        r"nice-to-have|"
        r"keys to success|"
        r"key success in the role"
        r")[ \t]*:?[ \t]*$"
    )

    remainder = text[match.end():]
    stop_match = stop_pattern.search(remainder)

    if stop_match:
        end = match.end() + stop_match.start()
    else:
        # Safety cap. Gemini still receives the full JD separately.
        end = min(len(text), start + 5000)

    return text[start:end].strip()


def fred_normalize_question_list(raw_questions):
    """
    Sanitizes Gemini's vetting-question output without rewriting its substance.
    """
    if not isinstance(raw_questions, list):
        return []

    cleaned = []

    for item in raw_questions:
        if isinstance(item, dict):
            question = (
                item.get("question")
                or item.get("text")
                or item.get("instruction")
                or ""
            )
        else:
            question = str(item)

        question = str(question).strip()

        # Strip duplicated numbering only. Do NOT rewrite the wording.
        question = re.sub(
            r"^\s*(?:question\s*)?\d+\s*[\.\)\:\-]\s*",
            "",
            question,
            flags=re.IGNORECASE,
        ).strip()

        if question and question.lower() not in {
            "n/a",
            "na",
            "none",
            "not applicable",
            "no",
        }:
            cleaned.append(question)

    return cleaned


def fred_vetting_response_required(question):
    """
    Determine whether an official Freddie Supplier Vetting item requires
    a candidate response.

    Conservative rule:
    - Explicitly optional / not-required items do NOT require a response.
    - Everything else defaults to REQUIRED.
    """

    text = str(question or "").strip().lower()

    if not text:
        return False

    optional_phrases = [
        "not required",
        "not mandatory",
        "optional",
        "encouraged to submit",
        "encouraged but not required",
        "may submit",
    ]

    if any(phrase in text for phrase in optional_phrases):
        return False

    return True


def fred_source_has_substantive_selection_intelligence(source_text):
    """
    Freddie-only deterministic safeguard.

    Returns True only when MSP/VNDLY notes or Spotlight transcript appear to
    contain substantive candidate-profile / selection guidance rather than
    routine staffing operations such as scheduling, offers, OOO status,
    backfills, reposting, or "JD updated" notices.
    """
    text = str(source_text or "").strip().lower()

    if not text:
        return False

    substantive_patterns = [
        r"\bmanager\b.{0,120}\b(?:need|needs|require|requires|required|want|wants|looking|focus|priority|critical|emphas)",
        r"\b(?:must|need|needs|required|requires|requirement|critical|priority)\b.{0,120}\b(?:candidate|experience|background|skill|technical|functional|business|domain|java|python|angular|sql|mortgage|finance|operations|frontend|backend|full[- ]?stack)",
        r"\b(?:candidate|candidates)\b.{0,120}\b(?:lacked|lack|missing|too|weak|stronger|depth|explain|demonstrate|show|not vetted|move forward|advance)",
        r"\b(?:interview|selection|shortlist)\b.{0,120}\b(?:feedback|lacked|missing|need|needs|require|stronger|depth|explain|demonstrate)",
        r"\b(?:not the focus|no longer required|not required|deemphasized|de-emphasized|instead of|rather than|do not use|don't use|does not use|not used|nice to have|helpful|optional|separate role)\b",
        r"\b(?:requirements?|reporting|stakeholders?|program|project|portfolio|product|operations?|communications?|design|legal|hr|recruiting|facilities|data|security)\b.{0,120}\b(?:priority|important|critical|required|must|need|needs|want|wants|looking|helpful|nice to have|not required|not used)",
        r"\b\d{1,3}\s*%\b.{0,80}\b(?:ui|frontend|front-end|backend|back-end|technical|functional|business)",
    ]

    return any(re.search(pattern, text, flags=re.IGNORECASE | re.DOTALL)
               for pattern in substantive_patterns)


def fred_cap_skill_years(value, total_experience):
    """
    Freddie-only hard ceiling: no skill/domain tenure may exceed the
    deterministic total professional experience.
    """
    value_text = str(value or "").strip()
    total_text = str(total_experience or "").strip()

    if not value_text or not total_text:
        return value_text

    value_match = re.search(r"(\d+)\s*\+\s*years", value_text, flags=re.IGNORECASE)
    total_match = re.search(r"(\d+)\s*\+\s*years", total_text, flags=re.IGNORECASE)

    if not value_match or not total_match:
        return value_text

    skill_years = int(value_match.group(1))
    total_years = int(total_match.group(1))

    if skill_years <= total_years:
        return value_text

    start, end = value_match.span(1)
    return value_text[:start] + str(total_years) + value_text[end:]


@st.cache_data(ttl=3600, show_spinner=False)
def fred_analyze_requisition(
    _api_key,
    job_description,
    msp_notes="",
    spotlight_transcript="",
    vndly_context=None,
):
    """
    Analyze the Freddie requisition ONCE and return structured job intelligence.

    Freddie-only design:
    - Official VNDLY JD + structured VNDLY fields establish formal requirements.
    - Dated Spotlight / MSP notes establish current manager emphasis and changes.
    - Internal Notes provide recruiter context, but cannot manufacture requirements.
    - Operational staffing noise is deliberately separated from selection intelligence.

    Cached by the JD/notes/transcript/context so Streamlit does not re-call Gemini
    whenever a recruiter types into another field.
    """
    vndly_context = vndly_context or {}

    if not job_description or not str(job_description).strip():
        return {
            "target_title": "",
            "explicit_must_have_requirements": [],
            "must_have_competency_groups": [],
            "required_requirements": [],
            "preferred_requirements": [],
            "domain_requirements": [],
            "current_manager_priorities": [],
            "selection_feedback": [],
            "role_clarifications": [],
            "role_evolution": [],
            "deemphasized_requirements": [],
            "explicitly_excluded_requirements": [],
            "deemphasized_or_negated_requirements": [],
            "supplier_vetting_questions": [],
        }

    vetting_hint = fred_extract_vetting_block(job_description)

    # Only the role-relevant structured VNDLY fields are sent to Gemini here.
    # Competition/admin fields remain available in the app but do not become
    # candidate requirements.
    matching_context = {
        "job_id": vndly_context.get("job_id", ""),
        "job_title": vndly_context.get("job_title", ""),
        "resource_manager": vndly_context.get("resource_manager", ""),
        "organization_unit": vndly_context.get("organization_unit", ""),
        "cost_center": vndly_context.get("cost_center", ""),
        "workday_cost_center": vndly_context.get("workday_cost_center", ""),
        "job_category": vndly_context.get("job_category", ""),
        "min_experience": vndly_context.get("min_experience", ""),
        "max_experience": vndly_context.get("max_experience", ""),
        "must_have_skills": vndly_context.get("must_have_skills", ""),
        "nice_to_have_skills": vndly_context.get("nice_to_have_skills", ""),
        "work_site_name": vndly_context.get("work_site_name", ""),
        "city": vndly_context.get("city", ""),
        "state": vndly_context.get("state", ""),
        "time_type": vndly_context.get("time_type", ""),
        "mnpi": vndly_context.get("mnpi", ""),
        "internal_notes": vndly_context.get("internal_notes", ""),
    }

    matching_context_json = json.dumps(
        matching_context,
        indent=2,
        ensure_ascii=False,
        default=str,
    )

    prompt = f"""
Return a valid JSON object ONLY.

You are analyzing a Freddie Mac contingent-worker requisition from Workday VNDLY.

Your job in this step is NOT to evaluate a candidate and NOT to write resume
content. Your only job is to convert the requisition into accurate structured
job intelligence that will later be used for candidate matching.

======================================================================
SOURCE HIERARCHY
======================================================================

You are receiving FOUR kinds of role intelligence:

1. OFFICIAL VNDLY JOB DESCRIPTION
2. STRUCTURED VNDLY JOB FIELDS
3. DATED MSP / VNDLY NOTES and SPOTLIGHT CALL TRANSCRIPTS
4. INTERNAL RECRUITER NOTES, if present

Use them differently.

FORMAL REQUIREMENTS:
- The Official VNDLY Job Description and populated structured VNDLY
  Must Have Skills / Nice To Have Skills / experience fields establish formal
  job requirements.
- Structured "Must Have Skills" are formal requirements when populated.
- Structured "Nice To Have Skills" are preferred requirements when populated.
- Do NOT promote a Nice To Have into a Must Have.
- First look in the JD for Freddie's explicitly labeled "Must Have
  Qualifications", "Must Have", "Must-Have", or equivalent section.
- Treat explicitly labeled Must Haves as primary formal matching criteria.
- Classify requirements at the INDIVIDUAL CLAUSE level, even when several
  requirements appear inside one paragraph.
- Wording such as "preferred", "nice to have", "a plus", "desired", or
  "preferred but not required" belongs under preferred_requirements even if
  physically located inside a Must Have paragraph.
- EXPLICIT OPTIONALITY OVERRIDES RESPONSIBILITY-SECTION MENTION: if Freddie or
  the hiring manager explicitly says a competency is "nice to have", "helpful",
  "not required", or otherwise optional, do NOT later promote that same
  competency to Required merely because it also appears inside Responsibilities.
  Example: if the header says Agile experience is nice to have and the body says
  the worker will support Agile planning, Agile remains Preferred unless a newer
  authoritative manager statement explicitly makes it mandatory.
- Only clearly mandatory items may appear in explicit_must_have_requirements.
- Scan the ENTIRE Official VNDLY Job Description for explicit mandatory language,
  not only the formally labeled Must Have section. Any statement anywhere in the
  JD using language such as "must have", "required", "minimum", "requires",
  "need(s)", or an explicit minimum-years requirement must be retained as a
  formal requirement when it clearly applies to the candidate profile.
- Example: if Responsibilities says "Must have 8+ years of experience", include
  that exact requirement in explicit_must_have_requirements even if it appears
  outside the labeled Must Have Qualifications section.
- If there is no explicit Must Have section, derive core required competencies
  from Required Qualifications and central recurring responsibilities.

ACTUAL HIRING-MANAGER / SELECTION EMPHASIS:
- Spotlight calls and role-relevant MSP/VNDLY clarifications can reveal which
  formal requirements matter most in practice.
- PROVENANCE GATE: current_manager_priorities, selection_feedback, and
  role_clarifications may ONLY contain information explicitly supported by the
  DATED MSP / VNDLY NOTES or DATED SPOTLIGHT CALL TRANSCRIPT.
- NEVER populate those three fields by copying, paraphrasing, or inferring from
  the Official VNDLY Job Description or structured VNDLY fields alone.
- If the notes/transcript contain only operational updates and no substantive
  manager/selection guidance, return [] for all three fields.
- Give especially high weight to explicit feedback explaining WHY prior
  candidates were not shortlisted, did not advance, or failed interviews, and
  what future submissions must demonstrate.
- Examples of high-signal selection intelligence include:
  "too IT-PM heavy", "needs stronger Angular", "lacked technical depth",
  "must be able to explain technologies on the resume", "mortgage required",
  "business operations PM is the critical requirement", or "model building is
  not the focus".
- Separate LOWER PRIORITY from TRUE EXCLUSION.
- If the manager says a JD item is not used by this team, not central, optional,
  not required, outside the day-to-day role, or belongs to a separate role, put
  it under `deemphasized_requirements`. This lowers hiring-manager emphasis but
  does NOT erase a genuine candidate/JD match that may still help MSP review.
- Example: JD says user stories but the manager says "we don't do user stories".
  Preserve User Stories as a formal JD artifact and put it under
  `deemphasized_requirements`. A candidate who genuinely has User Stories may
  still show it as secondary JD/MSP coverage, but it must not displace the
  manager's actual priorities.
- If the manager explicitly says candidates SHOULD NOT have/emphasize something,
  that it is disqualifying, unacceptable, or a negative profile signal, put it
  under `explicitly_excluded_requirements`. Only this field is treated as a true
  candidate-facing prohibition.
- If the manager says a prior PMO-setup role is separate and this person does not
  need to build a PMO from scratch, treat PMO setup as de-emphasized, not as a
  candidate disqualifier.
- A manager statement that something is helpful/nice-to-have belongs under
  current_manager_priorities or role_clarifications as appropriate, but does NOT
  become a formal Required requirement unless the manager explicitly says it is
  mandatory.
- Do NOT silently erase formal JD requirements. Keep formal requirements and
  current manager priorities as separate concepts.

======================================================================
RECENCY / ROLE-EVOLUTION RULE - CRITICAL
======================================================================

The MSP/VNDLY Notes and Spotlight Call Transcript may contain multiple dated
entries over the lifecycle of one requisition.

- Treat the date attached to each entry as meaningful.
- When two entries conflict, a NEWER explicit hiring-manager clarification or
  selection-feedback statement normally represents the current direction.
- Do NOT let an older Spotlight statement override a newer manager/MSP
  clarification.
- When a newer update merely reports scheduling or staffing activity and does
  NOT change the candidate profile, it must NOT override older substantive role
  guidance.
- Capture only SUBSTANTIVE candidate-profile changes over time in role_evolution.
- HARD PROVENANCE GATE: role_evolution may ONLY be populated when the DATED
  MSP / VNDLY NOTES or DATED SPOTLIGHT CALL TRANSCRIPT explicitly states WHAT
  changed in the desired candidate profile.
- Do NOT create a role_evolution item merely because a note says the job was
  updated, refreshed, reposted, re-released, backfilled, replaced, or that a
  prior offer was rejected. Those are operational/history facts unless the
  source itself states the substantive change.
- NEVER infer that an updated/reposted/backfill JD "reflects the latest candidate
  profile" or similar. If the source does not name the changed skill, weighting,
  scope, domain, seniority, responsibility, or manager expectation, return [].
- A role_evolution item is appropriate only when the source identifies a
  meaningful change in required skills, weighting, scope, domain background,
  seniority/experience, responsibilities, or manager expectations.
- Example of VALID evolution: "Manager clarified the role is 40% UI / 60%
  backend and now requires stronger Angular."
- Example of INVALID evolution: "Backfill release; JD has been updated" when no
  substantive change is described.
- current_manager_priorities must represent what appears to matter NOW after
  considering chronology.
- role_clarifications should capture concise current clarifications such as
  "40% UI / 60% backend" or "business operations PM, not IT PM".
- selection_feedback should capture evidence about why prior submissions or
  interviewees did or did not advance.

======================================================================
OPERATIONAL-NOISE FILTER - CRITICAL
======================================================================

MSP/VNDLY notes mix job-selection intelligence with staffing-process updates.

Do NOT turn any of the following into candidate requirements unless the wording
explicitly changes or clarifies the desired candidate profile:

- manager OOO / availability
- interview scheduling or placeholders
- debrief scheduling
- shortlisting deadlines by themselves
- offer status / offer processing
- candidate counts
- positions filled / slots added
- keeping candidates warm
- supplier administration
- vendor distribution timing
- role hold / closure risk
- routine follow-up status
- billing/rate administration

A sentence such as "interviews scheduled Monday" is operational noise.

A sentence such as "manager rejected candidates because they were too
backend-heavy and now requires stronger Angular" is HIGH-SIGNAL selection
intelligence.

======================================================================
INTERNAL NOTES RULE
======================================================================

Internal Notes are recruiter context, not an authoritative Freddie source.

Allowed:
- use them to resolve ambiguity,
- surface known recruiter context,
- improve emphasis when consistent with official/manager evidence.

Not allowed:
- manufacture a formal requirement,
- override an explicit hiring-manager statement with speculation,
- assign an unsupported skill to the role merely because an internal note
  speculates about it.

======================================================================
STRUCTURED VNDLY CONTEXT RULE
======================================================================

The structured VNDLY fields may help interpret the role.

- Job Title, Min/Max Experience, Must Have Skills, Nice To Have Skills are
  direct matching inputs.
- Organization Unit, Cost Center, Workday Cost Center, Job Category, Work Site,
  Time Type, and MNPI are CONTEXTUAL evidence only.
- Contextual metadata may help interpret an ambiguous JD, but it cannot create
  a candidate requirement absent from formal or manager guidance.
- Resource Manager is identification/context only.

======================================================================
MUST-HAVE COMPETENCY GROUPS
======================================================================

Create up to FOUR logical competency groups representing the highest-signal
current Must Have requirements.

Examples:
"Application Development (Python, Java, Spring Boot)"
"Data & Analytics (SQL, Snowflake, JSON)"
"ServiceNow Testing & ATF"
"Mortgage / Financial Services"

Rules:
- Group closely related Freddie requirements together.
- Preserve Freddie's exact terminology wherever practical.
- Do not invent a technology or competency.
- These are JOB requirements only. Do not evaluate a candidate in this step.
- A group may contain technologies, methodologies, domain expertise, or
  operational competencies.
- Rank groups using the current requirement hierarchy after applying substantive
  dated manager clarifications.

Return each group as:
{{
    "name": "",
    "terms": [],
    "reason": ""
}}

======================================================================
ROLE EVOLUTION FORMAT
======================================================================

Only include material changes or clarifications.

Each role_evolution item must be:
{{
    "date": "YYYY-MM-DD or best available date",
    "change": "Concise description of the substantive change/clarification"
}}

Do NOT add routine operational updates to role_evolution.

======================================================================
SUPPLIER VETTING QUESTIONS - ZERO INVENTION
======================================================================

Extract ONLY supplier vetting questions/instructions actually supplied by
Freddie/MSP in the Official Job Description.

Rules:
- Preserve substantive wording and original order.
- Remove only leading numbering such as "1." or "Question 1:".
- Correcting obvious whitespace is allowed.
- Do NOT rewrite, improve, expand, split, or invent questions.
- If Freddie provides one assessment instruction rather than conventional
  questions, return that instruction as ONE item.
- If the JD says Supplier Vetting Questions are N/A / None, return [].
- Do NOT mistake interview information, candidate-template questions,
  qualifications, or ordinary JD bullets for Supplier Vetting Questions.
- The standard candidate-template questions at the top of Freddie JDs
  (Former CW, work authorization, location, etc.) are NOT supplier vetting
  questions.

A deterministic parser found this possible vetting block. It is ONLY a hint;
validate it against the full JD:

--- POSSIBLE VETTING BLOCK ---
{vetting_hint}
--- END POSSIBLE VETTING BLOCK ---

======================================================================
OUTPUT FORMAT
======================================================================

Return ONLY:

{{
  "target_title": "",
  "explicit_must_have_requirements": [""],
  "must_have_competency_groups": [
    {{
      "name": "",
      "terms": [""],
      "reason": ""
    }}
  ],
  "required_requirements": [""],
  "preferred_requirements": [""],
  "domain_requirements": [""],
  "current_manager_priorities": [""],
  "selection_feedback": [""],
  "role_clarifications": [""],
  "role_evolution": [
    {{
      "date": "",
      "change": ""
    }}
  ],
  "deemphasized_requirements": [""],
  "explicitly_excluded_requirements": [""],
  "deemphasized_or_negated_requirements": [""],
  "supplier_vetting_questions": [""]
}}

======================================================================
INPUTS
======================================================================

STRUCTURED VNDLY JOB FIELDS:
{matching_context_json}

OFFICIAL VNDLY JOB DESCRIPTION:
{job_description}

DATED MSP / VNDLY NOTES:
{msp_notes}

DATED SPOTLIGHT CALL TRANSCRIPT:
{spotlight_transcript}
"""

    data = fred_generate_json(_api_key, prompt)

    required_keys = {
        "target_title",
        "explicit_must_have_requirements",
        "must_have_competency_groups",
        "required_requirements",
        "preferred_requirements",
        "domain_requirements",
        "current_manager_priorities",
        "selection_feedback",
        "role_clarifications",
        "role_evolution",
        "deemphasized_requirements",
        "explicitly_excluded_requirements",
        "deemphasized_or_negated_requirements",
        "supplier_vetting_questions",
    }

    missing_keys = required_keys - set(data.keys())

    if missing_keys:
        raise RuntimeError(
            "Freddie requisition analysis returned incomplete JSON. "
            f"Missing fields: {', '.join(sorted(missing_keys))}"
        )

    data["supplier_vetting_questions"] = fred_normalize_question_list(
        data.get("supplier_vetting_questions", [])
    )

    for list_key in [
        "explicit_must_have_requirements",
        "required_requirements",
        "preferred_requirements",
        "domain_requirements",
        "current_manager_priorities",
        "selection_feedback",
        "role_clarifications",
        "deemphasized_or_negated_requirements",
    ]:
        value = data.get(list_key, [])
        data[list_key] = value if isinstance(value, list) else []

    role_evolution = data.get("role_evolution", [])
    if not isinstance(role_evolution, list):
        role_evolution = []

    normalized_evolution = []
    for item in role_evolution:
        if not isinstance(item, dict):
            continue

        change = str(item.get("change", "") or "").strip()
        change_date = str(item.get("date", "") or "").strip()

        if change:
            normalized_evolution.append(
                {
                    "date": change_date,
                    "change": change,
                }
            )

    data["role_evolution"] = normalized_evolution

    # ================================================================
    # FREDDIE-ONLY PROVENANCE SAFEGUARD
    # ================================================================
    # Manager-priority / selection / clarification / evolution fields are
    # allowed only when MSP notes or the Spotlight transcript actually contain
    # substantive candidate-profile guidance. Formal JD requirements remain in
    # their own formal-requirement buckets.
    substantive_manager_source = (
        fred_source_has_substantive_selection_intelligence(msp_notes)
        or fred_source_has_substantive_selection_intelligence(spotlight_transcript)
    )

    if not substantive_manager_source:
        data["current_manager_priorities"] = []
        data["selection_feedback"] = []
        data["role_clarifications"] = []
        data["role_evolution"] = []

    groups = data.get("must_have_competency_groups", [])
    if not isinstance(groups, list):
        groups = []

    normalized_groups = []

    for group in groups[:4]:
        if not isinstance(group, dict):
            continue

        name = str(group.get("name", "")).strip()
        terms = group.get("terms", [])

        if not isinstance(terms, list):
            terms = []

        terms = [str(x).strip() for x in terms if str(x).strip()]

        if name:
            normalized_groups.append(
                {
                    "name": name,
                    "terms": terms,
                    "reason": str(group.get("reason", "")).strip(),
                }
            )

    data["must_have_competency_groups"] = normalized_groups

    # Backward-compatible mirror for any older downstream code/cache. v11 logic
    # uses the separated fields above.
    if not data.get("deemphasized_requirements") and data.get("deemphasized_or_negated_requirements"):
        data["deemphasized_requirements"] = list(
            data.get("deemphasized_or_negated_requirements", []) or []
        )
    data["deemphasized_or_negated_requirements"] = list(
        data.get("deemphasized_requirements", []) or []
    )

    return data


def fred_grounding_audit_prompt(
    resume_text,
    structured_req,
    job_description,
    msp_notes,
    spotlight_transcript,
    internal_notes,
    draft_data,
    calculated_total_experience,
    current_date,
    validator_errors=None,
    candidate_clarifications="",
):
    """Independent Freddie fact-check/prioritization pass for Summary + Skills."""
    error_section = ""
    if validator_errors:
        error_section = (
            "\nDETERMINISTIC VALIDATOR ERRORS FROM THE PREVIOUS PASS:\n- "
            + "\n- ".join(validator_errors)
            + "\nFix every listed problem.\n"
        )

    return f"""
Return valid JSON ONLY.

You are the FINAL FACT-CHECKER for a Freddie Mac VNDLY resume submission.
This is separate from the recruiter-writing pass. Accuracy and current hiring-
manager relevance are more important than preserving draft wording.

CURRENT DATE: {current_date}
PYTHON-CALCULATED TOTAL PROFESSIONAL EXPERIENCE:
{calculated_total_experience or "BLANK"}

SOURCE ROLES:
- ORIGINAL RESUME = PRIMARY proof for candidate experience, skills, tools,
  ownership, scope, accomplishments, employers, domains, and technologies.
- CANDIDATE-CONFIRMED CLARIFICATIONS may add factual experience/details that the
  resume omitted (for example reporting scope, older tool usage, portfolio scope,
  or leadership audience). Treat them as secondary candidate evidence, not as
  role evidence. Never use them to contradict dated resume facts.
- OFFICIAL CANDIDATE VETTING may also clarify depth, but theoretical answers do
  not become performed experience.
- SPOTLIGHT CALL + substantive MSP/VNDLY notes = highest-priority CURRENT hiring-
  team guidance for what to emphasize or de-emphasize.
- OFFICIAL JD + structured VNDLY fields = formal baseline requirements and
  automated matching terminology.
- INTERNAL NOTES = recruiter context only; never candidate evidence and never
  authority to override explicit Freddie guidance.

CRITICAL MANAGER-CLARIFICATION RULE:
Separate PRIORITY from PROHIBITION.
- Manager says a JD skill is not used / not central / not required: DE-EMPHASIZE
  it for hiring-manager persuasion, but do not automatically erase a truthful JD
  match that can help MSP shortlisting. It may remain as a secondary detail when
  space permits and it does not displace a higher-priority competency.
- Manager explicitly says candidates should NOT have/emphasize something or that
  it is a negative/disqualifier: OMIT it.

Examples:
- JD says User Stories but manager says "we don't do user stories" -> User
  Stories may remain as a secondary detail when the candidate genuinely has it,
  but it should not become one of the four primary competency rows unless other
  evidence makes it independently important.
- Manager says Power Platform familiarity is nice to have -> do not label it a
  Must Have and do not assign Power Platform unless candidate evidence proves it.
- Role wants executive reporting -> do not say candidate owned executive
  reporting unless the resume or candidate-confirmed clarification proves it.

ZERO-TOLERANCE CANDIDATE GROUNDING:
- Do not copy a tool, platform, methodology, responsibility, domain, or qualifier
  from the role sources and assign it to the candidate unless the resume proves it.
- Do not combine separate resume facts into a stronger composite claim.
- Do not upgrade support/participation into lead/ownership/management.
- Do not upgrade SharePoint into Power Platform, AWS into Redshift, Agile into
  PI Planning, or a general leadership interaction into executive reporting.
- If the candidate lacks a role requirement, omit it rather than fabricate it.

HIGH-IMPACT QUALIFIERS:
Words that materially change authority/scope/proficiency must be supported by the
specific cited candidate evidence for that claim: independent, lead/led/leading,
own/ownership, manage/managed/manager, oversight, architect, expert, advanced,
hands-on, enterprise-wide, organization-wide, global,
end-to-end, high-volume, large-scale, production.

SUMMARY:
- Exactly 4 sentences, one paragraph after finalization.
- Natural human recruiter voice, concise and evidence-based.
- Prefer factual scope over subjective intensifiers such as "extensive" or "deep".
- Use the deterministic total-career value exactly if a total-years statement is
  used; never attach that total to individual skills unless the resume proves it.
- Sentence 1: identity + strongest supported role anchors.
- Sentence 2: strongest concrete recent/repeated proof.
- Sentence 3: supporting execution/depth/impact.
- Sentence 4: truthful transferability to Freddie's current need; clearly keep
  target-role application separate from candidate past experience.
- For EACH sentence return 1-3 short VERBATIM candidate-evidence excerpts
  proving the candidate claims in that sentence.
- Avoid harmless recruiter-language leakage from role text. Generic glue words such
  as "provides", "functions", or "supports" are not meaningful candidate claims.

SKILLS TABLE:
- Return up to 4 rows; use 4 when four distinct defensible role-relevant rows
  exist. Never force an unsupported fourth row.
- Treat the four rows as scarce real estate.
- Spotlight/MSP current selection guidance controls emphasis first, then formal
  Must Haves, then explicit Preferred/Nice-to-Haves.
- Each row must be a distinct screening dimension.
- For each row return:
  * skill_area: concise high-level competency
  * area_evidence: 1-3 VERBATIM candidate-evidence excerpts proving the competency
  * role_evidence: 1-2 VERBATIM excerpts from the Official JD, Spotlight call,
    or MSP/VNDLY notes proving why the row matters to this requisition
  * details: 2-6 atomic items, each with ONE VERBATIM candidate-evidence excerpt
    supporting that detail
- Skill areas may be natural umbrella competency labels synthesized from multiple
  verified candidate facts. Details may use a concise natural label when the cited
  evidence clearly anchors the concept and the broader candidate record supplies
  the remaining context. Do not change the underlying scope, tool, domain, or
  authority level merely to mirror Freddie wording.
- Prefer the candidate's own terminology when a role synonym could imply a different
  scope. Example: keep "Product Roadmaps" when that is what the evidence proves; do
  not silently rewrite it as "Program Roadmaps" solely because the target is a
  program role.
- "Leadership Presentations" may describe presentations delivered to senior leaders;
  the noun "leadership" does not by itself claim that the candidate led a team.
- Do not use years in Skills rows.
- Do not use manager-explicitly-excluded artifacts.
- A merely de-emphasized JD skill may remain as a secondary detail for MSP/JD
  coverage if candidate evidence supports it and it does not displace a current
  manager priority.
- Prefer direct operational/repeated evidence over impressive but secondary work.

BULLET ORDER:
- Preserve the draft BULLET_ORDER unless a change is clearly needed for current
  Freddie priorities.
- Reorder only; never rewrite/omit/create work-history bullets.

{error_section}

STRUCTURED REQUISITION:
{json.dumps(fred_coerce_structured_req(structured_req), indent=2, ensure_ascii=False)}

OFFICIAL JD:
{job_description}

MSP / VNDLY NOTES:
{msp_notes}

SPOTLIGHT CALL TRANSCRIPT:
{spotlight_transcript}

INTERNAL NOTES:
{internal_notes}

CANDIDATE-CONFIRMED CLARIFICATIONS / FOLLOW-UP ANSWERS:
{candidate_clarifications or "NONE PROVIDED"}

DRAFT FROM RECRUITER-WRITING PASS:
{json.dumps(draft_data if isinstance(draft_data, dict) else {}, indent=2, ensure_ascii=False)}

ORIGINAL RESUME:
{resume_text}

OUTPUT SCHEMA:
{{
  "CANDIDATE_TITLE": "",
  "SUMMARY_SENTENCES": [
    {{"text": "", "evidence": [""]}},
    {{"text": "", "evidence": [""]}},
    {{"text": "", "evidence": [""]}},
    {{"text": "", "evidence": [""]}}
  ],
  "SKILLS": [
    {{
      "skill_area": "",
      "area_evidence": [""],
      "role_evidence": [""],
      "details": [
        {{"item": "", "evidence": ""}}
      ]
    }}
  ],
  "BULLET_ORDER": {{
    "1": [], "2": [], "3": [], "4": [], "5": [], "6": [], "7": []
  }}
}}
"""


def fred_normalize_audited_detail_records(value):
    if not isinstance(value, list):
        return []
    records = []
    for entry in value:
        if not isinstance(entry, dict):
            continue
        item = re.sub(r"\s+", " ", str(entry.get("item", "") or "")).strip()
        evidence = re.sub(r"\s+", " ", str(entry.get("evidence", "") or "")).strip()
        if item and evidence:
            records.append({"item": item, "evidence": evidence})
    return records


def fred_role_priority_quote_found(
    job_description,
    msp_notes,
    spotlight_transcript,
    quote,
):
    """
    Verify role-priority evidence against one authoritative Freddie source.

    Exact normalized quotes are preferred. A conservative fallback tolerates
    harmless Gemini shortening/ellipsis only when every substantive token in the
    cited excerpt appears in the SAME source. Candidate evidence remains strict.
    """
    quote_text = str(quote or "").strip()
    if not quote_text:
        return False

    sources = [job_description, spotlight_transcript, msp_notes]

    for source in sources:
        if source and fred_source_quote_found(source, quote_text):
            return True

    # Role evidence establishes WHY a row matters, not candidate history. Allow
    # harmless shortening or tense changes while still requiring the substantive
    # concepts to occur together in one authoritative source.
    low_value = {
        "all", "team", "role", "candidate", "person", "people", "current",
        "need", "needs", "needed", "want", "wants", "wanted", "level",
        "function", "functions", "provide", "provid", "support",
    }
    quote_tokens = [
        token for token in fred_signal_tokens(quote_text)
        if token not in low_value and len(token) >= 3
    ]
    if not quote_tokens:
        return False

    for source in sources:
        source_tokens = set(fred_signal_tokens(source))
        if not source_tokens:
            continue
        matched = sum(1 for token in quote_tokens if token in source_tokens)
        # For short evidence phrases require all substantive concepts. For longer
        # phrases tolerate one harmless omitted/reworded token.
        required = len(quote_tokens) if len(quote_tokens) <= 3 else max(3, len(quote_tokens) - 1)
        if matched >= required:
            return True

    return False


def fred_validate_grounded_tailoring(
    resume_text,
    structured_req,
    job_description,
    msp_notes,
    spotlight_transcript,
    data,
    candidate_clarifications="",
):
    """Deterministically audit Freddie Summary/Skills evidence after AI fact-check."""
    errors = []
    candidate_evidence_source = fred_candidate_evidence_text(
        resume_text, candidate_clarifications
    )
    structured_req = fred_coerce_structured_req(structured_req)
    role_text = "\n".join(
        [str(job_description or ""), str(msp_notes or ""), str(spotlight_transcript or "")]
    )

    sentences = data.get("SUMMARY_SENTENCES", []) if isinstance(data, dict) else []
    if not isinstance(sentences, list) or len(sentences) != 4:
        errors.append("SUMMARY must contain exactly four audited sentences.")
        sentences = sentences if isinstance(sentences, list) else []

    for idx, entry in enumerate(sentences[:4], start=1):
        if not isinstance(entry, dict):
            errors.append(f"SUMMARY sentence {idx} has invalid structure.")
            continue

        text = re.sub(r"\s+", " ", str(entry.get("text", "") or "")).strip()
        evidence = fred_string_list(entry.get("evidence", []))

        if not text:
            errors.append(f"SUMMARY sentence {idx} is blank.")
            continue
        if not evidence:
            errors.append(f"SUMMARY sentence {idx} has no candidate evidence.")
            continue

        valid_evidence = []
        for quote in evidence:
            if fred_candidate_evidence_quote_found(candidate_evidence_source, quote):
                valid_evidence.append(quote)
            else:
                errors.append(
                    f"SUMMARY sentence {idx} cites evidence not found in the supplied candidate sources: {quote[:120]}"
                )

        joined = " ".join(valid_evidence)
        unsupported_qualifiers = fred_unsupported_high_impact_claims(text, joined)
        if unsupported_qualifiers:
            errors.append(
                f"SUMMARY sentence {idx} contains unsupported high-impact claim(s): "
                + ", ".join(unsupported_qualifiers)
            )

        unsupported_identifiers = fred_explicit_identifiers_supported(
            text, joined, candidate_evidence_source
        )
        if unsupported_identifiers:
            errors.append(
                f"SUMMARY sentence {idx} contains unsupported named identifier(s): "
                + ", ".join(unsupported_identifiers)
            )

        negated_hits = fred_output_hits_excluded_requirement(text, structured_req)
        if negated_hits:
            errors.append(
                f"SUMMARY sentence {idx} emphasizes manager-explicitly-excluded concept(s): "
                + ", ".join(negated_hits)
            )

        # Sentence 4 may deliberately describe target-role transferability.
        if idx < 4:
            role_hits = fred_role_only_phrase_hits(
                text, role_text, candidate_evidence_source, joined
            )
            token_hits = fred_role_only_signal_tokens(
                text, role_text, candidate_evidence_source, joined
            )
            if role_hits:
                errors.append(
                    f"SUMMARY sentence {idx} contains role-source wording not "
                    "independently supported as candidate history: "
                    + "; ".join(role_hits)
                )
            if token_hits:
                errors.append(
                    f"SUMMARY sentence {idx} contains distinctive role-source "
                    "term(s) absent from candidate evidence: "
                    + ", ".join(token_hits)
                )

    skills = data.get("SKILLS", []) if isinstance(data, dict) else []
    if not isinstance(skills, list):
        skills = []
        errors.append("SKILLS has invalid structure.")

    if len(skills) > 4:
        errors.append("SKILLS returned more than four rows.")

    for idx, row in enumerate(skills[:4], start=1):
        if not isinstance(row, dict):
            errors.append(f"SKILL row {idx} has invalid structure.")
            continue

        area = re.sub(r"\s+", " ", str(row.get("skill_area", "") or "")).strip()
        area_evidence = fred_string_list(row.get("area_evidence", []))
        role_evidence = fred_string_list(row.get("role_evidence", []))
        details = fred_normalize_audited_detail_records(row.get("details", []))

        if not area:
            errors.append(f"SKILL row {idx} has no skill area.")
            continue
        if not area_evidence:
            errors.append(f"SKILL row {idx} has no candidate evidence.")
        else:
            for quote in area_evidence:
                if not fred_source_quote_found(candidate_evidence_source, quote):
                    errors.append(
                        f"SKILL row {idx} cites area evidence not found in the supplied candidate sources: {quote[:120]}"
                    )

        joined_area_evidence = " ".join(
            q for q in area_evidence if fred_candidate_evidence_quote_found(candidate_evidence_source, q)
        )

        if joined_area_evidence and not fred_claim_supported_by_evidence(
            area, candidate_evidence_source
        ):
            errors.append(
                f"SKILL row {idx} area is not sufficiently supported by its cited candidate evidence."
            )

        unsupported_area_qualifiers = fred_unsupported_high_impact_claims(
            area, joined_area_evidence
        )
        if unsupported_area_qualifiers:
            errors.append(
                f"SKILL row {idx} contains unsupported high-impact claim(s): "
                + ", ".join(unsupported_area_qualifiers)
            )

        unsupported_area_identifiers = fred_explicit_identifiers_supported(
            area, joined_area_evidence, candidate_evidence_source
        )
        if unsupported_area_identifiers:
            errors.append(
                f"SKILL row {idx} contains unsupported named identifier(s): "
                + ", ".join(unsupported_area_identifiers)
            )

        negated_hits = fred_output_hits_excluded_requirement(area, structured_req)
        if negated_hits:
            errors.append(
                f"SKILL row {idx} emphasizes manager-explicitly-excluded concept(s): "
                + ", ".join(negated_hits)
            )

        if not role_evidence:
            errors.append(f"SKILL row {idx} has no role-priority evidence.")
        else:
            for quote in role_evidence:
                if not fred_role_priority_quote_found(
                    job_description, msp_notes, spotlight_transcript, quote
                ):
                    errors.append(
                        f"SKILL row {idx} cites role-priority evidence not found in the JD/Spotlight/MSP sources: {quote[:120]}"
                    )

        if not details:
            errors.append(f"SKILL row {idx} has no grounded detail items.")
            continue

        for detail in details[:6]:
            item = detail["item"]
            evidence = detail["evidence"]

            if not fred_candidate_evidence_quote_found(candidate_evidence_source, evidence):
                errors.append(
                    f"SKILL row {idx} detail '{item}' cites evidence not found in the resume."
                )
                continue

            if not (
                fred_claim_supported_by_candidate_context(
                    item, evidence, candidate_evidence_source
                )
                or fred_claim_supported_by_evidence(
                    item, candidate_evidence_source
                )
            ):
                errors.append(
                    f"SKILL row {idx} detail '{item}' is not sufficiently supported by its cited candidate evidence."
                )

            unsupported_detail_qualifiers = fred_unsupported_high_impact_claims(
                item, evidence
            )
            if unsupported_detail_qualifiers:
                errors.append(
                    f"SKILL row {idx} detail '{item}' contains unsupported high-impact claim(s): "
                    + ", ".join(unsupported_detail_qualifiers)
                )

            unsupported_detail_identifiers = fred_explicit_identifiers_supported(
                item, evidence, candidate_evidence_source
            )
            if unsupported_detail_identifiers:
                errors.append(
                    f"SKILL row {idx} detail '{item}' contains unsupported identifier(s): "
                    + ", ".join(unsupported_detail_identifiers)
                )

            detail_negated = fred_output_hits_excluded_requirement(
                item, structured_req
            )
            if detail_negated:
                errors.append(
                    f"SKILL row {idx} detail '{item}' uses manager-explicitly-excluded concept(s): "
                    + ", ".join(detail_negated)
                )

    return errors


def fred_finalize_grounded_tailoring(data):
    """Convert the audited evidence-rich Freddie payload back to template keys."""
    output = {
        "CANDIDATE_TITLE": fred_clean_generated_text((data or {}).get("CANDIDATE_TITLE", "")),
        "SUMMARY": "",
        "BULLET_ORDER": (data or {}).get("BULLET_ORDER", {})
        if isinstance((data or {}).get("BULLET_ORDER", {}), dict)
        else {},
    }

    sentence_texts = []
    sentences = (data or {}).get("SUMMARY_SENTENCES", [])
    if isinstance(sentences, list):
        for entry in sentences[:4]:
            if isinstance(entry, dict):
                text = fred_clean_generated_text(entry.get("text", ""))
                if text:
                    sentence_texts.append(text)
    output["SUMMARY"] = " ".join(sentence_texts)

    skills = (data or {}).get("SKILLS", [])
    skills = skills if isinstance(skills, list) else []

    for idx in range(1, 5):
        if idx <= len(skills) and isinstance(skills[idx - 1], dict):
            row = skills[idx - 1]
            output[f"SKILL{idx}"] = fred_clean_generated_text(row.get("skill_area", ""))

            detail_names = []
            seen = set()
            for detail in fred_normalize_audited_detail_records(row.get("details", [])):
                item = fred_clean_generated_text(detail["item"])
                key = item.casefold()
                if not item or key in seen:
                    continue
                detail_names.append(item)
                seen.add(key)
                if len(detail_names) >= 6:
                    break
            output[f"SKILLDETAILS{idx}"] = ", ".join(detail_names)
        else:
            output[f"SKILL{idx}"] = ""
            output[f"SKILLDETAILS{idx}"] = ""

    return output


def fred_standardize_dates(date_str):
    """
    Standardize Freddie Mac work-history dates while preserving source precision:
    MMM YYYY - MMM YYYY
    MMM YYYY - CURRENT
    YYYY - YYYY
    or YYYY - CURRENT when the source provides years only.

    Never invent a month for a year-only source date.

    Examples:
    2026-02 - Current      -> FEB 2026 - CURRENT
    2021-08 - 2026-02      -> AUG 2021 - FEB 2026
    07/2019 - 08/2021      -> JUL 2019 - AUG 2021
    February 2026 - Present -> FEB 2026 - CURRENT
    """
    if not date_str:
        return ""

    text = str(date_str)

    # Normalize line breaks, spacing, and dash characters.
    text = text.replace("\r", " ").replace("\n", " ")
    text = text.replace("–", "-").replace("\u2014", "-")
    text = re.sub(r"\s+", " ", text).strip()

    month_names = {
        1: "JAN",
        2: "FEB",
        3: "MAR",
        4: "APR",
        5: "MAY",
        6: "JUN",
        7: "JUL",
        8: "AUG",
        9: "SEP",
        10: "OCT",
        11: "NOV",
        12: "DEC",
    }

    text_months = {
        "january": "JAN",
        "jan": "JAN",
        "february": "FEB",
        "feb": "FEB",
        "march": "MAR",
        "mar": "MAR",
        "april": "APR",
        "apr": "APR",
        "may": "MAY",
        "june": "JUN",
        "jun": "JUN",
        "july": "JUL",
        "jul": "JUL",
        "august": "AUG",
        "aug": "AUG",
        "september": "SEP",
        "sept": "SEP",
        "sep": "SEP",
        "october": "OCT",
        "oct": "OCT",
        "november": "NOV",
        "nov": "NOV",
        "december": "DEC",
        "dec": "DEC",
    }

    # Current / Present
    text = re.sub(
        r"\b(?:current|present)\b",
        "CURRENT",
        text,
        flags=re.IGNORECASE,
    )

    # YYYY-MM or YYYY/MM -> MMM YYYY
    def replace_year_month(match):
        year = int(match.group(1))
        month = int(match.group(2))
        return f"{month_names[month]} {year}"

    text = re.sub(
        r"\b((?:19|20)\d{2})[-/](0?[1-9]|1[0-2])\b",
        replace_year_month,
        text,
    )

    # MM/YYYY -> MMM YYYY
    def replace_month_year(match):
        month = int(match.group(1))
        year = int(match.group(2))
        return f"{month_names[month]} {year}"

    text = re.sub(
        r"\b(0?[1-9]|1[0-2])/((?:19|20)\d{2})\b",
        replace_month_year,
        text,
    )

    # Written month names -> three-letter uppercase months.
    month_pattern = (
        r"\b("
        r"January|Jan|February|Feb|March|Mar|April|Apr|May|"
        r"June|Jun|July|Jul|August|Aug|September|Sept|Sep|"
        r"October|Oct|November|Nov|December|Dec"
        r")\b"
    )

    text = re.sub(
        month_pattern,
        lambda m: text_months[m.group(1).lower()],
        text,
        flags=re.IGNORECASE,
    )

    # Normalize spacing around the range dash.
    text = re.sub(r"\s*-\s*", " - ", text)

    return text.strip()


def fred_calculate_total_experience(experience, current_date):
    """
    Calculate total professional experience deterministically from dated roles.

    Overlapping employment periods are merged so they are not double-counted.
    Month precision is preserved when supplied. Year-only dates are also usable
    with conservative calculation bounds (start year -> December, end year ->
    January) so the formatter does not silently lose legitimate older roles or
    invent months merely to increase tenure.
    """
    intervals = []

    for role in experience or []:
        if not isinstance(role, dict):
            continue

        parsed = fred_role_date_interval(
            role.get("Dates", ""),
            current_date,
        )

        if not parsed:
            continue

        start_idx, end_idx, _end_is_current, _end_year = parsed

        if end_idx < start_idx:
            continue

        intervals.append((start_idx, end_idx))

    if not intervals:
        return ""

    intervals.sort()
    merged = []

    for start, end in intervals:
        if not merged or start > merged[-1][1]:
            merged.append([start, end])
        else:
            merged[-1][1] = max(merged[-1][1], end)

    # Elapsed completed months. This intentionally mirrors ordinary date
    # difference logic rather than counting both endpoint months as full months.
    total_months = sum(
        max(0, end - start)
        for start, end in merged
    )

    completed_years = total_months // 12

    if completed_years < 1:
        return "<1 year"

    return f"{completed_years}+ years"


def fred_role_date_interval(dates, current_date):
    """
    Parse a Freddie-standardized role date range into conservative month indexes.

    Month/year dates use the stated month.
    Year-only dates never invent January for display; for calculations they use
    conservative bounds:
      - start year only -> December of that year
      - end year only   -> January of that year
    This avoids overstating tenure when month precision is unknown.
    """
    if not dates:
        return None

    text = str(dates).strip().upper()
    match = re.fullmatch(r"\s*(.+?)\s+-\s+(.+?)\s*", text)
    if not match:
        return None

    month_lookup = {
        "JAN": 1, "FEB": 2, "MAR": 3, "APR": 4, "MAY": 5, "JUN": 6,
        "JUL": 7, "AUG": 8, "SEP": 9, "OCT": 10, "NOV": 11, "DEC": 12,
    }

    def parse_part(value, is_start):
        value = value.strip().upper()

        if value in {"CURRENT", "PRESENT"}:
            return current_date.year, current_date.month, True

        m = re.fullmatch(
            r"(JAN|FEB|MAR|APR|MAY|JUN|JUL|AUG|SEP|OCT|NOV|DEC)\s+((?:19|20)\d{2})",
            value,
        )
        if m:
            return int(m.group(2)), month_lookup[m.group(1)], False

        m = re.fullmatch(r"((?:19|20)\d{2})", value)
        if m:
            year = int(m.group(1))
            # Conservative unknown-month assumptions for calculation only.
            return year, (12 if is_start else 1), False

        return None

    start = parse_part(match.group(1), True)
    end = parse_part(match.group(2), False)

    if not start or not end:
        return None

    sy, sm, _ = start
    ey, em, end_is_current = end

    start_idx = sy * 12 + (sm - 1)
    end_idx = ey * 12 + (em - 1)

    if end_idx < start_idx:
        return None

    return start_idx, end_idx, end_is_current, ey


def fred_skill_label_components(skill_label):
    """Return named supporting technologies from parentheses, if present."""
    label = str(skill_label or "").strip()
    if not label:
        return []

    components = []
    for part in re.findall(r"\(([^)]*)\)", label):
        for token in re.split(r"[,/&;+]", part):
            token = token.strip()
            if token:
                components.append(token)
    return components


def fred_role_text(role):
    """Flatten one structured role into searchable evidence text."""
    bullets = role.get("Bullets", []) or []
    if isinstance(bullets, str):
        bullets_text = bullets
    else:
        bullets_text = " ".join(str(x) for x in bullets)

    return " ".join(
        [
            str(role.get("Company", "") or ""),
            str(role.get("Title", "") or ""),
            bullets_text,
            str(role.get("Environment", "") or ""),
        ]
    ).casefold()


def fred_role_supports_skill_component(role, component):
    """Conservative alias check against one dated structured role."""
    role_text = fred_role_text(role)
    key = fred_normalize_skill_name(component)

    aliases = {
        "j2ee": ["j2ee", "java ee", "jee"],
        "microservices": ["microservice"],
        "spring boot": ["spring boot"],
        "spring framework": ["spring framework", "spring "],
        "java": ["java"],
        "oracle": ["oracle"],
        "db2": ["db2"],
        "sql server": ["sql server"],
        "microsoft sql server": ["sql server"],
        "sql": ["sql", "pl/sql"],
        "junit": ["junit"],
        "junit testing": ["junit"],
        "mockito": ["mockito"],
        "git": ["git"],
        "maven": ["maven"],
        "jenkins": ["jenkins"],
        "gradle": ["gradle"],
        "jdbc": ["jdbc", "java database connectivity"],
        "jsp": ["jsp", "jsps", "javaserver pages", "java server pages"],
        "servlets": ["servlet"],
        "banking": ["bank", "banking"],
        "investment": ["investment", "portfolio", "asset", "securit", "trade ", "trading"],
        "securities": ["securit", "trade ", "trading", "investment"],
    }

    terms = aliases.get(key)
    if terms:
        return fred_text_contains_any(role_text, terms)

    core = re.sub(r"[^a-z0-9+#./ -]", "", key).strip()
    return bool(core and len(core) >= 2 and fred_text_contains_term(role_text, core))


def fred_role_supports_skill_core(role, skill_label):
    """
    Determine whether a role supports the CORE competency represented by a
    Freddie Skills-table row. Parenthetical tools are examples/evidence, not an
    AND checklist that must all coexist in the same job.
    """
    label = re.sub(r"\([^)]*\)", "", str(skill_label or "")).strip().casefold()
    role_text = fred_role_text(role)

    # High-confidence competency families used by the Freddie skills table.
    family_aliases = [
        (["financial", "banking", "investment", "securities"],
         ["bank", "banking", "financial", "investment", "portfolio", "asset", "securit", "trade ", "trading", "retirement"]),
        (["relational database", "database", "sql"],
         ["sql", "pl/sql", "oracle", "db2", "sql server", "mysql", "postgresql", "sybase", "jdbc", "database"]),
        (["java", "j2ee", "spring"],
         ["java", "j2ee", "spring", "servlet", "jsp", "jdbc", "microservice"]),
        (["testing", "test", "quality"],
         ["junit", "mockito", "unit test", "testing", "test case"]),
        (["devops", "build", "deployment", "version control", "ci/cd"],
         ["git", "gitlab", "github", "maven", "gradle", "jenkins", "docker", "ci/cd", "build", "deploy"]),
        (["sharepoint"], ["sharepoint"]),
        (["communications", "content"], ["communication", "content", "newsletter", "publication", "writing", "editing"]),
        (["change management", "change communications", "adoption"], ["change", "adoption", "readiness", "implementation communication"]),
    ]

    for label_terms, evidence_terms in family_aliases:
        if any(term in label for term in label_terms):
            return fred_text_contains_any(role_text, evidence_terms)

    # Generic fallback: any explicitly named parenthetical component supports
    # the row as an AREA competency; we do not require every component at once.
    components = fred_skill_label_components(skill_label)
    if components and any(
        fred_role_supports_skill_component(role, component)
        for component in components
    ):
        return True

    # Last-resort lexical support for a concise core label.
    core_words = [
        w for w in re.findall(r"[a-z0-9+#.]+", label)
        if len(w) >= 4 and w not in {"development", "experience", "background", "expertise"}
    ]
    return bool(
        core_words
        and any(fred_text_contains_term(role_text, word) for word in core_words)
    )


def fred_filter_skill_role_indexes_by_label(
    skill_label,
    experience,
    role_indexes,
):
    """
    Build the dated role set for a Freddie Skills-table row based on the CORE
    competency, not an artificial requirement that every parenthetical tool
    appear in the same job.

    Gemini's proposed role indexes remain useful for domain/context judgments,
    but Python also adds any dated role that clearly supports the core area.
    """
    proposed = fred_clean_skill_role_indexes(role_indexes, len(experience))
    inferred = []

    for idx, role in enumerate(experience, start=1):
        if fred_role_supports_skill_core(role, skill_label):
            inferred.append(idx)

    # Preserve Gemini's domain-aware choices and add deterministic evidence.
    combined = []
    for idx in proposed + inferred:
        if idx not in combined:
            combined.append(idx)

    return combined


def fred_calculate_skill_years_from_roles(
    experience,
    role_indexes,
    current_date,
):
    """
    Deterministically calculate one Freddie Skills-row YEARS value from only the
    dated roles that Gemini identified as supporting the ENTIRE skill label.
    Overlapping roles are merged so months are not double-counted.
    """
    if not isinstance(role_indexes, list):
        return ""

    clean_indexes = []
    for value in role_indexes:
        try:
            idx = int(value)
        except (TypeError, ValueError):
            continue
        if 1 <= idx <= len(experience) and idx not in clean_indexes:
            clean_indexes.append(idx)

    if not clean_indexes:
        return ""

    intervals = []
    latest_year = None
    used_current = False

    for idx in clean_indexes:
        role = experience[idx - 1]
        parsed = fred_role_date_interval(
            role.get("Dates", ""),
            current_date,
        )
        if not parsed:
            continue

        start_idx, end_idx, end_is_current, end_year = parsed
        intervals.append((start_idx, end_idx))

        if end_is_current:
            used_current = True

        if latest_year is None or end_year > latest_year:
            latest_year = end_year

    if not intervals:
        return ""

    intervals.sort()
    merged = [list(intervals[0])]

    for start_idx, end_idx in intervals[1:]:
        last = merged[-1]
        if start_idx <= last[1] + 1:
            last[1] = max(last[1], end_idx)
        else:
            merged.append([start_idx, end_idx])

    # Inclusive role months.
    total_months = sum((end - start + 1) for start, end in merged)
    completed_years = total_months // 12

    last_used = "current" if used_current else str(latest_year or "").strip()

    if completed_years < 1:
        return f"<1 year, {last_used}" if last_used else "<1 year"

    return (
        f"{completed_years}+ years, {last_used}"
        if last_used
        else f"{completed_years}+ years"
    )


def fred_clean_skill_role_indexes(value, max_roles):
    """Return unique valid 1-based role indexes from Gemini output."""
    if not isinstance(value, list):
        return []

    result = []
    for item in value:
        try:
            idx = int(item)
        except (TypeError, ValueError):
            continue
        if 1 <= idx <= max_roles and idx not in result:
            result.append(idx)
    return result


def fred_normalize_candidate_positioning_title(title):
    """
    Freddie-only cleanup for the professional identity used in the Summary/VNDLY
    comments. This NEVER changes any title in the candidate's work history.

    Some resumes use a title line like:
        "Data Engineer, Python, Computer Vision & Data Pipelines"

    The occupational title is "Data Engineer"; the remaining text describes
    specialties. Returning them separately lets the Summary say:
        "Data Engineer specializing in Python, Computer Vision & Data Pipelines"
    instead of the incomplete:
        "Data Engineer, Computer Vision & Data Pipelines with 7+ years..."
    """
    raw = re.sub(r"\s+", " ", str(title or "")).strip(" ,|/:;-")
    if not raw:
        return "", ""

    # Occupational head words that usually make the left side a complete,
    # conventional market title on its own.
    occupation_pattern = re.compile(
        r"\b("
        r"engineer|developer|analyst|architect|scientist|specialist|"
        r"manager|director|administrator|consultant|owner|lead|"
        r"coordinator|writer|designer|tester|auditor|accountant|"
        r"strategist|advisor|adviser|scrum master|product owner|"
        r"project manager|program manager"
        r")\b",
        flags=re.IGNORECASE,
    )

    # Split only on separators commonly used to append skill/domain descriptors.
    parts = re.split(r"\s*(?:\||,|:)\s*", raw, maxsplit=1)

    if len(parts) != 2:
        return raw, ""

    left = parts[0].strip(" ,|/:;-")
    right = parts[1].strip(" ,|/:;-")

    if not left or not right:
        return raw, ""

    # Only split when the left side is already a complete occupational title.
    # This prevents altering legitimate titles that need their qualifier.
    if not occupation_pattern.search(left):
        return raw, ""

    # If the right side itself looks like another occupational title, preserve
    # the full original phrase rather than guessing.
    if occupation_pattern.search(right) and len(right.split()) <= 5:
        return raw, ""

    return left, right


def fred_repair_summary_identity_phrase(
    summary,
    candidate_name,
    raw_positioning_title,
    clean_positioning_title,
    specialty_text,
):
    """
    Repair only the Freddie Summary's candidate-identity phrase when the model
    copied an occupation + specialty heading too literally.

    Work-history titles are untouched.
    """
    if not summary:
        return ""

    cleaned = str(summary).strip()
    raw_title = str(raw_positioning_title or "").strip()
    clean_title = str(clean_positioning_title or "").strip()
    specialty = str(specialty_text or "").strip()

    if not raw_title or not clean_title or raw_title == clean_title:
        return cleaned

    # Best case: replace the exact awkward identity with a grammatical version.
    if specialty:
        natural_identity = (
            f"{clean_title} specializing in {specialty}"
        )
    else:
        natural_identity = clean_title

    # Replace only the first occurrence, which should be Sentence 1.
    cleaned = re.sub(
        re.escape(raw_title),
        natural_identity,
        cleaned,
        count=1,
        flags=re.IGNORECASE,
    )

    return cleaned


def fred_clean_summary(summary, candidate_name):
    """
    Final deterministic cleanup for Freddie Mac summaries.

    - Forces the candidate's name to use the same Title Case spelling
      extracted from the resume.
    - Removes/replaces recruiter-style phrases Gemini was explicitly
      instructed not to use.
    """
    if not summary:
        return ""

    cleaned = fred_clean_generated_text(summary)

    # ------------------------------------------------------------
    # 1. NORMALIZE CANDIDATE NAME CAPITALIZATION
    # ------------------------------------------------------------

    if candidate_name:
        normalized_name = str(candidate_name).strip()
        first_name = normalized_name.split()[0] if normalized_name else ""
        if normalized_name and first_name:
            cleaned = re.sub(
                r"^\s*" + re.escape(normalized_name) + r"\b",
                first_name,
                cleaned,
                count=1,
                flags=re.IGNORECASE,
            )
            cleaned = re.sub(
                r"^\s*" + re.escape(first_name) + r"\b",
                first_name,
                cleaned,
                count=1,
                flags=re.IGNORECASE,
            )

    # ------------------------------------------------------------
    # 2. REMOVE / REPLACE BANNED RECRUITER PHRASES
    # ------------------------------------------------------------

    phrase_replacements = {
        "proven history in": "experience in",
        "proven history of": "experience with",
        "proven history": "experience",
    }

    for old_phrase, new_phrase in phrase_replacements.items():
        cleaned = re.sub(
            re.escape(old_phrase),
            new_phrase,
            cleaned,
            flags=re.IGNORECASE,
        )

    # Clean up any accidental double spaces.
    cleaned = re.sub(r"\s{2,}", " ", cleaned).strip()

    return cleaned


def fred_safe_reorder_bullets(bullets, requested_order):
    """
    Reorder existing bullets ONLY.

    Gemini returns zero-based indexes. This function refuses to manufacture,
    delete, or rewrite any bullet. Invalid indexes are ignored and every
    unranked original bullet is appended in its original order.
    """
    if not isinstance(bullets, list):
        return bullets

    if not isinstance(requested_order, list):
        return bullets

    valid_order = []
    seen = set()

    for value in requested_order:
        try:
            idx = int(value)
        except (TypeError, ValueError):
            continue

        if 0 <= idx < len(bullets) and idx not in seen:
            valid_order.append(idx)
            seen.add(idx)

    if not valid_order:
        return bullets

    # Preserve every original bullet.
    valid_order.extend(
        idx for idx in range(len(bullets))
        if idx not in seen
    )

    return [bullets[idx] for idx in valid_order]


def fred_remove_table(table):
    tbl = table._element
    parent = tbl.getparent()

    if parent is not None:
        parent.remove(tbl)


def fred_find_table_containing(doc, text):
    target = str(text).lower()

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if target in cell.text.lower():
                    return table

    return None


def fred_get_reference_font(paragraph):
    """
    Capture the formatting already assigned to {{VETTING_QA}} in Word so
    generated Q&A inherits the template's typography.
    """
    if paragraph.runs:
        run = paragraph.runs[0]

        return {
            "name": run.font.name,
            "size": run.font.size,
            "color": (
                run.font.color.rgb
                if run.font.color and run.font.color.rgb
                else None
            ),
        }

    return {
        "name": None,
        "size": None,
        "color": None,
    }


def fred_apply_reference_font(run, reference, bold=None):
    if reference.get("name"):
        run.font.name = reference["name"]

    if reference.get("size"):
        run.font.size = reference["size"]

    if reference.get("color"):
        run.font.color.rgb = reference["color"]

    if bold is not None:
        run.bold = bold


def fred_insert_vetting_qa(cell, vetting_pairs):
    """
    Replace {{VETTING_QA}} with a dynamic number of bold questions and
    normal-weight candidate answers.
    """
    placeholder_paragraph = None

    for p in cell.paragraphs:
        if "{{vetting_qa}}" in p.text.lower():
            placeholder_paragraph = p
            break

    if placeholder_paragraph is None:
        return

    reference = fred_get_reference_font(placeholder_paragraph)

    # Clear the placeholder paragraph while keeping the paragraph itself.
    for run in list(placeholder_paragraph.runs):
        run._element.getparent().remove(run._element)

    placeholder_paragraph.paragraph_format.space_before = Pt(0)
    placeholder_paragraph.paragraph_format.space_after = Pt(2)

    for idx, pair in enumerate(vetting_pairs, start=1):
        question = str(pair.get("question", "")).strip()
        answer = str(pair.get("answer", "")).strip()

        if idx == 1:
            question_p = placeholder_paragraph
        else:
            question_p = cell.add_paragraph()

        question_p.paragraph_format.space_before = Pt(4 if idx > 1 else 0)
        question_p.paragraph_format.space_after = Pt(1)

        q_run = question_p.add_run(f"{idx}. {question}")
        fred_apply_reference_font(q_run, reference, bold=True)

        answer_p = cell.add_paragraph()
        answer_p.paragraph_format.left_indent = Pt(12)
        answer_p.paragraph_format.space_before = Pt(0)
        answer_p.paragraph_format.space_after = Pt(3)

        a_run = answer_p.add_run(answer)
        fred_apply_reference_font(a_run, reference, bold=False)


def process_freddie_word_doc(
    temp_path,
    mapping,
    vetting_pairs,
    out_path,
):
    """
    Freddie-specific Word processor.

    We intentionally keep this separate from process_word_doc() so the
    successful Fannie/other-client Q1-A1 behavior remains untouched.
    """
    doc = docx.Document(temp_path)
    expand_experience_placeholders(doc, mapping)
    ensure_experience_spacing(doc, mapping)

    # ================================================================
    # 1. SUPPLIER VETTING QUESTIONS
    # ================================================================

    vetting_table = fred_find_table_containing(doc, "{{VETTING_QA}}")

    if vetting_table is not None:
        if vetting_pairs:
            inserted = False

            for row in vetting_table.rows:
                for cell in row.cells:
                    if "{{vetting_qa}}" in cell.text.lower():
                        fred_insert_vetting_qa(cell, vetting_pairs)
                        inserted = True
                        break

                if inserted:
                    break
        else:
            # No official Freddie vetting questions = remove entire section.
            fred_remove_table(vetting_table)

    # ================================================================
    # 2. CERTIFICATIONS
    # ================================================================

    certifications = str(mapping.get("Certifications", "") or "").strip()

    if not certifications:
        cert_table = fred_find_table_containing(doc, "{{CERTIFICATIONS}}")

        if cert_table is not None:
            fred_remove_table(cert_table)

    # ================================================================
    # 3. EDUCATION
    # ================================================================

    has_any_education = any(
        str(mapping.get(f"School{i}", "") or "").strip()
        for i in range(1, 4)
    )

    education_table = fred_find_table_containing(doc, "{{School1}}")

    if education_table is not None:
        if not has_any_education:
            fred_remove_table(education_table)
        else:
            rows_to_delete = []

            for row in education_table.rows:
                row_text = " ".join(
                    cell.text for cell in row.cells
                ).lower()

                for i in range(1, 4):
                    if (
                        f"{{{{school{i}}}}}".lower() in row_text
                        and not str(mapping.get(f"School{i}", "") or "").strip()
                    ):
                        rows_to_delete.append(row)
                        break

            for row in rows_to_delete:
                try:
                    education_table._tbl.remove(row._tr)
                except Exception:
                    pass

    # ================================================================
    # 4. SKILLS - TWO-COLUMN SKILL AREA / SKILL DETAILS TABLE
    #    Remove unused rows instead of leaving blanks.
    # ================================================================

    skills_table = fred_find_table_containing(doc, "{{SKILL1}}")

    if skills_table is not None:
        rows_to_delete = []

        for row in skills_table.rows:
            row_text = " ".join(
                cell.text for cell in row.cells
            ).lower()

            for i in range(1, 5):
                if (
                    f"{{{{skill{i}}}}}".lower() in row_text
                    and not str(mapping.get(f"SKILL{i}", "") or "").strip()
                ):
                    rows_to_delete.append(row)
                    break

        for row in rows_to_delete:
            try:
                skills_table._tbl.remove(row._tr)
            except Exception:
                pass

    # ================================================================
    # 5. REPLACE TAGS IN BODY PARAGRAPHS
    # ================================================================

    paras_to_remove = []

    kill_keys = [
        "Company",
        "Title",
        "Bullets",
        "Dates",
        "Summary",
        "Skill",
        "Years",
        "Certification",
        "School",
    ]

    for p in list(doc.paragraphs):
        for key, value in mapping.items():
            tag = f"{{{{{key}}}}}"

            if tag.lower() not in p.text.lower():
                continue

            if not value or not str(value).strip():
                if any(k.lower() in key.lower() for k in kill_keys):
                    if p not in paras_to_remove:
                        paras_to_remove.append(p)
                else:
                    replace_tag_safely(p, tag, "")
                continue

            if "Bullets" in key:
                for run in p.runs:
                    if "\n" in run.text:
                        run.text = run.text.replace("\n", "")

                lines = [
                    line.strip()
                    for line in str(value).split("\n")
                    if line.strip()
                ]

                if not lines:
                    if p not in paras_to_remove:
                        paras_to_remove.append(p)
                    continue

                replace_tag_safely(p, tag, lines[0])

                curr_p = p

                for line in lines[1:]:
                    curr_p = insert_bullet_line_after(
                        curr_p,
                        line,
                    )

                # Preserve explicit Environment / Technologies lines only,
                # following the same rule as your existing formatter.
                num_match = re.search(r"\d+", key)

                if num_match:
                    env_val = mapping.get(
                        f"Environment{num_match.group()}"
                    )

                    if env_val and str(env_val).strip():
                        env_p = curr_p.insert_paragraph_before("")
                        curr_p._p.addnext(env_p._p)

                        try:
                            env_p.style = doc.styles["Normal"]
                        except Exception:
                            pass

                        env_p.paragraph_format.left_indent = Pt(0)
                        env_p.paragraph_format.space_after = Pt(0)

                        clean_env = re.sub(
                            r"^Environment\s*:\s*",
                            "",
                            str(env_val).strip(),
                            flags=re.IGNORECASE,
                        )

                        b_run = env_p.add_run("Environment: ")
                        b_run.bold = True
                        b_run.font.name = "Times New Roman"
                        b_run.font.size = Pt(12)

                        n_run = env_p.add_run(clean_env)
                        n_run.bold = False
                        n_run.font.name = "Times New Roman"
                        n_run.font.size = Pt(12)

                        curr_p = env_p

                spacer = curr_p.insert_paragraph_before("")
                curr_p._p.addnext(spacer._p)

                try:
                    spacer.style = doc.styles["Normal"]
                except Exception:
                    pass

                spacer.paragraph_format.space_after = Pt(0)
                spacer.paragraph_format.space_before = Pt(0)

            else:
                replace_tag_safely(
                    p,
                    tag,
                    str(value).strip(),
                )

    # ================================================================
    # 6. REPLACE TAGS INSIDE TABLE CELLS
    # ================================================================

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        tag = f"{{{{{key}}}}}"

                        if tag.lower() in p.text.lower():
                            replace_tag_safely(
                                p,
                                tag,
                                str(value) if value else "",
                            )

    # ================================================================
    # 7. DELETE EMPTY WORK-HISTORY PARAGRAPHS
    # ================================================================

    for p in paras_to_remove:
        try:
            delete_paragraph(p)
        except Exception:
            pass

    # Collapse extra blank lines created when optional Freddie sections were removed.
    collapse_extra_blank_paragraphs(doc)

    doc.save(out_path)

    return out_path


# ====================================================================
# --- 🟣 CLIENT APP 1.5: FREDDIE MAC 🟣 ---
# ====================================================================

def freddie_mac_app():
    TEMPLATE_FILENAME = "Freddie_Mac_Template.docx"

    st.title("Freddie Mac Precision Extractor")

    # ================================================================
    # API CONFIGURATION
    # ================================================================

    with st.sidebar:
        st.header("🔑 API Configuration")

        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input(
                "Gemini API Key",
                type="password",
                key="fred_api_key",
            )

            st.info(
                "Paste your Gemini API key here to run the tool."
            )

    # ================================================================
    # CANDIDATE INFORMATION
    # ================================================================

    st.header("📋 Candidate Information")

    # Initialize dependent fields before widgets are rendered.
    if "fred_site" not in st.session_state:
        st.session_state.fred_site = ""

    if "fred_elig" not in st.session_state:
        st.session_state.fred_elig = "N/A"

    def fred_sync_onsite():
        if st.session_state.get("fred_onsite") == "No":
            st.session_state.fred_site = "N/A"
        elif st.session_state.get("fred_site") == "N/A":
            st.session_state.fred_site = ""

    def fred_sync_former_cw():
        if st.session_state.get("fred_form") == "No":
            st.session_state.fred_elig = "N/A"
        elif st.session_state.get("fred_elig") == "N/A":
            st.session_state.fred_elig = "No"

    col1, col2 = st.columns(2)

    with col1:
        Current_Location = st.text_input(
            "Current Location:",
            key="fred_loc",
        )

        Auth_US = st.selectbox(
            "Do you currently possess unrestricted lawful authorization "
            "to work in the U.S indefinitely?",
            ["Yes", "No"],
            index=0,
            key="fred_auth",
        )

        with st.container(border=True):
            Available_Onsite = st.selectbox(
                "Available to be onsite (Yes / No):",
                ["Yes", "No"],
                index=0,
                key="fred_onsite",
                on_change=fred_sync_onsite,
            )

            Site = st.text_input(
                "If yes, at which locations:",
                key="fred_site",
            )

    with col2:
        Interview = st.text_input(
            "Interview Availability for next 7 days:",
            key="fred_int",
        )

        Sponsorship = st.selectbox(
            "Will you now or in the future require sponsorship for an "
            "immigration-related employment benefit?",
            ["Yes", "No"],
            index=1,
            key="fred_spon",
        )

        with st.container(border=True):
            Former_CW = st.selectbox(
                "Former CW/Employee?",
                ["Yes", "No"],
                index=1,
                key="fred_form",
                on_change=fred_sync_former_cw,
            )

            Eligibility = st.selectbox(
                "If Yes, Rehire Eligibility Check Completed?",
                ["Yes", "No", "N/A"],
                key="fred_elig",
            )

    # ================================================================
    # REQUISITION / JOB DESCRIPTION
    # ================================================================

    st.header("📝 Freddie Mac Requisition")

    # Freddie ONLY: private authenticated connection to the VNDLY/API
    # Google Sheet. Fannie/Deloitte/other client sheet logic remains untouched.
    SHEET_URL = (
        "https://docs.google.com/spreadsheets/d/"
        "1yMQmqYRl_wa6ynXZrORCMOFwHRPwC2GVFKCTGzMfAe4/"
        "edit?gid=0"
    )

    def load_freddie_reqs(url):
        try:
            conn = st.connection(
                "gsheets",
                type=GSheetsConnection,
            )

            df = conn.read(
                spreadsheet=url,
                worksheet="VNDLY Jobs",
                ttl=600,
            )

            return df.fillna("")

        except Exception as e:
            st.error(
                "❌ Unable to load the private Freddie Mac VNDLY sheet. "
                f"{str(e)}"
            )
            return pd.DataFrame()

    df_reqs = load_freddie_reqs(SHEET_URL)

    options = ["None"]
    req_mapping = {}

    if not df_reqs.empty:
        for idx, row in df_reqs.iterrows():
            req_id = str(row.get("Job ID", "")).replace(".0", "").strip()
            title = str(row.get("Job Title", "")).strip()
            manager = str(row.get("Resource Manager", "")).strip()
            status = str(row.get("Status", "")).strip()

            # MSP-only requisitions that were never distributed to Convergenz
            # remain in the VNDLY tracker but must not appear in the
            # Resume Formatter requisition dropdown.
            if status.lower() in {"don't have", "dont have"}:
                continue

            if not req_id:
                continue

            desc = str(row.get("Official Job Description", ""))
            notes = str(row.get("VNDLY / MSP Notes", ""))
            transcript = str(row.get("Spotlight Call Transcript", ""))
            internal_notes = str(row.get("Internal Notes", ""))

            # Clean VNDLY-native dropdown: Job ID - Job Title - Resource Manager.
            option_parts = [req_id, title, manager]
            option_str = " - ".join(
                [part.strip() for part in option_parts if part and part.strip()]
            )

            if option_str and option_str != "-":
                options.append(option_str)

                req_mapping[option_str] = {
                    "job_id": req_id,
                    "status": status,
                    "job_title": title,
                    "resource_manager": manager,
                    "job_posted_date": str(row.get("Job Posted Date", "")),
                    "distributed_to_convergenz": str(
                        row.get("Distributed to Convergenz", "")
                    ),
                    "distribution_lag_hours": row.get(
                        "Distribution Lag (Hours)",
                        "",
                    ),
                    "distribution_timing": str(
                        row.get("Distribution Timing", "")
                    ),
                    "open_positions": row.get("Open Positions", ""),
                    "total_positions": row.get("Total Positions", ""),
                    "pending_positions": row.get("Pending Positions", ""),
                    "max_submissions_per_vendor": row.get(
                        "Max Submissions / Vendor",
                        "",
                    ),
                    "total_supplier_candidates": row.get(
                        "Total Supplier Candidates",
                        "",
                    ),
                    "min_suppliers_submitted": row.get(
                        "Min Suppliers Submitted",
                        "",
                    ),
                    "closed_date": str(row.get("Closed Date", "")),
                    "reason_for_ending": str(row.get("Reason for Ending", "")),
                    "organization_unit": str(
                        row.get("Organization Unit", "")
                    ),
                    "cost_center": str(row.get("Cost Center", "")),
                    "workday_cost_center": str(
                        row.get("Workday Cost Center", "")
                    ),
                    "work_site_name": str(row.get("Work Site Name", "")),
                    "city": str(row.get("City", "")),
                    "state": str(row.get("State", "")),
                    "start_date": str(row.get("Start Date", "")),
                    "end_date": str(row.get("End Date", "")),
                    "min_experience": row.get("Min Experience", ""),
                    "max_experience": row.get("Max Experience", ""),
                    "must_have_skills": str(row.get("Must Have Skills", "")),
                    "nice_to_have_skills": str(
                        row.get("Nice To Have Skills", "")
                    ),
                    "rate": row.get("Rate", ""),
                    "max_bill_rate": row.get("Max Bill Rate", ""),
                    "created_on": str(row.get("Created On", "")),
                    "job_approved_date": str(
                        row.get("Job Approved Date", "")
                    ),
                    "status_date": str(row.get("Status Date", "")),
                    "job_category": str(row.get("Job Category", "")),
                    "time_type": str(row.get("Time Type", "")),
                    "single_sourced": str(row.get("Single Sourced", "")),
                    "intake_call_completed": str(
                        row.get("Intake Call Completed", "")
                    ),
                    "supplier_call_conducted": str(
                        row.get("Supplier Call Conducted", "")
                    ),
                    "mnpi": str(row.get("MNPI", "")),
                    "last_api_refresh": str(
                        row.get("Last API Refresh", "")
                    ),
                    "desc": desc,
                    "notes": notes,
                    "transcript": transcript,
                    "internal_notes": internal_notes,
                }

    if "fred_jd" not in st.session_state:
        st.session_state.fred_jd = ""

    if "fred_notes" not in st.session_state:
        st.session_state.fred_notes = ""

    if "fred_trans" not in st.session_state:
        st.session_state.fred_trans = ""

    if "fred_internal_notes" not in st.session_state:
        st.session_state.fred_internal_notes = ""

    def update_fred_jd_text():
        selected = st.session_state.get(
            "fred_req_selector",
            "None",
        )

        # Clear answers whenever the requisition changes so an answer
        # from Job A can NEVER accidentally survive into Job B.
        for key in list(st.session_state.keys()):
            if key.startswith("fred_vetting_answer_"):
                del st.session_state[key]

        if (
            selected != "None"
            and selected in req_mapping
        ):
            req = req_mapping[selected]

            st.session_state.fred_jd = (
                req.get("desc", "").strip()
            )

            st.session_state.fred_notes = (
                req.get("notes", "").strip()
            )

            st.session_state.fred_trans = (
                req.get("transcript", "").strip()
            )

            st.session_state.fred_internal_notes = (
                req.get("internal_notes", "").strip()
            )

        else:
            st.session_state.fred_jd = ""
            st.session_state.fred_notes = ""
            st.session_state.fred_trans = ""
            st.session_state.fred_internal_notes = ""

    selected_req = st.selectbox(
        "🔍 Select a Freddie Mac Requisition (Type to search):",
        options=options,
        key="fred_req_selector",
        on_change=update_fred_jd_text,
    )

    selected_vndly_context = (
        req_mapping.get(selected_req, {})
        if selected_req != "None"
        else {}
    )

    # Operational VNDLY metadata is useful to the recruiter, but most of it
    # should not become candidate claims. Keep it in a collapsed snapshot.
    if selected_vndly_context:
        with st.expander(
            "📊 VNDLY Requisition Snapshot",
            expanded=False,
        ):
            snapshot_lines = []

            def add_snapshot(label, value):
                value = str(value or "").strip()
                if value:
                    snapshot_lines.append(f"**{label}:** {value}")

            add_snapshot("Job ID", selected_vndly_context.get("job_id"))
            add_snapshot("Status", selected_vndly_context.get("status"))
            add_snapshot("Job Title", selected_vndly_context.get("job_title"))
            add_snapshot(
                "Resource Manager",
                selected_vndly_context.get("resource_manager"),
            )
            add_snapshot(
                "Organization Unit",
                selected_vndly_context.get("organization_unit"),
            )
            add_snapshot(
                "Cost Center",
                selected_vndly_context.get("cost_center"),
            )
            add_snapshot(
                "Work Site",
                selected_vndly_context.get("work_site_name"),
            )
            add_snapshot(
                "Location",
                ", ".join(
                    [
                        value
                        for value in [
                            str(selected_vndly_context.get("city", "")).strip(),
                            str(selected_vndly_context.get("state", "")).strip(),
                        ]
                        if value
                    ]
                ),
            )
            add_snapshot(
                "Positions",
                (
                    f"{selected_vndly_context.get('open_positions', '')} open / "
                    f"{selected_vndly_context.get('total_positions', '')} total"
                ).strip(),
            )
            add_snapshot(
                "Max Submissions / Vendor",
                selected_vndly_context.get("max_submissions_per_vendor"),
            )
            add_snapshot(
                "Total Supplier Candidates",
                selected_vndly_context.get("total_supplier_candidates"),
            )
            add_snapshot(
                "Distributed to Convergenz",
                selected_vndly_context.get("distributed_to_convergenz"),
            )
            add_snapshot(
                "End Date",
                selected_vndly_context.get("end_date"),
            )
            add_snapshot(
                "Max Bill Rate",
                selected_vndly_context.get("max_bill_rate"),
            )

            if snapshot_lines:
                st.markdown("\n\n".join(snapshot_lines))

    Job_Description = st.text_area(
        "1. Official VNDLY Job Description:",
        height=150,
        key="fred_jd",
    )

    MSP_Notes = st.text_area(
        "2. VNDLY / MSP Notes (Optional):",
        height=120,
        key="fred_notes",
    )

    Call_Transcript = st.text_area(
        "3. Spotlight Call Transcript "
        "(Optional - Hiring Manager Priority):",
        height=120,
        key="fred_trans",
    )

    # Internal Notes remain recruiter context. They are visible/editable so the
    # recruiter knows exactly what context is being supplied to the analyzer.
    Internal_Notes = st.text_area(
        "4. Internal Notes (Optional - Recruiter Context):",
        height=80,
        key="fred_internal_notes",
    )

    # Use the currently displayed Internal Notes value in the analysis context,
    # allowing a recruiter to add/adjust context without editing the source sheet.
    selected_vndly_context = dict(selected_vndly_context)
    selected_vndly_context["internal_notes"] = Internal_Notes

    # ================================================================
    # STRUCTURED REQUISITION ANALYSIS
    # ================================================================

    req_analysis = {
        "target_title": "",
        "explicit_must_have_requirements": [],
        "must_have_competency_groups": [],
        "required_requirements": [],
        "preferred_requirements": [],
        "domain_requirements": [],
        "current_manager_priorities": [],
        "selection_feedback": [],
        "role_clarifications": [],
        "role_evolution": [],
        "deemphasized_requirements": [],
        "explicitly_excluded_requirements": [],
        "deemphasized_or_negated_requirements": [],
        "supplier_vetting_questions": [],
    }

    analysis_error = None

    if (
        API_KEY
        and Job_Description
        and Job_Description.strip()
    ):
        try:
            with st.spinner(
                "Analyzing Freddie Mac requirements "
                "and vetting questions..."
            ):
                req_analysis = fred_analyze_requisition(
                    API_KEY,
                    Job_Description,
                    MSP_Notes,
                    Call_Transcript,
                    selected_vndly_context,
                )

        except Exception as e:
            analysis_error = str(e)

            st.error(
                "❌ Could not analyze the Freddie Mac requisition. "
                f"{analysis_error}"
            )

    # Optional transparency for recruiters. Show the analysis whenever ANY
    # substantive requirement intelligence was found; some Freddie roles have
    # no formally labeled Must-Have block but still contain important manager
    # priorities, clarifications, selection feedback, or role evolution.
    has_req_intelligence = any(
        req_analysis.get(key)
        for key in [
            "explicit_must_have_requirements",
            "required_requirements",
            "preferred_requirements",
            "current_manager_priorities",
            "selection_feedback",
            "role_clarifications",
            "role_evolution",
            "deemphasized_requirements",
            "explicitly_excluded_requirements",
            "deemphasized_or_negated_requirements",
            "must_have_competency_groups",
        ]
    )

    if has_req_intelligence:
        with st.expander(
            "🎯 Freddie Requirement Analysis",
            expanded=False,
        ):
            st.markdown("**Explicit Must-Have Requirements:**")
            for requirement in req_analysis[
                "explicit_must_have_requirements"
            ]:
                st.markdown(f"- {requirement}")

            current_priorities = req_analysis.get(
                "current_manager_priorities",
                [],
            )
            if current_priorities:
                st.markdown("**Current Manager Priorities:**")
                for item in current_priorities:
                    st.markdown(f"- {item}")

            clarifications = req_analysis.get(
                "role_clarifications",
                [],
            )
            if clarifications:
                st.markdown("**Latest Role Clarifications:**")
                for item in clarifications:
                    st.markdown(f"- {item}")

            deemphasized = req_analysis.get("deemphasized_requirements", [])
            if deemphasized:
                st.markdown("**De-emphasized / Secondary for Manager Review:**")
                for item in deemphasized:
                    st.markdown(f"- {item}")

            excluded = req_analysis.get("explicitly_excluded_requirements", [])
            if excluded:
                st.markdown("**Explicitly Excluded / Negative Profile Signals:**")
                for item in excluded:
                    st.markdown(f"- {item}")

            selection_feedback = req_analysis.get(
                "selection_feedback",
                [],
            )
            if selection_feedback:
                st.markdown("**Selection / Interview Feedback:**")
                for item in selection_feedback:
                    st.markdown(f"- {item}")

            evolution = req_analysis.get(
                "role_evolution",
                [],
            )
            if evolution:
                st.markdown("**Role Evolution:**")
                for item in evolution:
                    if isinstance(item, dict):
                        date_label = str(item.get("date", "") or "").strip()
                        change = str(item.get("change", "") or "").strip()
                        if change:
                            prefix = f"{date_label}: " if date_label else ""
                            st.markdown(f"- {prefix}{change}")

            groups = req_analysis.get(
                "must_have_competency_groups",
                [],
            )

            if groups:
                st.markdown("**Planned Skills Priorities:**")

                for idx, group in enumerate(
                    groups,
                    start=1,
                ):
                    terms = ", ".join(
                        group.get("terms", [])
                    )

                    if terms:
                        st.markdown(
                            f"{idx}. **{group.get('name', '')}** "
                            f"- {terms}"
                        )
                    else:
                        st.markdown(
                            f"{idx}. **{group.get('name', '')}**"
                        )

    # ================================================================
    # DYNAMIC OFFICIAL SUPPLIER VETTING QUESTIONS
    # ================================================================

    st.header("🎤 Supplier Vetting Questions")

    vetting_questions = req_analysis.get(
        "supplier_vetting_questions",
        [],
    )

    vetting_answers = []
    vetting_required_flags = []

    if analysis_error:
        st.warning(
            "Vetting questions cannot be safely determined until "
            "the requisition analysis succeeds."
        )

    elif not Job_Description.strip():
        st.info(
            "Select a Freddie Mac requisition above. "
            "Any official supplier vetting questions will appear here "
            "automatically."
        )

    elif not vetting_questions:
        st.success(
            "✅ No official Supplier Vetting Questions were identified "
            "for this requisition."
        )

    else:
        st.caption(
            "Questions below were extracted from Freddie Mac's VNDLY "
            "requisition. Enter the candidate's direct responses. "
            "The question wording is intentionally read-only."
        )

        safe_req_key = re.sub(
            r"[^A-Za-z0-9_\-]+",
            "_",
            selected_req or "manual"
        )
        
        for idx, question in enumerate(
            vetting_questions,
            start=1,
        ):
            response_required = fred_vetting_response_required(question)
            vetting_required_flags.append(response_required)

            if response_required:
                st.markdown(
                    f"**{idx}. {question}**"
                )
                response_label = f"Candidate Response {idx}"
                response_placeholder = (
                    "Paste the candidate's direct response here..."
                )
            else:
                st.markdown(
                    f"**{idx}. {question}**  \n"
                    "*Optional - Freddie/MSP does not require a response.*"
                )
                response_label = f"Candidate Response {idx} (Optional)"
                response_placeholder = (
                    "Optional - leave blank if nothing is being submitted."
                )

            answer = st.text_area(
                response_label,
                height=100,
                key=f"fred_vetting_answer_{safe_req_key}_{idx}",
                placeholder=response_placeholder,
            )

            vetting_answers.append(answer)

    # ================================================================
    # OPTIONAL CANDIDATE-CONFIRMED FOLLOW-UP DETAILS
    # ================================================================

    st.header("✅ Candidate-Confirmed Clarifications (Optional)")
    st.caption(
        "Paste direct candidate follow-up answers or recruiter-verified candidate facts "
        "that are not already clear on the resume. Examples: executive reporting scope, "
        "portfolio size, older tool experience, leadership audience, or other factual "
        "clarifications. These are treated as secondary candidate evidence, never as role/JD evidence."
    )
    Candidate_Clarifications_Raw = st.text_area(
        "Candidate-Confirmed Follow-Up Answers / Clarifications:",
        height=140,
        key="fred_candidate_clarifications",
        placeholder="Optional. Paste direct candidate responses or verified screening notes here...",
    )
    Candidate_Clarifications = fred_clean_candidate_clarifications(
        Candidate_Clarifications_Raw
    )

    # ================================================================
    # FILE UPLOAD
    # ================================================================

    st.header("📂 File Uploads")

    resume_file = st.file_uploader(
        "📤 Upload the candidate's resume...",
        type=["pdf", "docx", "doc"],
        key="fred_res",
    )

    # ================================================================
    # GENERATE SUBMISSION
    # ================================================================

    if st.button(
        "🚀 Generate Freddie Mac Submission",
        type="primary",
    ):

        # ------------------------------------------------------------
        # PRE-GENERATION VALIDATION
        # ------------------------------------------------------------

        validation_errors = []

        if not API_KEY:
            validation_errors.append(
                "Gemini API Key is missing."
            )

        if not Job_Description.strip():
            validation_errors.append(
                "A Freddie Mac requisition / Job Description "
                "must be selected."
            )

        if analysis_error:
            validation_errors.append(
                "The requisition analysis did not complete successfully."
            )

        if not Current_Location.strip():
            validation_errors.append(
                "Current Location is required."
            )

        if (
            Available_Onsite == "Yes"
            and (
                not Site.strip()
                or Site.strip().upper() == "N/A"
            )
        ):
            validation_errors.append(
                "Enter the location(s) where the candidate "
                "can be onsite."
            )

        if not Interview.strip():
            validation_errors.append(
                "Interview Availability for the next 7 days "
                "is required."
            )

        if (
            Former_CW == "Yes"
            and Eligibility == "N/A"
        ):
            validation_errors.append(
                "Rehire Eligibility cannot be N/A when the "
                "candidate is a former Freddie Mac CW/employee."
            )

        if not resume_file:
            validation_errors.append(
                "Upload the candidate's resume."
            )

        if not os.path.exists(TEMPLATE_FILENAME):
            validation_errors.append(
                f"The template '{TEMPLATE_FILENAME}' was not found."
            )

        if vetting_questions:
            missing_answers = [
                str(idx)
                for idx, (answer, response_required) in enumerate(
                    zip(
                        vetting_answers,
                        vetting_required_flags,
                    ),
                    start=1,
                )
                if response_required and not str(answer).strip()
            ]

            if missing_answers:
                validation_errors.append(
                    "Candidate responses are required for the mandatory "
                    "Supplier Vetting Question(s). Missing response(s): "
                    + ", ".join(missing_answers)
                )

        if validation_errors:
            for error in validation_errors:
                st.error(f"❌ {error}")

            return

        # ------------------------------------------------------------
        # PROCESS
        # ------------------------------------------------------------

        with st.spinner(
            "Building optimized Freddie Mac submission..."
        ):
            resume_path = None

            try:
                # ====================================================
                # SAVE / EXTRACT ORIGINAL RESUME
                # ====================================================

                safe_original_name = os.path.basename(
                    resume_file.name
                )

                resume_path = (
                    f"temp_freddie_{safe_original_name}"
                )

                with open(resume_path, "wb") as f:
                    f.write(
                        resume_file.getbuffer()
                    )

                raw_text = extract_text(resume_path)

                # ====================================================
                # PASS 1 - STRUCTURED RESUME EXTRACTION
                # ====================================================

                extraction_prompt = f"""
Return a valid JSON object ONLY.

You are extracting a candidate's ORIGINAL resume into a clean structured
format for a Freddie Mac VNDLY submission.

CRITICAL:
- This is an extraction step, NOT a tailoring step.
- Preserve the candidate's factual work history.
- Do NOT add JD terminology.
- Do NOT rewrite accomplishments to sound more relevant.
- Do NOT invent technologies, metrics, employers, dates, or responsibilities.

RULES:

1. NAME
- Pull from the resume header.
- Title Case.
- If completely missing/unreadable, use the filename if possible.

2. COMPANY
- Extract the primary END-CLIENT / employer name actually represented by the
  resume.
- Strip geographic locations that are merely appended to the company line.
- If a contracting vendor and end-client are clearly shown together, preserve
  the primary end-client used by the resume experience.
- Do not guess an end-client that is not stated.

3. EDUCATION
- Extract school and degree exactly enough to preserve meaning.

4. EXPERIENCE
Return EVERY dated professional role shown anywhere in the resume, including
condensed or abbreviated entries under headings such as "Earlier Experience",
"Additional Experience", or similar sections. Do not stop after the first seven
roles. Older abbreviated roles may have an empty Bullets array and blank
Environment; they must still be returned when Company/Title/Dates are present so
total career experience can be calculated correctly.

For each role return:
- Company
- Title
- Bullets as an ARRAY
- Environment
- Dates

TITLE:
- Remove only employment-type modifiers such as "(Contract)" or "- Contract"
  when they are merely suffixes.
- Do not upgrade or alter seniority.

BULLETS:
- Preserve every substantive original bullet.
- Light cleanup of malformed bullet characters/spacing is allowed.
- Keep bullets in ORIGINAL ORDER in this extraction response.
- Do NOT rewrite bullets to match Freddie's JD.

ENVIRONMENT:
- Populate ONLY when the original resume explicitly has an "Environment:"
  or "Technologies:" line.
- Never compile an Environment line from ordinary bullets.

DATES:
- Preserve the source dates faithfully enough for later date calculations.
- NEVER invent a month that the source resume does not provide.
- If a role is shown only as years (for example "2018 - 2020"), return only
  those years. Do NOT convert it to "January 2018 - January 2020", "01/2018",
  or any other assumed month.
- Mixed precision must also stay mixed (for example "2018 - March 2020").

5. CERTIFICATIONS
- Extract certifications into one comma-separated string.
- Preserve completion status exactly when the resume distinguishes completed
  certifications from credentials that are still "In Progress", "Pursuing",
  "Expected", "Pending", or otherwise not yet earned.
- NEVER present an in-progress or pending credential as completed.
- For an unfinished credential, append the status directly to the credential
  name using this format:
  "PMP - In Progress"
  "CBAP - In Progress"
- If the resume explicitly labels a group of certifications as "In Progress",
  apply "- In Progress" to every credential under that heading until the next
  clear section heading or status change.
- Do not invent an in-progress status for completed credentials.
- Do not remove a credential solely because it is in progress; preserve it with
  the explicit status.
- If none are found, return "".

CERTIFICATION TRUTH CHECK:
Before returning Certifications, compare every credential against the original
resume and confirm that no credential labeled In Progress/Pursuing/Pending has
been output as though it were completed.

OUTPUT:

{{
  "FullName": "",
  "Certifications": "",
  "Education": [
    {{
      "School": "",
      "Degree": ""
    }}
  ],
  "Experience": [
    {{
      "Company": "",
      "Title": "",
      "Bullets": [
        ""
      ],
      "Environment": "",
      "Dates": ""
    }}
  ]
}}

FILENAME:
{resume_file.name}

ORIGINAL RESUME:
{raw_text}
"""

                data = fred_generate_json(
                    API_KEY,
                    extraction_prompt,
                )

                name = str(
                    data.get(
                        "FullName",
                        "",
                    )
                ).strip().title()

                certifications = str(
                    data.get(
                        "Certifications",
                        "",
                    )
                ).strip()

                education = data.get(
                    "Education",
                    [],
                )

                if not isinstance(
                    education,
                    list,
                ):
                    education = []

                experience = data.get(
                    "Experience",
                    [],
                )

                if not isinstance(
                    experience,
                    list,
                ):
                    experience = []

                # Normalize ALL extracted dated roles first. Older abbreviated
                # "Earlier Experience" entries are retained here for accurate
                # Preserve every dated role for both total-career experience math
                # and the generated Freddie work-history section.
                normalized_experience = []

                for role in experience:
                    if not isinstance(
                        role,
                        dict,
                    ):
                        continue

                    bullets = role.get(
                        "Bullets",
                        [],
                    )

                    if isinstance(
                        bullets,
                        str,
                    ):
                        bullets = [
                            x.strip()
                            for x in bullets.split("\n")
                            if x.strip()
                        ]

                    if not isinstance(
                        bullets,
                        list,
                    ):
                        bullets = []

                    bullets = [
                        re.sub(
                            r"^[\s\-\•\*\d\.\)]+",
                            "",
                            str(x),
                        ).strip()
                        for x in bullets
                        if str(x).strip()
                    ]

                    normalized_experience.append(
                        {
                            "Company": str(
                                role.get(
                                    "Company",
                                    "",
                                )
                            ).strip(),
                            "Title": re.sub(
                                r"\s*\(.*$",
                                "",
                                str(
                                    role.get(
                                        "Title",
                                        "",
                                    )
                                ),
                            ).strip(),
                            "Bullets": bullets,
                            "Environment": str(
                                role.get(
                                    "Environment",
                                    "",
                                )
                            ).strip(),
                            "Dates": fred_standardize_dates(
                                str(
                                    role.get(
                                        "Dates",
                                        "",
                                    )
                                ).strip()
                            ),
                        }
                    )

                all_experience_for_years = normalized_experience
                experience = normalized_experience

                # ====================================================
                # OFFICIAL VETTING Q&A STRUCTURE
                # ====================================================

                vetting_pairs = [
                    {
                        "question": question,
                        "answer": answer.strip(),
                    }
                    for question, answer in zip(
                        vetting_questions,
                        vetting_answers,
                    )
                    if str(answer).strip()
                ]

                vetting_for_ai = "\n\n".join(
                    [
                        f"Q{idx}: {pair['question']}\n"
                        f"A{idx}: {pair['answer']}"
                        for idx, pair in enumerate(
                            vetting_pairs,
                            start=1,
                        )
                    ]
                )

                if not vetting_for_ai:
                    vetting_for_ai = (
                        "No candidate-supplied Supplier Vetting responses "
                        "were provided or required for this requisition."
                    )

                # ====================================================
                # PASS 2 - CANDIDATE MATCH / SUMMARY / SKILLS /
                #          SAFE BULLET PRIORITIZATION
                # ====================================================

                structured_req = json.dumps(
                    req_analysis,
                    indent=2,
                    ensure_ascii=False,
                )

                structured_exp = json.dumps(
                    experience,
                    indent=2,
                    ensure_ascii=False,
                )

                current_date_obj = date.today()
                current_date = current_date_obj.isoformat()

                calculated_total_experience = fred_calculate_total_experience(
                    all_experience_for_years,
                    current_date_obj,
                )

                tailoring_prompt = f"""
Return a valid JSON object ONLY.

You are an elite Senior Technical Recruiter preparing a Freddie Mac
contingent-worker submission through Workday VNDLY.

CURRENT DATE:
{current_date}

PYTHON-CALCULATED TOTAL PROFESSIONAL EXPERIENCE:
{calculated_total_experience}

STRUCTURED VNDLY EXPERIENCE RANGE:
Minimum: {selected_vndly_context.get("min_experience", "") or "BLANK"}
Maximum: {selected_vndly_context.get("max_experience", "") or "BLANK"}

The total professional experience value above was calculated deterministically
from the candidate's dated work history.

CRITICAL:
- PYTHON-CALCULATED TOTAL PROFESSIONAL EXPERIENCE is available for accuracy,
  but it is NOT mandatory to lead with or even state total years in the Summary.
- Use the total-years figure only when it materially strengthens the shortlist case.
- If structured VNDLY experience guidance shows the candidate is materially
  above the target range, DO NOT state the total-career years figure in the
  Summary. Use a natural authority phrase such as "extensive experience".
- If you do state total career experience, use the exact deterministic value.
- Do NOT independently recalculate it, round it up, or alter its wording.
- If the calculated value is blank, do not invent a total-years figure.

There are THREE sequential audiences for this submission:

1. VNDLY / AUTOMATED MATCHING
   Make the candidate's genuine match to Freddie's highest-priority
   requirements easy for an automated matching system to detect.

2. MSP HUMAN SHORTLIST REVIEW
   Assume an MSP recruiter understands the requisition and staffing process
   but may NOT have the hiring manager's depth of technical knowledge.
   The MSP reviewer should be able to determine within approximately
   20-30 seconds:
   - Does this candidate clearly possess Freddie's most important Must Haves?
   - Is the required specialized experience genuinely demonstrated?
   - Is the required business/domain background present?
   - Is there concrete evidence behind the keywords?
   - Is there any obvious reason NOT to shortlist the candidate?

3. FREDDIE MAC HIRING MANAGER REVIEW
   If shortlisted, preserve enough specific technical and operational evidence
   for the hiring manager to immediately understand the candidate's depth and
   want to interview them.

DUAL-AUDIENCE WRITING PRINCIPLE:
- Do NOT dumb down or remove meaningful technical evidence.
- Instead, TRANSLATE specialized technical work into clear functional language
  that an MSP recruiter can understand, then support it with the exact
  technologies, methodologies, and domain terminology that matter to Freddie.
- The MSP recruiter should understand WHY the experience is relevant without
  needing to understand every underlying technology.
- The hiring manager should still see HOW the candidate performed the work and
  which technologies were actually used.
- Optimize for:
  clear match first,
  technical proof second,
  promotional language never.

======================================================================
HIERARCHY OF ROLE REQUIREMENTS
======================================================================

Use the ALREADY-ANALYZED structured requisition intelligence.

For EMPHASIS and PRESENT-DAY selection strategy, use this order:

1. Latest explicit hiring-manager correction / selection feedback that
   substantively changes or clarifies what future candidates must demonstrate
2. Current hiring-manager priorities / latest substantive Spotlight or MSP
   clarification
3. Freddie's formal Must Have Qualifications and populated structured VNDLY
   Must Have Skills
4. Detailed Required Qualifications and core recurring responsibilities
5. Preferred / Nice-to-Have qualifications

CRITICAL:
- "Latest" matters only when the newer information is substantive role or
  selection guidance. A newer scheduling/offer/OOO update does NOT override an
  older substantive manager requirement.
- Use selection_feedback and role_clarifications to understand why previous
  candidates failed and what must be made obvious in a new submission.
- Use role_evolution to understand how the target profile changed over time.
- Distinguish DE-EMPHASIZED from EXCLUDED. A manager statement such as
  "we don't use user stories" lowers that concept's hiring-manager priority but
  does not automatically remove a genuine candidate/JD match from MSP-facing
  coverage. Keep it secondary when useful and truthful.
- Only `explicitly_excluded_requirements` are true candidate-facing prohibitions.
- Formal Must Haves still require strong truthful coverage unless authoritative
  manager guidance clearly changes the requirement.
- Operational staffing noise must have ZERO influence on candidate claims.
- Internal recruiter notes may guide emphasis but may not manufacture a
  requirement or override explicit Freddie guidance.

The requisition has ALREADY been analyzed into structured job intelligence.
Do not independently reinvent its hierarchy.

STRUCTURED REQUISITION:
{structured_req}

======================================================================
VNDLY / MATCHING STRATEGY
======================================================================

When the candidate genuinely possesses a Freddie Must Have:

- Use Freddie's exact skill/technology/domain terminology when supported.
- Prefer the exact Freddie term over a synonym when both are truthful.
- Do NOT merely repeat keywords.
- Pair important terms with contextual evidence of how the candidate actually
  used them.
- Favor demonstrated skills in work-experience bullets over a standalone skills
  inventory.
- Favor recent, repeated, production, operational, and hands-on use.
- Domain requirements such as mortgage, financial services, appraisal,
  ServiceNow IRM/GRC, fixed income, etc. can be as important as technologies.

Do not engage in keyword stuffing.

======================================================================
CANDIDATE POSITIONING TITLE - NATURAL MARKET TITLE
======================================================================

Before writing the Summary, choose ONE concise professional title for the
candidate and return it as CANDIDATE_TITLE.

PURPOSE:
This is the natural functional title used to describe the candidate in the
submission narrative. It is NOT automatically Freddie's VNDLY Job Title.

CRITICAL DISTINCTION:
- Freddie's VNDLY Job Title may be a VMS/rate-card/taxonomy label selected for
  budgeting, job-family, or system purposes.
- Treat the VNDLY Job Title and STRUCTURED REQUISITION target_title as CLUES
  about the role, not as wording that must be copied into the candidate summary.
- Do NOT default to awkward system labels such as:
  "Developer-Java/J2EE Specialist",
  "Developer-Python Specialist",
  "Project Manager-Senior",
  or similar VMS/rate-card constructions.
- Preserve an official VNDLY title only when it is already a normal, natural
  market-facing professional title AND accurately describes the candidate.

HOW TO CHOOSE CANDIDATE_TITLE:
Use ALL relevant evidence together:
1. What function the current role actually needs, based on the JD plus current
   manager/MSP/Spotlight guidance.
2. What the candidate has actually done in dated work history.
3. The candidate's recent/repeated professional titles and responsibilities.
4. A more natural role title stated inside the body of the JD, when present.
5. The VNDLY Job Title only as supporting context, never as the default wording.

TITLE QUALITY RULES:
- Prefer a conventional external-market title a recruiter or hiring manager
  would naturally say aloud.
- Usually 2-5 words.
- Prefer one clear occupational lane or at most two closely related functional
  lanes. Do not create a broad mash-up title that stacks several senior
  specialties when a simpler target-relevant identity describes the candidate.
- CANDIDATE_TITLE must be a grammatically complete occupational title that can
  stand on its own after the words "[First Name] is a...".
- Do NOT copy a resume heading or job title literally when it combines the
  occupation with appended specialty descriptors after a comma, pipe, slash,
  colon, or similar separator.
- Example source/current title:
  "Data Engineer, Python, Computer Vision & Data Pipelines"
  Preferred CANDIDATE_TITLE:
  "Data Engineer"
  Then describe the specialties naturally in Sentence 1, for example:
  "Vamshidhar is a Data Engineer specializing in Python, computer vision, and
  data pipelines..."
- Another acceptable construction is:
  "[First Name] is a Data Engineer with 7+ years of professional experience in
  Python, computer vision, and data pipelines..."
- Never produce an incomplete construction such as:
  "[First Name] is a Data Engineer, Computer Vision & Data Pipelines with..."
- Keep it functional and specific enough to communicate the candidate's lane.
- Use seniority only when BOTH the candidate's demonstrated career level and
  the target role support it.
- Do not inflate to Lead, Principal, Architect, Manager, etc. merely because the
  candidate once held that title.
- Do not downgrade an established senior candidate merely because the VNDLY
  rate-card title omits seniority.
- Avoid awkward VMS punctuation/order such as "Developer-Java/J2EE Specialist."
- Do not force the target role's function onto a candidate whose actual work
  materially differs from it.
- When the JD itself contains a natural role heading that is more descriptive
  than the VNDLY system title, strongly prefer that heading when the candidate
  genuinely supports it.

EXAMPLE:
VNDLY Job Title: "Developer-Java/J2EE Specialist"
JD body role heading: "Senior Java Developer"
Candidate: 18+ years of Java/Spring/J2EE development

Preferred CANDIDATE_TITLE:
"Senior Java Developer"

Not:
"Developer-Java/J2EE Specialist"

SUMMARY CONSISTENCY RULE:
- If Sentence 1 identifies the candidate by a professional title, use the exact
  CANDIDATE_TITLE you selected.
- Do not introduce a competing title elsewhere in the Summary.

======================================================================
SUMMARY - EXACTLY 4 SENTENCES
======================================================================

Write exactly FOUR sentences as ONE paragraph.

EVIDENCE-FIRST TRANSFERABILITY RULE - CRITICAL:
The goal is NOT to make the candidate's past sound identical to the Freddie job.
The goal is to show:
1. WHAT the candidate actually did,
2. HOW they did it / which supported tools or methods they used,
3. WHAT impact, outcome, ownership, or operational value resulted when the
   resume provides it, and
4. WHY that proven experience transfers to the work Freddie needs done.

Keep PAST EVIDENCE and TARGET-ROLE APPLICATION grammatically separate.

BAD:
"Arunabh specialized in data-modeling workflow testing and automated quality
gates."

when the resume only shows enterprise QA, Python automation, deployment
validation, and CI/CD.

GOOD:
"Arunabh's enterprise QA and Python automation experience provides a strong
foundation for building and executing structured test cases around Freddie's
data-modeling platform."

The GOOD version explains transferability without rewriting the candidate's
history.

Do not convert:
- a Freddie requirement into a past candidate accomplishment,
- a target platform/domain into an employer context the resume does not show,
- related experience into direct hands-on ownership,
- two unrelated resume facts into one stronger composite claim.

Write exactly FOUR sentences as ONE paragraph.

The Summary serves TWO human readers in sequence:

1. An MSP recruiter who knows Freddie's requirements but may not be deeply
   technical.
2. The Freddie Mac hiring manager who can evaluate the technical depth.

The first half of the Summary must make the shortlist decision easy.
The second half must provide the concrete proof that preserves hiring-manager
credibility.

----------------------------------------------------------------------
Sentence 1 - IDENTITY + STRONGEST DIFFERENTIATOR
----------------------------------------------------------------------

- Use the candidate's FIRST NAME.
- Use the exact CANDIDATE_TITLE selected under the Natural Market Title rules
  above when identifying the candidate by title.
- CANDIDATE_TITLE is the occupation only. If relevant specialties need to appear
  in Sentence 1, connect them with natural grammar such as "specializing in",
  "focused on", or "with experience in". Do not append specialty phrases to the
  title with a comma or pipe and then continue directly into "with X+ years".
- Do NOT copy the VNDLY Job Title merely because it is present in the structured
  requisition.
- Do not relabel a Data Engineer as a Data Scientist solely because Freddie's
  requisition uses that title. Target-role terminology may be used only when
  the candidate's actual work clearly supports that functional identity.
- You MAY use PYTHON-CALCULATED TOTAL PROFESSIONAL EXPERIENCE when it
  strengthens the shortlist case, but do not force it into Sentence 1.
- When the candidate is far above a stated target experience range, prefer
  "extensive experience" or an equally natural authority phrase over a large
  number that could create an unnecessary overqualification concern.
- If you do state total career experience, use the exact deterministic value and
  never independently calculate or alter it.
- CRITICAL GRAMMAR RULE: The total-career experience figure must NOT be written
  in a way that implies the candidate has that same number of years with every
  skill named later in the sentence.
- Do NOT write:
  "7+ years of experience in Python, SQL, and Excel"
  unless EACH named skill independently supports 7+ years.
- Prefer:
  "7+ years of professional experience, including 4+ years of Python scripting
  and extensive SQL and Excel data analysis"
  when the individual skills have different tenures.
- Give the MSP reviewer an immediate checklist-level view of the candidate's
  match to Freddie's highest-priority requirements.
- Prioritize approximately 2-4 of the most important supported Must Have
  competencies rather than trying to mention everything.
- Keep this sentence understandable without requiring deep technical knowledge.

----------------------------------------------------------------------
Sentence 2 - CONCRETE RECENT PROOF
----------------------------------------------------------------------

- Use the current/most-recent employer when it provides strong relevant proof.
- HOWEVER, if an earlier employer provides materially stronger evidence for
  Freddie's central need, use that stronger experience instead.
- Explain the candidate's most unusually relevant experience in clear
  functional language.
- Focus first on WHAT the candidate actually did and WHY it is relevant to
  Freddie's need.
- Do not require the MSP reviewer to interpret specialized technical jargon to
  understand the match.
- If a specialized technical concept is important, translate its practical
  meaning.

Example principle:

Instead of relying only on:
"parsed bounding boxes, labels, confidence scores, and metadata"

prefer a construction such as:
"worked directly with appraisal PDFs and property images, validating
information extracted by computer vision models against the source documents"

when that is supported by the resume.

The technical specifics can then appear in Sentence 3.

----------------------------------------------------------------------
Sentence 3 - SUPPORTING DEPTH + IMPACT
----------------------------------------------------------------------

- Now provide the concrete technical proof behind the plain-English match.
- Explain HOW the relevant work was performed.
- Naturally incorporate Freddie's most important exact technologies,
  methodologies, operational concepts, or domain terminology when supported.
- Favor demonstrated hands-on use over standalone skills-list mentions.
- Emphasize production ownership, validation, testing, troubleshooting,
  reconciliation, scale, measurable impact, or operational responsibility when
  relevant.
- Use technical terminology where it adds evidence rather than complexity.
- Do NOT create a dense keyword list.
- Normally use no more than 4 material named technologies/tools in this
  sentence. A fifth is allowed only when necessary to describe one coherent
  workflow accurately.

----------------------------------------------------------------------
Sentence 4 - TRANSFER TO THIS FREDDIE ROLE
----------------------------------------------------------------------

- Tie together the strongest remaining evidence that increases confidence in
  the shortlist decision.
- When Freddie requires or strongly values domain experience such as mortgage,
  financial services, appraisal, capital markets, fixed income, ServiceNow
  IRM/GRC, etc., make that domain connection explicit when genuinely supported.
- State the specific value the candidate's demonstrated experience indicates
  they can provide to THIS Freddie role.
- This is the ONE sentence where target-role language may be used explicitly to
  explain transferability, even when the candidate did not previously work in
  that exact Freddie environment.
- Make the grammar unmistakable: prior experience "provides a foundation for",
  "transfers to", "supports", or "can be applied to" the Freddie need.
- NEVER phrase the Freddie target responsibility as though the candidate
  already performed it at a prior employer.
- The conclusion should make the candidate's relevance obvious rather than
  merely praise the candidate.
- Keep it evidence based and grounded in prior execution.
- Do not say "meets the requirements", "great fit", "ideal candidate",
  "strong candidate", or similar evaluative conclusions.
- Do not mention city, onsite requirement, interview availability, rate,
  authorization, or sponsorship.

======================================================================
MSP TRANSLATION RULE
======================================================================

For specialized technical experience:

PLAIN-LANGUAGE RELEVANCE FIRST -> TECHNICAL PROOF SECOND.

The Summary should answer both:

"What does this experience mean for the work Freddie needs done?"

and

"What specific evidence proves the candidate can do it?"

Do not assume that an MSP reviewer understands why a technology is relevant
merely because the technology name appears.

Do not remove important technical terminology simply to make the Summary
easier to read.

Translate the relevance; preserve the evidence.

======================================================================
VETTING-ANSWER INTERACTION
======================================================================

If official Freddie Supplier Vetting Questions exist:

- Assume the MSP reviewer may read those responses before or alongside the
  Summary.
- Do NOT waste Summary space repeating detailed vetting answers.
- Use the Summary to make the CONCLUSION of the candidate's match immediately
  visible.
- Let the vetting answers provide additional depth and confirmation.
- Candidate-supplied vetting responses remain secondary evidence and may not
  override contradictory resume evidence.

======================================================================
STYLE
======================================================================

- Human recruiter voice.
- NEVER use an em dash. Use commas, parentheses, colons, semicolons, or standard hyphens instead.
- Direct, factual, concise.
- Evidence over praise.
- Clear enough for a knowledgeable staffing professional who is not a subject
  matter expert in the underlying technology.
- Technically credible enough for the Freddie hiring manager.
- Do NOT make the Summary sound simplistic or generic merely to accommodate
  the MSP reader.
- Do NOT use promotional or generic recruiter language such as:
  "highly accomplished"
  "highly experienced"
  "highly skilled"
  "strong background"
  "excellent fit"
  "ideal candidate"
  "perfect candidate"
  "strong candidate"
  "directly aligns"
  "directly aligning"
  "aligns with"
  "aligns perfectly"
  "positions them uniquely"
  "exceptionally well prepared"
  "proven history"
- Do not say that the candidate "meets Freddie Mac's requirements" or
  "directly aligns with Freddie Mac's core requirements." Demonstrate the
  alignment through specific evidence instead.
- Do not use "Additionally", "Furthermore", or "Moreover".
- Do not simply restate the JD.
- When stating total professional experience, use the conservative format
  "X+ years" based on completed years as of the CURRENT DATE above.
- TOTAL EXPERIENCE VS. SKILL TENURE:
  A total-career figure such as "7+ years" describes the candidate's overall
  professional experience only. Never allow sentence grammar to convert that
  number into claimed tenure for named skills whose evidence supports fewer
  years.

FINAL SUMMARY TEST:

Before returning the Summary, ask:

1. Could an MSP recruiter who understands the requisition but is not deeply
   technical explain in plain English WHY this candidate should be shortlisted?

2. Can that reviewer quickly identify Freddie's most important supported
   Must Haves without searching through technical prose?

3. Does the Summary provide concrete evidence behind those claims rather than
   merely saying the candidate is qualified?

4. Would the Freddie hiring manager still see meaningful technical,
   operational, and domain-specific depth?

5. Is every candidate claim defensible against the resume?

6. Did I accidentally turn a Freddie requirement, manager comment, or target
   environment into a statement about the candidate's past experience?

7. Does the final sentence explain TRANSFERABILITY rather than fabricate
   previous performance of the target-role work?

If any answer is no, revise the Summary before returning it.

======================================================================
HIGH-PRIORITY REQUIREMENT EVIDENCE STANDARD
======================================================================

The importance of a Freddie requirement does NOT lower the evidence threshold.
It raises it.

When a requirement is a Must Have, manager priority, or specifically emphasized
in MSP/VNDLY notes:

- Require direct resume evidence before presenting the candidate as possessing
  that competency.
- Do NOT upgrade related, adjacent, administrative, coordination, support,
  oversight, or exposure-level experience into hands-on ownership merely
  because the requirement is important to Freddie.
- Do NOT merge two separate resume facts into a stronger composite claim unless
  the resume explicitly connects them.

Examples:

If the resume says:
"Ensured published documents were stored in SharePoint libraries"

you may accurately describe:
- SharePoint exposure
- use of SharePoint libraries
- supporting document availability in SharePoint

but you may NOT automatically claim:
- SharePoint content management
- building SharePoint sites/pages
- publishing content in SharePoint
- administering SharePoint
- hands-on SharePoint development

unless the resume provides that evidence.

Likewise, if one bullet says the candidate created job aids and another says
documents were stored in SharePoint, do NOT combine those facts into:
"created and maintained job aids in SharePoint"
unless the resume explicitly connects the job aids to SharePoint.

For high-priority requirements, distinguish carefully between:

1. DIRECT HANDS-ON EVIDENCE
   The resume explicitly shows the candidate personally performed the required
   work.

2. RELATED / PARTIAL EVIDENCE
   The resume shows adjacent, supporting, coordinating, oversight, or
   exposure-level experience but does not prove the full requirement.

3. NO EVIDENCE
   The resume does not support the requirement.

Only DIRECT HANDS-ON EVIDENCE may be stated as full possession of the
requirement in SUMMARY or SKILLS.

RELATED / PARTIAL EVIDENCE may be described accurately at its demonstrated
level, but must never be upgraded into the full Freddie requirement.

When evidence is partial, it is better to omit the requirement from the Skills
table than to overstate the candidate.

======================================================================
SKILLS TABLE - EXACTLY 4 HIGH-SIGNAL ROWS WHEN SUPPORTED
======================================================================

The Freddie template has FOUR Skills rows and TWO columns:

LEFT COLUMN:
SKILL AREA

RIGHT COLUMN:
SKILL DETAILS

PURPOSE:

This table is the MSP reviewer's FASTEST VISUAL CHECKLIST of the candidate's
match to the CURRENT Freddie requisition.

It should borrow the clean grouping philosophy used by the Fannie template:
one concise competency area on the left, with the most relevant supporting
tools, technologies, methodologies, or domain capabilities on the right.

There are NO years-of-experience or year-last-used values in the Freddie
Skills table.

----------------------------------------------------------------------
ROW FORMAT
----------------------------------------------------------------------

For every populated row return:

SKILL#:
A concise functional competency / skill area.

SKILLDETAILS#:
A comma-separated list of approximately 2-6 highly relevant supporting items.

Examples of the desired structure:

SKILL1:
"Core Java & J2EE Development"

SKILLDETAILS1:
"Java, Spring Framework, J2EE, Servlets, JSP, JDBC"

SKILL2:
"Relational Database Development"

SKILLDETAILS2:
"Oracle, DB2, SQL Server"

SKILL3:
"Java Unit Testing"

SKILLDETAILS3:
"JUnit, Mockito"

SKILL4:
"Financial Services Application Development"

SKILLDETAILS4:
"Investment Data, Portfolio Data, Transaction Processing"

Other role-dependent examples:

"Programming & Backend"
"Python, Django, Flask, FastAPI"

"Cloud Development"
"AWS Lambda, S3, RDS, Secrets Manager"

"DevOps & CI/CD"
"Jenkins, Docker, Kubernetes, Helm"

"Data & Database Systems"
"PostgreSQL, MySQL, Oracle, AuroraDB"

"SAFe Agile Framework"
"PI Planning, Scrum of Scrums, ART Leadership Sync"

"Agile Project Management"
"Jira, Confluence, Azure DevOps"

"Stakeholder & Executive Reporting"
"Tableau, Power BI, Executive Dashboards"

These examples illustrate FORMAT and grouping philosophy only.
Do NOT force any example terminology onto a candidate or requisition.

----------------------------------------------------------------------
HOW TO CHOOSE THE FOUR ROWS
----------------------------------------------------------------------

Use the full Freddie requisition intelligence already supplied to you.

Priority order:

1. Latest substantive hiring-manager selection feedback / correction.
2. Current manager priorities and substantive Spotlight guidance.
3. Current substantive MSP / VNDLY clarification that affects what candidates
   must demonstrate.
4. Formal Must Have requirements from the Official JD.
5. Structured VNDLY Must Have skills when populated.
6. Core recurring responsibilities / required domain expertise.
7. Formal Preferred / Nice To Have requirements.
8. Structured VNDLY Nice To Have skills when populated.

The final four rows should represent the candidate's FOUR strongest,
most useful shortlist signals for THIS requisition.

FREDDIE MSP COVERAGE RULE:
- Do not automatically delete a supported JD skill merely because the manager
  says the team does not use it or it is not central.
- If that skill remains useful for MSP/JD keyword recognition, keep it as a
  secondary detail inside a broader row when doing so costs little and does not
  displace a higher-priority manager requirement.
- Example: if User Stories is in the JD, the candidate genuinely has User Story
  experience, and the manager says "we don't do user stories", User Stories may
  still appear as a detail under Business Requirements/Product Analysis while
  the row itself emphasizes the requirements/process work the manager actually
  needs.
- Only a true `explicitly_excluded_requirements` item should be automatically
  omitted.

Do NOT mechanically convert Freddie's four planned competency groups into
four resume rows. Those groups are prioritization guidance. Build the final
rows around the candidate's actual evidence and the requirements that matter
most to the current requisition.

----------------------------------------------------------------------
MSP-SCAN STRATEGY
----------------------------------------------------------------------

The left column should make the requirement cluster recognizable immediately.

Prefer:
"Core Java & J2EE Development"

over:
"Programming"

when Java/J2EE is central to the role.

Prefer:
"Mortgage Data & Analytics"

over:
"Domain Expertise"

when mortgage experience is specifically important.

Prefer:
"ServiceNow IRM/GRC Testing"

over:
"Quality Assurance"

when ServiceNow IRM/GRC is the actual screening requirement.

The right column then provides the exact supporting terms that make the match
easy to see.

The reviewer should NOT have to dig through the work history to discover the
candidate's most important supported Freddie keywords.

----------------------------------------------------------------------
CANDIDATE-EVIDENCE RULE - ABSOLUTE
----------------------------------------------------------------------

Every SKILL AREA and every SKILL DETAILS item must be genuinely supported by
the candidate's resume or official candidate-supplied Freddie vetting answers.

The JD, Spotlight transcript, MSP notes, Internal Notes, structured Must/Nice
fields, and requirement analysis tell you WHAT TO PRIORITIZE.

They do NOT create candidate experience.

Rules:

- Never copy a Freddie technology, methodology, responsibility, business
  process, or domain term into the candidate's Skills table unless candidate
  evidence supports it.
- Never infer a specific technology from a broader one.
- Never convert exposure/coordination/oversight into hands-on expertise.
- Do not combine unrelated resume facts into a stronger candidate claim.
- Do not use a vague broad category to disguise a missing central Must Have.
- If the candidate lacks one Freddie requirement, do not fabricate it merely to
  make the table look complete.
- Move to the next highest-value relevant competency the candidate genuinely
  possesses.

If fewer than four defensible, role-relevant competency groups exist, leave
the unused SKILL# and SKILLDETAILS# values blank. The Freddie Word processor
will remove the unused row.

----------------------------------------------------------------------
GROUPING RULES
----------------------------------------------------------------------

Each row should represent ONE coherent competency family.

Good:
"Java Unit Testing" | "JUnit, Mockito"

Good:
"Relational Database Development" | "Oracle, DB2, SQL Server"

Good:
"Change Communications" | "Business Communications, Job Aids, SharePoint Content"

Avoid:
"Development & Testing & Cloud & Databases" | "Java, JUnit, AWS, Oracle"

because it mixes unrelated screening dimensions.

Avoid unnecessary duplication between rows.

For example, if:
"Core Java & J2EE Development" already contains "Spring Framework"

do not create a separate:
"Spring Development" | "Spring Boot"

unless Spring itself is independently important enough to deserve one of only
four rows and doing so would not crowd out a more important supported Freddie
requirement.

----------------------------------------------------------------------
SKILL AREA RULES
----------------------------------------------------------------------

- Usually 2-6 words.
- Clear, professional, conventional terminology.
- Specific enough that the MSP can understand why the row matters.
- Do not write sentences.
- Do not include years.
- Do not include "current".
- Do not put a long tool inventory in the left column.
- Do not simply repeat the VNDLY/rate-card job title.

----------------------------------------------------------------------
SKILL DETAILS RULES
----------------------------------------------------------------------

- Maximum 6 items. Use fewer when fewer than 6 materially strengthen the row.
- Comma-separated plain text.
- Prefer exact Freddie/JD terminology when the candidate truly supports it.
- Include technologies, platforms, methodologies, functional capabilities, or
  domain concepts that substantively strengthen the SKILL AREA.
- Prioritize the most important requirement-specific items rather than the
  candidate's entire technology inventory.
- Do not add weak filler merely to reach a certain number of items.
- Do not include years, proficiency ratings, or "current".
- Do not repeat the exact same detail across multiple rows unless unavoidable.

----------------------------------------------------------------------
FINAL MSP / HIRING-MANAGER CHECK
----------------------------------------------------------------------

Review the four rows together and ask:

1. If an MSP recruiter looked ONLY at this table and Freddie's current
   requirements, would the candidate's strongest genuine matches be obvious?

2. Are Freddie's most important supported Must Haves visible using recognizable
   terminology?

3. Does each row represent a materially different screening dimension?

4. Did a broad secondary strength consume a row that should instead show a more
   important supported Must Have?

5. Would the hiring manager still consider every term technically and
   factually defensible against the resume?

6. Is the table easy to scan without looking cluttered?

If not, revise the four rows before returning them.

======================================================================
OFFICIAL FREDDIE VETTING ANSWERS
======================================================================

Freddie required that these responses come directly from the candidate.

They are SECONDARY supporting evidence.

Allowed:
- Clarify technical depth.
- Clarify how something already reflected in the work history was used.
- Clarify scope, approach, or operational context.

Not allowed:
- Use an answer to invent an employer/date.
- Calculate years solely from a candidate's unsupported statement.
- Convert a theoretical answer into performed experience.
- Override contradictory resume evidence.

OFFICIAL CANDIDATE-SUPPLIED VETTING Q&A:
{vetting_for_ai}

CANDIDATE-CONFIRMED CLARIFICATIONS / FOLLOW-UP ANSWERS:
{Candidate_Clarifications or "NONE PROVIDED"}

Candidate-confirmed clarifications are SECONDARY candidate evidence. They may
add factual details the resume omitted, but may not contradict dated resume facts.

======================================================================
CANDIDATE TRUTH SOURCES
======================================================================

The original resume is PRIMARY. Candidate-confirmed clarifications and official
Freddie vetting responses are SECONDARY factual sources. Role/JD/manager text is
never candidate evidence.

ZERO-TOLERANCE RULES:

1. Never assign a JD technology to the candidate merely because Freddie asks
   for it.

2. Never infer a specific tool from a broader category.
   Example:
   Resume says AWS -> you may NOT invent Redshift.

3. Never upgrade exposure into ownership.

4. Never fabricate a metric, scale, accomplishment, mortgage domain,
   certification, platform, methodology, or responsibility.

5. Technologies that appear only in the JD/Spotlight/MSP notes describe the
   ROLE, not the candidate.

6. Every candidate claim in SUMMARY and SKILLS must survive a side-by-side
   audit against the resume and any supplied candidate-confirmed clarifications,
   with official vetting responses used only as factual secondary support.

======================================================================
EXPERIENCE PRIORITY / OPERATIONAL OWNERSHIP
======================================================================

Prefer:
- Things the candidate actually built, supported, maintained, tested,
  troubleshot, validated, deployed, operated, analyzed, or owned.
- Recurring direct responsibility.
- Recent relevant work.
- Production support and issue resolution when relevant.
- Operational ownership.
- Compliance/audit/validation responsibility when relevant.
- Evidence covering multiple Freddie requirements at once.

Avoid overweighting:
- One-time transformation projects when the target role is operational.
- Technologies listed only in a skills inventory.
- Generic management language.
- Impressive but irrelevant projects.

======================================================================
SAFE WORK-HISTORY BULLET PRIORITIZATION
======================================================================

You are NOT allowed to rewrite the work-history bullets.

Instead, for each role, rank the EXISTING ORIGINAL BULLETS by relevance to
this Freddie requisition.

Return zero-based bullet indexes.

Example:
If a role has 4 bullets and bullets 2, 0, 3, 1 are the best order:
[2, 0, 3, 1]

Rules:
- Do NOT omit bullets.
- Do NOT create bullets.
- Do NOT merge bullets.
- Do NOT rewrite bullets.
- Rank only the existing indexes.
- Current hiring-manager priorities outrank merely de-emphasized JD artifacts.
  A truthful de-emphasized item may remain visible for MSP/JD coverage, but do
  not move that bullet ahead of stronger bullets proving current manager priorities.
- If the original order is already best, return the original indexes.
- If a role has no bullets, return [].

STRUCTURED ORIGINAL EXPERIENCE:
{structured_exp}

======================================================================
OUTPUT FORMAT
======================================================================

Return ONLY:

{{
  "CANDIDATE_TITLE": "",
  "SUMMARY": "",
  "SKILL1": "",
  "SKILLDETAILS1": "",
  "SKILL2": "",
  "SKILLDETAILS2": "",
  "SKILL3": "",
  "SKILLDETAILS3": "",
  "SKILL4": "",
  "SKILLDETAILS4": "",
  "BULLET_ORDER": {{
    "1": [],
    "2": [],
    "3": [],
    "4": [],
    "5": [],
    "6": [],
    "7": []
  }}
}}

======================================================================
FULL ORIGINAL RESUME
======================================================================

{raw_text}
"""

                draft_summary_data = fred_generate_json(
                    API_KEY,
                    tailoring_prompt,
                )

                # ====================================================
                # PASS 2B - INDEPENDENT FREDDIE GROUNDING / PRIORITY AUDIT
                # ====================================================
                audited_summary_data = fred_generate_json(
                    API_KEY,
                    fred_grounding_audit_prompt(
                        resume_text=raw_text,
                        structured_req=req_analysis,
                        job_description=Job_Description,
                        msp_notes=MSP_Notes,
                        spotlight_transcript=Call_Transcript,
                        internal_notes=Internal_Notes,
                        draft_data=draft_summary_data,
                        calculated_total_experience=calculated_total_experience,
                        current_date=current_date,
                        candidate_clarifications=Candidate_Clarifications,
                    ),
                )

                freddie_grounding_errors = fred_validate_grounded_tailoring(
                    raw_text,
                    req_analysis,
                    Job_Description,
                    MSP_Notes,
                    Call_Transcript,
                    audited_summary_data,
                    candidate_clarifications=Candidate_Clarifications,
                )

                if freddie_grounding_errors:
                    audited_summary_data = fred_generate_json(
                        API_KEY,
                        fred_grounding_audit_prompt(
                            resume_text=raw_text,
                            structured_req=req_analysis,
                            job_description=Job_Description,
                            msp_notes=MSP_Notes,
                            spotlight_transcript=Call_Transcript,
                            internal_notes=Internal_Notes,
                            draft_data=audited_summary_data,
                            calculated_total_experience=calculated_total_experience,
                            current_date=current_date,
                            validator_errors=freddie_grounding_errors,
                            candidate_clarifications=Candidate_Clarifications,
                        ),
                    )

                    freddie_grounding_errors = fred_validate_grounded_tailoring(
                        raw_text,
                        req_analysis,
                        Job_Description,
                        MSP_Notes,
                        Call_Transcript,
                        audited_summary_data,
                        candidate_clarifications=Candidate_Clarifications,
                    )

                if freddie_grounding_errors:
                    hard_grounding_errors = [
                        error for error in freddie_grounding_errors
                        if fred_grounding_error_is_hard(error)
                    ]

                    if hard_grounding_errors:
                        audited_summary_data = fred_generate_json(
                            API_KEY,
                            fred_grounding_audit_prompt(
                                resume_text=raw_text,
                                structured_req=req_analysis,
                                job_description=Job_Description,
                                msp_notes=MSP_Notes,
                                spotlight_transcript=Call_Transcript,
                                internal_notes=Internal_Notes,
                                draft_data=audited_summary_data,
                                calculated_total_experience=calculated_total_experience,
                                current_date=current_date,
                                validator_errors=hard_grounding_errors,
                                candidate_clarifications=Candidate_Clarifications,
                            ),
                        )
                        final_errors = fred_validate_grounded_tailoring(
                            raw_text,
                            req_analysis,
                            Job_Description,
                            MSP_Notes,
                            Call_Transcript,
                            audited_summary_data,
                            candidate_clarifications=Candidate_Clarifications,
                        )
                        remaining_hard = [
                            error for error in final_errors
                            if fred_grounding_error_is_hard(error)
                        ]
                        if remaining_hard:
                            audited_summary_data = fred_drop_hard_grounding_items(
                                audited_summary_data,
                                remaining_hard,
                            )
                            print(
                                "Freddie hard grounding items removed after repair: "
                                + " | ".join(remaining_hard[:8])
                            )
                    else:
                        # Citation-shape, lexical-similarity, and generic competency
                        # diagnostics are useful repair guidance, but they must not
                        # destroy coherent recruiter-quality prose or block the run.
                        print(
                            "Freddie soft grounding diagnostics accepted after repair: "
                            + " | ".join(freddie_grounding_errors[:8])
                        )

                summary_data = fred_finalize_grounded_tailoring(
                    audited_summary_data
                )

                # ====================================================
                # VALIDATE CANDIDATE POSITIONING TITLE / SUMMARY / SKILLS
                # ====================================================

                raw_candidate_positioning_title = str(
                    summary_data.get(
                        "CANDIDATE_TITLE",
                        "",
                    )
                ).strip()

                (
                    candidate_positioning_title,
                    candidate_positioning_specialty,
                ) = fred_normalize_candidate_positioning_title(
                    raw_candidate_positioning_title
                )

                final_summary = str(
                    summary_data.get(
                        "SUMMARY",
                        "",
                    )
                ).strip()

                final_summary = fred_repair_summary_identity_phrase(
                    final_summary,
                    name,
                    raw_candidate_positioning_title,
                    candidate_positioning_title,
                    candidate_positioning_specialty,
                )

                final_summary = fred_clean_summary(
                    final_summary,
                    name,
                )
                final_summary = fred_suppress_total_career_years(
                    final_summary,
                    calculated_total_experience,
                    selected_vndly_context,
                )

                if not final_summary:
                    raise RuntimeError(
                        "The Freddie optimization step returned no "
                        "candidate Summary."
                    )

                # ====================================================
                # APPLY SAFE BULLET REORDERING
                # ====================================================

                bullet_orders = summary_data.get(
                    "BULLET_ORDER",
                    {},
                )

                if not isinstance(
                    bullet_orders,
                    dict,
                ):
                    bullet_orders = {}

                for idx, role in enumerate(
                    experience,
                    start=1,
                ):
                    original_bullets = role.get(
                        "Bullets",
                        [],
                    )

                    requested_order = bullet_orders.get(
                        str(idx),
                        [],
                    )

                    role["Bullets"] = (
                        fred_safe_reorder_bullets(
                            original_bullets,
                            requested_order,
                        )
                    )

                # ====================================================
                # FREDDIE SKILLS TABLE
                # ====================================================
                # Freddie does not require skill-by-skill years or year-last-used.
                # The template now uses two columns:
                #   SKILL#        = concise competency / skill area
                #   SKILLDETAILS# = supporting tools, technologies, methodologies,
                #                   or domain capabilities
                #
                # Keeping these separate avoids artificial tenure calculations
                # and makes the first-page skills section easier for MSP reviewers
                # to scan while preserving hiring-manager credibility.

                # ====================================================
                # FREDDIE-ONLY SKILL DETAILS HARD CAP
                # ====================================================
                # The prompt asks for no more than six details per row. Enforce
                # that deterministically so long AI outputs cannot clutter the
                # two-column Skills table.
                for skill_num in range(1, 5):
                    details_key = f"SKILLDETAILS{skill_num}"
                    summary_data[details_key] = fred_cap_skill_details(
                        summary_data.get(details_key, ""),
                        max_items=6,
                    )

                # ====================================================
                # BUILD VNDLY SUBMISSION SUMMARY + RECOMMENDED SKILLS
                # ====================================================

                vndly_package = {
                    "candidate_positioning_title": candidate_positioning_title,
                    "vndly_summary": "",
                    "final_recommended_vndly_skills": [],
                    "exact_must_have_skills": [],
                    "exact_nice_to_have_skills": [],
                    "required_vndly_skills_from_job_intelligence": [],
                    "manager_priority_vndly_skills": [],
                    "preferred_vndly_skills_from_job_intelligence": [],
                    "jd_coverage_vndly_skills": [],
                    "additional_vndly_skills": [],
                }
                vndly_package_error = None

                try:
                    with st.spinner(
                        "Building VNDLY comments and exact skill recommendations..."
                    ):
                        vndly_package = fred_build_vndly_candidate_package(
                            API_KEY,
                            name,
                            raw_text,
                            req_analysis,
                            selected_vndly_context,
                            vetting_for_ai,
                            final_summary,
                            calculated_total_experience,
                            candidate_positioning_title,
                            candidate_submission_facts={
                                "current_location": Current_Location,
                                "available_onsite": Available_Onsite,
                                "onsite_locations": Site,
                                "former_freddie_cw_or_employee": Former_CW,
                                "rehire_eligibility_check_completed": Eligibility,
                                "interview_availability": Interview,
                            },
                            candidate_clarifications=Candidate_Clarifications,
                        )
                except Exception as package_error:
                    vndly_package_error = str(package_error)

                # ====================================================
                # BUILD WORD MAPPING
                # ====================================================

                mapping = {
                    "FullName": name,
                    "Location": Current_Location,
                    "Remote": Available_Onsite,
                    "FormerFM": Former_CW,
                    "FormerCW": Former_CW,
                    "Eligibility": Eligibility,
                    "Authorization": Auth_US,
                    "Sponsorship": Sponsorship,
                    "Site": Site,
                    "Interview": Interview,
                    "Certifications": certifications,
                    "SUMMARY": final_summary,
                    "SKILL1": str(
                        summary_data.get(
                            "SKILL1",
                            "",
                        )
                    ).strip(),
                    "SKILLDETAILS1": str(
                        summary_data.get(
                            "SKILLDETAILS1",
                            "",
                        )
                    ).strip(),
                    "SKILL2": str(
                        summary_data.get(
                            "SKILL2",
                            "",
                        )
                    ).strip(),
                    "SKILLDETAILS2": str(
                        summary_data.get(
                            "SKILLDETAILS2",
                            "",
                        )
                    ).strip(),
                    "SKILL3": str(
                        summary_data.get(
                            "SKILL3",
                            "",
                        )
                    ).strip(),
                    "SKILLDETAILS3": str(
                        summary_data.get(
                            "SKILLDETAILS3",
                            "",
                        )
                    ).strip(),
                    "SKILL4": str(
                        summary_data.get(
                            "SKILL4",
                            "",
                        )
                    ).strip(),
                    "SKILLDETAILS4": str(
                        summary_data.get(
                            "SKILLDETAILS4",
                            "",
                        )
                    ).strip(),
                }

                # Education
                for i in range(1, 4):
                    if i <= len(education):
                        edu_item = education[i - 1]

                        if not isinstance(
                            edu_item,
                            dict,
                        ):
                            edu_item = {}

                        mapping[f"School{i}"] = clean_school(
                            str(
                                edu_item.get(
                                    "School",
                                    "",
                                )
                            )
                        )

                        mapping[f"Degree{i}"] = str(
                            edu_item.get(
                                "Degree",
                                "",
                            )
                        ).strip()

                    else:
                        mapping[f"School{i}"] = ""
                        mapping[f"Degree{i}"] = ""

                # Work History
                for i in range(1, max(7, len(experience)) + 1):
                    if i <= len(experience):
                        role = experience[i - 1]

                        mapping[f"Company{i}"] = (
                            role.get(
                                "Company",
                                "",
                            )
                        )

                        mapping[f"Title{i}"] = (
                            role.get(
                                "Title",
                                "",
                            )
                        )

                        mapping[f"Bullets{i}"] = clean_bullets(
                            role.get(
                                "Bullets",
                                [],
                            )
                        )

                        mapping[f"Environment{i}"] = (
                            role.get(
                                "Environment",
                                "",
                            )
                        )

                        mapping[f"Dates{i}"] = (
                            role.get(
                                "Dates",
                                "",
                            )
                        )

                    else:
                        mapping[f"Company{i}"] = ""
                        mapping[f"Title{i}"] = ""
                        mapping[f"Bullets{i}"] = ""
                        mapping[f"Environment{i}"] = ""
                        mapping[f"Dates{i}"] = ""

                # ====================================================
                # GENERATE WORD DOCUMENT
                # ====================================================

                safe_name = (
                    re.sub(
                        r"[^A-Za-z0-9_\-]+",
                        "_",
                        name,
                    ).strip("_")
                    or "Candidate"
                )

                out_file = (
                    f"Submission_FreddieMac_{safe_name}.docx"
                )

                process_freddie_word_doc(
                    TEMPLATE_FILENAME,
                    mapping,
                    vetting_pairs,
                    out_file,
                )

                # ====================================================
                # VNDLY SUBMISSION PACKAGE - READY TO COPY
                # ====================================================

                st.subheader("📨 VNDLY Submission Details")

                if vndly_package_error:
                    st.warning(
                        "⚠️ The resume was generated successfully, but the "
                        "VNDLY comments/skill package could not be created. "
                        f"Package error: {vndly_package_error}"
                    )

                st.markdown("**VNDLY Comments**")
                if vndly_package.get("vndly_summary", ""):
                    st.code(
                        vndly_package.get("vndly_summary", ""),
                        language=None,
                    )
                else:
                    st.caption("No VNDLY comments were generated.")

                ranked_vndly_skills = vndly_package.get(
                    "final_recommended_vndly_skills",
                    [],
                ) or []

                st.markdown("**Recommended VNDLY Skills**")
                st.caption(
                    "Ranked from highest to lowest submission value. Select every skill below "
                    "that is available in VNDLY. There is no artificial eight-skill cap; the "
                    "tool stops when additional skills would add noise rather than shortlist value."
                )

                source_labels = {
                    "exact_structured_must": "Exact Structured Must",
                    "exact_structured_nice": "Exact Structured Nice",
                    "required_job_intelligence": "Required",
                    "manager_priority": "Manager Priority",
                    "preferred_job_intelligence": "Preferred",
                    "additional_high_value": "Additional",
                }

                if ranked_vndly_skills:
                    for index, item in enumerate(
                        ranked_vndly_skills,
                        start=1,
                    ):
                        skill_name = str(
                            item.get("skill_name", "") or ""
                        ).strip()
                        category = str(
                            item.get("source_category", "") or ""
                        ).strip().lower()
                        label = source_labels.get(
                            category,
                            "Recommended",
                        )
                        st.markdown(
                            f"{index}. `{skill_name}` - {label}"
                        )
                else:
                    st.caption(
                        "No VNDLY catalogue skills were supported strongly enough "
                        "to recommend."
                    )

                # ====================================================
                # AUTOMATIC FREDDIE SUBMISSION EMAIL
                # ====================================================

                if not vndly_package_error and vndly_package.get(
                    "vndly_summary",
                    "",
                ):
                    email_key = fred_generation_email_key(
                        name,
                        selected_vndly_context.get("job_id", ""),
                        raw_text,
                        vndly_package.get("vndly_summary", ""),
                    )

                    sent_keys = st.session_state.setdefault(
                        "fred_sent_email_keys",
                        set(),
                    )

                    if email_key not in sent_keys:
                        try:
                            fred_send_submission_email(
                                name,
                                selected_vndly_context,
                                vndly_package,
                            )
                            sent_keys.add(email_key)
                            st.success(
                                "📧 VNDLY submission details were emailed automatically."
                            )
                        except Exception as email_error:
                            st.warning(
                                "⚠️ The resume was generated successfully, but the "
                                "automatic Freddie submission email could not be sent. "
                                f"Email error: {email_error}"
                            )
                    else:
                        st.info(
                            "📧 This exact Freddie candidate/job submission was already "
                            "emailed during the current app session, so a duplicate "
                            "email was not sent."
                        )
                else:
                    st.warning(
                        "⚠️ Automatic email was skipped because the VNDLY "
                        "submission package was not available."
                    )

                with open(
                    out_file,
                    "rb",
                ) as file:
                    st.download_button(
                        label="⬇️ Download Generated Document",
                        data=file,
                        file_name=out_file,
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.wordprocessingml.document"
                        ),
                        type="primary",
                    )

                st.success(
                    "✅ Freddie Mac submission is ready."
                )

            except Exception as e:
                st.error(
                    f"❌ Freddie Mac processing failed: {str(e)}"
                )

            finally:
                if (
                    resume_path
                    and os.path.exists(resume_path)
                ):
                    try:
                        os.remove(resume_path)
                    except Exception:
                        pass


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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="per_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="per_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="per_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="per_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="per_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="per_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="per_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="per_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="per_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="per_a5")

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
                        - For 'Dates', preserve the date precision actually shown on the original resume. If month + year are shown, format as "MM/YYYY to MM/YYYY" (or "MM/YYYY to Present"). If ONLY years are shown, preserve them as "YYYY to YYYY" (or "YYYY to Present"). NEVER invent January, "01", or any other month when the source provides only a year. Mixed-precision dates must remain mixed rather than filling in a missing month.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="co_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="co_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="co_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="co_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="co_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="co_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="co_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="co_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="co_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="co_a5")

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
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="ad_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="ad_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="ad_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="ad_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="ad_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="ad_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="ad_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="ad_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="ad_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="ad_a5")

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
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="cb_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="cb_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="cb_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="cb_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="cb_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="cb_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="cb_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="cb_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="cb_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="cb_a5")

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
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="bn_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="bn_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="bn_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="bn_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="bn_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="bn_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="bn_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="bn_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="bn_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="bn_a5")

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
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
    
    # Row 1
    row1_c1, row1_c2 = st.columns(2)
    Question_1 = row1_c1.text_input("Question 1", key="dg_q1")
    Answer_1 = row1_c2.text_area("Answer 1", height=68, key="dg_a1")

    # Row 2
    row2_c1, row2_c2 = st.columns(2)
    Question_2 = row2_c1.text_input("Question 2", key="dg_q2")
    Answer_2 = row2_c2.text_area("Answer 2", height=68, key="dg_a2")

    # Row 3
    row3_c1, row3_c2 = st.columns(2)
    Question_3 = row3_c1.text_input("Question 3", key="dg_q3")
    Answer_3 = row3_c2.text_area("Answer 3", height=68, key="dg_a3")

    # Row 4
    row4_c1, row4_c2 = st.columns(2)
    Question_4 = row4_c1.text_input("Question 4", key="dg_q4")
    Answer_4 = row4_c2.text_area("Answer 4", height=68, key="dg_a4")

    # Row 5
    row5_c1, row5_c2 = st.columns(2)
    Question_5 = row5_c1.text_input("Question 5", key="dg_q5")
    Answer_5 = row5_c2.text_area("Answer 5", height=68, key="dg_a5")

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
                        - Return "Pursuing" if the resume indicates ongoing study, contains the word "Pursuing", "In-progress", or lists a graduation date in the future (relative to October 2026).
                    4. Experience: Company, Title, Bullets (LIST), Environment (String, optional), Dates.
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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
                    for i in range(1, max(7, len(exp)) + 1):
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
def process_deloitte_doc(template_path, mapping, resume_path, output_path):
    import docx
    import re
    from docx.shared import Pt, Inches
    from docx.enum.text import WD_TAB_ALIGNMENT
    from copy import deepcopy

    doc = docx.Document(template_path)
    orig_doc = docx.Document(resume_path)

    # 1. Delete empty skill rows BEFORE replacing text
    for tbl in doc.tables:
        rows_to_delete = []
        for row in tbl.rows:
            cell_text = row.cells[0].text
            for i in range(1, 9):
                if f"{{{{SKILL{i}}}}}" in cell_text:
                    if not mapping.get(f"SKILL{i}", "").strip():
                        rows_to_delete.append(row)
        for row in rows_to_delete:
            # Safer XML removal method for table rows
            row._element.getparent().remove(row._element)

    # 2. Replace placeholders while preserving formatting
    for p in doc.paragraphs:
        # Explicitly rebuild the header lines to maintain the bold prefix
        if "Candidate Legal Name:" in p.text and "{{FullName}}" in p.text:
            p.text = ""
            r1 = p.add_run("Candidate Legal Name: ")
            r1.bold = True
            p.add_run(str(mapping.get("FullName", "")))
            continue
        if "Preferred Name:" in p.text and "{{PreferredName}}" in p.text:
            p.text = ""
            r1 = p.add_run("Preferred Name: ")
            r1.bold = True
            p.add_run(str(mapping.get("PreferredName", "")))
            continue
        if "Current Location:" in p.text and "{{Location}}" in p.text:
            p.text = ""
            r1 = p.add_run("Current Location: ")
            r1.bold = True
            p.add_run(str(mapping.get("Location", "")))
            continue
        if "Upcoming Scheduled Time off:" in p.text and "{{TimeOff}}" in p.text:
            p.text = ""
            r1 = p.add_run("Upcoming Scheduled Time off: ")
            r1.bold = True
            p.add_run(str(mapping.get("TimeOff", "")))
            continue

        # For all other paragraphs (like Summary points)
        for k, v in mapping.items():
            placeholder = f"{{{{{k}}}}}"
            if placeholder in p.text:
                replaced = False
                for run in p.runs:
                    if placeholder in run.text:
                        run.text = run.text.replace(placeholder, str(v))
                        replaced = True
                # Fallback if the placeholder got split across multiple runs by Word
                if not replaced:
                    p.text = p.text.replace(placeholder, str(v))

    # Replace placeholders in table cells
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for k, v in mapping.items():
                        placeholder = f"{{{{{k}}}}}"
                        if placeholder in p.text:
                            replaced = False
                            for run in p.runs:
                                if placeholder in run.text:
                                    run.text = run.text.replace(placeholder, str(v))
                                    replaced = True
                            if not replaced:
                                p.text = p.text.replace(placeholder, str(v))

    # --- XML SCRUBBER FUNCTION ---
    def clean_xml_element(element):
        for hl in element.xpath('.//w:hyperlink'):
            parent = hl.getparent()
            if parent is not None:
                idx = parent.index(hl)
                for child in list(hl):
                    parent.insert(idx, child)
                    idx += 1
                parent.remove(hl)
        for drawing in element.xpath('.//w:drawing | .//w:pict | .//w:object'):
            parent = drawing.getparent()
            if parent is not None:
                parent.remove(drawing)
        for bm in element.xpath('.//w:bookmarkStart | .//w:bookmarkEnd'):
            parent = bm.getparent()
            if parent is not None:
                parent.remove(bm)
        for sectPr in element.xpath('.//w:sectPr'):
            parent = sectPr.getparent()
            if parent is not None:
                parent.remove(sectPr)
        return element
    # ----------------------------------

    # 3. Extract the "Rest of Resume"
    body_elements = orig_doc._body._body
    start_copying = False
    elements_to_copy = []

    for element in body_elements:
        try:
            text = "".join(element.itertext())
        except Exception:
            text = ""
            
        if "Links:" in text or "Technical Skills" in text:
            start_copying = True
        
        if start_copying:
            clean_el = clean_xml_element(deepcopy(element))
            elements_to_copy.append(clean_el)

    # 4. Insert extracted elements perfectly at {{RESUME_BODY}}
    for p in doc.paragraphs:
        if "{{RESUME_BODY}}" in p.text:
            parent = p._element.getparent()
            index = parent.index(p._element)
            for el in elements_to_copy:
                parent.insert(index, el)
                index += 1
            parent.remove(p._element)
            break

    # 5. STRICT FORMATTING & TAB STOP FIX
    def apply_strict_formatting(paragraph):
        # A. Set Tab Stop to 7.5" Right-Aligned and remove right indent margins
        try:
            paragraph.paragraph_format.tab_stops.clear_all()
            paragraph.paragraph_format.tab_stops.add_tab_stop(Inches(7.5), WD_TAB_ALIGNMENT.RIGHT)
            paragraph.paragraph_format.right_indent = Inches(0)
        except Exception:
            pass

        # B. Clean up runs: Set Arial 10pt and collapse multiple tabs/spaces
        for run in paragraph.runs:
            run.font.name = 'Arial'
            run.font.size = Pt(10)
            if run.text:
                # Convert 3 or more spaces to a single tab
                run.text = re.sub(r' {3,}', '\t', run.text)
                # Collapse multiple tabs into a single tab
                run.text = re.sub(r'\t{2,}', '\t', run.text)
                # Remove spaces that are directly next to tabs
                run.text = re.sub(r' \t|\t ', '\t', run.text)

        # C. Cross-Run Cleanup: Catch tabs that are split across formatting boundaries
        for i in range(len(paragraph.runs) - 1):
            if paragraph.runs[i].text and paragraph.runs[i].text.endswith('\t'):
                j = i + 1
                while j < len(paragraph.runs) and paragraph.runs[j].text.startswith('\t'):
                    paragraph.runs[j].text = paragraph.runs[j].text.lstrip('\t')
                    j += 1

    # Apply the formatting sweep across the entire generated document
    for p in doc.paragraphs:
        apply_strict_formatting(p)

    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    apply_strict_formatting(p)

    # Collapse extra blank body paragraphs left by removed/empty sections.
    collapse_extra_blank_paragraphs(doc)
    
    doc.save(output_path)
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

                    # Deloitte later reopens the source resume with python-docx
                    # so it can copy the original resume content/XML into the
                    # Deloitte template. PDFs therefore need a temporary DOCX
                    # version for that step.
                    resume_docx_path = resume_path

                    if os.path.splitext(resume_path)[1].lower() == ".pdf":
                        resume_docx_path = (
                            os.path.splitext(resume_path)[0] +
                            "_converted.docx"
                        )

                        pdf_converter = Converter(resume_path)

                        try:
                            pdf_converter.convert(
                                resume_docx_path,
                                start=0,
                                end=None
                            )
                        finally:
                            pdf_converter.close()

                    # Keep extracting text from the ORIGINAL uploaded file.
                    # extract_text() already handles PDFs correctly.
                    raw_text = extract_text(resume_path)
                    client = genai.Client(api_key=API_KEY)

                    prompt = f"""
                    Return a valid JSON object ONLY.

                    RULES:
                    1. Name: Pull from top/header (Title Case). If the name is completely missing or unreadable in the text, extract it from the provided FILENAME.
                    2. Company Name Cleaning: Extract ONLY the primary end-client name. You MUST physically strip out any geographic locations (e.g., ", DELAWARE", ", MD") and strip out any contracting agencies/vendors.
                    3. Education: Extract School and Degree.
                    4. Experience: Company, Title, Bullets (LIST), Environment (String, optional), Dates.
                        - For 'Dates', preserve the date precision from the original resume. If a role lists only years, return only those years (for example, "2018 - 2020"). NEVER invent January, "01", or any other month. If month + year are present, preserve them; mixed-precision ranges must stay mixed.
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

                    # --- UPDATED MAPPING TO SUPPORT 8 SKILLS + COMPANIES + 5 SUMMARIES ---
                    mapping = {
                        "FullName": final_name, "PreferredName": Preferred_Name, 
                        "TimeOff": Time_Off, "Location": Current_Location_City_ST,
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
                    for i in range(1, max(7, len(exp)) + 1):
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
                            - SUMMARY1 (The Anchor): Use their FIRST NAME. State their total progressive professional experience, their core title, and dominant expertise solving the PRIMARY technical "must-have" of the Job Description. (Format years strictly rounded DOWN to the nearest whole year as 'X+ years').
                            - SUMMARY2 (The Alignment): NO PRONOUNS. Start directly with an action verb or adverb (e.g., "Recently architected..."). Explicitly name their CURRENT or most recent employer and frame their most relevant project as a 1:1 match for the manager's challenge.
                            - SUMMARY3 (The Execution): NO PRONOUNS. Start directly with an action verb (e.g., "Engineered..."). Weave the specific tools/methodologies into an "Execution Statement" highlighting complexity, scale, or business impact.
                            - SUMMARY4 (Secondary Expertise): NO PRONOUNS. Start directly with a verb (e.g., "Possesses secondary expertise in..."). Highlight a secondary skill, architecture, or methodology requested in the JD that they possess.
                            - SUMMARY5 (The Closer): NO PRONOUNS. Start directly with a verb (e.g., "Will deliver immediate value by..."). State what specific value they will deliver on Day 1 based on their past execution track record.

                            CRITICAL ATS HACK: Across these 5 bullets, you MUST seamlessly embed exact phrases from the Job Description to maximize the Fieldglass match score.

                            ========================
                            STYLE & TONE RULES (STRICT)
                            ========================
                            - TELEGRAPHIC STYLE: Bullet 1 MUST use the candidate's name. Bullets 2, 3, 4, and 5 MUST strictly drop all pronouns (he/she/they) and names, starting directly with an action verb.
                            - Write like a human pitching to a colleague. Confident, direct, and factual.
                            - TECH MATCHING: Strictly align the tools you highlight with the JD. 
                            - LOCATION NEUTRAL: Never mention the physical location (e.g., Reston, VA, onsite, hybrid) in the summary.
                            - LEADERSHIP VERBS: Use high-authority active verbs (e.g., 'Engineered', 'Optimized', 'Architected').
                            - Do NOT use generic filler: "strong background," "highly experienced," "positions them uniquely."
                            - Do NOT repeat or restate the job description.

                            ========================
                            SKILLS SECTION (INTELLIGENT SEMANTIC MATCHING)
                            ========================
                            - You are evaluating the candidate against the exact skills listed in the "Skills Matrix:" input.
                            - INTELLIGENT MATCHING: You MUST look for semantic equivalents, synonyms, or related frameworks in the resume (e.g., if the matrix asks for "React.js", you MUST credit the candidate if they have "React Native", "React", etc.).
                            - Output the EXACT skill name as it is written in the Skills Matrix input, but calculate the Years and Companies based on your intelligent matching of their resume.
                            - If the provided Skills Matrix only lists 4 skills, you MUST ONLY output those 4 skills and leave SKILL5 through SKILL8 completely empty ("").
                            - You MUST extract and populate the 'Years' and 'Companies' fields for EVERY single skill you evaluate. Do not leave 'Years' or 'Companies' blank if the candidate has experience with the tool/concept.

                            Format: "Skill Area"
                            Years format: "X+ years" ONLY. (CRITICAL: DO NOT add ", current" or any years/dates to this string. Only output the number of years).
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
                            - Use October 2026 as the current date.
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

                    # Deloitte needs a Word-compatible source here because
                    # process_deloitte_doc() copies XML/content from the
                    # candidate's original resume into the template.
                    process_deloitte_doc(
                        TEMPLATE_FILENAME,
                        mapping,
                        resume_docx_path,
                        out_file
                    )
                    
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

                        if (
                            resume_docx_path != resume_path
                            and os.path.exists(resume_docx_path)
                        ):
                            os.remove(resume_docx_path)
                    except:
                        pass

                except Exception as e:
                    st.error(f"❌ Process Failed: {str(e)}")


# ====================================================================
# --- ⚓ CLIENT APP 2: NAVY FEDERAL CREDIT UNION ⚓ ---
# ====================================================================

def navy_generate_json(api_key, prompt, max_attempts=6):
    """Navy Federal Gemini JSON helper using the Freddie retry pattern."""
    client = genai.Client(api_key=api_key)

    models_to_try = [
        "gemini-2.5-flash",
        "gemini-3-flash-preview",
        "gemini-3.1-flash-lite-preview",
        "gemini-pro-latest",
    ]

    last_error = None

    for attempt in range(max_attempts):
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
                ),
            )

            data = repair_and_load_json(response.text)

            if data:
                return data

        except Exception as e:
            last_error = e

            if "503" in str(e) and attempt < max_attempts - 1:
                time.sleep(2 ** ((attempt % 2) + 1))
                continue

            if attempt < max_attempts - 1:
                time.sleep(1)
                continue

    raise RuntimeError(
        "Gemini could not return usable Navy Federal JSON after retries. "
        f"Last error: {last_error}"
    )


def navy_normalize_completion_status(education_item, current_year=None):
    """
    Determine whether an education entry is completed without defaulting to Yes.

    The AI extraction supplies explicit status/date evidence, but this deterministic
    pass overrides any contradictory result when the resume says the credential is
    still in progress or shows a future completion year.
    """
    if not isinstance(education_item, dict):
        return ""

    if current_year is None:
        current_year = date.today().year

    school = str(education_item.get("School", "") or "").strip()
    degree = str(education_item.get("Degree", "") or "").strip()
    completion = str(
        education_item.get("CompletionStatus", "")
        or education_item.get("Completed", "")
        or ""
    ).strip()
    completion_date = str(
        education_item.get("CompletionDate", "")
        or education_item.get("EducationDate", "")
        or ""
    ).strip()

    combined = " ".join(
        value for value in [school, degree, completion, completion_date] if value
    )
    normalized = combined.casefold()

    unfinished_patterns = [
        r"\bin\s+progress\b",
        r"\bpursuing\b",
        r"\bcurrently\s+enrolled\b",
        r"\bexpected\b",
        r"\banticipated\b",
        r"\bpending\b",
        r"\bongoing\b",
        r"\bnot\s+completed\b",
        r"\bnot\s+yet\s+completed\b",
    ]

    if any(re.search(pattern, normalized) for pattern in unfinished_patterns):
        return "No"

    future_years = [
        int(match)
        for match in re.findall(r"\b(?:19|20)\d{2}\b", completion_date)
        if int(match) > current_year
    ]

    if future_years:
        return "No"

    completion_key = completion.casefold()

    if completion_key in {
        "no",
        "n",
        "incomplete",
        "in progress",
        "pursuing",
        "pending",
    }:
        return "No"

    if completion_key in {
        "yes",
        "y",
        "completed",
        "complete",
        "graduated",
        "earned",
        "awarded",
    }:
        return "Yes"

    # A past completion/graduation year is affirmative evidence when the model
    # extracted it specifically as the completion date rather than an attendance
    # date. If the resume is ambiguous, leave the field blank instead of guessing.
    past_or_current_years = [
        int(match)
        for match in re.findall(r"\b(?:19|20)\d{2}\b", completion_date)
        if int(match) <= current_year
    ]

    if past_or_current_years:
        return "Yes"

    return ""


def navy_set_cell_text_preserve_format(cell, value):
    """Replace simple cell text while preserving the template paragraph styling."""
    value = str(value or "")

    if not cell.paragraphs:
        cell.add_paragraph(value)
        return

    paragraph = cell.paragraphs[0]

    if paragraph.runs:
        paragraph.runs[0].text = value
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(value)

    for extra_paragraph in cell.paragraphs[1:]:
        for run in extra_paragraph.runs:
            run.text = ""


def navy_normalize_evidence_text(value):
    """
    Normalize text only for deterministic grounding checks.

    This deliberately ignores punctuation/case/whitespace differences while
    preserving the actual words and numbers. It is NOT used to invent semantic
    equivalence between the JD and the candidate's resume.
    """
    value = str(value or "")
    value = (
        value.replace("\u00a0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    value = value.casefold()
    value = re.sub(r"[^a-z0-9+#./'\-]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def navy_normalize_quote_match_text(value):
    """
    Normalize text for exact-source quote verification while ignoring harmless
    punctuation differences introduced by Gemini (for example, adding a final
    period to an otherwise verbatim JD sentence).

    IMPORTANT: This does NOT perform synonym or semantic matching. The same
    words/numbers must still appear in the same order; only punctuation,
    whitespace, dash style, and capitalization are ignored.
    """
    value = str(value or "")
    value = (
        value.replace("\u00a0", " ")
        .replace("\u2018", "'")
        .replace("\u2019", "'")
        .replace("\u201c", '"')
        .replace("\u201d", '"')
        .replace("\u2013", "-")
        .replace("\u2014", "-")
    )
    value = value.casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    value = re.sub(r"\s+", " ", value).strip()
    return value


def navy_evidence_snippet_found(resume_text, snippet):
    """
    Return True only when the cited evidence is actually present in the uploaded
    resume. Punctuation-only differences are tolerated, but paraphrases are not.
    """
    normalized_resume = navy_normalize_quote_match_text(resume_text)
    normalized_snippet = navy_normalize_quote_match_text(snippet)

    if not normalized_snippet or len(normalized_snippet) < 6:
        return False

    return normalized_snippet in normalized_resume


def navy_clean_summary_style(value):
    """
    Deterministically remove common AI-sounding resume phrasing from Navy Federal
    summary bullets and standardize whole-year experience anchors to X+ years.
    """
    value = re.sub(r"\s+", " ", str(value or "")).strip()

    # "Progressive experience" is valid English but reads formulaic/AI-generated
    # in recruiter summaries. Navy Federal uses the simpler human phrasing.
    value = re.sub(
        r"\bprogressive\s+(?:professional\s+)?experience\b",
        "experience",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(
        r"\bprogressive\s+career\b",
        "career",
        value,
        flags=re.IGNORECASE,
    )

    # Standardize a leading whole-year anchor to the requested recruiter style.
    value = re.sub(
        r"^\s*(\d+)\s+years?\b(?!\s*\+)",
        lambda m: f"{m.group(1)}+ years",
        value,
        count=1,
        flags=re.IGNORECASE,
    )
    value = re.sub(
        r"^\s*(\d+)\+\s+year\b",
        lambda m: f"{m.group(1)}+ years",
        value,
        count=1,
        flags=re.IGNORECASE,
    )

    return value


def navy_string_list(value):
    """Normalize an AI JSON value into a clean list of non-empty strings."""
    if isinstance(value, list):
        return [
            re.sub(r"\s+", " ", str(item or "")).strip()
            for item in value
            if str(item or "").strip()
        ]

    if value is None:
        return []

    cleaned = re.sub(r"\s+", " ", str(value)).strip()
    return [cleaned] if cleaned else []


def navy_skill_detail_records(value):
    """
    Normalize skill-detail evidence records.

    Expected AI shape:
      [{"item": "NIST SP 800-53", "evidence": "exact resume quote"}]
    """
    if not isinstance(value, list):
        return []

    records = []

    for entry in value:
        if not isinstance(entry, dict):
            continue

        item = re.sub(
            r"\s+",
            " ",
            str(
                entry.get("item")
                or entry.get("name")
                or entry.get("value")
                or ""
            ),
        ).strip()

        evidence = re.sub(
            r"\s+",
            " ",
            str(entry.get("evidence") or ""),
        ).strip()

        if item and evidence:
            records.append(
                {
                    "item": item,
                    "evidence": evidence,
                }
            )

    return records


def navy_signal_tokens(value):
    """
    Create conservative lexical tokens for candidate-evidence checks.

    Very common connector words are ignored. Lightweight suffix normalization
    allows "track" to support "tracking" while still requiring the substantive
    words in a skill detail to appear in the cited resume evidence.
    """
    stop_words = {
        "a", "an", "and", "as", "at", "by", "for", "from", "in", "into",
        "of", "on", "or", "the", "to", "using", "with", "within",
        "management", "support", "process", "processes", "program", "programs",
    }

    def root(token):
        token = token.casefold().strip()

        # Safe derivational normalization for common resume terminology.
        # These are grammatical variants, not semantic substitutions.
        derivational_roots = {
            "compliance": "compli",
            "compliant": "compli",
            "comply": "compli",
            "auditor": "audit",
            "auditors": "audit",
            "auditing": "audit",
            "audits": "audit",
            "assessment": "assess",
            "assessments": "assess",
            "assessing": "assess",
            "assessed": "assess",
            "validation": "valid",
            "validate": "valid",
            "validated": "valid",
            "validating": "valid",
            "remediation": "remedi",
            "remediate": "remedi",
            "remediated": "remedi",
            "remediating": "remedi",
            "analysis": "analy",
            "analyze": "analy",
            "analyzed": "analy",
            "analytical": "analy",
            "policy": "polic",
            "policies": "polic",
            "procedure": "procedur",
            "procedures": "procedur",
            "tracking": "track",
            "tracked": "track",
            "tracks": "track",
            "controls": "control",
            "controlled": "control",
            "governance": "govern",
            "governing": "govern",
            "regulation": "regulat",
            "regulations": "regulat",
            "regulatory": "regulat",
        }

        if token in derivational_roots:
            return derivational_roots[token]

        # Preserve acronyms/numerical framework identifiers naturally.
        if len(token) <= 3:
            return token

        for suffix in ("ing", "ed", "es", "s"):
            if token.endswith(suffix) and len(token) - len(suffix) >= 4:
                return token[:-len(suffix)]

        return token

    raw_tokens = re.findall(
        r"[A-Za-z0-9]+(?:[./+-][A-Za-z0-9]+)*",
        str(value or ""),
    )

    output = []

    for token in raw_tokens:
        lowered = token.casefold()

        if lowered in stop_words:
            continue

        normalized = root(lowered)

        if normalized and normalized not in output:
            output.append(normalized)

    return output


def navy_skill_item_supported_by_evidence(item, evidence):
    """
    Deterministically require the substantive words in each right-column skill
    detail to be present in its cited resume evidence.

    This blocks JD-only details such as GLBA/FFIEC from surviving merely because
    the model cited some unrelated compliance sentence.
    """
    item_tokens = navy_signal_tokens(item)
    evidence_tokens = set(navy_signal_tokens(evidence))

    if not item_tokens:
        return False

    matched = sum(1 for token in item_tokens if token in evidence_tokens)

    # One-token items (e.g. "FFIEC", "ServiceNow") must match exactly.
    if len(item_tokens) == 1:
        return matched == 1

    # Multi-token conceptual methods may legitimately summarize the exact resume
    # wording (for example, "Security Policies & Procedures" from evidence that
    # says "information security policy ... procedures"). Require a clear
    # majority of the substantive words rather than near-verbatim overlap. Named
    # one-token tools/frameworks still require an exact match above.
    required = max(1, int((len(item_tokens) * 0.60) + 0.9999))
    return matched >= required


def navy_detail_compound_claim_supported(item, evidence):
    """
    Protect against combining individually supported words into a stronger
    compound claim that the cited resume sentence does not actually make.

    Right-column details are supposed to be atomic. When Gemini joins concepts
    with &, /, "and", or a comma, require the combined wording to be supported
    as one coherent concept by the SAME cited resume excerpt. Otherwise the
    detail is rejected and the repair pass must split it into resume-native
    atomic items.

    Examples:
    - "SOC 1 & SOC 2 Audits" is NOT supported by one sentence that says
      "SOC 2 and SOX audits" plus another sentence elsewhere about SOC 1 reports.
    - "Defect Tracking & Remediation" is NOT supported when defect tracking and
      remediation plans come from different resume bullets.
    - A literal combined phrase already present in the cited resume excerpt is
      allowed.
    """
    item_text = re.sub(r"\s+", " ", str(item or "")).strip()
    evidence_text = re.sub(r"\s+", " ", str(evidence or "")).strip()

    if not item_text or not evidence_text:
        return False

    # Only invoke the stricter compound rule when the generated detail visibly
    # joins multiple concepts. Hyphens inside one term are not treated as joins.
    has_compound_connector = bool(
        re.search(r"\s(?:&|/|and)\s|,", item_text, flags=re.IGNORECASE)
        or re.search(r"\b[A-Za-z0-9]+/[A-Za-z0-9]+\b", item_text)
    )

    if not has_compound_connector:
        return True

    item_tokens = navy_signal_tokens(item_text)
    evidence_tokens = navy_signal_tokens(evidence_text)

    if not item_tokens:
        return False

    # Every substantive token in a compound detail must be supported by the
    # SAME evidence excerpt. This is intentionally stricter than the general
    # 60% conceptual threshold used for simple atomic details.
    if not all(token in set(evidence_tokens) for token in item_tokens):
        return False

    # Named/versioned phrases such as SOC 1, SOC 2, NIST 800-53, ISO 27001,
    # PCI DSS, etc. must each be visible in the same evidence excerpt. This
    # prevents a generic acronym match from turning SOC 1 report review into
    # SOC 1 audit experience.
    named_phrases = []
    for pattern in [
        r"\b[A-Z]{2,}(?:\s+[A-Z]{1,4})?\s+\d+(?:[-./]\d+)?\b",
        r"\b[A-Z]{2,}(?:\s+[A-Z]{2,})+\b",
    ]:
        for match in re.findall(pattern, item_text):
            cleaned = re.sub(r"\s+", " ", match).strip()
            if cleaned and cleaned not in named_phrases:
                named_phrases.append(cleaned)

    evidence_norm = navy_normalize_evidence_text(evidence_text)
    for phrase in named_phrases:
        phrase_norm = navy_normalize_evidence_text(phrase)
        if phrase_norm and phrase_norm not in evidence_norm:
            return False

    return True


def navy_jd_evidence_snippet_found(jd_text, snippet):
    """
    Return True only when a cited JD excerpt is actually present in the JD.

    Gemini may add/remove harmless punctuation around an otherwise verbatim
    quote, so punctuation and whitespace are ignored here. Word order and the
    actual wording still have to match the Fieldglass description.
    """
    normalized_jd = navy_normalize_quote_match_text(jd_text)
    normalized_snippet = navy_normalize_quote_match_text(snippet)

    if not normalized_snippet or len(normalized_snippet) < 6:
        return False

    return normalized_snippet in normalized_jd


def navy_role_priority_evidence_snippet_found(jd_text, internal_notes, snippet):
    """
    Verify that role-priority evidence is traceable to either current Internal
    Notes or the Official Fieldglass JD.

    Role-priority evidence establishes RELEVANCE ONLY; it never proves candidate
    experience. Because Gemini can occasionally shorten a long source sentence,
    omit intervening list items, or insert an ellipsis while preserving the
    source meaning, this check is intentionally a little more tolerant than the
    candidate-resume evidence check.

    Validation order:
    1. Prefer an exact normalized source substring.
    2. Otherwise allow an abridged/reordered excerpt only when EVERY substantive
       token in the cited evidence is present in the SAME source (Internal Notes
       or JD). This tolerates harmless quote shortening but does not allow a new
       tool, requirement, qualifier, framework, or other substantive concept to
       be invented.

    Candidate claims remain protected separately by the much stricter resume
    evidence validator.
    """
    snippet_text = str(snippet or "").strip()
    if not snippet_text:
        return False

    sources = [
        str(internal_notes or "").strip(),
        str(jd_text or "").strip(),
    ]

    # Best case: Gemini returned an actual contiguous source excerpt.
    for source in sources:
        if source and navy_jd_evidence_snippet_found(source, snippet_text):
            return True

    # Conservative fallback for harmlessly abbreviated role-priority quotes.
    # Use the same lexical normalization used elsewhere in Navy validation so
    # grammatical variants do not create false failures. ALL substantive tokens
    # must still be traceable to one source; we never merge support across Notes
    # and JD.
    snippet_tokens = navy_signal_tokens(
        navy_normalize_quote_match_text(snippet_text)
    )

    if not snippet_tokens:
        return False

    for source in sources:
        if not source:
            continue

        source_tokens = set(
            navy_signal_tokens(
                navy_normalize_quote_match_text(source)
            )
        )

        if source_tokens and all(
            token in source_tokens
            for token in snippet_tokens
        ):
            return True

    return False


def navy_explicit_identifiers(value):
    """
    Extract concrete identifiers that should never be invented by an umbrella
    skill label: acronyms, mixed-case product names, and framework/version tokens.

    Examples: FFIEC, GLBA, NIST, ServiceNow, ZenGRC, BitSight, 800-53.
    Generic title-cased words such as Security, Risk, Policies, or Procedures are
    intentionally ignored because high-level categories may summarize evidence.
    """
    text = str(value or "")
    identifiers = []

    # All-caps acronyms / named standards. Ignore very common presentation words
    # that can safely appear as general resume terminology.
    for token in re.findall(r"\b[A-Z][A-Z0-9./+-]{1,}\b", text):
        if token not in identifiers:
            identifiers.append(token)

    # Mixed-case product/vendor names such as ServiceNow, BitSight, ZenGRC.
    for token in re.findall(r"\b[A-Z][a-z]+[A-Z][A-Za-z0-9]*\b", text):
        if token not in identifiers:
            identifiers.append(token)

    # Version/control identifiers containing digits and separators.
    for token in re.findall(r"\b(?:[A-Za-z]{1,8}\s*)?\d{2,4}(?:[-./]\d+[A-Za-z]?)?\b", text):
        clean = re.sub(r"\s+", " ", token).strip()
        if clean and clean not in identifiers:
            identifiers.append(clean)

    return identifiers


def navy_explicit_identifiers_supported(value, evidence_text, resume_text):
    """
    Return unsupported concrete identifiers from an output label/item.

    The evidence snippets must come from the resume already; this extra check
    prevents an umbrella phrase from smuggling in a named regulation/tool that
    the resume never contained.
    """
    evidence_norm = navy_normalize_evidence_text(evidence_text)
    resume_norm = navy_normalize_evidence_text(resume_text)
    unsupported = []

    for identifier in navy_explicit_identifiers(value):
        ident_norm = navy_normalize_evidence_text(identifier)
        if not ident_norm:
            continue
        if ident_norm not in evidence_norm and ident_norm not in resume_norm:
            unsupported.append(identifier)

    return unsupported



def navy_unsupported_high_impact_claims(output_text, evidence_text):
    """
    Return high-impact qualifiers/authority claims used in generated candidate
    wording that are not themselves supported by the cited resume evidence.

    This is intentionally ROLE-AGNOSTIC. Navy Federal requisitions span technical,
    operational, HR, legal, facilities, design, finance, risk, data, and many
    other disciplines, so the validator protects claim strength rather than
    hard-coding any one job family's terminology.

    The purpose is to stop a JD adjective/verb from hitchhiking onto an otherwise
    supported noun phrase. Example: evidence may support "vendor risk assessments"
    but not "independent vendor risk assessments"; the unsupported qualifier
    "independent" must therefore be removed.

    Only the evidence excerpts cited for the specific Summary/Skill claim are
    considered. If the resume supports the qualifier elsewhere, Gemini can fix the
    claim by citing that exact supporting excerpt on its repair pass.
    """
    output = str(output_text or "")
    evidence = str(evidence_text or "")

    # Each tuple:
    # (human-readable label, generated-output pattern, acceptable evidence pattern)
    #
    # Evidence patterns intentionally allow grammatical variants of the SAME
    # underlying claim (lead/led/leading, own/owned/ownership, etc.) but do not
    # substitute adjacent concepts (support != lead, participate != own).
    rules = [
        (
            "independent/independently",
            r"\bindependent(?:ly)?\b",
            r"\bindependent(?:ly)?\b",
        ),
        (
            "Second Line / 2LOD",
            r"\b(?:2lod|second[\s-]+line(?:\s+of\s+defen[cs]e)?)\b",
            r"\b(?:2lod|second[\s-]+line(?:\s+of\s+defen[cs]e)?)\b",
        ),
        (
            "Three Lines model",
            r"\b(?:three|3)[\s-]+lines?(?:\s+of\s+defen[cs]e|\s+model)?\b",
            r"\b(?:three|3)[\s-]+lines?(?:\s+of\s+defen[cs]e|\s+model)?\b",
        ),
        (
            "enterprise-wide",
            r"\benterprise[\s-]+wide\b",
            r"\benterprise[\s-]+wide\b",
        ),
        (
            "organization-wide",
            r"\b(?:organization|organisation)[\s-]+wide\b",
            r"\b(?:organization|organisation)[\s-]+wide\b",
        ),
        (
            "global/globally",
            r"\bglobal(?:ly)?\b",
            r"\bglobal(?:ly)?\b",
        ),
        (
            "lead/led/leading/leadership",
            r"\b(?:lead|leads|leading|led|leadership)\b",
            r"\b(?:lead|leads|leading|led|leadership)\b",
        ),
        (
            "ownership/owned",
            r"\b(?:own|owns|owned|owning|ownership)\b",
            r"\b(?:own|owns|owned|owning|ownership)\b",
        ),
        (
            "manage/managed/managing/manager",
            # Do NOT treat the noun "management" by itself as an authority
            # qualifier. Across Navy Federal roles, legitimate disciplines such
            # as Data Management, Project Management, Risk Management, Change
            # Management, Content Management, and Metadata Management use the
            # word without implying people-management or ownership authority.
            r"\b(?:manage|manages|managed|managing|manager)\b",
            r"\b(?:manage|manages|managed|managing|manager)\b",
        ),
        (
            "oversight/oversee",
            r"\b(?:oversee|oversees|oversaw|overseeing|oversight)\b",
            r"\b(?:oversee|oversees|oversaw|overseeing|oversight)\b",
        ),
        (
            "architect/architected/architecture",
            r"\b(?:architect|architects|architected|architecting|architecture)\b",
            r"\b(?:architect|architects|architected|architecting|architecture)\b",
        ),
        (
            "expert/expertise",
            r"\b(?:expert|expertise)\b",
            r"\b(?:expert|expertise)\b",
        ),
        (
            "advanced",
            r"\badvanced\b",
            r"\badvanced\b",
        ),
        (
            "extensive",
            r"\bextensive\b",
            r"\bextensive\b",
        ),
        (
            "deep",
            r"\bdeep\b",
            r"\bdeep\b",
        ),
        (
            "strong",
            r"\bstrong\b",
            r"\bstrong\b",
        ),
        (
            "hands-on",
            r"\bhands?[\s-]+on\b",
            r"\bhands?[\s-]+on\b",
        ),
        (
            "cross-functional",
            r"\bcross[\s-]+functional\b",
            r"\bcross[\s-]+functional\b",
        ),
        (
            "senior",
            r"\bsenior\b",
            r"\bsenior\b",
        ),
        (
            "principal",
            r"\bprincipal\b",
            r"\bprincipal\b",
        ),
        (
            "strategic/strategy",
            r"\b(?:strategic|strategy)\b",
            r"\b(?:strategic|strategy)\b",
        ),
        (
            "regulatory/regulation",
            r"\b(?:regulatory|regulation|regulations)\b",
            r"\b(?:regulatory|regulation|regulations)\b",
        ),
        (
            "enforce/enforcement",
            r"\b(?:enforce|enforces|enforced|enforcing|enforcement)\b",
            r"\b(?:enforce|enforces|enforced|enforcing|enforcement)\b",
        ),
        (
            "validate/validation",
            r"\b(?:validate|validates|validated|validating|validation)\b",
            r"\b(?:validate|validates|validated|validating|validation)\b",
        ),
        (
            "approve/approval",
            r"\b(?:approve|approves|approved|approving|approval)\b",
            r"\b(?:approve|approves|approved|approving|approval)\b",
        ),
        (
            "high-volume",
            r"\bhigh[\s-]+volume\b",
            r"\bhigh[\s-]+volume\b",
        ),
        (
            "large-scale",
            r"\blarge[\s-]+scale\b",
            r"\blarge[\s-]+scale\b",
        ),
        (
            "mission-critical",
            r"\bmission[\s-]+critical\b",
            r"\bmission[\s-]+critical\b",
        ),
        (
            "end-to-end",
            r"\bend[\s-]+to[\s-]+end\b",
            r"\bend[\s-]+to[\s-]+end\b",
        ),
        (
            "full lifecycle",
            r"\bfull[\s-]+life[\s-]*cycle\b",
            r"\bfull[\s-]+life[\s-]*cycle\b",
        ),
        (
            "production",
            r"\bproduction\b",
            r"\bproduction\b",
        ),
        (
            "sole/primary ownership",
            r"\b(?:sole|primary)\s+(?:owner|ownership|responsibility)\b",
            r"\b(?:sole|primary)\s+(?:owner|ownership|responsibility)\b",
        ),
    ]

    unsupported = []
    seen = set()

    for label, output_pattern, evidence_pattern in rules:
        if not re.search(output_pattern, output, flags=re.IGNORECASE):
            continue

        if re.search(evidence_pattern, evidence, flags=re.IGNORECASE):
            continue

        key = label.casefold()
        if key not in seen:
            unsupported.append(label)
            seen.add(key)

    return unsupported


def navy_jd_only_phrase_hits(output_text, jd_text, resume_text, evidence_text=""):
    """
    Detect suspicious multi-word phrases copied from the JD into candidate-owned
    output when that wording is not grounded in the resume.

    This is a deterministic backstop against phrases such as "effective challenge"
    or "Issue Owners" leaking from the requisition into the candidate profile.

    Natural resume-grounded paraphrases are allowed when the SPECIFIC evidence
    cited for the same claim contains all substantive concepts. Concrete named
    identifiers and high-impact qualifiers remain protected by separate stricter
    validators, so this evidence-aware allowance does not permit unsupported
    tools, frameworks, authority, scale, or proficiency claims.
    """
    output_norm = navy_normalize_evidence_text(output_text)
    jd_norm = navy_normalize_evidence_text(jd_text)
    resume_norm = navy_normalize_evidence_text(resume_text)

    if not output_norm or not jd_norm:
        return []

    output_words = output_norm.split()

    stop_words = {
        "a", "an", "and", "as", "at", "be", "by", "for", "from", "in", "is",
        "of", "on", "or", "the", "to", "with", "within", "years", "year",
        "experience", "candidate", "role",
    }

    hits = []
    seen = set()

    # Prefer longer overlaps first so a 4-word JD leak is reported once rather
    # than as every possible 2-word sub-phrase.
    for size in (5, 4, 3, 2):
        if len(output_words) < size:
            continue

        for idx in range(0, len(output_words) - size + 1):
            phrase_words = output_words[idx:idx + size]
            phrase = " ".join(phrase_words)

            if phrase in seen:
                continue

            signal_words = [
                word for word in phrase_words
                if word not in stop_words and len(word) >= 4
            ]

            # Require a genuinely meaningful phrase, not a content word plus a
            # connector such as "assessments and". This preserves protection for
            # JD-only phrases such as "effective challenge", "Issue Owners", or
            # "Second Line of Defense" while avoiding low-information false hits.
            if len(signal_words) < 2:
                continue

            if phrase_words[0] in stop_words or phrase_words[-1] in stop_words:
                continue

            if phrase in jd_norm and phrase not in resume_norm:
                # List punctuation/connectors can make the same resume-backed
                # technology sequence look different after normalization. For
                # example, a resume may say "Advanced SQL, Python, PySpark"
                # while the JD/output says "SQL, Python and PySpark". That is
                # not a JD-only candidate claim. Before flagging the phrase,
                # remove low-information connector words from BOTH the phrase
                # and resume and require the remaining substantive sequence to
                # appear contiguously in the resume. This is intentionally a
                # sequence check, not a bag-of-words check, so unrelated words
                # scattered elsewhere in the resume cannot manufacture support.
                phrase_core = [
                    word for word in phrase_words
                    if word not in stop_words
                ]
                resume_core = [
                    word for word in resume_norm.split()
                    if word not in stop_words
                ]

                core_is_resume_backed = False
                if phrase_core:
                    core_len = len(phrase_core)
                    for core_idx in range(0, len(resume_core) - core_len + 1):
                        if resume_core[core_idx:core_idx + core_len] == phrase_core:
                            core_is_resume_backed = True
                            break

                if core_is_resume_backed:
                    continue

                # Also allow a natural paraphrase when the exact resume evidence
                # cited for THIS claim contains every substantive concept in the
                # phrase. Example: "data deletion controls" is fairly supported
                # by evidence saying "data deletion requirements, control
                # procedures" even though the exact JD phrase is absent. This is
                # scoped to the claim's cited evidence rather than the whole resume
                # so unrelated concepts elsewhere cannot manufacture support.
                phrase_signal_tokens = navy_signal_tokens(phrase)
                evidence_signal_tokens = set(
                    navy_signal_tokens(evidence_text)
                )

                if (
                    phrase_signal_tokens
                    and evidence_signal_tokens
                    and all(
                        token in evidence_signal_tokens
                        for token in phrase_signal_tokens
                    )
                ):
                    continue

                # Do not add shorter phrases that are fully contained in an
                # already-reported longer phrase.
                if any(phrase in existing for existing in hits):
                    continue

                hits.append(phrase)
                seen.add(phrase)

    return hits[:12]


def navy_validate_grounded_output(resume_text, jd_text, internal_notes, data):
    """
    Deterministically validate the audited Navy Federal summary/skills package.

    The document is allowed to proceed only when:
    - every Summary bullet cites real resume evidence,
    - every skill row cites real resume evidence,
    - every skill row cites real role-priority evidence from Internal Notes or
      the Official Fieldglass JD,
    - every right-column skill detail cites real resume evidence and is
      lexically supported by that evidence,
    - no suspicious JD-only multi-word phrase survives,
    - high-impact qualifiers/authority claims are independently supported by the
      cited resume evidence,
    - AI-sounding "progressive experience" wording is absent.
    """
    errors = []

    for i in range(1, 5):
        summary_key = f"SUMMARY{i}"
        evidence_key = f"SUMMARY{i}_EVIDENCE"

        summary = navy_clean_summary_style(
            data.get(summary_key, "")
        )
        data[summary_key] = summary

        if not summary:
            errors.append(f"{summary_key} is blank.")
            continue

        if re.search(r"\bprogressive\b", summary, flags=re.IGNORECASE):
            errors.append(
                f"{summary_key} still contains the prohibited word 'progressive'."
            )

        evidence_items = navy_string_list(
            data.get(evidence_key, [])
        )

        if not evidence_items:
            errors.append(
                f"{summary_key} has no resume evidence."
            )
        else:
            for evidence in evidence_items:
                if not navy_evidence_snippet_found(
                    resume_text,
                    evidence,
                ):
                    errors.append(
                        f"{summary_key} cites evidence not found in the resume: "
                        f"{evidence[:120]}"
                    )

        unsupported_qualifiers = navy_unsupported_high_impact_claims(
            summary,
            " ".join(evidence_items),
        )

        if unsupported_qualifiers:
            errors.append(
                f"{summary_key} contains high-impact claim(s) not supported by "
                "its cited resume evidence: "
                + ", ".join(unsupported_qualifiers)
            )

        jd_hits = navy_jd_only_phrase_hits(
            summary,
            jd_text,
            resume_text,
            " ".join(evidence_items),
        )

        if jd_hits:
            errors.append(
                f"{summary_key} contains JD-only wording not found in the resume: "
                + "; ".join(jd_hits)
            )

    for i in range(1, 5):
        skill_key = f"SKILL{i}"
        skill_evidence_key = f"SKILL{i}_EVIDENCE"
        skill_role_evidence_key = f"SKILL{i}_ROLE_EVIDENCE"
        details_key = f"SKILLDETAILS{i}"

        skill = re.sub(
            r"\s+",
            " ",
            str(data.get(skill_key, "") or ""),
        ).strip()
        data[skill_key] = skill

        if not skill:
            errors.append(f"{skill_key} is blank.")
            continue

        skill_evidence = navy_string_list(
            data.get(skill_evidence_key, [])
        )

        if not skill_evidence:
            errors.append(
                f"{skill_key} has no resume evidence."
            )
        else:
            for evidence in skill_evidence:
                if not navy_evidence_snippet_found(
                    resume_text,
                    evidence,
                ):
                    errors.append(
                        f"{skill_key} cites evidence not found in the resume: "
                        f"{evidence[:120]}"
                    )

            # Left-column Skills are intentionally high-level umbrella categories.
            # Do NOT demand near-verbatim lexical overlap with the evidence; that
            # creates false failures for legitimate summaries such as
            # "Security Policies & Procedures." Instead, verify that any concrete
            # named identifiers embedded in the category are actually resume-backed.
            unsupported_identifiers = navy_explicit_identifiers_supported(
                skill,
                " ".join(skill_evidence),
                resume_text,
            )

            if unsupported_identifiers:
                errors.append(
                    f"{skill_key} contains unsupported named identifier(s): "
                    + ", ".join(unsupported_identifiers)
                )

        # Each scarce Skills row must also prove WHY it deserves one of the
        # four slots. In v7 this role-priority evidence may come from substantive
        # Internal Notes (supplier-call / hiring-team context) OR the Official
        # Fieldglass JD. It establishes relevance only and can never establish
        # candidate experience.
        skill_role_evidence = navy_string_list(
            data.get(skill_role_evidence_key, [])
        )

        if not skill_role_evidence:
            errors.append(
                f"{skill_key} has no role-priority evidence from Internal Notes or the Official Fieldglass JD."
            )
        else:
            for role_evidence in skill_role_evidence:
                if not navy_role_priority_evidence_snippet_found(
                    jd_text,
                    internal_notes,
                    role_evidence,
                ):
                    errors.append(
                        f"{skill_key} cites role-priority evidence not found in Internal Notes or the Official Fieldglass JD: "
                        f"{role_evidence[:120]}"
                    )

        unsupported_skill_qualifiers = navy_unsupported_high_impact_claims(
            skill,
            " ".join(skill_evidence),
        )

        if unsupported_skill_qualifiers:
            errors.append(
                f"{skill_key} contains high-impact claim(s) not supported by "
                "its cited resume evidence: "
                + ", ".join(unsupported_skill_qualifiers)
            )

        skill_jd_hits = navy_jd_only_phrase_hits(
            skill,
            jd_text,
            resume_text,
            " ".join(skill_evidence),
        )

        # High-level competency labels may legitimately paraphrase resume wording.
        # Only treat JD overlap as suspicious when the cited resume evidence does
        # NOT independently support the concept. This still blocks formal JD-only
        # labels such as "Second Line of Defense" when the resume lacks them.
        if skill_jd_hits and not navy_skill_item_supported_by_evidence(
            skill,
            " ".join(skill_evidence),
        ):
            errors.append(
                f"{skill_key} contains JD-only wording not supported by its "
                "resume evidence: " + "; ".join(skill_jd_hits)
            )

        detail_records = navy_skill_detail_records(
            data.get(details_key, [])
        )

        if not detail_records:
            errors.append(
                f"{details_key} has no grounded detail records."
            )
            continue

        verified_detail_records = []
        rejected_detail_items = []

        for detail in detail_records[:6]:
            item = detail["item"]
            evidence = detail["evidence"]
            reject_reason = ""

            if not navy_evidence_snippet_found(
                resume_text,
                evidence,
            ):
                reject_reason = "evidence was not found in the resume"
            else:
                unsupported_identifiers = navy_explicit_identifiers_supported(
                    item,
                    evidence,
                    resume_text,
                )

                if unsupported_identifiers:
                    reject_reason = (
                        "unsupported named identifier(s): "
                        + ", ".join(unsupported_identifiers)
                    )

            if not reject_reason and not navy_skill_item_supported_by_evidence(
                item,
                evidence,
            ):
                reject_reason = "wording was not sufficiently supported by its evidence"

            if not reject_reason:
                unsupported_detail_qualifiers = navy_unsupported_high_impact_claims(
                    item,
                    evidence,
                )
                if unsupported_detail_qualifiers:
                    reject_reason = (
                        "high-impact claim(s) were not supported by the same "
                        "resume evidence: "
                        + ", ".join(unsupported_detail_qualifiers)
                    )

            if not reject_reason and not navy_detail_compound_claim_supported(
                item,
                evidence,
            ):
                reject_reason = (
                    "compound wording joined concepts that were not all supported "
                    "by the same resume evidence"
                )

            # At this point the detail has real resume evidence, no unsupported
            # named identifiers, and sufficient lexical/conceptual support from
            # that evidence. A matching JD phrase is therefore relevance, not a
            # hallucination, and should not cause a false rejection.

            if reject_reason:
                rejected_detail_items.append(
                    f"{item} ({reject_reason})"
                )
                continue

            verified_detail_records.append(detail)

        # Safe-prune isolated questionable details rather than failing the entire
        # resume when the row still has enough fully grounded content. If fewer
        # than three verified details survive, trigger the automatic AI repair
        # pass so the row can be rebuilt from resume-native evidence.
        data[details_key] = verified_detail_records

        if len(verified_detail_records) < 3:
            error_text = (
                f"{details_key} has only {len(verified_detail_records)} verified "
                "detail item(s); at least 3 are required."
            )

            if rejected_detail_items:
                error_text += (
                    " Rejected: "
                    + "; ".join(rejected_detail_items[:4])
                )

            errors.append(error_text)

    return errors


def navy_finalize_grounded_output(data):
    """
    Convert the audited/evidence-rich JSON into the simple strings expected by
    the Word template. Internal evidence fields never appear in the resume.
    """
    output = {}

    for i in range(1, 5):
        output[f"SUMMARY{i}"] = navy_clean_summary_style(
            data.get(f"SUMMARY{i}", "")
        )

        output[f"SKILL{i}"] = re.sub(
            r"\s+",
            " ",
            str(data.get(f"SKILL{i}", "") or ""),
        ).strip()

        detail_records = navy_skill_detail_records(
            data.get(f"SKILLDETAILS{i}", [])
        )

        detail_names = []
        seen = set()

        for record in detail_records:
            item = re.sub(
                r"\s+",
                " ",
                str(record.get("item", "") or ""),
            ).strip()

            key = item.casefold()

            if not item or key in seen:
                continue

            detail_names.append(item)
            seen.add(key)

            if len(detail_names) >= 6:
                break

        output[f"SKILLDETAILS{i}"] = ", ".join(detail_names)

    return output


def navy_grounding_audit_prompt(
    resume_text,
    jd_text,
    internal_notes,
    draft_data,
    current_date,
    validator_errors=None,
):
    """
    Build the independent Navy Federal evidence-audit prompt.

    This pass is intentionally separate from the recruiter-writing pass. Its job
    is not to make the candidate sound better; its job is to remove/rewrite any
    claim that the uploaded resume does not support.
    """
    error_section = ""

    if validator_errors:
        error_section = (
            "\nDETERMINISTIC VALIDATOR ERRORS FROM THE PREVIOUS PASS:\n- "
            + "\n- ".join(validator_errors)
            + "\nYou MUST fix every listed problem.\n"
        )

    return f"""
Return a valid JSON object ONLY.

You are the FINAL FACT-CHECKER for a Navy Federal Credit Union Fieldglass
resume formatter.

CURRENT DATE:
{current_date}

YOUR ONLY PURPOSE:
Make the four Summary bullets and four Skills rows accurate, grounded, and
optimized first for the NAVY FEDERAL HIRING MANAGER'S interview decision, while
preserving truthful Official-JD overlap for possible Fieldglass ranking/scoring.
Never assign JD-only or Internal-Notes-only experience to the candidate.

This is a compliance/fact-audit step, NOT a creative-writing step.

SOURCE ROLES:
- ORIGINAL RESUME = the ONLY source of truth for candidate experience, skills,
  tools, frameworks, methodologies, employers, scope, ownership, metrics, and
  domain knowledge.
- INTERNAL NOTES = HIGHEST-PRIORITY relevance source when substantive. Use
  supplier-call/hiring-team context to determine what will actually drive the
  interview decision, including must-haves, prior-candidate gaps, day-to-day
  work, and technical/business weighting.
- OFFICIAL FIELDGLASS JD = baseline relevance source and secondary potential
  matching/ranking signal.
- INTERNAL NOTES and the JD are NEVER candidate evidence.

ABSOLUTE RULE:
If a word/phrase comes from the JD or Internal Notes and describes the
CANDIDATE, it may be used only when the same capability is explicitly
demonstrated in the Original Resume. Do not infer that a related skill is the
same skill.

If Internal Notes say a resume-proven skill is the #1 priority even when the JD
omits it, KEEP that skill prominent. The Notes prove role priority and the resume
proves candidate capability.

UNIVERSAL FACT-CHECK EXAMPLES:
- JD names a tool/framework/module/regulation that the resume does not -> DO NOT
  output it as candidate experience.
- JD says "lead" while the resume says "supported" -> keep "supported"; do not
  upgrade it to "lead".
- JD says "enterprise-wide" while the resume proves only project/team scope ->
  omit "enterprise-wide".
- JD says "advanced" or "expert" for a technology while the resume only lists or
  uses it -> do not claim an unsupported proficiency level.
- JD says "end-to-end ownership" while the resume proves participation in only
  part of the lifecycle -> do not claim end-to-end ownership.
- JD says "high-volume" or "large-scale" while the resume gives no scale evidence
  -> omit the scale qualifier.
- Resume-supported adjacent experience may still be highlighted using the
  resume's own accurate wording.

HIGH-IMPACT QUALIFIER RULE — UNIVERSAL:
Qualifiers that change LEVEL, AUTHORITY, SCOPE, PROFICIENCY, SCALE, or FORMAL
RESPONSIBILITY are candidate facts and need their own support. This includes,
without limitation: independent/independently, formal line-of-defense language,
lead/led/leading, own/owned/ownership, manage/managed/manager, oversight,
architect/architected/architecture, senior, principal, expert/expertise,
advanced, extensive, deep, strong, hands-on, cross-functional, enterprise-wide,
organization-wide, global, strategic, regulatory, enforcement, validation,
approval, high-volume, large-scale, mission-critical, end-to-end, full
lifecycle, and production.

For every Summary bullet, Skill category, and right-column detail, if one of
these stronger qualifiers appears, the evidence cited for THAT SAME claim must
support the qualifier itself. A supported noun phrase does not make an
unsupported modifier true.

IMPORTANT MANAGEMENT DISTINCTION:
- Treat verbs/titles that assert authority (managed, managing, manager) as claims
  that need direct resume support.
- Do NOT treat the standalone discipline noun "management" as an authority
  claim. Phrases such as Data Management, Project Management, Risk Management,
  Change Management, Content Management, Metadata Management, Vendor Management,
  and similar domain labels may be used when the resume supports that discipline.

SUMMARY STYLE:
- Exactly 4 bullets.
- One sentence per bullet.
- Do NOT use the word "progressive" anywhere.
- Prefer natural recruiter language such as "10+ years of experience..." rather
  than "10 years of progressive experience".
- Do NOT start with the candidate's name.
- No "strong background", "highly experienced", "seasoned professional",
  "excellent fit", "aligns well", or similar filler.
- If using a total-years statement, calculate only from dated work history,
  round DOWN to completed whole years, and format as "X+ years".
- Do not turn the target role's responsibilities into claims that the candidate
  already performed them.

SUMMARY EVIDENCE:
For EACH summary bullet, return 1-3 short VERBATIM excerpts copied from the
Original Resume that together support every substantive candidate claim in that
bullet.
- Evidence must be exact source wording, not paraphrased.
- Do not cite the JD.
- If you cannot support a bullet with exact resume evidence, rewrite the bullet.

SKILL ROWS:
- Exactly 4 rows.
- Left column = concise high-level competency.
- A high-level category may summarize multiple resume bullets, but it must be a
  fair umbrella for work explicitly demonstrated in the resume.
- Do not use a formal discipline/title from the JD merely because it sounds
  adjacent to the candidate's experience.

FOUR-SLOT PRIORITY RULE — CRITICAL:
- Treat the four rows as scarce space. Do NOT merely choose four valid skills.
- First identify substantive Internal-Notes must-haves, prior-candidate gaps,
  interview criteria, day-to-day priorities, and technical/business weighting.
- Then identify the Official JD's core responsibilities and must-haves.
- Map all resume-supported competencies against those priorities and use the top
  four DISTINCT competencies, prioritizing Internal Notes first when substantive.
- A resume-supported manager priority missing from the JD MAY take a row.
- Core day-to-day requirements outrank generic qualifications, adjacent tools,
  nice-to-haves, and impressive-but-secondary experience.
- For EACH row, return SKILL#_ROLE_EVIDENCE with 1-2 short VERBATIM excerpts from
  either Internal Notes or the Official Fieldglass JD showing why that row
  deserves one of the four slots.
- Each ROLE_EVIDENCE excerpt must be ONE CONTIGUOUS source clause/sentence copied
  exactly. Do NOT insert "...", omit words from the middle of a source sentence,
  splice two distant fragments together, or rewrite the source. If the original
  sentence is long, choose a shorter contiguous clause that still proves the
  priority.
- Role-priority evidence proves RELEVANCE ONLY and must never be used as
  candidate evidence.

SKILL ROW EVIDENCE:
For EACH left-column Skill, return 1-3 VERBATIM resume excerpts supporting that
category. The left-column label is an umbrella category, so it may reasonably
paraphrase/summarize multiple cited resume excerpts. It does NOT need to be a
verbatim phrase from the resume. However, never put a named tool, regulation,
framework, or formal discipline in the category unless the resume supports it.

RIGHT-COLUMN DETAILS:
- 3-6 items per row.
- Every item must be a tool, technology, framework, methodology, process, or
  domain method explicitly named or directly demonstrated in the resume.
- USE RESUME-NATIVE WORDING whenever possible. If the resume says "defects",
  prefer "Defect Tracking" rather than renaming it "Issue Tracking" just because
  the JD uses the word issue.
- KEEP EACH DETAIL ATOMIC. Do not combine two concepts with "&", "/", "and",
  or a comma unless ONE cited resume excerpt supports the FULL combined claim.
- Never combine individually supported words into a stronger relationship the
  resume does not make. Example: a resume that says "SOC 2 and SOX audits" and
  elsewhere says "SOC 1 and 2 reports" does NOT support "SOC 1 & SOC 2 Audits".
  Use truthful atomic items such as "SOC 2 Audits" and "SOC 1/2 Evidence Review"
  only when the cited resume wording supports each item.
- If one resume bullet supports defect tracking and another supports remediation
  plans, output "Defect Tracking" and "Risk Treatment Plans" separately rather
  than "Issue Tracking & Remediation".
- Return each detail as an object containing:
  "item": the concise resume skill/detail
  "evidence": one VERBATIM resume excerpt that explicitly supports that item
- Do not return an item when you cannot quote one resume excerpt that supports
  the item.
- Do not preserve a stronger qualifier merely because the underlying skill/item
  is supported. The qualifier itself must be supported by that same evidence.

HIRING-MANAGER + FIELDGLASS RULE:
Hiring-team priorities from substantive Internal Notes drive emphasis first.
Exact JD language remains useful as secondary Fieldglass matching/ranking
coverage ONLY when the same language/capability is genuinely supported by the
resume. Truth outranks both sources.

{error_section}

OUTPUT SCHEMA:
{{
  "SUMMARY1": "",
  "SUMMARY1_EVIDENCE": [""],
  "SUMMARY2": "",
  "SUMMARY2_EVIDENCE": [""],
  "SUMMARY3": "",
  "SUMMARY3_EVIDENCE": [""],
  "SUMMARY4": "",
  "SUMMARY4_EVIDENCE": [""],

  "SKILL1": "",
  "SKILL1_EVIDENCE": [""],
  "SKILL1_ROLE_EVIDENCE": [""],
  "SKILLDETAILS1": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL2": "",
  "SKILL2_EVIDENCE": [""],
  "SKILL2_ROLE_EVIDENCE": [""],
  "SKILLDETAILS2": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL3": "",
  "SKILL3_EVIDENCE": [""],
  "SKILL3_ROLE_EVIDENCE": [""],
  "SKILLDETAILS3": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL4": "",
  "SKILL4_EVIDENCE": [""],
  "SKILL4_ROLE_EVIDENCE": [""],
  "SKILLDETAILS4": [
    {{"item": "", "evidence": ""}}
  ]
}}

DRAFT TO AUDIT/CORRECT:
{json.dumps(draft_data, indent=2, ensure_ascii=False)}

OFFICIAL FIELDGLASS JOB DESCRIPTION:
{jd_text}

INTERNAL NOTES:
{internal_notes if str(internal_notes or "").strip() else "No internal notes were provided."}

ORIGINAL RESUME:
{resume_text}
"""



def process_navy_federal_word_doc(temp_path, mapping, out_path):
    """
    Navy Federal Word processor.

    Uses the same expandable work-history mechanics as Freddie so resumes with
    more than seven roles remain complete, while adding Navy-specific education
    completion handling for the template's literal Yes/No column.
    """
    doc = docx.Document(temp_path)
    expand_experience_placeholders(doc, mapping)
    ensure_experience_spacing(doc, mapping)

    # ================================================================
    # 1. CERTIFICATIONS
    # ================================================================
    certifications = str(mapping.get("Certifications", "") or "").strip()

    if not certifications:
        cert_table = fred_find_table_containing(doc, "{{CERTIFICATIONS}}")
        if cert_table is not None:
            fred_remove_table(cert_table)

    # ================================================================
    # 2. EDUCATION + COMPLETION STATUS
    # ================================================================
    has_any_education = any(
        str(mapping.get(f"School{i}", "") or "").strip()
        for i in range(1, 4)
    )

    education_table = fred_find_table_containing(doc, "{{School1}}")

    if education_table is not None:
        if not has_any_education:
            fred_remove_table(education_table)
        else:
            rows_to_delete = []

            for row in education_table.rows:
                row_text = " ".join(cell.text for cell in row.cells).lower()

                for i in range(1, 4):
                    school_tag = f"{{{{school{i}}}}}".lower()

                    if school_tag not in row_text:
                        continue

                    if not str(mapping.get(f"School{i}", "") or "").strip():
                        rows_to_delete.append(row)
                    elif len(row.cells) >= 3:
                        navy_set_cell_text_preserve_format(
                            row.cells[2],
                            mapping.get(f"Completed{i}", ""),
                        )

                    break

            for row in rows_to_delete:
                try:
                    education_table._tbl.remove(row._tr)
                except Exception:
                    pass

    # ================================================================
    # 3. SKILLS — TWO-COLUMN SKILL AREA / DETAILS TABLE
    # ================================================================
    skills_table = fred_find_table_containing(doc, "{{SKILL1}}")

    if skills_table is not None:
        rows_to_delete = []

        for row in skills_table.rows:
            row_text = " ".join(cell.text for cell in row.cells).lower()

            for i in range(1, 5):
                if (
                    f"{{{{skill{i}}}}}".lower() in row_text
                    and not str(mapping.get(f"SKILL{i}", "") or "").strip()
                ):
                    rows_to_delete.append(row)
                    break

        for row in rows_to_delete:
            try:
                skills_table._tbl.remove(row._tr)
            except Exception:
                pass

    # ================================================================
    # 4. REPLACE TAGS IN BODY PARAGRAPHS
    # ================================================================
    paras_to_remove = []

    kill_keys = [
        "Company",
        "Title",
        "Bullets",
        "Dates",
        "Summary",
        "Skill",
        "Certification",
        "School",
        "Degree",
    ]

    for p in list(doc.paragraphs):
        for key, value in mapping.items():
            tag = f"{{{{{key}}}}}"

            if tag.lower() not in p.text.lower():
                continue

            if not value or not str(value).strip():
                if any(k.lower() in key.lower() for k in kill_keys):
                    if p not in paras_to_remove:
                        paras_to_remove.append(p)
                else:
                    replace_tag_safely(p, tag, "")
                continue

            if "Bullets" in key:
                for run in p.runs:
                    if "\n" in run.text:
                        run.text = run.text.replace("\n", "")

                lines = [
                    line.strip()
                    for line in str(value).split("\n")
                    if line.strip()
                ]

                if not lines:
                    if p not in paras_to_remove:
                        paras_to_remove.append(p)
                    continue

                replace_tag_safely(p, tag, lines[0])
                curr_p = p

                for line in lines[1:]:
                    curr_p = insert_bullet_line_after(curr_p, line)

                spacer = curr_p.insert_paragraph_before("")
                curr_p._p.addnext(spacer._p)

                try:
                    spacer.style = doc.styles["Normal"]
                except Exception:
                    pass

                spacer.paragraph_format.space_after = Pt(0)
                spacer.paragraph_format.space_before = Pt(0)

            else:
                replace_tag_safely(p, tag, str(value).strip())

    # ================================================================
    # 5. REPLACE TAGS INSIDE TABLE CELLS
    # ================================================================
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    for key, value in mapping.items():
                        tag = f"{{{{{key}}}}}"

                        if tag.lower() in p.text.lower():
                            replace_tag_safely(
                                p,
                                tag,
                                str(value) if value else "",
                            )

    # ================================================================
    # 6. DELETE EMPTY WORK-HISTORY PARAGRAPHS
    # ================================================================
    for p in paras_to_remove:
        try:
            delete_paragraph(p)
        except Exception:
            pass

    collapse_extra_blank_paragraphs(doc)
    doc.save(out_path)
    return out_path


def navy_federal_app():
    TEMPLATE_FILENAME = "Navy_Federal.docx"

    st.title("Navy Federal Precision Extractor")

    # ================================================================
    # API CONFIGURATION
    # ================================================================
    with st.sidebar:
        st.header("🔑 API Configuration")

        if "API_KEY" in st.secrets:
            API_KEY = st.secrets["API_KEY"]
            st.success("✅ API Key loaded from Secrets")
        else:
            API_KEY = st.text_input(
                "Gemini API Key",
                type="password",
                key="navy_api_key",
            )
            st.info("Paste your Gemini API key here to run the tool.")

    # ================================================================
    # CANDIDATE INFORMATION
    # ================================================================
    st.header("📋 Candidate Information")

    col1, col2 = st.columns(2)

    with col1:
        Full_Legal_Name = st.text_input(
            "Full Legal Name",
            key="navy_full_name",
        )

        Work_Location = st.text_input(
            "Work Location",
            key="navy_location",
        )

    with col2:
        Work_Arrangement = st.selectbox(
            "Hybrid/Onsite/Remote",
            ["Hybrid", "Onsite", "Remote"],
            index=0,
            key="navy_remote",
        )

        Work_Authorization = st.selectbox(
            "Work Authorization",
            [
                "Authorized to work in the U.S. and will not require visa sponsorship now or in the future",
                "Authorized to work in the U.S. and may require visa sponsorship in the future",
            ],
            index=0,
            key="navy_work_auth",
        )

    # ================================================================
    # NAVY FEDERAL REQUISITION
    # ================================================================
    st.header("📝 Navy Federal Requisition")

    SHEET_URL = (
        "https://docs.google.com/spreadsheets/d/"
        "1swKaDhGdlmO0ILv7tngkRCbsK-7jF89cdCpf5bypzwM/"
        "edit?usp=sharing"
    )

    @st.cache_data(ttl=600)
    def load_navy_federal_reqs(url):
        """Load the Navy Federal worksheet from the shared req tracker."""
        # First use the same public-sheet approach already used by Fannie/Deloitte.
        try:
            base_url = url.split("/edit")[0]
            csv_url = (
                base_url
                + "/gviz/tq?tqx=out:csv&sheet=Navy%20Federal"
            )
            df = pd.read_csv(csv_url)
            return df.fillna("")
        except Exception:
            pass

        # Fallback to the configured Streamlit Google Sheets connection.
        try:
            conn = st.connection(
                "gsheets",
                type=GSheetsConnection,
            )
            df = conn.read(
                spreadsheet=url,
                worksheet="Navy Federal",
                ttl=600,
            )
            return df.fillna("")
        except Exception as e:
            st.error(
                "❌ Unable to load the Navy Federal requisition sheet. "
                f"{str(e)}"
            )
            return pd.DataFrame()

    df_reqs = load_navy_federal_reqs(SHEET_URL)

    options = ["None"]
    req_mapping = {}

    if not df_reqs.empty:
        for _, row in df_reqs.iterrows():
            req_id = str(row.get("ID", "")).replace(".0", "").strip()
            vms_id = str(row.get("VMS ID", "")).replace(".0", "").strip()
            title = str(row.get("Job Title", "")).strip()
            manager = str(row.get("Manager", "")).strip()
            status = str(row.get("Status", "")).strip()
            client_company = str(row.get("Client Company", "")).strip()
            description = str(row.get("Description", "") or "").strip()
            published_description = str(
                row.get("Published Description", "") or ""
            ).strip()

            if not any([req_id, vms_id, title]):
                continue

            option_parts = [req_id, vms_id, title, manager]
            option_str = " - ".join(
                part for part in option_parts if part
            )

            if option_str and option_str != "-":
                options.append(option_str)
                req_mapping[option_str] = {
                    "id": req_id,
                    "vms_id": vms_id,
                    "job_title": title,
                    "manager": manager,
                    "client_company": client_company,
                    "status": status,
                    "date_added": str(row.get("Date Added", "") or "").strip(),
                    "published_description": published_description,
                    # Per Annabelle's workflow, Column I / Description is the
                    # official Fieldglass job description used by the formatter.
                    "description": description,
                }

    if "navy_jd" not in st.session_state:
        st.session_state.navy_jd = ""

    if "navy_internal_notes" not in st.session_state:
        st.session_state.navy_internal_notes = ""

    def update_navy_jd_text():
        selected = st.session_state.get("navy_req_selector", "None")

        if selected != "None" and selected in req_mapping:
            st.session_state.navy_jd = req_mapping[selected].get(
                "description",
                "",
            ).strip()
            # Internal Notes are intentionally recruiter-entered only for now.
            st.session_state.navy_internal_notes = ""
        else:
            st.session_state.navy_jd = ""
            st.session_state.navy_internal_notes = ""

    selected_req = st.selectbox(
        "🔍 Select a Navy Federal Requisition (Type to search):",
        options=options,
        key="navy_req_selector",
        on_change=update_navy_jd_text,
    )

    selected_req_context = (
        req_mapping.get(selected_req, {})
        if selected_req != "None"
        else {}
    )

    Job_Description = st.text_area(
        "1. Official Fieldglass Job Description:",
        height=175,
        key="navy_jd",
    )

    Internal_Notes = st.text_area(
        "2. Internal Notes (Optional - Recruiter Context):",
        height=100,
        key="navy_internal_notes",
    )

    # ================================================================
    # FILE UPLOAD
    # ================================================================
    st.header("📂 File Uploads")

    resume_file = st.file_uploader(
        "📤 Upload the candidate's resume...",
        type=["pdf", "docx", "doc"],
        key="navy_res",
    )

    if st.button(
        "🚀 Generate Navy Federal Submission",
        type="primary",
        key="navy_generate",
    ):
        validation_errors = []

        if not API_KEY:
            validation_errors.append("Enter the Gemini API Key.")

        if not Full_Legal_Name.strip():
            validation_errors.append("Enter the candidate's Full Legal Name.")

        if not Work_Location.strip():
            validation_errors.append("Enter the candidate's Work Location.")

        if not Job_Description.strip():
            validation_errors.append(
                "Select a Navy Federal requisition or paste the Official Fieldglass Job Description."
            )

        if not resume_file:
            validation_errors.append("Upload the candidate's resume.")

        if not os.path.exists(TEMPLATE_FILENAME):
            validation_errors.append(
                f"The template '{TEMPLATE_FILENAME}' was not found."
            )

        if validation_errors:
            for error in validation_errors:
                st.error(f"❌ {error}")
            return

        with st.spinner("Building hiring-manager-focused Navy Federal submission..."):
            resume_path = None

            try:
                safe_original_name = os.path.basename(resume_file.name)
                resume_path = f"temp_navy_{safe_original_name}"

                with open(resume_path, "wb") as f:
                    f.write(resume_file.getbuffer())

                raw_text = extract_text(resume_path)
                current_date_obj = date.today()
                current_date = current_date_obj.isoformat()

                # ====================================================
                # PASS 1 — FACTUAL RESUME EXTRACTION
                # ====================================================
                extraction_prompt = f"""
Return a valid JSON object ONLY.

You are extracting a candidate's ORIGINAL resume into a clean structured
format for a Navy Federal Credit Union Fieldglass submission.

CURRENT DATE:
{current_date}

CRITICAL:
- This is an extraction step, NOT a tailoring step.
- Preserve the candidate's factual work history.
- Do NOT add job-description terminology.
- Do NOT rewrite accomplishments to sound more relevant.
- Do NOT invent technologies, metrics, employers, dates, responsibilities,
  education completion, or certifications.

RULES:

1. NAME
- Extract the candidate name from the resume header when available.
- The user-entered Full Legal Name will control the final template, so do not
  alter or "correct" that user-entered value.

2. COMPANY
- Extract the primary employer/end-client actually represented by the resume.
- Strip geographic locations that are merely appended to the company line.
- Do not guess an end-client that is not stated.

3. EDUCATION
For every education entry return:
- School
- Degree
- CompletionStatus: ONLY "Yes", "No", or ""
- CompletionDate: graduation/completion/expected-completion date or year ONLY;
  do not put attendance-start dates here.

COMPLETION RULES:
- "Yes" only when the resume explicitly indicates the degree/credential was
  completed/earned/graduated OR provides a completion/graduation year that is
  not in the future.
- "No" when the resume says In Progress, Pursuing, Expected, Anticipated,
  Pending, Currently Enrolled, or otherwise clearly unfinished, or when the
  completion date/year is in the future.
- If completion cannot be established from the resume, return "" rather than
  guessing Yes.

4. EXPERIENCE
Return EVERY dated professional role shown anywhere in the resume, including
condensed/abbreviated entries under Earlier Experience or similar sections.
Do not stop after seven roles.

For every role return:
- Company
- Title
- Bullets as an ARRAY
- Dates

TITLE:
- Remove only obvious employment-type suffixes such as "(Contract)" or
  "- Contract" when they are not part of the actual title.
- Do not upgrade seniority.

BULLETS:
- Preserve every substantive original bullet in original order.
- Light cleanup of malformed bullet characters/spacing is allowed.
- Do NOT rewrite bullets to match the Navy Federal JD.

DATES:
- Preserve the source date precision.
- NEVER invent a month when the source provides only a year.
- Mixed-precision ranges must stay mixed.

5. CERTIFICATIONS
- Extract certifications into one comma-separated string.
- Preserve explicit In Progress / Pursuing / Pending status rather than
  presenting unfinished certifications as completed.
- If none are found, return "".

OUTPUT:
{{
  "FullName": "",
  "Certifications": "",
  "Education": [
    {{
      "School": "",
      "Degree": "",
      "CompletionStatus": "",
      "CompletionDate": ""
    }}
  ],
  "Experience": [
    {{
      "Company": "",
      "Title": "",
      "Bullets": [""],
      "Dates": ""
    }}
  ]
}}

FILENAME:
{resume_file.name}

ORIGINAL RESUME:
{raw_text}
"""

                data = navy_generate_json(
                    API_KEY,
                    extraction_prompt,
                )

                certifications = str(
                    data.get("Certifications", "") or ""
                ).strip()

                education = data.get("Education", [])
                if not isinstance(education, list):
                    education = []

                experience = data.get("Experience", [])
                if not isinstance(experience, list):
                    experience = []

                normalized_experience = []

                for role in experience:
                    if not isinstance(role, dict):
                        continue

                    bullets = role.get("Bullets", [])
                    if isinstance(bullets, str):
                        bullets = [
                            line.strip()
                            for line in bullets.split("\n")
                            if line.strip()
                        ]
                    if not isinstance(bullets, list):
                        bullets = []

                    clean_role = {
                        "Company": str(role.get("Company", "") or "").strip(),
                        "Title": re.sub(
                            r"\s*(?:\(|-|–|—)\s*(?:contract|consultant|contractor)\b.*$",
                            "",
                            str(role.get("Title", "") or "").strip(),
                            flags=re.IGNORECASE,
                        ).strip(),
                        "Bullets": [
                            re.sub(
                                r"^[\s\-\•\*\d\.\)]+",
                                "",
                                str(item),
                            ).strip()
                            for item in bullets
                            if str(item).strip()
                        ],
                        "Dates": standardize_dates(
                            str(role.get("Dates", "") or "").strip()
                        ),
                    }

                    if any(
                        clean_role.get(key)
                        for key in ["Company", "Title", "Dates", "Bullets"]
                    ):
                        normalized_experience.append(clean_role)

                experience = normalized_experience

                # ====================================================
                # PASS 2 — NAVY FEDERAL HIRING-MANAGER SUMMARY + SKILLS OPTIMIZATION
                # ====================================================
                structured_experience = json.dumps(
                    experience,
                    indent=2,
                    ensure_ascii=False,
                )

                tailoring_prompt = f"""
Return a valid JSON object ONLY.

You are a Senior Recruiter preparing a Navy Federal Credit Union
contingent-worker resume submission through Fieldglass.

CURRENT DATE:
{current_date}

There are TWO optimization goals, but they are NOT equal:

1. PRIMARY — NAVY FEDERAL HIRING-MANAGER INTERVIEW SELECTION
The MSP appears to pass submitted candidates through rather than use shortlisting
as the main gate. Optimize the four Summary bullets and four Skills rows primarily
to make the hiring manager want to interview the candidate.

2. SECONDARY — FIELDGLASS DISCOVERABILITY / POSSIBLE RANKING OR SCORING
The exact downstream Fieldglass AI behavior is unknown. Preserve useful exact JD
terminology when the resume genuinely supports it so the candidate remains easy
to recognize in any automated matching/ranking layer. Never chase JD keywords at
the expense of current hiring-team priorities.

SOURCE HIERARCHY FOR ROLE PRIORITIES:
1. INTERNAL NOTES — highest priority when substantive. Treat supplier-call,
   hiring-manager, MSP, recruiter, or account-team notes about must-haves,
   prior-candidate failures, interview criteria, real day-to-day work,
   technical/business weighting, preferred backgrounds, or de-emphasized JD
   items as the most current signal of what will drive an interview decision.
2. OFFICIAL FIELDGLASS JOB DESCRIPTION — baseline responsibilities,
   qualifications, technical/domain terminology, and secondary Fieldglass
   matching/ranking coverage.

If Internal Notes clarify, narrow, reorder, or contradict the JD, use the Notes
to decide EMPHASIS. Do not erase truthful relevant JD coverage unnecessarily,
but do not let generic/stale JD language outrank current hiring-team guidance.

SOURCE HIERARCHY FOR CANDIDATE FACTS:
1. ORIGINAL RESUME ONLY.

THE JD AND INTERNAL NOTES ARE NEVER CANDIDATE EVIDENCE. They tell you WHAT
MATTERS, not WHAT THE CANDIDATE HAS DONE.

If Internal Notes identify a priority that is absent from the JD, you MAY still
feature it prominently in the Summary or Skills when the uploaded resume proves
it. Example: if Notes say Tableau is the #1 priority and the resume proves
Tableau, emphasize Tableau even if the JD does not mention it. If the resume does
not prove Tableau, do not assign Tableau to the candidate.

======================================================================
HARD SOURCE FIREWALL — ZERO TOLERANCE
======================================================================

Before writing ANY candidate claim, mentally ask:
"Could I point to the uploaded resume and show the exact experience that proves
this?"

If NO, DO NOT WRITE IT.

This applies to:
- skills
- tools
- frameworks
- regulations
- methodologies
- responsibilities
- ownership
- scope
- metrics
- titles
- domains
- stakeholder types
- governance models
- issue/risk terminology
- audit/compliance terminology

NEVER upgrade a related concept into the JD's more specific concept.

UNIVERSAL EXAMPLES ACROSS NAVY FEDERAL JOB TYPES:
- A JD requests a named tool, platform/module, framework, regulation, programming
  language, design system, equipment type, or methodology that the resume does
  not show -> DO NOT assign it to the candidate.
- A JD says "lead", "owned", "managed", "independent", "enterprise-wide",
  "advanced", "expert", "high-volume", "end-to-end", "production", or another
  stronger scope/proficiency/authority qualifier -> use that qualifier ONLY when
  the resume evidence independently supports that level of claim.
- "Supported", "assisted", "participated", or "worked with" does NOT automatically
  become "led", "owned", "managed", "architected", or "independently performed".
- A broad platform does NOT prove a specific module or product feature.
- A general technical/business capability does NOT prove a formal discipline,
  certification, regulatory framework, governance model, or specialized method.
- A candidate can be highly relevant without matching every JD term. Use the
  strongest truthful resume-native wording instead of upgrading the claim.

HIGH-IMPACT QUALIFIER RULE — ALL JOB FAMILIES:
Treat qualifiers that change LEVEL, AUTHORITY, SCOPE, PROFICIENCY, SCALE, or
FORMAL RESPONSIBILITY as candidate facts in their own right. Examples include
independent/independently, lead/led/leading, own/owned/ownership,
manage/managed/manager, oversight, architect/architected/architecture, senior, principal,
expert/expertise, advanced, extensive, deep, strong, hands-on,
cross-functional, enterprise-wide, organization-wide, global, strategic,
regulatory, enforcement, validation, approval, high-volume, large-scale,
mission-critical, end-to-end, full lifecycle, production, and formal
line-of-defense terminology.

Every such qualifier used in a Summary or Skills claim must be supported by the
resume evidence cited for THAT SAME claim. Do not let a JD adjective or authority
verb hitchhike onto an otherwise supported noun phrase.

IMPORTANT MANAGEMENT DISTINCTION:
- Verbs/titles that assert authority (managed, managing, manager) require direct
  resume support.
- The standalone noun "management" is often the name of a legitimate discipline,
  not an authority claim. Do NOT treat Data Management, Project Management, Risk
  Management, Change Management, Content Management, Metadata Management, Vendor
  Management, or similar domain labels as people-management/ownership claims.
  They may be used when the resume supports that discipline.

When the JD contains a term the resume does not support, simply omit that term.
An honest partial match is better than a fabricated exact match.

======================================================================
SUMMARY — EXACTLY 4 BULLET POINTS
======================================================================

Create exactly four concise resume bullets mapped to SUMMARY1-SUMMARY4.

STYLE:
- Do NOT start with the candidate's name.
- Do NOT use the word "progressive" anywhere.
- Do NOT use "progressive experience", "progressive professional experience",
  "seasoned professional", "highly experienced", "strong background",
  "excellent candidate", "great fit", "aligns well", or similar AI/recruiter
  filler.
- When years can be calculated responsibly, use natural phrasing such as
  "10+ years of experience..." or "10+ years in..." — never
  "10 years of progressive experience".
- One strong sentence per bullet.
- No pronouns unless unavoidable.
- Do not mention location, work authorization, or availability.

BULLET 1 — EXPERIENCE ANCHOR
- Calculate total professional experience ONLY from dated work history.
- Round DOWN to completed whole years and format as "X+ years".
- Pair that with the candidate's most relevant resume-proven functional
  identity/expertise for this requisition.
- If dates are too ambiguous, omit the years rather than guessing.

BULLET 2 — PRIMARY HIRING-TEAM NEED
- Select the strongest recent/recurring resume responsibility that proves the
  #1 actual hiring-team priority.
- When substantive Internal Notes identify a must-have, screening failure, or
  day-to-day priority, use that before a lower-priority generic JD item.
- Use only the scope and terminology the resume supports.

BULLET 3 — EXECUTION / CONTROLS / TOOLS / IMPACT
- Show how the candidate performed relevant work.
- Use only tools, frameworks, methodologies, processes, scale, and outcomes
  actually supported by the resume.
- Prefer repeated operational ownership over isolated projects.

BULLET 4 — SECONDARY DIFFERENTIATOR
- Use the strongest additional resume-proven requirement, domain experience,
  stakeholder capability, process expertise, or complementary technical skill.
- Do NOT convert a target-role responsibility into past candidate experience.

For EACH bullet, also return SUMMARY#_EVIDENCE containing 1-3 short VERBATIM
excerpts copied from the ORIGINAL RESUME that support every substantive claim in
the bullet. Evidence must be exact resume wording, not a paraphrase.

======================================================================
SKILL & COMPETENCY REQUIREMENTS — EXACTLY 4 ROWS
======================================================================

SKILL# — LEFT COLUMN
- Concise high-level competency/category.
- Choose categories that matter to the JD AND are substantially demonstrated by
  the resume.
- The category may be a natural umbrella label summarizing multiple resume
  bullets; it does not need to appear verbatim in the resume.
- Do not borrow a named framework, regulation, tool, or formal JD discipline
  merely because it is adjacent to the candidate's background.
- Return SKILL#_EVIDENCE with 1-3 short VERBATIM resume excerpts proving the
  category.
- Return SKILL#_ROLE_EVIDENCE with 1-2 short VERBATIM excerpts proving why this
  competency deserves one of the four slots. Each excerpt may come from either
  substantive Internal Notes OR the Official Fieldglass JD.
- Each ROLE_EVIDENCE excerpt must be ONE CONTIGUOUS source clause/sentence copied
  exactly. Do NOT use ellipses, delete intervening words, combine separate source
  fragments, or paraphrase the quote. If a source sentence is long, select a
  shorter contiguous clause that still proves the priority.
- Prefer Internal Notes when they directly state a stronger/current must-have,
  screening problem, interview criterion, or day-to-day priority.
- Role-priority evidence is for prioritization ONLY; it never proves candidate
  experience.

FOUR-SLOT PRIORITY RULE — CRITICAL:
- These four rows are scarce real estate. Do NOT return merely four truthful
  competencies. Return the FOUR MOST INTERVIEW-RELEVANT resume-supported
  competencies for the actual hiring team's current priorities.
- First rank substantive Internal-Notes must-haves, prior-candidate failures,
  interview criteria, day-to-day needs, and technical/business weighting.
- Then rank the Official JD's core responsibilities and explicit must-haves.
- Map those priorities against all resume-supported candidate competencies and
  select the top four DISTINCT intersections.
- A resume-supported hiring-team priority missing from the JD MAY take a Skills
  row. Core manager priorities outrank generic JD language and nice-to-haves.
- Preserve useful JD overlap where truthful so possible Fieldglass ranking or
  scoring still sees the candidate's genuine baseline match.

SKILLDETAILS# — RIGHT COLUMN
- Return a JSON ARRAY of 3-6 objects.
- Each object MUST have:
  "item": a concise Tool / Technology / Methodology / Process / domain method
  "evidence": a VERBATIM resume excerpt explicitly supporting that exact item
- Every item must be supported by its own evidence.
- USE RESUME-NATIVE WORDING. Never rename a resume concept to a more attractive
  JD term unless the resume itself supports that JD term.
- KEEP EACH DETAIL ATOMIC: one evidence excerpt must support the full item. Split
  combined concepts into separate items when different bullets are needed.
- Do NOT combine concepts with "&", "/", "and", or commas unless ONE cited
  resume excerpt supports the entire combined claim. Individually supported words
  may not be fused into a stronger claim. Example: "SOC 2 and SOX audits" plus
  "SOC 1 and 2 reports" does NOT equal "SOC 1 & SOC 2 Audits".
- Do not use a JD-only skill.
- Do not use years used / last used.
- Do not repeat the left-column category wording in the right column.

OPERATIONAL OWNERSHIP RULE:
Prefer recurring responsibilities and direct ownership over isolated projects.
Unless the role is explicitly implementation-focused, favor day-to-day
execution, issue/problem resolution, reporting, controls, compliance,
operations, stakeholder support, and measurable outcomes.

FINAL SELF-AUDIT BEFORE OUTPUT:
1. Is every framework/regulation/tool in the output present or directly
   demonstrated in the resume?
2. Did any phrase sneak in only because it appeared in the JD?
3. Did I accidentally convert "related experience" into a more specific formal
   discipline?
4. Would every candidate claim survive a side-by-side resume audit?
5. Is every evidence excerpt copied verbatim from the resume?
6. Did I avoid the word "progressive"?
7. Did substantive Internal Notes drive emphasis before generic JD language?
8. Are these truly the FOUR BEST INTERVIEW-DRIVING resume-supported competency
   rows for the current hiring-team priorities, while retaining truthful JD
   coverage for possible Fieldglass ranking/scoring?
9. Does every right-column detail describe exactly what its ONE cited resume
   excerpt proves, without combining separate facts into a stronger claim?
10. Does every high-impact qualifier (lead/owned/managed/independent/advanced/
    expert/enterprise-wide/high-volume/end-to-end/production/etc.) have direct
    support in the evidence cited for that SAME claim?

OUTPUT:
{{
  "SUMMARY1": "",
  "SUMMARY1_EVIDENCE": [""],
  "SUMMARY2": "",
  "SUMMARY2_EVIDENCE": [""],
  "SUMMARY3": "",
  "SUMMARY3_EVIDENCE": [""],
  "SUMMARY4": "",
  "SUMMARY4_EVIDENCE": [""],

  "SKILL1": "",
  "SKILL1_EVIDENCE": [""],
  "SKILL1_ROLE_EVIDENCE": [""],
  "SKILLDETAILS1": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL2": "",
  "SKILL2_EVIDENCE": [""],
  "SKILL2_ROLE_EVIDENCE": [""],
  "SKILLDETAILS2": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL3": "",
  "SKILL3_EVIDENCE": [""],
  "SKILL3_ROLE_EVIDENCE": [""],
  "SKILLDETAILS3": [
    {{"item": "", "evidence": ""}}
  ],

  "SKILL4": "",
  "SKILL4_EVIDENCE": [""],
  "SKILL4_ROLE_EVIDENCE": [""],
  "SKILLDETAILS4": [
    {{"item": "", "evidence": ""}}
  ]
}}

OFFICIAL FIELDGLASS JOB DESCRIPTION:
{Job_Description}

INTERNAL NOTES (OPTIONAL RECRUITER CONTEXT):
{Internal_Notes if Internal_Notes.strip() else "No internal notes were provided."}

STRUCTURED ORIGINAL EXPERIENCE:
{structured_experience}

FULL ORIGINAL RESUME:
{raw_text}
"""

                # Pass 2A: recruiter-written draft with explicit resume evidence.
                draft_data = navy_generate_json(
                    API_KEY,
                    tailoring_prompt,
                )

                # Pass 2B: independent fact-checking pass. This pass is not asked
                # to sell the candidate; it is asked to remove/rewrite unsupported
                # claims and provide exact resume evidence for everything retained.
                audited_data = navy_generate_json(
                    API_KEY,
                    navy_grounding_audit_prompt(
                        resume_text=raw_text,
                        jd_text=Job_Description,
                        internal_notes=Internal_Notes,
                        draft_data=draft_data,
                        current_date=current_date,
                    ),
                )

                grounding_errors = navy_validate_grounded_output(
                    raw_text,
                    Job_Description,
                    Internal_Notes,
                    audited_data,
                )

                if grounding_errors:
                    # One automatic repair pass using the deterministic validator's
                    # exact failures. If this still cannot be proven from the resume,
                    # fail closed instead of generating a potentially fabricated file.
                    audited_data = navy_generate_json(
                        API_KEY,
                        navy_grounding_audit_prompt(
                            resume_text=raw_text,
                            jd_text=Job_Description,
                            internal_notes=Internal_Notes,
                            draft_data=audited_data,
                            current_date=current_date,
                            validator_errors=grounding_errors,
                        ),
                    )

                    grounding_errors = navy_validate_grounded_output(
                        raw_text,
                        Job_Description,
                        Internal_Notes,
                        audited_data,
                    )

                if grounding_errors:
                    raise RuntimeError(
                        "The Navy Federal grounding safety check could not verify "
                        "all generated Summary/Skills content against the uploaded "
                        "resume, so no document was created. The formatter already "
                        "attempted an automatic correction pass. Verification "
                        "issues: "
                        + " | ".join(grounding_errors[:8])
                    )

                summary_data = navy_finalize_grounded_output(
                    audited_data
                )

                missing_summary = [
                    str(i)
                    for i in range(1, 5)
                    if not str(
                        summary_data.get(f"SUMMARY{i}", "") or ""
                    ).strip()
                ]

                missing_skills = [
                    str(i)
                    for i in range(1, 5)
                    if not str(
                        summary_data.get(f"SKILL{i}", "") or ""
                    ).strip()
                    or not str(
                        summary_data.get(f"SKILLDETAILS{i}", "") or ""
                    ).strip()
                ]

                if missing_summary:
                    raise RuntimeError(
                        "The Navy Federal verified optimization step did not return "
                        "all four required Summary bullets. Missing: "
                        + ", ".join(missing_summary)
                    )

                if missing_skills:
                    raise RuntimeError(
                        "The Navy Federal verified optimization step did not return "
                        "all four required Skills rows. Missing: "
                        + ", ".join(missing_skills)
                    )

                # ====================================================
                # BUILD WORD MAPPING
                # ====================================================
                mapping = {
                    "FullName": Full_Legal_Name.strip(),
                    "Location": Work_Location.strip(),
                    "Remote": Work_Arrangement,
                    "WorkAuthorization": Work_Authorization,
                    "Certifications": certifications,
                    "SUMMARY1": str(summary_data.get("SUMMARY1", "") or "").strip(),
                    "SUMMARY2": str(summary_data.get("SUMMARY2", "") or "").strip(),
                    "SUMMARY3": str(summary_data.get("SUMMARY3", "") or "").strip(),
                    "SUMMARY4": str(summary_data.get("SUMMARY4", "") or "").strip(),
                    "SKILL1": str(summary_data.get("SKILL1", "") or "").strip(),
                    "SKILLDETAILS1": str(summary_data.get("SKILLDETAILS1", "") or "").strip(),
                    "SKILL2": str(summary_data.get("SKILL2", "") or "").strip(),
                    "SKILLDETAILS2": str(summary_data.get("SKILLDETAILS2", "") or "").strip(),
                    "SKILL3": str(summary_data.get("SKILL3", "") or "").strip(),
                    "SKILLDETAILS3": str(summary_data.get("SKILLDETAILS3", "") or "").strip(),
                    "SKILL4": str(summary_data.get("SKILL4", "") or "").strip(),
                    "SKILLDETAILS4": str(summary_data.get("SKILLDETAILS4", "") or "").strip(),
                }

                # Education — never default completion to Yes.
                for i in range(1, 4):
                    if i <= len(education) and isinstance(education[i - 1], dict):
                        edu_item = education[i - 1]

                        mapping[f"School{i}"] = clean_school(
                            str(edu_item.get("School", "") or "")
                        )
                        mapping[f"Degree{i}"] = str(
                            edu_item.get("Degree", "") or ""
                        ).strip()
                        mapping[f"Completed{i}"] = navy_normalize_completion_status(
                            edu_item,
                            current_year=current_date_obj.year,
                        )
                    else:
                        mapping[f"School{i}"] = ""
                        mapping[f"Degree{i}"] = ""
                        mapping[f"Completed{i}"] = ""

                # Work History — preserve EVERY extracted role and rely on the
                # shared dynamic template expander for Company8/9/etc.
                for i in range(1, max(7, len(experience)) + 1):
                    if i <= len(experience):
                        role = experience[i - 1]
                        mapping[f"Company{i}"] = str(
                            role.get("Company", "") or ""
                        ).strip()
                        mapping[f"Title{i}"] = str(
                            role.get("Title", "") or ""
                        ).strip()
                        mapping[f"Bullets{i}"] = clean_bullets(
                            role.get("Bullets", [])
                        )
                        mapping[f"Dates{i}"] = str(
                            role.get("Dates", "") or ""
                        ).strip()
                    else:
                        mapping[f"Company{i}"] = ""
                        mapping[f"Title{i}"] = ""
                        mapping[f"Bullets{i}"] = ""
                        mapping[f"Dates{i}"] = ""

                safe_name = (
                    re.sub(
                        r"[^A-Za-z0-9_\-]+",
                        "_",
                        Full_Legal_Name.strip(),
                    ).strip("_")
                    or "Candidate"
                )

                out_file = f"Submission_NavyFederal_{safe_name}.docx"

                process_navy_federal_word_doc(
                    TEMPLATE_FILENAME,
                    mapping,
                    out_file,
                )

                with open(out_file, "rb") as file:
                    st.download_button(
                        label="⬇️ Download Generated Document",
                        data=file,
                        file_name=out_file,
                        mime=(
                            "application/vnd.openxmlformats-"
                            "officedocument.wordprocessingml.document"
                        ),
                        type="primary",
                    )

                st.success("✅ Navy Federal submission is ready.")

            except Exception as e:
                st.error(
                    f"❌ Navy Federal processing failed: {str(e)}"
                )

            finally:
                if resume_path and os.path.exists(resume_path):
                    try:
                        os.remove(resume_path)
                    except Exception:
                        pass


# ====================================================================
# --- MAIN ROUTING LOGIC (The Dropdown Page) ---
# ====================================================================

st.title("Resume Formatter")
st.markdown("Please select your client account below to access the customized formatter.")

client_selection = st.selectbox(
    "Select Client Account or Tool:", 
    ["Fannie Mae", "Freddie Mac", "Navy Federal", "Deloitte", "Peraton", "Capital One", "ADUSA", "CBRE", "BNSF", "Dallas Generic", "PDF to Word"]
)

st.divider()

if client_selection == "Fannie Mae":
    fannie_mae_app()
elif client_selection == "Freddie Mac":
    freddie_mac_app()
elif client_selection == "Navy Federal":
    navy_federal_app()
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
