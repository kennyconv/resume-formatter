import streamlit as st
import google.generativeai as genai
import docx
import PyPDF2
import json

# --- Extraction Logic ---

def extract_text(uploaded_file):
    text = ""
    if uploaded_file.name.lower().endswith('.pdf'):
        pdf = PyPDF2.PdfReader(uploaded_file)
        for page in pdf.pages:
            text += page.extract_text() + "\n"
    else:
        doc = docx.Document(uploaded_file)
        text = "\n".join([p.text for p in doc.paragraphs])
    return text

def ai_extraction(raw_text, api_key):
    genai.configure(api_key=api_key)
    # Using a stable path to avoid 404 errors
    model = genai.GenerativeModel('gemini-1.5-flash')
    
    prompt = f"""
    You are a data extraction tool. Extract the following from the resume into JSON:
    1. Full Name
    2. Education: List of objects with "School" and "Degree"
    3. Experience: List of objects with "Company", "Title", and "Dates" (MMM YYYY - MMM YYYY or Current)

    JSON Structure:
    {{
        "FullName": "",
        "Education": [{{"School": "", "Degree": ""}}],
        "Experience": [{{"Company": "", "Title": "", "Dates": ""}}]
    }}

    RESUME TEXT:
    {raw_text}
    """
    
    response = model.generate_content(
        prompt,
        generation_config={"response_mime_type": "application/json"}
    )
    return json.loads(response.text)

# --- UI Layout ---
st.set_page_config(page_title="Resume Data Extractor", layout="wide")
st.title("📄 Resume Data Extractor")

# Sidebar for API Key
api_key = st.sidebar.text_input("Gemini API Key", type="password")

# Split the screen into two columns
left_col, right_col = st.columns(2)

with left_col:
    st.subheader("Data Entry & Upload")
    uploaded_file = st.file_uploader("Upload Resume (PDF or DOCX)", type=["pdf", "docx"])
    
    loc_input = st.text_input("Current Location (City, ST)")
    remote_input = st.selectbox("Remote or Onsite", ["Remote", "Onsite"])
    former_fm = st.selectbox("Former FM?", [
        "N", 
        "Y - Per CRC, this candidate is eligible for rehire"
    ])
    link_input = st.text_input("LinkedIn Profile/GitHub/Portfolio Link")
    
    st.write("---")
    st.markdown("**Supplier Technical Interview Results**")
    qa_pairs = []
    for i in range(1, 6):
        q = st.text_input(f"Question {i}", key=f"q{i}")
        a = st.text_area(f"Answer {i}", key=f"a{i}")
        if q or a:
            qa_pairs.append({"q": q, "a": a, "num": i})

    format_button = st.button("Extract & Format")

with right_col:
    st.subheader("Extracted Data")
    if format_button:
        if not api_key:
            st.error("Please enter an API Key in the sidebar.")
        elif not uploaded_file:
            st.error("Please upload a resume.")
        else:
            try:
                # 1. AI Extraction
                raw_resume_text = extract_text(uploaded_file)
                extracted_data = ai_extraction(raw_resume_text, api_key)
                
                # 2. Display Candidate Info
                st.text_input("Name:", value=extracted_data.get("FullName", "Not Found"))
                st.text_input("Current Location: (City and State only)", value=loc_input)
                st.text_input("Remote or Onsite:", value=remote_input)
                st.text_input("Former FM FTE or Contractor Y/N: (If yes, add CRC approval)", value=former_fm)
                st.text_input("LinkedIn Profile/GitHub/Portfolio Link", value=link_input)
                
                # 3. Display Education
                for idx, edu in enumerate(extracted_data.get("Education", []), 1):
                    st.text_input(f"School{idx}", value=edu.get("School", ""))
                    st.text_input(f"Degree{idx}", value=edu.get("Degree", ""))
                
                # 4. Display Work History
                for idx, exp in enumerate(extracted_data.get("Experience", []), 1):
                    st.text_input(f"Company{idx}", value=exp.get("Company", ""))
                    st.text_input(f"Title{idx}", value=exp.get("Title", ""))
                    st.text_input(f"Dates{idx}", value=exp.get("Dates", ""))
                
                # 5. Display Interview Results
                for pair in qa_pairs:
                    if pair["q"]:
                        st.text_input(f"Question {pair['num']}", value=pair["q"])
                    if pair["a"]:
                        st.text_input(f"Answer {pair['num']}", value=pair["a"])
                
                st.success("Extraction Complete!")
                
            except Exception as e:
                st.error(f"Error during extraction: {str(e)}")
    else:
        st.info("Upload a resume and click 'Extract & Format' to see results here.")
