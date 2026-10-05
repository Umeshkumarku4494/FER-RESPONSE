import os
import re
import io
import time
import zipfile
import difflib
from datetime import datetime
from dateutil.relativedelta import relativedelta
from typing import List, Optional

import streamlit as st
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

from pydantic import BaseModel, Field
from google import genai
from google.genai import types
from dotenv import load_dotenv, set_key

# Auto-load saved credentials
ENV_PATH = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path=ENV_PATH)

# ===========================================================================
# 1. Pydantic Schemas for Strict Structured JSON Output
# ===========================================================================

class FeatureComparisonRow(BaseModel):
    feature_name: str = Field(description="Limitation element of Claim 1")
    present_invention: str = Field(description="Disclosure in present Claim 1")
    d1_disclosure: str = Field(description="Disclosure in D1 or 'Not disclosed'")
    d2_disclosure: str = Field(description="Disclosure in D2 or 'Not disclosed'")
    d3_disclosure: Optional[str] = Field(default="Not disclosed", description="Disclosure in D3 or 'Not disclosed'")
    d4_disclosure: Optional[str] = Field(default="Not disclosed", description="Disclosure in D4 or 'Not disclosed'")


class ClaimElement(BaseModel):
    claim_number: int = Field(description="Claim number")
    original_text: str = Field(description="Verbatim text of as-filed claim")
    amended_text: str = Field(description="Amended text with reference numerals in parentheses")
    basis_in_spec: str = Field(description="Section 59(1) specification support basis")


class FERStructuredOutput(BaseModel):
    app_no: str
    filing_date: str
    fer_date: str
    controller_name: str
    examiner_name: str
    applicant_name: str
    invention_title: str
    brief_summary: str
    comparison_table: List[FeatureComparisonRow]
    inventive_step_argument: str
    sufficiency_argument: str
    clarity_argument: str
    amended_claims: List[ClaimElement]


# ===========================================================================
# 2. In-Memory Text Extraction
# ===========================================================================

def extract_text_from_upload(uploaded_file, max_pages: int = 6, max_chars: int = 18000) -> str:
    name = uploaded_file.name.lower()
    content = uploaded_file.read()
    uploaded_file.seek(0)

    text = ""
    if name.endswith(".pdf"):
        with fitz.open(stream=content, filetype="pdf") as doc:
            pages = min(len(doc), max_pages)
            text = "\n".join([doc[i].get_text("text") for i in range(pages)])
    elif name.endswith((".docx", ".doc")):
        doc = Document(io.BytesIO(content))
        text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    else:
        text = content.decode("utf-8", errors="ignore")

    return text[:max_chars]


# ===========================================================================
# 3. AI Analysis Engine (Adaptive Cloud Tier)
# ===========================================================================

def run_ai_pipeline(fer_txt: str, spec_txt: str, claims_txt: str, prior_arts_txt: str, api_key: str) -> FERStructuredOutput:
    client = genai.Client(api_key=api_key)
    candidate_models = [
        "gemini-flash-lite-latest",
        "gemini-2.5-flash-lite",
        "gemini-flash-latest",
        "gemini-3.8-flash"
    ]

    prompt = f"""
You are a registered Indian Patent Agent drafting an official First Examination Report (FER) response for the Indian Patent Office (IPO).

INPUT DATA:
=== FIRST EXAMINATION REPORT (FER) ===
{fer_txt}

=== AS-FILED COMPLETE SPECIFICATION ===
{spec_txt}

=== AS-FILED CLAIMS ===
{claims_txt}

=== CITED PRIOR ART DOCUMENTS ===
{prior_arts_txt}

MANDATES:
1. Extract metadata accurately: App No, Filing Date, FER Date, Controller Name, Examiner Name, Applicant, Title.
2. Brief Summary: Formulate technical summary highlighting novel structural combinations.
3. Section 2(1)(ja) - Inventive Step:
   - Break Claim 1 into key technical limitations.
   - Build a comparison table matching Claim 1 features against cited art (D1-D4).
   - Rebut obviousness and hindsight reconstruction.
4. Section 10(4) Sufficiency & Section 10(5) Scope/Clarity traversals.
5. Amend Claims (1 to N):
   - Add drawing reference numerals in parentheses to all features in claims using the complete specification.
   - Remove open-ended terms and positively define technical features.
   - Provide explicit Section 59(1) basis in the specification.
6. Return strictly valid JSON adhering to the specified schema.
"""

    last_error = None
    for model_name in candidate_models:
        for attempt in range(3):
            try:
                response = client.models.generate_content(
                    model=model_name,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json",
                        response_schema=FERStructuredOutput,
                        temperature=0.1
                    )
                )
                return FERStructuredOutput.model_validate_json(response.text)
            except Exception as e:
                last_error = e
                time.sleep(4 * (attempt + 1))

    raise RuntimeError(f"Generation error: {last_error}")


# ===========================================================================
# 4. Word Document Generators (.docx)
# ===========================================================================

def apply_text_style(run, size_pt=11, bold=False, italic=False, underline=False, strike=False, color_rgb=(0, 0, 0)):
    run.font.name = "Calibri"
    run.font.size = Pt(size_pt)
    run.font.color.rgb = RGBColor(*color_rgb)
    run.bold = bold
    run.italic = italic
    run.underline = underline
    run.font.strike = strike

def add_heading(doc, text, size=11, bold=True, space_before=12, space_after=4):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(space_before)
    p.paragraph_format.space_after = Pt(space_after)
    r = p.add_run(text)
    apply_text_style(r, size_pt=size, bold=bold)
    return p

def add_paragraph(doc, text, size=11, space_after=6):
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.15
    r = p.add_run(text)
    apply_text_style(r, size_pt=size)
    return p

def add_signature(doc, agent_name: str, regn_no: str, applicant_name: str):
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(20)
    p.paragraph_format.space_after = Pt(4)
    today_str = datetime.now().strftime("%B %d, %Y")
    r_date = p.add_run(f"Dated this {today_str}\n\n\n")
    apply_text_style(r_date, size_pt=11)
    r_sig = p.add_run(f"Digitally Signed\n{agent_name}\n(Regn. No.: {regn_no})\nAgent for the Applicant\n{applicant_name}")
    apply_text_style(r_sig, size_pt=11, bold=True)

def build_main_reply_stream(data: FERStructuredOutput, agent_name: str, regn_no: str) -> io.BytesIO:
    doc = Document()
    for s in doc.sections:
        s.top_margin = s.bottom_margin = s.left_margin = s.right_margin = Inches(1.0)

    clean_date = data.fer_date.replace("/", "-").strip()
    try:
        fer_dt = datetime.strptime(clean_date, "%d-%m-%Y")
        due_date_str = (fer_dt + relativedelta(months=6)).strftime("%d/%m/%Y")
    except Exception:
        due_date_str = "Within six months of FER dispatch"

    add_heading(doc, "BEFORE THE CONTROLLER OF PATENTS\nTHE PATENT OFFICE, DELHI", size=12, bold=True)
    add_paragraph(doc, "Boudhik Sampada Bhawan, Plot No. 32, Sector 14, Dwarka, New Delhi – 110078")
    add_paragraph(doc, f"Dated: {datetime.now().strftime('%B %d, %Y')}")
    add_paragraph(doc, f"KIND ATTN: {data.controller_name.upper()}, CONTROLLER OF PATENTS\nExaminer: {data.examiner_name}")

    tbl = doc.add_table(rows=5, cols=2)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    rows = [
        ("Application No.", data.app_no),
        ("Filing Date", data.filing_date),
        ("FER Date", data.fer_date),
        ("Statutory Due Date (6 Months)", due_date_str),
        ("Applicant", data.applicant_name)
    ]
    for idx, (k, v) in enumerate(rows):
        r = tbl.rows[idx]
        apply_text_style(r.cells[0].paragraphs[0].add_run(k), size_pt=10, bold=True)
        apply_text_style(r.cells[1].paragraphs[0].add_run(v), size_pt=10)

    add_heading(doc, f"Subject: Reply to First Examination Report dated {data.fer_date} in respect of Patent Application No. {data.app_no}", size=11, bold=True)
    add_paragraph(doc, "Dear Sir,\n\nPlease find below the point-wise submissions on behalf of the Applicant in response to the objections raised by the Learned Controller.")

    add_heading(doc, "BRIEF SUMMARY OF THE INVENTION", size=11, bold=True)
    add_paragraph(doc, data.brief_summary)

    add_heading(doc, "RESPONSE TO OBJECTIONS UNDER SECTION 2(1)(ja) – INVENTIVE STEP", size=11, bold=True)
    add_paragraph(doc, data.inventive_step_argument)

    add_heading(doc, "Feature-by-Feature Prior Art Comparison Table (Claim 1)", size=10, bold=True)
    comp_tbl = doc.add_table(rows=1, cols=4)
    comp_tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, title in enumerate(["Feature of Claim 1", "Present Invention", "D1 Disclosure", "D2 / D3 / D4 Disclosure"]):
        apply_text_style(comp_tbl.rows[0].cells[i].paragraphs[0].add_run(title), size_pt=10, bold=True)

    for row_obj in data.comparison_table:
        row = comp_tbl.add_row()
        d_comb = f"D2: {row_obj.d2_disclosure}\nD3: {row_obj.d3_disclosure}\nD4: {row_obj.d4_disclosure}".strip()
        vals = [row_obj.feature_name, row_obj.present_invention, row_obj.d1_disclosure, d_comb]
        for c_idx, val in enumerate(vals):
            apply_text_style(row.cells[c_idx].paragraphs[0].add_run(val), size_pt=9.5)

    add_heading(doc, "RESPONSE TO OBJECTIONS UNDER SECTION 10(4) – SUFFICIENCY OF DISCLOSURE", size=11, bold=True)
    add_paragraph(doc, data.sufficiency_argument)

    add_heading(doc, "RESPONSE TO OBJECTIONS REGARDING SCOPE, CLARITY AND CONCISENESS", size=11, bold=True)
    add_paragraph(doc, data.clarity_argument)

    add_heading(doc, "RESPONSE TO PART-III: FORMAL REQUIREMENTS", size=11, bold=True)
    add_paragraph(doc, (
        "1. Statement and Undertaking (Form 3): Updated Form 3 is submitted herewith.\n"
        "2. Signatures on Forms: All requisite forms (Forms 1, 3, 5, 9, 18), specification, and drawings have been digitally signed by the authorized Patent Agent under Rule 8.\n"
        "3. Reference Numerals: The Claims and Abstract have been amended to incorporate drawing reference numerals in parentheses in accordance with Rule 13(4) and Rule 13(7).\n"
        "4. Section 59(1) Compliance: The amended claims are strictly supported by the as-filed specification."
    ))

    add_heading(doc, "PRAYER", size=11, bold=True)
    add_paragraph(doc, (
        "In view of the above submissions, the Applicant respectfully submits that all objections have been overcome.\n\n"
        "The Learned Controller is respectfully requested to accept this reply and allow the patent application to proceed to grant.\n\n"
        "In case the Learned Controller has any further objections, it is requested that an opportunity of being heard under Section 14 be granted before taking any adverse decision."
    ))

    add_signature(doc, agent_name, regn_no, data.applicant_name)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream

def build_claims_stream(data: FERStructuredOutput, agent_name: str, regn_no: str, marked: bool = False) -> io.BytesIO:
    doc = Document()
    title = "MARKED-UP COPY OF AMENDED CLAIMS" if marked else "CLEAN COPY OF AMENDED CLAIMS"
    sub = "Showing additions underlined and deletions struck through (Section 59 Compliance)" if marked else "In accordance with Rule 14 of The Patents Rules, 2003"

    add_heading(doc, title, size=13, bold=True)
    add_paragraph(doc, f"Indian Patent Application No.: {data.app_no}\n{sub}")
    add_heading(doc, "WE CLAIM:", size=11, bold=True)

    for item in data.amended_claims:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.15
        r_num = p.add_run(f"{item.claim_number}. ")
        apply_text_style(r_num, size_pt=11, bold=True)

        if not marked:
            apply_text_style(p.add_run(item.amended_text), size_pt=11)
        else:
            orig_words = item.original_text.split()
            amend_words = item.amended_text.split()
            matcher = difflib.SequenceMatcher(None, orig_words, amend_words)
            for op, a0, a1, b0, b1 in matcher.get_opcodes():
                orig_part = " ".join(orig_words[a0:a1]) + " "
                amend_part = " ".join(amend_words[b0:b1]) + " "
                if op == "equal":
                    apply_text_style(p.add_run(orig_part), size_pt=11)
                elif op == "delete":
                    apply_text_style(p.add_run(orig_part), size_pt=11, color_rgb=(180, 0, 0), strike=True)
                elif op == "insert":
                    apply_text_style(p.add_run(amend_part), size_pt=11, color_rgb=(0, 120, 0), underline=True)
                elif op == "replace":
                    apply_text_style(p.add_run(orig_part), size_pt=11, color_rgb=(180, 0, 0), strike=True)
                    apply_text_style(p.add_run(amend_part), size_pt=11, color_rgb=(0, 120, 0), underline=True)

            p_basis = doc.add_paragraph()
            p_basis.paragraph_format.space_after = Pt(6)
            apply_text_style(p_basis.add_run(f"[Basis under Section 59(1): {item.basis_in_spec}]"), size_pt=9.5, italic=True, color_rgb=(80, 80, 80))

    add_signature(doc, agent_name, regn_no, data.applicant_name)
    stream = io.BytesIO()
    doc.save(stream)
    stream.seek(0)
    return stream


# ===========================================================================
# 5. Streamlit SaaS Front-End with Session State Persistence
# ===========================================================================

st.set_page_config(
    page_title="IP Docket AI — Patent Prosecution Suite",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Plus+Jakarta+Sans:wght@400;500;600;700&display=swap');
    html, body, [class*="css"] {
        font-family: 'Plus Jakarta Sans', sans-serif;
    }
    .hero-container {
        padding: 1.8rem 2.2rem;
        border-radius: 18px;
        background: linear-gradient(135deg, #0F172A 0%, #1E293B 50%, #0F766E 100%);
        color: white;
        margin-bottom: 2rem;
    }
    .hero-title {
        font-size: 2rem;
        font-weight: 700;
        margin-bottom: 0.4rem;
    }
    .hero-subtitle {
        color: #94A3B8;
        font-size: 0.95rem;
        margin: 0;
    }
    .doc-card {
        background: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 14px;
        padding: 1.2rem;
        margin-bottom: 1.2rem;
    }
    .doc-card-header {
        font-weight: 600;
        font-size: 1.05rem;
        color: #0F172A;
    }
    .badge-ready {
        background: #DCFCE7;
        color: #166534;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.2rem 0.6rem;
        border-radius: 9999px;
    }
    .badge-pending {
        background: #FEF3C7;
        color: #92400E;
        font-size: 0.75rem;
        font-weight: 600;
        padding: 0.2rem 0.6rem;
        border-radius: 9999px;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-container">
    <div class="hero-title">⚡ IP Counsel Suite &bull; FER Response Architect</div>
    <p class="hero-subtitle">
        Automated parsing, prior art feature-mapping, and Section 59-compliant response generation 
        conforming to the Patents Act, 1970 and Patent Rules, 2003.
    </p>
</div>
""", unsafe_allow_html=True)

# -------------------------------------------------------------
# Auto-Load & Permanent Save of API Key
# -------------------------------------------------------------
stored_key = ""
try:
    if "GEMINI_API_KEY" in st.secrets:
        stored_key = st.secrets["GEMINI_API_KEY"]
except Exception:
    pass

if not stored_key:
    stored_key = os.getenv("GEMINI_API_KEY", "")

with st.sidebar:
    st.markdown("### 🏛️ Practitioner Credentials")
    agent_name = st.text_input("Registered Patent Agent", value="Dr. UMESHKUMAR K.U")
    regn_no = st.text_input("Registration ID", value="IN/PA 3226")
    st.markdown("---")
    
    st.markdown("### 🔑 Engine Authorization")
    api_key = st.text_input(
        "Gemini API Key", 
        value=stored_key, 
        type="password", 
        help="Once entered, this key is saved to your local disk and remembered permanently."
    )
    
    # If the user typed/changed a new key, save it permanently to .env file
    if api_key and api_key != stored_key:
        try:
            if not os.path.exists(ENV_PATH):
                open(ENV_PATH, "w").close()
            set_key(ENV_PATH, "GEMINI_API_KEY", api_key)
            os.environ["GEMINI_API_KEY"] = api_key
            st.toast("✅ API Key saved permanently!", icon="💾")
        except Exception:
            pass

    if api_key:
        st.markdown('<span class="badge-ready">✓ API Key Saved & Active</span>', unsafe_allow_html=True)
    else:
        st.markdown('<span class="badge-pending">! Enter Key Once to Save</span>', unsafe_allow_html=True)
        
    st.markdown("---")
    st.caption("Jurisdiction: Indian Patent Office (IPO)<br>Branch Off: Delhi / Mumbai / Kolkata / Chennai", unsafe_allow_html=True)

col_left, col_right = st.columns(2, gap="medium")

with col_left:
    st.markdown('<div class="doc-card"><div class="doc-card-header">📄 1. Official Examination Report</div></div>', unsafe_allow_html=True)
    fer_file = st.file_uploader("Upload FER Document", type=["pdf"], key="fer")

    st.markdown('<div class="doc-card"><div class="doc-card-header">📘 2. Complete Specification</div></div>', unsafe_allow_html=True)
    spec_file = st.file_uploader("Upload Specification File", type=["docx", "doc", "pdf"], key="spec")

with col_right:
    st.markdown('<div class="doc-card"><div class="doc-card-header">📋 3. As-Filed Claims</div></div>', unsafe_allow_html=True)
    claims_file = st.file_uploader("Upload Pending Claims", type=["docx", "doc", "pdf"], key="claims")

    st.markdown('<div class="doc-card"><div class="doc-card-header">📚 4. Cited Prior Art References</div></div>', unsafe_allow_html=True)
    pa_files = st.file_uploader("Upload Cited Prior Art (D1, D2, D3, D4)", type=["pdf", "docx"], accept_multiple_files=True, key="pa")

st.markdown("<br>", unsafe_allow_html=True)

btn_col1, btn_col2, btn_col3 = st.columns([1, 2, 1])
with btn_col2:
    generate_clicked = st.button("Generate Complete Response Package", use_container_width=True)

# -------------------------------------------------------------
# Trigger Generation and Store in Session State
# -------------------------------------------------------------
if generate_clicked:
    if not api_key:
        st.error("Authentication required: Please enter your Gemini API Key in the sidebar.")
    elif not (fer_file and spec_file and claims_file and pa_files):
        st.warning("Missing documentation: Please upload all four document slots before generating.")
    else:
        status_box = st.status("Analyzing prosecution dossier...", expanded=True)
        try:
            status_box.write("Extracting claims, specification embodiments, and FER citations...")
            fer_txt = extract_text_from_upload(fer_file, max_pages=8, max_chars=20000)
            spec_txt = extract_text_from_upload(spec_file, max_pages=6, max_chars=18000)
            claims_txt = extract_text_from_upload(claims_file, max_pages=4, max_chars=12000)

            pa_text_parts = []
            for idx, pa in enumerate(pa_files, start=1):
                pa_content = extract_text_from_upload(pa, max_pages=3, max_chars=6000)
                pa_text_parts.append(f"--- PRIOR ART D{idx} ({pa.name}) ---\n{pa_content}")
            prior_arts_txt = "\n\n".join(pa_text_parts)

            status_box.write("Formulating Section 2(1)(ja) non-obviousness distinctions and claim charts...")
            structured_data = run_ai_pipeline(fer_txt, spec_txt, claims_txt, prior_arts_txt, api_key)

            status_box.write("Rendering redline diff engine & preparing DOCX packages...")
            main_docx = build_main_reply_stream(structured_data, agent_name, regn_no)
            clean_docx = build_claims_stream(structured_data, agent_name, regn_no, marked=False)
            marked_docx = build_claims_stream(structured_data, agent_name, regn_no, marked=True)

            app_id = re.sub(r"[^a-zA-Z0-9]", "", structured_data.app_no)

            zip_stream = io.BytesIO()
            with zipfile.ZipFile(zip_stream, "w", zipfile.ZIP_DEFLATED) as zip_file:
                zip_file.writestr(f"{app_id}_FER_Reply.docx", main_docx.getvalue())
                zip_file.writestr(f"{app_id}_Clean_Amended_Claims.docx", clean_docx.getvalue())
                zip_file.writestr(f"{app_id}_Marked_Amended_Claims.docx", marked_docx.getvalue())
            zip_stream.seek(0)

            # Persist results in session_state so downloads don't wipe the screen
            st.session_state["fer_results"] = {
                "structured_data": structured_data,
                "app_id": app_id,
                "main_docx_bytes": main_docx.getvalue(),
                "clean_docx_bytes": clean_docx.getvalue(),
                "marked_docx_bytes": marked_docx.getvalue(),
                "zip_bytes": zip_stream.getvalue()
            }

            status_box.update(label="Response compilation complete!", state="complete", expanded=False)

        except Exception as ex:
            status_box.update(label="Compilation Error", state="error", expanded=True)
            st.error(f"Execution failed: {str(ex)}")

# -------------------------------------------------------------
# Display Persistent Results
# -------------------------------------------------------------
if "fer_results" in st.session_state:
    res = st.session_state["fer_results"]
    structured_data = res["structured_data"]
    app_id = res["app_id"]

    st.markdown("---")
    st.markdown(f"### 📦 Deliverables Ready: Application No. `{structured_data.app_no}`")

    # Interactive Previews
    tab1, tab2, tab3 = st.tabs(["Prior Art Comparison Chart", "Inventive Step Traversal", "Amended Claims Preview"])
    
    with tab1:
        table_rows = []
        for r in structured_data.comparison_table:
            table_rows.append({
                "Claim 1 Feature": r.feature_name,
                "Present Invention": r.present_invention,
                "D1": r.d1_disclosure,
                "D2 / D3 / D4": f"{r.d2_disclosure} | {r.d3_disclosure}"
            })
        st.dataframe(table_rows, use_container_width=True)

    with tab2:
        st.markdown(structured_data.inventive_step_argument)

    with tab3:
        for c in structured_data.amended_claims[:3]:
            st.markdown(f"**Claim {c.claim_number}:** {c.amended_text}")
            st.caption(f"Basis under Section 59(1): {c.basis_in_spec}")
        if len(structured_data.amended_claims) > 3:
            st.info(f"+ {len(structured_data.amended_claims) - 3} additional amended claims included in downloaded files.")

    st.markdown("<br>", unsafe_allow_html=True)
    
    # Download Buttons (Persist seamlessly without reloads)
    d1, d2, d3, d4 = st.columns(4)
    with d1:
        st.download_button(
            "📄 Official FER Reply (.docx)", 
            data=res["main_docx_bytes"], 
            file_name=f"{app_id}_FER_Reply.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            use_container_width=True
        )
    with d2:
        st.download_button(
            "📋 Clean Claims (.docx)", 
            data=res["clean_docx_bytes"], 
            file_name=f"{app_id}_Clean_Claims.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            use_container_width=True
        )
    with d3:
        st.download_button(
            "📝 Marked Claims (.docx)", 
            data=res["marked_docx_bytes"], 
            file_name=f"{app_id}_Marked_Claims.docx", 
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
            use_container_width=True
        )
    with d4:
        st.download_button(
            "📦 Complete ZIP Archive", 
            data=res["zip_bytes"], 
            file_name=f"{app_id}_FER_Package.zip", 
            mime="application/zip", 
            use_container_width=True
        )