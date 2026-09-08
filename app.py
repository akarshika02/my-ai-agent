import os
import sys
import time
import importlib.util
from concurrent.futures import ThreadPoolExecutor
import pandas as pd
import requests
import streamlit as st

# --- Base Directory Resolution ---
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# --- Dynamically Load run 1.py ---
RUN_1_PATH = os.path.join(BASE_DIR, "run 1.py")
if not os.path.exists(RUN_1_PATH):
    RUN_1_PATH = "run 1.py"

spec = importlib.util.spec_from_file_location("run_1", RUN_1_PATH)
run_1 = importlib.util.module_from_spec(spec)
spec.loader.exec_module(run_1)

def get_template_df():
    """
    Finds, loads, or auto-generates the Blank File template DataFrame so the agent never fails.
    """
    # Session state uploaded file
    if st.session_state.get("custom_blank_file_path") and os.path.exists(st.session_state["custom_blank_file_path"]):
        try:
            return pd.read_excel(st.session_state["custom_blank_file_path"]), st.session_state["custom_blank_file_path"]
        except Exception:
            pass
            
    candidates = [
        os.path.join(BASE_DIR, "Blank File.xlsx"),
        "Blank File.xlsx",
        os.path.join(BASE_DIR, "uploaded_template_Blank File.xlsx"),
        "uploaded_template_Blank File.xlsx"
    ]
    
    for path in candidates:
        if os.path.exists(path):
            try:
                df = pd.read_excel(path)
                if not df.columns.empty:
                    return df, path
            except Exception:
                continue
                
    # Auto-generate Blank File.xlsx if missing anywhere
    fallback_path = os.path.join(BASE_DIR, "Blank File.xlsx")
    headers = getattr(run_1, "EXCEL_HEADERS", [])
    if not headers:
        headers = [
            "Date", "Year", "Month", "Date Posted", "Job ID LinkedIn", "Job ID Dexcom Page",
            "Job Posting site", "Location", "Country", "US / OUS", "Business Function",
            "Department", "Department from Dexcom Company Website", "Confidence Level",
            "Job Title", "Fresh / reposted", "Dexcom Offers / Why Dexcom? / What you'll get",
            "Summary / Position Summary / Meet the team / Role Summary",
            "Essential Duties and Responsibilities / Where you come in",
            "Supervisory Responsibilities",
            "Required Qualifications / What makes you successful / Requirements / Essential Capabilities / Competencies / About you",
            "Preferred Qualifications / Key Competencies", "Education Requirements and Experience",
            "Travel Required", "Workplace Type", "Functional Description",
            "Functional / Business Knowledge", "Scope", "Judgement", "Language Skills",
            "Physical Demands", "Work Environment", "Points to Note", "Management",
            "Field Sales", "Pay / Non-Exempt Salary Details / Commercial Salary Details / Exempt Salary Details",
            "Shifts", "Direct URL of the Dexcom Job Page"
        ]
    df_auto = pd.DataFrame(columns=headers)
    try:
        df_auto.to_excel(fallback_path, index=False)
    except Exception:
        pass
    return df_auto, "Auto-generated Blank File template"

def safe_save_excel(df: pd.DataFrame, default_name: str = "job_postings_final.xlsx") -> tuple:
    """
    Safely save DataFrame to Excel, attempting fallback filenames if file is locked by Excel.
    """
    targets = [
        os.path.join(BASE_DIR, default_name),
        default_name,
        os.path.join(BASE_DIR, "job_postings.xlsx"),
        "job_postings.xlsx",
        os.path.join(BASE_DIR, f"job_postings_output_{int(time.time())}.xlsx")
    ]
    
    warning_msg = None
    for target in targets:
        try:
            df.to_excel(target, index=False)
            if os.path.basename(target) != default_name:
                warning_msg = f"⚠️ `{default_name}` is open in Excel. Output saved as `{os.path.basename(target)}` instead."
            return target, warning_msg
        except PermissionError:
            continue
        except Exception:
            continue
            
    timestamp_file = os.path.join(BASE_DIR, f"job_postings_output_{int(time.time())}.xlsx")
    df.to_excel(timestamp_file, index=False)
    return timestamp_file, f"⚠️ Output saved as `{os.path.basename(timestamp_file)}`."

# --- Page Configuration ---
st.set_page_config(
    page_title="LinkedIn Job Posting Extraction Agent",
    page_icon="🤖",
    layout="wide",
    initial_sidebar_state="expanded"
)

# --- Custom Styling ---
st.markdown("""
<style>
    .main-header {
        text-align: center;
        margin-bottom: 1.5rem;
        padding: 1.5rem;
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%);
        border: 1px solid #334155;
        border-radius: 16px;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.4);
    }
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        background: linear-gradient(135deg, #38bdf8, #818cf8, #c084fc);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.4rem;
    }
    .main-subtitle {
        color: #94a3b8;
        font-size: 1.05rem;
    }
    .log-box {
        font-family: 'Fira Code', 'Courier New', Courier, monospace;
        background-color: #090d16;
        color: #38bdf8;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #334155;
        height: 250px;
        overflow-y: auto;
        font-size: 0.88rem;
        line-height: 1.5;
    }
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown("""
<div class="main-header">
    <div class="main-title">🤖 LinkedIn Job Posting Extraction Agent</div>
    <div class="main-subtitle">Upload your input Excel file and click process. The agent automatically converts, extracts parameters, and generates your downloadable Excel file.</div>
</div>
""", unsafe_allow_html=True)

# Load template DataFrame cleanly
df_template_found, template_source = get_template_df()

# --- Sidebar Controls & Custom Template Uploader ---
with st.sidebar:
    st.header("⚙️ Agent Settings")
    st.success(f"✓ **Template Loaded:** `{os.path.basename(template_source)}` ({len(df_template_found.columns)} cols)")
    
    with st.expander("📤 Optional: Upload Custom Blank Template", expanded=False):
        uploaded_custom_template = st.file_uploader("Upload custom template (.xlsx)", type=["xlsx", "xls"], key="custom_tpl_uploader")
        if uploaded_custom_template is not None:
            custom_path = os.path.join(BASE_DIR, f"uploaded_template_{uploaded_custom_template.name}")
            with open(custom_path, "wb") as f:
                f.write(uploaded_custom_template.getbuffer())
            st.session_state["custom_blank_file_path"] = custom_path
            st.rerun()

# --- Session State Setup ---
if "extracted_df" not in st.session_state:
    st.session_state["extracted_df"] = None
if "output_excel_saved" not in st.session_state:
    st.session_state["output_excel_saved"] = None
if "excel_warning" not in st.session_state:
    st.session_state["excel_warning"] = None

# --- Main Interface ---
st.subheader("1️⃣ Upload Input Excel File")
uploaded_file = st.file_uploader("Choose an Excel file (.xlsx, .xls)", type=["xlsx", "xls"], key="single_file_uploader")

input_file_path = None
if uploaded_file is not None:
    input_file_path = os.path.join(BASE_DIR, f"uploaded_{uploaded_file.name}")
    with open(input_file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    st.success(f"✓ Uploaded: `{uploaded_file.name}`")
else:
    default_candidates = [
        os.path.join(BASE_DIR, "Input1.xlsx"),
        "Input1.xlsx"
    ]
    for cand in default_candidates:
        if os.path.exists(cand):
            input_file_path = cand
            st.info(f"💡 Default input file ready: `{os.path.basename(cand)}`")
            break

if input_file_path and os.path.exists(input_file_path):
    df_preview = pd.read_excel(input_file_path)
    st.markdown(f"**Input Rows Found:** `{len(df_preview)} records`")
    with st.expander("👁️ View Input File Preview", expanded=False):
        st.dataframe(df_preview.head(5), use_container_width=True)
        
    st.markdown("---")
    
    process_btn = st.button("🚀 Process & Extract All Jobs", type="primary", use_container_width=True)
    
    if process_btn:
        st.markdown("### 🔄 Processing Steps & Progress")
        
        step_status = st.empty()
        log_box = st.empty()
        progress_bar = st.progress(0.0)
        
        m1, m2, m3, m4 = st.columns(4)
        met1 = m1.empty()
        met2 = m2.empty()
        met3 = m3.empty()
        met4 = m4.empty()
        
        logs = []
        
        def add_log(msg):
            logs.append(f"[{time.strftime('%H:%M:%S')}] {msg}")
            log_box.markdown(f"<div class='log-box'>{'<br>'.join(logs[-15:])}</div>", unsafe_allow_html=True)
            
        # STEP A: Convert Excel to TXT
        step_status.markdown("🔄 **Step 1/4:** Converting Excel input file to `output.txt`...")
        add_log(f"Reading input file `{os.path.basename(input_file_path)}` ({len(df_preview)} rows)...")
        
        output_txt_path = os.path.join(BASE_DIR, "output.txt")
        df_preview.to_csv(output_txt_path, sep="\t", index=False)
        add_log("✓ Successfully converted Excel input to tab-delimited format.")
        time.sleep(0.4)
        
        # STEP B: Load Blank File Template
        step_status.markdown("📋 **Step 2/4:** Loading output template structure...")
        df_blank, template_source_used = get_template_df()
        template_cols = list(df_blank.columns)
        add_log(f"✓ Loaded output template from `{os.path.basename(template_source_used)}` ({len(template_cols)} parameter columns).")
        time.sleep(0.4)
        
        # STEP C: Parse TXT and Run Extraction Code
        step_status.markdown("⚡ **Step 3/4:** Executing `run 1.py` extraction code on job URLs...")
        entries = []
        with open(output_txt_path, "r", encoding="utf-8", errors="ignore") as f:
            for line in f:
                l_str = line.strip()
                if not l_str:
                    continue
                if ("linkedin_webpages" in l_str.lower() or "job_link" in l_str.lower()) and "http" not in l_str.lower():
                    continue
                entries.append(l_str)
                
        total_items = len(entries)
        add_log(f"Starting parameter extraction on {total_items} LinkedIn job link(s)...")
        
        met1.metric("Processed Jobs", f"0 / {total_items}")
        met2.metric("Both Sites Matched", "0")
        met3.metric("LinkedIn Only", "0")
        met4.metric("Fresh Jobs", "0")
        
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept-Language': 'en-US,en;q=0.9'
        }
        
        records_indexed = []
        
        def process_single(item):
            idx, link_line = item
            time.sleep((idx % 3) * 0.3)
            row_data = run_1.process_job_link(link_line, headers, debug=False)
            return idx, row_data
            
        with ThreadPoolExecutor(max_workers=min(3, max(1, total_items))) as executor:
            indexed_items = list(enumerate(entries, start=1))
            
            for completed_count, (idx, row_data) in enumerate(executor.map(process_single, indexed_items), start=1):
                records_indexed.append((idx, row_data))
                
                curr_rows = [r[1] for r in records_indexed]
                both_cnt = sum(1 for r in curr_rows if r.get("Job Posting site") == "Both")
                li_cnt = sum(1 for r in curr_rows if r.get("Job Posting site") == "LinkedIn")
                
                msg = f"[{completed_count}/{total_items}] Extracted: {row_data.get('Job Title', 'N/A')} | Location: {row_data.get('Location', 'N/A')} | Dept: {row_data.get('Department from Dexcom Company Website', 'N/A')}"
                add_log(msg)
                
                progress_bar.progress(completed_count / total_items)
                met1.metric("Processed Jobs", f"{completed_count} / {total_items}")
                met2.metric("Both Sites Matched", both_cnt)
                met3.metric("LinkedIn Only", li_cnt)
                
        # Sort and update fresh/reposted status
        records_indexed.sort(key=lambda x: x[0])
        final_records = [r[1] for r in records_indexed]
        run_1.update_fresh_reposted_status(final_records)
        
        fresh_cnt = sum(1 for r in final_records if r.get("Fresh / reposted") == "Fresh/New Job Posted")
        met4.metric("Fresh Jobs", fresh_cnt)
        
        # STEP D: Populate Template & Output Excel File
        step_status.markdown("📥 **Step 4/4:** Formatting output matching template structure...")
        df_extracted = pd.DataFrame(final_records)
        
        for col in template_cols:
            if col not in df_extracted.columns:
                df_extracted[col] = "-"
                
        final_cols = [c for c in template_cols if c in df_extracted.columns]
        if "LinkedIn URL" in df_extracted.columns and "LinkedIn URL" not in final_cols:
            final_cols.append("LinkedIn URL")
            
        df_final = df_extracted[final_cols]
        
        if hasattr(df_final, 'map'):
            df_final = df_final.map(run_1.sanitize_excel_value)
        else:
            df_final = df_final.applymap(run_1.sanitize_excel_value)
            
        st.session_state["extracted_df"] = df_final
        
        saved_file, warning_text = safe_save_excel(df_final, "job_postings_final.xlsx")
        st.session_state["output_excel_saved"] = saved_file
        st.session_state["excel_warning"] = warning_text
        
        step_status.markdown("✅ **Extraction Completed Successfully!**")
        add_log(f"🎉 Successfully populated {len(df_final)} rows into `{os.path.basename(saved_file)}`!")
        st.balloons()
        
# --- Download & Results Section ---
if st.session_state["extracted_df"] is not None:
    st.markdown("---")
    st.subheader("🎉 Extracted Results & Download")
    
    if st.session_state.get("excel_warning"):
        st.warning(st.session_state["excel_warning"])
        
    excel_path = st.session_state.get("output_excel_saved") or os.path.join(BASE_DIR, "job_postings_final.xlsx")
    if os.path.exists(excel_path):
        with open(excel_path, "rb") as f:
            bytes_data = f.read()
            
        st.download_button(
            label=f"📥 Download Extracted Job Postings Excel File ({os.path.basename(excel_path)})",
            data=bytes_data,
            file_name=os.path.basename(excel_path),
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
        
    st.markdown("#### 🔍 Extracted Data Preview")
    df_out = st.session_state["extracted_df"]
    search_q = st.text_input("🔍 Search within extracted output:", "")
    if search_q:
        filtered_df = df_out[df_out.apply(lambda r: r.astype(str).str.contains(search_q, case=False).any(), axis=1)]
    else:
        filtered_df = df_out
        
    st.dataframe(filtered_df, use_container_width=True, height=450)
else:
    st.info("👉 Upload an Excel file above and click **Process & Extract All Jobs** to begin.")
