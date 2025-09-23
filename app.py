import streamlit as st
import os
from pathlib import Path
from typing import Dict

# ========= Helpers ========= #
# Create starter Excel if missing
def ensure_excel_present(path: Path):
    try:
        if not path.exists():
            from openpyxl import Workbook
            path.parent.mkdir(parents=True, exist_ok=True)
            wb = Workbook()
            ws = wb.active
            ws.title = "Fill me"
            ws["A1"] = "Fill required inputs then press Launch."
            wb.save(path)
    except Exception:
        pass

# Return Excel file bytes
def get_excel_bytes(excel_path: Path) -> bytes:
    ensure_excel_present(excel_path)
    return excel_path.read_bytes()

# ========= Config ========= #
st.set_page_config(page_title="BRIXS Reports Downloader", layout="wide")

SCRIPT_DIR = Path(__file__).parent.resolve()
scripts: Dict[str, Path] = {
    "📋 Budget Comparison": SCRIPT_DIR / "Budget_comparison.py",
    "📋 Trial/Balance/Income/12Month Statement/Budget Comparison(with PTD)": SCRIPT_DIR / "financial_analytics.py",
    "📋 General Ledger": SCRIPT_DIR / "gl_analytics.py",
    "📋 Property Residential": SCRIPT_DIR / "residential.py",
    "📋 Affordable Receivable Report(Include/Exclude)": SCRIPT_DIR / "affordable_receivable_report.py",
    "📋 Affordable Rent Roll with Lease Charges": SCRIPT_DIR / "affordable_report.py",
    "📂 Consolidated Report": SCRIPT_DIR / "consolidation.py",
}

excel_files: Dict[str, str] = {
    "📋 Budget Comparison": "Budget_comparison.xlsx",
    "📋 Trial/Balance/Income/12Month Statement/Budget Comparison(with PTD)": "financial_analytics.xlsx",
    "📋 General Ledger": "gl_analytics.xlsx",
    "📋 Property Residential": "residential.xlsx",
    "📋 Affordable Receivable Report(Include/Exclude)": "affordable_receivable_report.xlsx",
    "📋 Affordable Rent Roll with Lease Charges": "affordable_report.xlsx",
    "📂 Consolidated Report": "consolidation.xlsx",
}

ACCENTS = [
    {"c500": "#6366F1", "c400": "#818CF8", "glow": "rgba(99,102,241,.35)"},
    {"c500": "#A855F7", "c400": "#C084FC", "glow": "rgba(168,85,247,.35)"},
    {"c500": "#F97316", "c400": "#FB923C", "glow": "rgba(249,115,22,.35)"},
]

if "btn_status" not in st.session_state:
    st.session_state.btn_status = {label: "" for label in scripts.keys()}

# ========= Page Header ========= #
st.markdown(
    """
    <div style="background:#E2E8F0;padding:20px;border-radius:10px;margin-bottom:20px;">
    <h1>BRIXS Reports Downloader</h1>
    <p>Select a report: download the Excel template, upload the filled file, then run the script inside Streamlit Cloud.</p>
    </div>
    """,
    unsafe_allow_html=True
)

# ========= Cards Loop ========= #
for idx, (label, script_path) in enumerate(scripts.items()):
    accent = ACCENTS[idx % len(ACCENTS)]
    state = st.session_state.btn_status.get(label, "")

    # Titles
    st.subheader(label)

    # 1. Download Excel template
    excel_path = SCRIPT_DIR / excel_files[label]
    excel_bytes = get_excel_bytes(excel_path)
    st.download_button(
        label=f"📥 Download {excel_files[label]}",
        data=excel_bytes,
        file_name=excel_files[label],
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f"download_{idx}"
    )

    # 2. Upload filled Excel file
    uploaded_file = st.file_uploader(
        f"Upload filled {excel_files[label]}",
        type=["xlsx"],
        key=f"upload_{idx}"
    )
    if uploaded_file:
        user_filled_path = SCRIPT_DIR / f"user_{excel_files[label]}"
        with open(user_filled_path, "wb") as f:
            f.write(uploaded_file.read())
        st.success(f"Uploaded {excel_files[label]} successfully!")

    # 3. Run script inside Streamlit with UTF-8 encoding
    run_button = st.button(
        f"▶ Run {Path(script_path).name}",
        key=f"run_{idx}"
    )
    if run_button:
        with st.spinner(f"Running {script_path.name}…"):
            try:
                # ✅ Force UTF-8 to avoid 'charmap' error
                with open(script_path, "r", encoding="utf-8") as f:
                    code = f.read()
                exec(code, globals())
                st.success(f"Finished running {script_path.name}")
                st.session_state.btn_status[label] = "started"
            except Exception as e:
                st.error(f"❌ Error running {script_path.name}: {e}")
                st.session_state.btn_status[label] = ""

    # Status line
    st.write(f"**Status:** {st.session_state.btn_status.get(label, '') or 'Idle'}")
    st.markdown("---")
