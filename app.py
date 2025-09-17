import streamlit as st
import os
import platform
import time
import random
import subprocess
from pathlib import Path
from typing import Dict, Any

# Optional: create a starter workbook if the target Excel file is missing
def ensure_excel_present(path: Path):
    try:
        if not path.exists():
            from openpyxl import Workbook  # lightweight, common dep in your env
            path.parent.mkdir(parents=True, exist_ok=True)
            wb = Workbook()
            ws = wb.active
            ws.title = "Fill me"
            ws["A1"] = "Fill required inputs then press Launch."
            wb.save(path)
    except Exception:
        # If openpyxl isn't available, silently skip creation; Open Excel will error gracefully
        pass

# Cross-platform “open file with default app”
def open_system_file(path: Path):
    try:
        if platform.system() == "Windows":
            os.startfile(str(path))  # type: ignore[attr-defined]
        elif platform.system() == "Darwin":
            subprocess.run(["open", str(path)], check=False)
        else:  # Linux and others
            subprocess.run(["xdg-open", str(path)], check=False)
        return True, None
    except Exception as e:
        return False, str(e)

# =========================
# Config & Data (yours)
# =========================
st.set_page_config(page_title="Belsara Automation Launcher", layout="wide")

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
    {"name": "indigo", "c500": "#6366F1", "c400": "#818CF8", "glow": "rgba(99,102,241,.35)"},
    {"name": "purple", "c500": "#A855F7", "c400": "#C084FC", "glow": "rgba(168,85,247,.35)"},
    {"name": "orange", "c500": "#F97316", "c400": "#FB923C", "glow": "rgba(249,115,22,.35)"},
]

LABELS = list(scripts.keys())
ID_TO_LABEL = {f"id{i}": label for i, label in enumerate(LABELS)}
LABEL_TO_ID = {label: i for i, label in enumerate(LABELS)}

# =========================
# Session state
# =========================
if "btn_status" not in st.session_state:
    st.session_state.btn_status = {label: "" for label in scripts.keys()}
if "pending_launch_id" not in st.session_state:
    st.session_state.pending_launch_id = None
if "launch_modal_shown" not in st.session_state:
    st.session_state.launch_modal_shown = False
if "last_launch_ts" not in st.session_state:
    st.session_state.last_launch_ts = None

def set_status(label: str, status: str):
    st.session_state.btn_status[label] = status

# =========================
# Query params helpers
# =========================
def _get_all_query_params() -> Dict[str, Any]:
    try:
        return {k: v for k, v in dict(st.query_params).items()}
    except Exception:
        return st.experimental_get_query_params()

def _set_query_params(**params):
    try:
        st.query_params.clear()
        for k, v in params.items():
            st.query_params[k] = v
    except Exception:
        st.experimental_set_query_params(**params)

def _clear_launch_query_param():
    qp = _get_all_query_params()
    if "launch" in qp:
        qp.pop("launch", None)
        _set_query_params(**qp)

def _get_launch_param():
    try:
        val = st.query_params.get("launch", None)
        if isinstance(val, list):
            val = val[0] if val else None
        return val
    except Exception:
        qp = st.experimental_get_query_params()
        return (qp.get("launch", [None]) or [None])[0]

# =========================
# Styles (unchanged)
# =========================
particles = []
for _ in range(22):
    left = f"{random.random() * 100:.2f}%"
    top = f"{random.random() * 100:.2f}%"
    delay = f"{random.random() * 5:.2f}s"
    dur = f"{10 + random.random() * 20:.2f}s"
    size = f"{random.randint(6,10)}px"
    particles.append(
        f'<div class="particle" style="left:{left}; top:{top}; animation-delay:{delay}; '
        f'animation-duration:{dur}; width:{size}; height:{size};"></div>'
    )

st.markdown(
    f"""
<style>
html, body {{ height: 100%; }}
[data-testid="stAppViewContainer"] {{
  background: linear-gradient(135deg, #e0f2fe 0%, #ffffff 100%);
}}
.main > div {{ padding-top: 0.75rem; }}

.belsara-hero {{
  position: relative;
  background: linear-gradient(145deg, #e2e8f0, #cbd5e1);
  color: #0f172a; border-radius: 20px; padding: 24px 28px; margin: 10px 0 18px;
  border: 1px solid rgba(15,23,42,.16);
  box-shadow: 0 8px 24px rgba(15,23,42,.12), inset 0 1px 0 rgba(255,255,255,.7);
  overflow: hidden;
}}
.belsara-hero:before {{
  content:"";
  position:absolute; inset: -2px;
  background: conic-gradient(from 180deg at 50% 50%, rgba(255,255,255,.06), rgba(255,255,255,0), rgba(255,255,255,.06));
  filter: blur(24px); opacity:.35; pointer-events:none;
}}
.belsara-hero h1 {{ margin: 0; font-size: 28px; letter-spacing:.2px; color: #000000; }}
.belsara-hero p {{ margin: 4px 0 0; opacity: 0.8; }}

.grid {{ display: grid; grid-template-columns: repeat(2, 1fr); gap: 20px; width: 100%; max-width: 1200px; margin: 0 auto; }}
.grid-item {{ display: flex; flex-direction: column; }}

@media (max-width: 1100px) {{ .grid {{ grid-template-columns: repeat(2, 1fr); gap: 15px; }} }}
@media (max-width: 730px)  {{ .grid {{ grid-template-columns: 1fr; gap: 10px; }} }}

.glass-card {{
  position:relative; border-radius: 16px; padding: 14px;
  backdrop-filter: blur(12px);
  background: linear-gradient(135deg, rgba(224,242,254,.95) 0%, rgba(255,255,255,.95) 100%);
  border: 1px solid rgba(15,23,42,.08);
  box-shadow: 0 25px 60px rgba(2,132,199,.08), inset 0 1px 0 rgba(255,255,255,.6);
  transition: transform .25s ease, box-shadow .25s ease, border-color .25s ease;
  overflow: hidden;
  margin: 0;
  box-sizing: border-box;
  min-width: 0;
  min-height: 260px;
  display: flex;
  flex-direction: column;
}}
.glass-card:hover {{ transform: translateY(-4px); box-shadow: 0 20px 60px rgba(0,0,0,.55); }}
.glass-card .accent-ring {{
  position:absolute; inset:-2px; border-radius: 20px; pointer-events:none;
  background: radial-gradient(600px 180px at 10% 10%, var(--accentGlow), transparent 60%),
              radial-gradient(600px 180px at 90% 90%, var(--accentGlow), transparent 60%);
  opacity: .25; transition: opacity .3s ease;
}}
.glass-card:hover .accent-ring {{ opacity: .45; }}

.card-title {{
  margin: 0 0 8px 0; font-size: 16px; color: #0f172a; font-weight: 700; letter-spacing:.2px;
  text-shadow: 0 1px 0 rgba(0,0,0,.35);
  line-height: 1.25;
  min-height: calc(1.25em * 2);
  display: -webkit-box; -webkit-line-clamp: 2; -webkit-box-orient: vertical; overflow: hidden;
}}
.card-head {{ flex: 1 1 auto; }}
.card-desc {{ color: #334155; font-size: 13px; margin-bottom: 14px; min-height: 32px; line-height: 1.25; }}

.fancy-btn {{
  position: relative; display:block; width: 100%; text-decoration:none;
  background: linear-gradient(145deg, #e2e8f0, #cbd5e1);
  border: 1px solid rgba(15,23,42,.16); color: #0f172a;
  border-radius: 12px; padding: 10px 14px; font-weight: 700; letter-spacing:.2px;
  box-shadow: 0 8px 24px rgba(15,23,42,.12);
  transition: transform .2s ease, box-shadow .2s ease, border-color .2s ease, color .2s ease, opacity .2s ease;
  overflow: hidden; cursor: pointer; margin-top: auto;
}}
.fancy-btn:hover {{ transform: translateY(-2px) scale(1.01); border-color: var(--accent400); box-shadow: 0 20px 60px var(--accentShadow); }}
.fancy-btn:active {{ transform: translateY(0) scale(.99); }}
.fancy-btn:before {{
  content: ""; position:absolute; inset:0;
  background: linear-gradient(90deg, rgba(255,255,255,0), var(--accent400) 35%, rgba(255,255,255,0) 70%);
  transform: translateX(-120%); opacity:.22; pointer-events:none;
}}
.fancy-btn:hover:before {{ transform: translateX(120%); transition: transform 1.15s ease; }}
.fancy-btn.disabled {{ opacity: .6; pointer-events: none; }}

.btn-row {{ display:flex; align-items:center; gap:10px; justify-content: space-between; }}
.btn-left {{ display:flex; align-items:center; gap:10px; }}
.btn-icon {{ display:flex; align-items:center; justify-content:center; width:36px; height:36px; border-radius:10px;
  background: linear-gradient(145deg, var(--accent400), transparent 70%); box-shadow: inset 0 0 0 1px rgba(255,255,255,.08); }}
.btn-text-1 {{ color: var(--accent400); font-weight: 800; line-height: 1.05; }}
.btn-text-2 {{ color: #000000; font-size: 12.5px; margin-top:-2px; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; display:block; }}
.arrow {{ opacity:.45; transition: transform .2s ease, opacity .2s ease; color: var(--accent400); }}
.fancy-btn:hover .arrow {{ transform: translateX(4px); opacity:.95; }}

.badge-dot {{ width:8px; height:8px; border-radius:999px; display:inline-block; margin-right:6px; }}
.dot-green {{ background:#22C55E; box-shadow: 0 0 0 4px rgba(34,197,94,.15); animation: pulse 1.5s ease infinite; }}
.dot-yellow {{ background:#FACC15; box-shadow: 0 0 0 4px rgba(250,204,21,.15); }}
.dot-gray {{ background:#6B7280; box-shadow: 0 0 0 4px rgba(107,114,128,.12); }}

@keyframes pulse {{
  0% {{ transform: scale(1); opacity: .9; }}
  50% {{ transform: scale(1.25); opacity: .55; }}
  100% {{ transform: scale(1); opacity: .9; }}
}}
.particles {{ position: fixed; inset: 0; pointer-events: none; z-index: 0; overflow: hidden; }}
.particle {{ position: absolute; background: rgba(129,140,248,.32); border-radius: 999px; filter: blur(0.5px); animation: float linear infinite; }}
@keyframes float {{
  0% {{ transform: translateY(0) rotate(0deg); opacity: .7; }}
  100% {{ transform: translateY(-100vh) rotate(360deg); opacity: 0; }}
}}
</style>
<div class="particles">{''.join(particles)}</div>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
<div class="belsara-hero">
  <h1>
    <span style="font-family:'Comic Sans MS','Comic Sans',cursive;">BRIXS</span> Reports Downloader
  </h1>
  <p>Select a bot to launch. Each opens in a new terminal.</p>
</div>
""",
    unsafe_allow_html=True,
)



# =========================
# Launch helpers
# =========================
def launch_script(label: str, script_path: Path):
    try:
        if platform.system() == "Windows":
            os.system(f'start cmd /k python "{script_path}"')
        elif platform.system() == "Darwin":
            os.system(f'osascript -e \'tell app "Terminal" to do script "python3 \\"{script_path}\\""\'' )
        elif platform.system() == "Linux":
            os.system(f'gnome-terminal -- python3 "{script_path}"')
        else:
            st.error("❌ Unsupported OS")
            return
        set_status(label, "started")
        st.toast(f"🎯 Launched {label} in a new terminal window.", icon="✅")
    except Exception as e:
        set_status(label, "")
        st.error(f"❌ Failed to launch {label}: {e}")

def show_fill_modal(label: str, excel_file: str, script_path: Path):
    """
    Minimal 2-button modal for Consolidation (no header/description/Open Excel).
    Full dialog for all other reports.
    """
    excel_path = SCRIPT_DIR / excel_file
    ensure_excel_present(excel_path)

    # Robust detection: match by script file name or excel file name (not the label)
    is_consolidated = (
        script_path.name.lower().startswith("consolidation")
        or excel_file.lower().startswith("consolidation")
    )

    try:
        # Streamlit sometimes shows a default header text; we hide it via CSS.
        dialog_title = " "  # keep non-empty, then hide with CSS below

        @st.dialog(dialog_title)
        def _modal():
            # Always inject CSS to hide the dialog header title completely
            st.markdown(
                """
                <style>
                  /* Hide the dialog title/header text */
                  [data-testid="stDialog"] h1, 
                  [data-testid="stDialog"] h2 { display:none !important; }
                </style>
                """,
                unsafe_allow_html=True,
            )

            if is_consolidated:
                # Minimal: ONLY Launch now + Cancel
                spacer_left, c1, c2, spacer_right = st.columns([1, 1, 1, 1])
                with c1:
                    if st.button("Launch now", use_container_width=True, key=f"launch_min_{LABEL_TO_ID.get(label, 'cons')}"):
                        set_status(label, "connecting")
                        st.toast("🔗 Connecting…", icon="⏳")
                        time.sleep(0.4)
                        launch_script(label, script_path)
                        _clear_launch_query_param()
                        st.session_state.pending_launch_id = None
                        st.session_state.launch_modal_shown = True
                        st.rerun()
                with c2:
                    if st.button("Cancel", type="secondary", use_container_width=True, key=f"cancel_min_{LABEL_TO_ID.get(label, 'cons')}"):
                        set_status(label, "")
                        _clear_launch_query_param()
                        st.session_state.pending_launch_id = None
                        st.session_state.launch_modal_shown = True
                        st.rerun()

            else:
                # Original full flow for other reports
                st.markdown(
                    f"""
<div style="display:flex; gap:12px; align-items:flex-start;">
  <div style="font-size:28px;">📝</div>
  <div>
    <div style="font-weight:700; font-size:16px; margin-bottom:6px;">Please fill the Excel sheet first</div>
    <div>
      <b>{excel_file}</b><br/>
      You will find this file in the launcher folder.
    </div>
  </div>
</div>
                    """,
                    unsafe_allow_html=True,
                )
                st.divider()
                c1, c2, c3 = st.columns(3)
                with c1:
                    if st.button("Open Excel", key=f"modal_open_{LABEL_TO_ID[label]}"):
                        ok, err = open_system_file(excel_path)
                        if ok:
                            st.toast(f"🧾 Opening {excel_file}…", icon="📂")
                        else:
                            st.error(f"Couldn't open {excel_file}: {err or 'unknown error'}")

                with c2:
                    if st.button("Launch now", key=f"modal_launch_{LABEL_TO_ID[label]}"):
                        set_status(label, "connecting")
                        st.toast("🔗 Connecting…", icon="⏳")
                        time.sleep(0.5)
                        launch_script(label, script_path)
                        _clear_launch_query_param()
                        st.session_state.pending_launch_id = None
                        st.session_state.launch_modal_shown = True
                        st.rerun()

                with c3:
                    if st.button("Cancel", type="secondary", key=f"modal_cancel_{LABEL_TO_ID[label]}"):
                        set_status(label, "")
                        _clear_launch_query_param()
                        st.session_state.pending_launch_id = None
                        st.session_state.launch_modal_shown = True
                        st.rerun()

        _modal()

    except Exception:
        # Fallback behavior: keep minimal for consolidation; full for others
        if is_consolidated:
            set_status(label, "connecting")
            st.toast("Launching consolidated report…", icon="ℹ️")
            time.sleep(0.4)
            launch_script(label, script_path)
            _clear_launch_query_param()
        else:
            st.toast(f"📝 Please fill {excel_file} (in folder).", icon="ℹ️")
            ok, err = open_system_file(SCRIPT_DIR / excel_file)
            if not ok:
                st.error(f"Couldn't open {excel_file}: {err or 'unknown error'}")
            set_status(label, "connecting")
            time.sleep(0.6)
            launch_script(label, script_path)
            _clear_launch_query_param()
            
# =========================
# Handle URL-triggered launches
# =========================
qp = _get_all_query_params()
launch_id = qp.get("launch")
ts = qp.get("ts")
if isinstance(launch_id, list):
    launch_id = launch_id[0] if launch_id else None
if isinstance(ts, list):
    ts = ts[0] if ts else None

if launch_id:
    _clear_launch_query_param()
    if ts and ts != st.session_state.last_launch_ts:
        st.session_state.last_launch_ts = ts
        label = ID_TO_LABEL.get(launch_id)
        if label:
            excel_file = excel_files.get(label, "the corresponding Excel file")
            show_fill_modal(label, excel_file, scripts[label])

# =========================
# Cards grid
# =========================
for idx, (label, script_path) in enumerate(scripts.items()):
    accent = ACCENTS[idx % len(ACCENTS)]
    state = st.session_state.btn_status.get(label, "")

    btn_text_1 = "started!" if state == "started" else ("Connecting…" if state == "connecting" else "Launch")
    btn_text_2 = (
        "Successfully started"
        if state == "started"
        else ("Please wait a moment" if state == "connecting" else f"Runs {Path(script_path).name}")
    )
    arrow = "✓" if state == "started" else ("…" if state == "connecting" else "→")
    is_disabled = state != ""
    nonce = int(time.time() * 1000) + idx
    href = "#" if is_disabled else f"?launch=id{LABEL_TO_ID[label]}&ts={nonce}"

    if label == "📂 Consolidated Report":
        cols = st.columns([1, 2, 1])
        target_col = cols[1]
    else:
        if idx % 2 == 0:
            cols = st.columns(2)
        target_col = cols[idx % 2]
    with target_col:
        dot_cls = "dot-green" if state == "started" else ("dot-yellow" if state == "connecting" else "dot-gray")
        card_desc = (
            "This will consolidate all the reports." if label == "📂 Consolidated Report"
            else "This will download the files for the above Report Type  ."
        )
        st.markdown(
            f"""
<div class="glass-card" style="--accent500:{accent['c500']}; --accent400:{accent['c400']}; --accentGlow:{accent['glow']}; --accentShadow:{accent['glow']}">
  <div class="accent-ring"></div>
  <div class="card-head">
    <h3 class="card-title">{label}</h3>
    <div class="card-desc">{card_desc}</div>
  </div>
  <a class="fancy-btn{' disabled' if is_disabled else ''}" href="{href}" target="_self" aria-disabled="{str(is_disabled).lower()}">
    <span class="btn-row">
      <span class="btn-left">
        <span class="btn-icon">{'✓' if state=='started' else ('⏳' if state=='connecting' else '▶')}</span>
        <span>
          <div class="btn-text-1">{btn_text_1}</div>
          <div class="btn-text-2">{btn_text_2}</div>
        </span>
      </span>
      <span class="arrow">{arrow}</span>
    </span>
  </a>
</div>
            """,
            unsafe_allow_html=True,
        )
        st.markdown(
            f"""<div style="margin-top:10px; color: rgba(0,0,0,0.8); font-size: 12.5px;">
  <span class="badge-dot {dot_cls}"></span>
  <span>Status:</span> <b style="color:#000000;">{state}</b>
</div>""",
            unsafe_allow_html=True,
        )
