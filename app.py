from pathlib import Path
import shutil
import uuid
import streamlit as st
from engine import process_workbook

BASE = Path(__file__).resolve().parent
CSS_PATH = BASE / "static" / "style.css"


def safe_unlink(path: Path) -> None:
    try:
        path.unlink()
    except FileNotFoundError:
        pass


st.set_page_config(
    page_title="TDO Store Matching",
    layout="wide",
)

# Load exact style.css + Streamlit layout bindings to mirror index.html 1:1
base_css = CSS_PATH.read_text(encoding="utf-8") if CSS_PATH.exists() else ""

st.markdown(
    f"""
    <style>
    {base_css}

    /* Streamlit container & widget overrides to match exact index.html UI */
    #MainMenu, header[data-testid="stHeader"], footer {{
        display: none !important;
    }}
    .stApp {{
        background: radial-gradient(circle at 20% 10%, #1242b0 0, transparent 30%), linear-gradient(135deg, #061b60, #020a25 72%) !important;
        font-family: Inter, Segoe UI, Arial, sans-serif !important;
        color: #fff !important;
    }}
    .stApp::before {{
        content: "";
        position: fixed;
        inset: 0;
        background-image: radial-gradient(#18a8ff 1px, transparent 1px);
        background-size: 28px 28px;
        opacity: .13;
        pointer-events: none;
    }}
    .block-container {{
        position: relative;
        max-width: 1200px !important;
        margin: auto !important;
        padding: 42px 28px !important;
    }}
    /* Right column styled as the exact .card container */
    div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2) {{
        background: linear-gradient(145deg, #ffffffee, #eaf1ffdf) !important;
        color: #10193a !important;
        border: 1px solid #fff !important;
        padding: 38px !important;
        border-radius: 24px !important;
        box-shadow: 0 30px 80px #0008 !important;
        backdrop-filter: blur(18px) !important;
        text-align: center !important;
    }}
    /* File uploader styled as .drop */
    [data-testid="stFileUploader"] {{
        margin: 16px 0 6px !important;
    }}
    [data-testid="stFileUploader"] section {{
        min-height: 150px !important;
        border: 2px dashed #77a7ef !important;
        border-radius: 16px !important;
        background: #f7faff !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        flex-direction: column !important;
        padding: 18px !important;
    }}
    [data-testid="stFileUploader"] section * {{
        color: #0956ed !important;
    }}
    [data-testid="stFileUploader"] small {{
        color: #7c87a6 !important;
    }}
    [data-testid="stFileUploader"] button {{
        background: #e8f0fe !important;
        color: #0956ed !important;
        border: 1px solid #77a7ef !important;
        border-radius: 8px !important;
        padding: 6px 14px !important;
        font-weight: 700 !important;
        width: auto !important;
    }}
    /* Execute analysis button styled as button */
    div.stButton > button {{
        width: 100% !important;
        border: 0 !important;
        border-radius: 12px !important;
        padding: 15px !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        color: #fff !important;
        background: linear-gradient(90deg, #078a2a, #00c853) !important;
        cursor: pointer !important;
        margin-top: 6px !important;
    }}
    div.stButton > button:disabled {{
        opacity: .45 !important;
        cursor: not-allowed !important;
    }}
    /* Download button styled as .download */
    div[data-testid="stDownloadButton"] > button {{
        width: 100% !important;
        border: 0 !important;
        border-radius: 12px !important;
        padding: 15px !important;
        font-weight: 800 !important;
        font-size: 16px !important;
        color: #fff !important;
        background: linear-gradient(90deg, #074fd6, #0aa9ef) !important;
        cursor: pointer !important;
        margin-top: 8px !important;
    }}
    .card-header h3 {{
        font-size: 28px;
        margin: 20px 0 5px;
        color: #10193a;
    }}
    .card-header .muted {{
        color: #66708e;
        margin-bottom: 0;
    }}
    .status-box {{
        min-height: 28px;
        margin: 14px 0 8px;
        color: #334;
        font-size: 14px;
        font-weight: 600;
    }}
    .metrics-strip {{
        display: flex;
        justify-content: space-around;
        gap: 8px;
        margin: 8px 0 12px;
        padding: 10px;
        border-radius: 12px;
        background: #eef4ff;
        border: 1px solid #cfe0ff;
        color: #10193a;
        font-size: 13px;
    }}
    .metrics-strip span b {{
        color: #074fd6;
        font-size: 15px;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# Exact background orbs + header from index.html
st.markdown(
    """
    <div class="orb a"></div>
    <div class="orb b"></div>
    <header style="margin-bottom: 36px;">
        <div class="brand">NIQ</div>
        <div>
            <h1>Trade Dimensions Operations</h1>
            <p>Store Matching Automation Engine</p>
        </div>
    </header>
    """,
    unsafe_allow_html=True,
)

# Exact 1.15fr / 0.85fr hero grid layout
col_left, col_right = st.columns([1.15, 0.85], gap="large")

with col_left:
    st.markdown(
        """
        <div class="copy" style="padding-top: 28px;">
            <span class="eyebrow">QUALITY ANALYSIS</span>
            <h2>Upload. Validate. Match.<br><strong>Download with confidence.</strong></h2>
        </div>
        """,
        unsafe_allow_html=True,
    )

with col_right:
    st.markdown(
        """
        <div class="card-header">
            <div class="shield">✓</div>
            <h3>Upload your workbook</h3>
            <p class="muted">Excel .xlsx, one sheet or multiple sheets supported</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    uploaded = st.file_uploader(
        "Choose Excel file",
        type=["xlsx"],
        accept_multiple_files=False,
        label_visibility="collapsed",
        key="xlsx_uploader",
    )

    # Reset previous output if user changes or removes the uploaded file
    current_sig = (uploaded.name, uploaded.size) if uploaded is not None else None
    if st.session_state.get("last_file_sig") != current_sig:
        st.session_state["last_file_sig"] = current_sig
        st.session_state.pop("output_bytes", None)
        st.session_state.pop("result_info", None)
        st.session_state.pop("error_text", None)

    is_valid_xlsx = bool(uploaded is not None and uploaded.name.lower().endswith(".xlsx"))

    if uploaded is not None and not is_valid_xlsx:
        st.markdown('<div class="status-box" style="color:#c62828;">Upload an .xlsx workbook.</div>', unsafe_allow_html=True)
    elif is_valid_xlsx:
        st.markdown(f'<div class="status-box">Selected: {uploaded.name}</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="status-box">or drag and drop here</div>', unsafe_allow_html=True)

    run_btn = st.button("Execute analysis", disabled=not is_valid_xlsx)

    if run_btn and is_valid_xlsx:
        uploads = BASE / "uploads"
        outputs = BASE / "outputs"
        uploads.mkdir(exist_ok=True)
        outputs.mkdir(exist_ok=True)
        token = uuid.uuid4().hex
        inp = uploads / f"{token}.xlsx"
        out = outputs / f"{token}_Processed_Output.xlsx"

        with inp.open("wb") as f:
            shutil.copyfileobj(uploaded, f)

        try:
            with st.spinner("Running analysis..."):
                result = process_workbook(inp, out)
            safe_unlink(inp)
            st.session_state["output_bytes"] = out.read_bytes()
            safe_unlink(out)
            st.session_state["result_info"] = {
                "red_rows": result.red_rows,
                "orange_rows": result.orange_rows,
                "pair_groups": result.pair_groups,
            }
            st.session_state["error_text"] = None
        except Exception as e:
            safe_unlink(inp)
            safe_unlink(out)
            st.session_state.pop("output_bytes", None)
            st.session_state.pop("result_info", None)
            st.session_state["error_text"] = str(e) or repr(e)

    if st.session_state.get("error_text"):
        st.markdown(
            f'<div class="status-box" style="color:#c62828;">{st.session_state["error_text"]}</div>',
            unsafe_allow_html=True,
        )

    if st.session_state.get("output_bytes") and st.session_state.get("result_info"):
        info = st.session_state["result_info"]
        st.markdown(
            f"""
            <div class="status-box" style="color:#078a2a;">Analysis complete. Ready to download.</div>
            <div class="metrics-strip">
                <span>Red Rows: <b>{info["red_rows"]}</b></span>
                <span>Orange Rows: <b>{info["orange_rows"]}</b></span>
                <span>Pair Groups: <b>{info["pair_groups"]}</b></span>
            </div>
            """,
            unsafe_allow_html=True,
        )
        st.download_button(
            label="Download processed workbook",
            data=st.session_state["output_bytes"],
            file_name="Processed_Output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )