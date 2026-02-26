"""
Excel VBA to Python Converter — Streamlit Frontend

Launch with:
    streamlit run streamlit_app.py
"""
from __future__ import annotations

import logging
import os
import tempfile
from pathlib import Path

import streamlit as st

from config import get_config
from data_exporter import DataExporter
from formula_extractor import FormulaExtractor
from llm_converter import VBAToPythonConverter, ConversionResult
from vba_extractor import VBAExtractor
from workbook_analyzer import WorkbookAnalyzer

# ---------------------------------------------------------------------------
# Configuration & Logging
# ---------------------------------------------------------------------------
cfg = get_config()

# Constants
_MIME_PYTHON = "text/x-python"

logging.basicConfig(
    level=getattr(logging, cfg.LOG_LEVEL, logging.INFO),
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

VBA_EXTENSIONS = cfg.VBA_EXTENSIONS
DATA_EXTENSIONS = cfg.DATA_EXTENSIONS
ALL_EXTENSIONS = cfg.ALL_EXTENSIONS

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _allowed(filename: str, extensions: set[str]) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in extensions


def _save_temp(uploaded_file) -> str:
    """Write the uploaded file to a temp path and return it."""
    suffix = Path(uploaded_file.name).suffix
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp.write(uploaded_file.getvalue())
    tmp.close()
    return tmp.name


def _cleanup(path: str) -> None:
    try:
        if os.path.exists(path):
            os.remove(path)
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Page Config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Excel VBA → Python Converter",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ---------------------------------------------------------------------------
# Custom CSS
# ---------------------------------------------------------------------------
st.markdown(
    """
    <style>
    /* code blocks */
    .stCodeBlock { border-radius: 8px; }
    /* metric cards */
    div[data-testid="stMetric"] {
        background-color: #262730;
        padding: 12px 16px;
        border-radius: 8px;
    }
    /* tighter tabs */
    .stTabs [data-baseweb="tab-list"] { gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        padding: 8px 16px;
        border-radius: 6px 6px 0 0;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------------------------------------------------------------------
# Sidebar — Settings
# ---------------------------------------------------------------------------
with st.sidebar:
    st.image("https://img.icons8.com/color/96/microsoft-excel-2019.png", width=64)
    st.title("Settings")

    provider_choice = st.selectbox(
        "Conversion Engine",
        options=["offline", "anthropic", "openai"],
        index=0,
        format_func=lambda p: {
            "offline": "⚡ Offline (rule-based, no API key)",
            "anthropic": "🤖 Anthropic Claude (API key required)",
            "openai": "🤖 OpenAI GPT (API key required)",
        }.get(p, p),
        help="Choose 'Offline' for free, instant conversion without an API key.",
    )

    target_library = st.selectbox(
        "Target Python Library",
        options=["pandas", "polars"],
        index=0,
        help="Library used for data operations in converted code.",
    )

    if provider_choice == "offline":
        st.success("**Engine:** Offline (rule-based)  \nNo API key needed.")
    else:
        provider_display = provider_choice.capitalize()
        st.info(f"**LLM Provider:** {provider_display}  \n**Model:** `{cfg.LLM_MODEL}`")

    st.divider()
    st.caption(
        f"Max upload size: **{cfg.MAX_FILE_SIZE_MB} MB**  \n"
        f"VBA extensions: {', '.join(sorted(VBA_EXTENSIONS))}  \n"
        f"Data extensions: {', '.join(sorted(DATA_EXTENSIONS))}"
    )

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.title("🔄 Excel VBA → Python Converter")
st.markdown(
    "Upload an Excel workbook containing **VBA macros** or **formulas** and convert "
    "them to clean, idiomatic Python powered by LLMs."
)

# ---------------------------------------------------------------------------
# File Uploader
# ---------------------------------------------------------------------------
uploaded_file = st.file_uploader(
    "Upload your Excel file",
    type=sorted(ALL_EXTENSIONS),
    help="Supported formats: " + ", ".join(f".{e}" for e in sorted(ALL_EXTENSIONS)),
)

if uploaded_file is None:
    st.info("👆 Upload an Excel file to get started.")
    st.stop()

filename = uploaded_file.name
ext = filename.rsplit(".", 1)[1].lower() if "." in filename else ""

# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------
tab_vba, tab_formulas, tab_data, tab_full = st.tabs(
    ["📝 VBA Extraction & Conversion", "📐 Formula Extraction", "📊 Data Export", "🔍 Full Analysis"]
)

# ===========================  TAB 1 — VBA  =================================
with tab_vba:
    st.header("VBA Extraction & Conversion")

    if ext not in VBA_EXTENSIONS:
        st.warning(
            f"`.{ext}` files typically don't contain VBA macros. "
            "Try a **.xlsm**, **.xls**, **.xlsb**, **.xla**, or **.xlam** file."
        )

    if st.button("🔍 Extract VBA Modules", key="extract_vba"):
        filepath = _save_temp(uploaded_file)
        try:
            with st.spinner("Extracting VBA modules…"):
                extractor = VBAExtractor(filepath)
                modules = extractor.extract_all()

            if not modules:
                st.warning("No VBA modules found in this file.")
            else:
                st.success(f"Found **{len(modules)}** VBA module(s).")
                st.session_state["vba_modules"] = modules
        except Exception as exc:
            st.error(f"Extraction failed: {exc}")
        finally:
            _cleanup(filepath)

    # Show extracted modules --------------------------------------------------
    modules: list[dict] = st.session_state.get("vba_modules", [])

    if modules:
        st.divider()

        # Convert-all button
        col_header, col_btn = st.columns([3, 1])
        with col_header:
            st.subheader("Extracted Modules")
        with col_btn:
            convert_all = st.button("⚡ Convert All", key="convert_all_vba")

        if convert_all:
            converter = VBAToPythonConverter(provider=provider_choice)
            progress = st.progress(0, text="Converting…")
            converted: list[dict] = []

            for idx, mod in enumerate(modules):
                try:
                    result: ConversionResult = converter.convert_with_result(
                        mod["code"],
                        mod.get("name", "module"),
                        target_library,
                    )
                    converted.append({**mod, "result": result})
                except Exception as exc:
                    converted.append({
                        **mod,
                        "result": ConversionResult(
                            success=False, python_code="", error=str(exc)
                        ),
                    })
                progress.progress(
                    (idx + 1) / len(modules),
                    text=f"Converted {idx + 1}/{len(modules)}",
                )

            st.session_state["converted_modules"] = converted
            progress.empty()
            st.success(f"Converted **{len(converted)}** module(s).")

        # Display each module -------------------------------------------------
        converted_modules: list[dict] = st.session_state.get("converted_modules", [])

        for i, mod in enumerate(modules):
            with st.expander(
                f"📄 {mod.get('name', f'Module {i+1}')} — {mod.get('type', 'Unknown')}",
                expanded=(i == 0),
            ):
                vba_col, py_col = st.columns(2)

                with vba_col:
                    st.markdown("**VBA Code**")
                    st.code(mod["code"], language="vb", line_numbers=True)

                with py_col:
                    # Check if we have a converted version
                    conv = next(
                        (
                            c
                            for c in converted_modules
                            if c.get("name") == mod.get("name")
                        ),
                        None,
                    )

                    if conv and conv["result"].success:
                        st.markdown("**Python Code**")
                        st.code(
                            conv["result"].python_code,
                            language="python",
                            line_numbers=True,
                        )
                        if conv["result"].conversion_notes:
                            with st.popover("📝 Conversion Notes"):
                                for note in conv["result"].conversion_notes:
                                    st.markdown(f"- {note}")
                        if conv["result"].tokens_used:
                            st.caption(f"Tokens used: {conv['result'].tokens_used}")

                        st.download_button(
                            "⬇️ Download .py",
                            data=conv["result"].python_code,
                            file_name=f"{mod.get('name', 'module')}.py",
                            mime=_MIME_PYTHON,
                            key=f"dl_mod_{i}",
                        )
                    elif conv and not conv["result"].success:
                        st.error(f"Conversion failed: {conv['result'].error}")
                    else:
                        # Single-module convert button
                        if st.button(
                            "🔄 Convert to Python", key=f"convert_single_{i}"
                        ):
                            with st.spinner("Converting…"):
                                try:
                                    converter = VBAToPythonConverter(provider=provider_choice)
                                    res = converter.convert_with_result(
                                        mod["code"],
                                        mod.get("name", "module"),
                                        target_library,
                                    )
                                    if res.success:
                                        st.markdown("**Python Code**")
                                        st.code(
                                            res.python_code,
                                            language="python",
                                            line_numbers=True,
                                        )
                                        st.download_button(
                                            "⬇️ Download .py",
                                            data=res.python_code,
                                            file_name=f"{mod.get('name', 'module')}.py",
                                            mime=_MIME_PYTHON,
                                            key=f"dl_single_{i}",
                                        )
                                    else:
                                        st.error(f"Conversion failed: {res.error}")
                                except Exception as exc:
                                    st.error(f"Error: {exc}")


# ===========================  TAB 2 — FORMULAS  ============================
with tab_formulas:
    st.header("Formula Extraction & Conversion")

    if ext not in DATA_EXTENSIONS:
        st.warning(f"`.{ext}` is not a supported data format.")

    if st.button("📐 Extract Formulas", key="extract_formulas"):
        filepath = _save_temp(uploaded_file)
        try:
            with st.spinner("Extracting formulas…"):
                fx = FormulaExtractor(filepath)
                formulas = fx.extract_all_formulas()
                stats = fx.get_formula_statistics(formulas)

            if not formulas:
                st.warning("No formulas found.")
            else:
                st.success(f"Found **{len(formulas)}** formula(s).")
                st.session_state["formulas"] = formulas
                st.session_state["formula_stats"] = stats
        except Exception as exc:
            st.error(f"Extraction failed: {exc}")
        finally:
            _cleanup(filepath)

    formulas = st.session_state.get("formulas", [])
    stats = st.session_state.get("formula_stats", {})

    if formulas:
        st.divider()

        # Statistics ----------------------------------------------------------
        if stats:
            st.subheader("Statistics")
            stat_cols = st.columns(4)
            stat_cols[0].metric("Total Formulas", stats.get("total_formulas", 0))
            stat_cols[1].metric("Unique Functions", stats.get("unique_functions_used", 0))
            stat_cols[2].metric("Sheets", stats.get("sheets_with_formulas", 0))
            stat_cols[3].metric("Array Formulas", stats.get("array_formulas", 0))

            top_funcs = stats.get("most_common_functions", [])
            if top_funcs:
                st.markdown("**Most-used functions:** " + ", ".join(
                    f"`{fn}` ({cnt})" for fn, cnt in top_funcs[:10]
                ))

        # Table of formulas ---------------------------------------------------
        st.subheader("Formulas")
        for idx, f in enumerate(formulas):
            with st.expander(
                f"{f.sheet_name}!{f.cell_address} — `{f.formula[:60]}{'…' if len(f.formula) > 60 else ''}`",
                expanded=False,
            ):
                st.code(f.formula, language="excel")
                st.caption(
                    f"**Type:** {f.formula_type}  |  "
                    f"**Functions:** {', '.join(f.contains_functions) or '—'}  |  "
                    f"**Deps:** {', '.join(f.dependencies) or '—'}"
                )

                if st.button(
                    "🔄 Convert to Python", key=f"convert_formula_{idx}"
                ):
                    with st.spinner("Converting formula…"):
                        try:
                            converter = VBAToPythonConverter(provider=provider_choice)
                            py = converter.convert_formula(
                                f.formula, f.cell_address, f.sheet_name
                            )
                            st.code(py, language="python", line_numbers=True)
                        except Exception as exc:
                            st.error(f"Conversion failed: {exc}")


# ===========================  TAB 3 — DATA EXPORT  =========================
with tab_data:
    st.header("Data Export")

    if ext not in DATA_EXTENSIONS:
        st.warning(f"`.{ext}` is not a supported data format.")

    if st.button("📊 Export Sheet Data", key="export_data"):
        filepath = _save_temp(uploaded_file)
        try:
            with st.spinner("Exporting sheet data…"):
                exporter = DataExporter(filepath)
                result = exporter.export_all_sheets()

            st.session_state["export_result"] = result
            st.success(
                f"Exported **{len(result.sheet_data)}** sheet(s)."
            )
        except Exception as exc:
            st.error(f"Export failed: {exc}")
        finally:
            _cleanup(filepath)

    export_result = st.session_state.get("export_result")

    if export_result:
        st.divider()

        # Metadata ------------------------------------------------------------
        meta = export_result.metadata
        if meta:
            mcols = st.columns(3)
            mcols[0].metric("Sheets", meta.get("total_sheets", "—"))
            mcols[1].metric("Total Rows", meta.get("total_rows", "—"))
            mcols[2].metric("Total Columns", meta.get("total_columns", "—"))

        # Generated Python code -----------------------------------------------
        st.subheader("Generated Python Code")
        st.code(export_result.python_code, language="python", line_numbers=True)
        st.download_button(
            "⬇️ Download data_loader.py",
            data=export_result.python_code,
            file_name="data_loader.py",
            mime=_MIME_PYTHON,
            key="dl_data_code",
        )

        # Preview DataFrames --------------------------------------------------
        st.subheader("Sheet Previews")
        for sd in export_result.sheet_data:
            with st.expander(f"📄 {sd.sheet_name} — {sd.data_range}", expanded=False):
                st.dataframe(sd.dataframe, use_container_width=True)
                st.caption(
                    f"**Rows:** {len(sd.dataframe)}  |  "
                    f"**Columns:** {len(sd.dataframe.columns)}  |  "
                    f"**Header detected:** {'Yes' if sd.has_header else 'No'}"
                )


# ===========================  TAB 4 — FULL ANALYSIS  =======================
with tab_full:
    st.header("Full Workbook Analysis")
    st.markdown(
        "Run **all** analysis steps at once: VBA extraction, formula extraction, "
        "data export, and dependency mapping — then generate a complete Python "
        "recreation script."
    )

    if st.button("🔍 Analyze Workbook", key="analyze_all", type="primary"):
        filepath = _save_temp(uploaded_file)
        try:
            with st.spinner("Running full workbook analysis — this may take a moment…"):
                analyzer = WorkbookAnalyzer(filepath)
                analysis = analyzer.analyze_complete()
                report = analyzer.generate_analysis_report(analysis)

            st.session_state["analysis"] = analysis
            st.session_state["analysis_report"] = report
            st.success("Analysis complete!")
        except Exception as exc:
            st.error(f"Analysis failed: {exc}")
        finally:
            _cleanup(filepath)

    analysis = st.session_state.get("analysis")
    report = st.session_state.get("analysis_report")

    if analysis:
        st.divider()

        # Summary metrics -----------------------------------------------------
        st.subheader("Summary")
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("VBA Modules", len(analysis.vba_modules) if analysis.vba_modules else 0)
        m2.metric("Formulas", len(analysis.formulas) if analysis.formulas else 0)
        m3.metric(
            "Sheets",
            len(analysis.data_export.sheet_data) if analysis.data_export else 0,
        )
        m4.metric("Has VBA", "Yes" if analysis.has_vba else "No")

        # Report --------------------------------------------------------------
        if report:
            st.subheader("Analysis Report")
            st.text(report)

        # Generated Python script ---------------------------------------------
        if analysis.python_script:
            st.subheader("Generated Python Script")
            st.code(analysis.python_script, language="python", line_numbers=True)
            st.download_button(
                "⬇️ Download workbook_recreation.py",
                data=analysis.python_script,
                file_name="workbook_recreation.py",
                mime=_MIME_PYTHON,
                key="dl_full_script",
            )
