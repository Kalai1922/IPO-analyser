import io
import os
import pandas as pd
import plotly.express as px
import streamlit as st
from google import genai
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(
    page_title="AI-Powered Hybrid IPO Evaluation Engine",
    page_icon="🤖",
    layout="wide",
)

# ---------------------------------------------------------
# SIDEBAR: API KEY & TARGET COMPANY INPUT
# ---------------------------------------------------------
with st.sidebar:
    st.header("🔑 API Configuration")

    default_key = ""
    try:
        default_key = (
            st.secrets.get("GEMINI_API_KEY", "")
            or st.secrets.get("GOOGLE_API_KEY", "")
            or ""
        )
    except:
        pass

    user_api_key = st.text_input(
        "Gemini API Key",
        value=default_key,
        type="password",
        help="Paste your Gemini API key from Google AI Studio here.",
    )

    active_key = (
        user_api_key
        or os.environ.get("GEMINI_API_KEY")
        or os.environ.get("GOOGLE_API_KEY")
    )
    client = genai.Client(api_key=active_key) if active_key else None

    st.markdown("---")
    st.header("🔍 Target Company")
    input_company = st.text_input(
        "Company / IPO Name",
        "Jio Platforms Limited",
        help="Enter any upcoming IPO or corporation.",
    )

    auto_research = st.button("🚀 Auto-Research with AI", type="primary")

    st.markdown("---")
    st.header("📌 Issue Parameters")
    company_name = st.text_input(
        "Confirmed Name", value=st.session_state.get("res_name", input_company)
    )
    sector = st.text_input(
        "Sector",
        value=st.session_state.get("res_sector", "Telecom & Digital Services"),
    )
    issue_size = st.number_input(
        "Issue Size (₹ Cr)",
        min_value=1.0,
        value=float(st.session_state.get("res_size", 40000.0)),
    )
    fresh_issue_pct = st.slider("Fresh Issue Share (%)", 0, 100, 50)
    ofs_pct = 100 - fresh_issue_pct

    st.markdown("---")
    st.header("⚙️ Framework Weights")
    w_biz = st.slider("Business Quality Weight", 0.0, 0.5, 0.25, 0.05)
    w_fin = st.slider("Financial Health Weight", 0.0, 0.5, 0.30, 0.05)
    w_val = st.slider("Valuation Discipline Weight", 0.0, 0.5, 0.20, 0.05)
    w_gov = st.slider("Governance & Proceeds Weight", 0.0, 0.5, 0.15, 0.05)
    w_sen = st.slider("Market Sentiment Weight", 0.0, 0.5, 0.10, 0.05)

st.title("🤖 AI-Powered Hybrid IPO Evaluation Engine")
st.caption(
    "Type an IPO or company name, paste your API key in the sidebar, and auto-populate metrics using AI."
)

# ---------------------------------------------------------
# AI RESEARCH LOGIC (GEMINI API)
# ---------------------------------------------------------
if auto_research:
    if not client:
        st.error(
            "⚠️ Please enter a valid Gemini API Key in the sidebar input box!"
        )
    else:
        with st.spinner(
            f"Analyzing financials, business model, and valuations for {input_company}..."
        ):
            prompt = f"""
            You are an expert financial analyst. Provide a structured evaluation for the company/IPO: '{input_company}'.
            Return ONLY a raw string formatted as key-value pairs separated by '::', with each pair on a new line:
            SECTOR:: <industry sector>
            ISSUE_SIZE:: <approximate issue size in crores as a number only, e.g. 5000>
            STRENGTHS:: <1-2 sentences summarizing core business strengths and growth drivers>
            RISKS:: <1-2 sentences summarizing key execution or valuation risks>
            BIZ_SCORE:: <suggested score from 1 to 10 for business quality>
            FIN_SCORE:: <suggested score from 1 to 10 for financial health>
            VAL_SCORE:: <suggested score from 1 to 10 for valuation reasonableness>
            GOV_SCORE:: <suggested score from 1 to 10 for governance/proceeds>
            SEN_SCORE:: <suggested score from 1 to 10 for market sentiment/GMP>
            """
            try:
                # Updated to gemini-1.5-flash for universal compatibility
                response = client.models.generate_content(
                    model="gemini-1.5-flash", contents=prompt
                )
                text = response.text

                lines = text.strip().split("\n")
                res_data = {}
                for line in lines:
                    if "::" in line:
                        k, v = line.split("::", 1)
                        res_data[k.strip()] = v.strip()

                st.session_state["res_name"] = input_company
                st.session_state["res_sector"] = res_data.get(
                    "SECTOR", "Diversified"
                )
                try:
                    st.session_state["res_size"] = float(
                        "".join(
                            ch
                            for ch in res_data.get("ISSUE_SIZE", "500")
                            if ch.isdigit() or ch == "."
                        )
                    )
                except:
                    st.session_state["res_size"] = 500.0

                st.session_state["res_strengths"] = res_data.get(
                    "STRENGTHS", "Strong market position."
                )
                st.session_state["res_risks"] = res_data.get(
                    "RISKS", "Competitive pressures and margin risks."
                )

                st.session_state["s_biz"] = float(
                    res_data.get("BIZ_SCORE", 7)
                )
                st.session_state["s_fin"] = float(
                    res_data.get("FIN_SCORE", 7)
                )
                st.session_state["s_val"] = float(
                    res_data.get("VAL_SCORE", 6)
                )
                st.session_state["s_gov"] = float(
                    res_data.get("GOV_SCORE", 7)
                )
                st.session_state["s_sen"] = float(
                    res_data.get("SEN_SCORE", 7)
                )

                st.success(
                    "✅ AI Research completed! Sliders and notes have been populated."
                )
                st.rerun()
            except Exception as e:
                st.error(f"Error connecting to Gemini API: {e}")

# ---------------------------------------------------------
# MAIN TABS
# ---------------------------------------------------------
tab_eval, tab_peers, tab_summary = st.tabs(
    ["📝 Phase Scoring", "⚔️ Peer Comparison", "📈 Executive Dashboard & Export"]
)

# ---------------------------------------------------------
# TAB 1: PHASE SCORING
# ---------------------------------------------------------
with tab_eval:
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("Phase 1: Business Quality & Moat")
        score_biz = st.slider(
            "Business Quality Score",
            1.0,
            10.0,
            float(st.session_state.get("s_biz", 7.0)),
        )

        st.subheader("Phase 2: Financial Health & Performance")
        score_fin = st.slider(
            "Financial Health Score",
            1.0,
            10.0,
            float(st.session_state.get("s_fin", 7.0)),
        )

        st.subheader("Phase 3: Valuation Discipline")
        score_val = st.slider(
            "Valuation Discipline Score",
            1.0,
            10.0,
            float(st.session_state.get("s_val", 6.5)),
        )

    with col2:
        st.subheader("Phase 4: Governance & Proceeds")
        score_gov = st.slider(
            "Governance & Proceeds Score",
            1.0,
            10.0,
            float(st.session_state.get("s_gov", 7.5)),
        )

        st.subheader("Phase 5: Market Sentiment & Demand")
        score_sen = st.slider(
            "Market Sentiment Score",
            1.0,
            10.0,
            float(st.session_state.get("s_sen", 7.0)),
        )

        st.markdown("---")
        st.subheader("📋 Qualitative Notes & AI Findings")
        notes_biz = st.text_area(
            "Key Strengths",
            st.session_state.get(
                "res_strengths",
                "Strong market share, high scalability, robust digital ecosystem.",
            ),
        )
        notes_risks = st.text_area(
            "Key Risks",
            st.session_state.get(
                "res_risks",
                "Regulatory shifts, intense competition, high valuation multiples.",
            ),
        )

# ---------------------------------------------------------
# COMPUTATIONS
# ---------------------------------------------------------
final_score = (
    (score_biz * w_biz)
    + (score_fin * w_fin)
    + (score_val * w_val)
    + (score_gov * w_gov)
    + (score_sen * w_sen)
) / (w_biz + w_fin + w_val + w_gov + w_sen)

if final_score >= 8.0:
    verdict = "🟢 HIGH CONVICTION (APPLY)"
elif final_score >= 6.0:
    verdict = "🟡 SELECTIVE / LISTING GAIN ONLY"
else:
    verdict = "🔴 AVOID / WEAK FUNDAMENTALS"

# ---------------------------------------------------------
# TAB 2: PEER COMPARISON TABLE
# ---------------------------------------------------------
with tab_peers:
    st.subheader("⚔️ Relative Peer Valuation Matrix")
    st.caption("Customize peer metrics to verify relative positioning.")

    default_peers_data = {
        "Company": [company_name, "Industry Peer A", "Industry Peer B"],
        "P/E / EV-EBITDA": [32.5, 41.0, 24.2],
        "ROE / ROCE (%)": [18.2, 21.5, 14.1],
        "Revenue Growth (%)": [15.0, 12.0, 9.5],
        "Debt/Equity": [0.35, 0.20, 0.55],
    }

    peer_df = pd.DataFrame(default_peers_data)
    edited_peer_df = st.data_editor(
        peer_df, num_rows="dynamic", use_container_width=True
    )

# ---------------------------------------------------------
# TAB 3: EXECUTIVE DASHBOARD & EXPORT
# ---------------------------------------------------------
with tab_summary:
    st.subheader(f"Summary Scorecard: {company_name}")

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Weighted Overall Score", f"{final_score:.2f} / 10")
    m2.metric("Verdict", verdict)
    m3.metric("Issue Size", f"₹{issue_size:,.0f} Cr")
    m4.metric("Fresh / OFS Ratio", f"{fresh_issue_pct}% / {ofs_pct}%")

    st.markdown("---")

    c1, c2 = st.columns([1, 1])

    with c1:
        st.markdown("### Pillar Breakdown")
        radar_df = pd.DataFrame(
            dict(
                Pillar=[
                    "Business Quality",
                    "Financial Health",
                    "Valuation",
                    "Governance",
                    "Sentiment",
                ],
                Score=[score_biz, score_fin, score_val, score_gov, score_sen],
            )
        )
        fig = px.line_polar(
            radar_df,
            r="Score",
            theta="Pillar",
            line_close=True,
            range_r=[0, 10],
            title="5-Pillar Structural Strength",
        )
        fig.update_traces(fill="toself", fillcolor="rgba(13, 110, 253, 0.2)")
        st.plotly_chart(fig, use_container_width=True)

    with c2:
        st.markdown("### Score Summary Table")
        summary_scores_df = pd.DataFrame(
            {
                "Pillar": [
                    "Business Quality",
                    "Financial Health",
                    "Valuation Discipline",
                    "Governance & Proceeds",
                    "Market Sentiment",
                ],
                "Raw Score (/10)": [
                    score_biz,
                    score_fin,
                    score_val,
                    score_gov,
                    score_sen,
                ],
                "Weight": [w_biz, w_fin, w_val, w_gov, w_sen],
                "Weighted Score": [
                    score_biz * w_biz,
                    score_fin * w_fin,
                    score_val * w_val,
                    score_gov * w_gov,
                    score_sen * w_sen,
                ],
            }
        )
        st.dataframe(summary_scores_df, use_container_width=True, hide_index=True)

        st.markdown("**Qualitative Summary:**")
        st.info(f"**Strengths:** {notes_biz}\n\n**Risks:** {notes_risks}")

    # ---------------------------------------------------------
    # EXCEL EXPORT ENGINE (OPENPYXL)
    # ---------------------------------------------------------
    st.markdown("---")
    st.subheader("📥 Export Complete Analysis to Excel Template")


    def generate_formatted_excel():
        output = io.BytesIO()
        wb = Workbook()

        ws1 = wb.active
        ws1.title = "Executive Summary"
        ws1.views.sheetView[0].showGridLines = True

        header_fill = PatternFill(
            start_color="1F4E78", end_color="1F4E78", fill_type="solid"
        )
        header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
        title_font = Font(name="Calibri", size=16, bold=True, color="1F4E78")
        bold_font = Font(name="Calibri", size=11, bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        ws1["A1"] = f"IPO EVALUATION REPORT: {company_name.upper()}"
        ws1["A1"].font = title_font
        ws1["A2"] = f"Sector: {sector} | Issue Size: ₹{issue_size} Cr"
        ws1["A2"].font = Font(name="Calibri", size=11, italic=True)

        ws1["A4"] = "Metric"
        ws1["B4"] = "Value"
        for col in ["A", "B"]:
            ws1[f"{col}4"].fill = header_fill
            ws1[f"{col}4"].font = header_font

        overview_data = [
            ("Weighted Score", f"{final_score:.2f} / 10"),
            ("Final Verdict", verdict),
            ("Fresh Issue Share", f"{fresh_issue_pct}%"),
            ("OFS Share", f"{ofs_pct}%"),
        ]

        for idx, (k, v) in enumerate(overview_data, start=5):
            ws1[f"A{idx}"] = k
            ws1[f"B{idx}"] = v
            ws1[f"A{idx}"].font = bold_font
            ws1[f"A{idx}"].border = thin_border
            ws1[f"B{idx}"].border = thin_border

        ws1["A11"] = "Pillar"
        ws1["B11"] = "Score (/10)"
        ws1["C11"] = "Weight"
        ws1["D11"] = "Weighted Contribution"
        for col in ["A", "B", "C", "D"]:
            ws1[f"{col}11"].fill = header_fill
            ws1[f"{col}11"].font = header_font

        for row_idx, row in summary_scores_df.iterrows():
            r = row_idx + 12
            ws1[f"A{r}"] = row["Pillar"]
            ws1[f"B{r}"] = row["Raw Score (/10)"]
            ws1[f"C{r}"] = row["Weight"]
            ws1[f"D{r}"] = row["Weighted Score"]
            for col in ["A", "B", "C", "D"]:
                ws1[f"{col}{r}"].border = thin_border

        ws1["A18"] = "Key Strengths:"
        ws1["A18"].font = bold_font
        ws1["B18"] = notes_biz

        ws1["A19"] = "Key Risks:"
        ws1["A19"].font = bold_font
        ws1["B19"] = notes_risks

        ws2 = wb.create_sheet(title="Peer Comparison")
        ws2.views.sheetView[0].showGridLines = True

        ws2["A1"] = "PEER VALUATION & RELATIVE METRICS"
        ws2["A1"].font = title_font

        for r in dataframe_to_rows(edited_peer_df, index=False, header=True):
            ws2.append(r)

        for cell in ws2[3]:
            cell.fill = header_fill
            cell.font = header_font

        for ws in [ws1, ws2]:
            for col in ws.columns:
                max_len = max(len(str(cell.value or "")) for cell in col)
                col_letter = col[0].column_letter
                ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

        wb.save(output)
        return output.getvalue()


    excel_data = generate_formatted_excel()

    st.download_button(
        label=f"📄 Download {company_name} IPO Report (.xlsx)",
        data=excel_data,
        file_name=f"{company_name.lower().replace(' ', '_')}_ipo_analysis.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
