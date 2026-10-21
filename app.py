import io
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.dataframe import dataframe_to_rows

st.set_page_config(
    page_title="Hybrid IPO Evaluation Engine", page_icon="📊", layout="wide"
)

# ---------------------------------------------------------
# STYLING & CUSTOM CSS
# ---------------------------------------------------------
st.markdown(
    """
    <style>
    .metric-card {
        background-color: #f8f9fa;
        border: 1px solid #e9ecef;
        padding: 15px;
        border-radius: 8px;
        text-align: center;
    }
    .stProgress > div > div > div > div {
        background-color: #0d6efd;
    }
    </style>
""",
    unsafe_allow_html=True,
)

st.title("📊 Hybrid IPO Evaluation & Analytics Engine")
st.caption("Structured 5-Phase Analysis Framework | Streamlit + Excel Export")

# ---------------------------------------------------------
# SIDEBAR: GENERAL IPO DETAILS
# ---------------------------------------------------------
with st.sidebar:
    st.header("📌 Issue Overview")
    company_name = st.text_input("Company Name", "TechFab Engineers Ltd")
    sector = st.text_input("Sector / Industry", "Precision Engineering")
    issue_size = st.number_input("Issue Size (₹ Cr)", min_value=1.0, value=500.0)
    fresh_issue_pct = st.slider("Fresh Issue Share (%)", 0, 100, 70)
    ofs_pct = 100 - fresh_issue_pct
    st.info(f"OFS Share: **{ofs_pct}%**")

    st.markdown("---")
    st.header("⚙️ Framework Weights")
    w_biz = st.slider("Business Quality Weight", 0.0, 0.5, 0.25, 0.05)
    w_fin = st.slider("Financial Health Weight", 0.0, 0.5, 0.30, 0.05)
    w_val = st.slider("Valuation Discipline Weight", 0.0, 0.5, 0.20, 0.05)
    w_gov = st.slider("Governance & Proceeds Weight", 0.0, 0.5, 0.15, 0.05)
    w_sen = st.slider("Market Sentiment Weight", 0.0, 0.5, 0.10, 0.05)

    total_weight = w_biz + w_fin + w_val + w_gov + w_sen
    if abs(total_weight - 1.0) > 0.01:
        st.warning(f"⚠️ Total weight is {total_weight:.2f}. Ideal sum is 1.00.")

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
        s_biz_model = st.slider(
            "Business Model Clarity & Sustainability", 1, 10, 8
        )
        s_ind_growth = st.slider("Industry Growth & Tailwinds", 1, 10, 7)
        s_moat = st.slider(
            "Competitive Advantage / Entry Barriers", 1, 10, 7
        )
        s_diversification = st.slider(
            "Client / Revenue Diversification", 1, 10, 6
        )
        score_biz = (
            s_biz_model + s_ind_growth + s_moat + s_diversification
        ) / 4.0

        st.subheader("Phase 2: Financial Health & Performance")
        s_rev_growth = st.slider("3-Yr Revenue & PAT Growth Rate", 1, 10, 8)
        s_margins = st.slider("EBITDA / PAT Margin Profile", 1, 10, 7)
        s_returns = st.slider("Return Ratios (ROE >15%, ROCE >15%)", 1, 10, 8)
        s_leverage = st.slider("Debt Profile & Cash Generation", 1, 10, 8)
        score_fin = (
            s_rev_growth + s_margins + s_returns + s_leverage
        ) / 4.0

        st.subheader("Phase 3: Valuation Discipline")
        s_pe_val = st.slider("Post-IPO P/E vs Industry Peers", 1, 10, 7)
        s_pb_ev = st.slider("P/B & EV/EBITDA Relative Comfort", 1, 10, 6)
        s_margin_safety = st.slider("Margin of Safety at Upper Band", 1, 10, 6)
        score_val = (s_pe_val + s_pb_ev + s_margin_safety) / 3.0

    with col2:
        st.subheader("Phase 4: Governance & Proceeds")
        s_proceeds = st.slider(
            "Use of Proceeds (Debt Reduction / Expansion)", 1, 10, 8
        )
        s_ofs = st.slider("OFS Portion Comfort Level", 1, 10, 7)
        s_promoter = st.slider("Promoter Track Record & Integrity", 1, 10, 8)
        score_gov = (s_proceeds + s_ofs + s_promoter) / 3.0

        st.subheader("Phase 5: Market Sentiment & Demand")
        s_qib = st.slider("QIB / NII Institutional Demand", 1, 10, 8)
        s_gmp = st.slider("Grey Market Premium (GMP) Sentiment", 1, 10, 7)
        score_sen = (s_qib + s_gmp) / 2.0

        st.markdown("---")
        st.subheader("📋 Qualitative Notes & Key Observations")
        notes_biz = st.text_area(
            "Key Strengths",
            "Strong order book, niche manufacturing, debt reduction via fresh issue.",
        )
        notes_risks = st.text_area(
            "Key Risks",
            "Customer concentration in top 3 clients, raw material price volatility.",
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
    verdict_color = "green"
elif final_score >= 6.0:
    verdict = "🟡 SELECTIVE / LISTING GAIN ONLY"
    verdict_color = "orange"
else:
    verdict = "🔴 AVOID / WEAK FUNDAMENTALS"
    verdict_color = "red"

# ---------------------------------------------------------
# TAB 2: PEER COMPARISON TABLE
# ---------------------------------------------------------
with tab_peers:
    st.subheader("⚔️ Relative Peer Valuation Matrix")
    st.caption("Input peer metrics to verify relative valuation positioning.")

    default_peers_data = {
        "Company": [company_name, "Peer A (Market Leader)", "Peer B (Mid-Cap)"],
        "P/E Ratio": [22.5, 31.0, 18.2],
        "ROE (%)": [20.2, 22.5, 14.1],
        "ROCE (%)": [24.4, 25.1, 16.5],
        "PAT Margin (%)": [11.7, 13.2, 7.8],
        "Debt/Equity": [0.35, 0.12, 0.65],
    }

    peer_df = pd.DataFrame(default_peers_data)
    edited_peer_df = st.data_editor(peer_df, num_rows="dynamic", use_container_width=True)

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

        # Sheet 1: Dashboard
        ws1 = wb.active
        ws1.title = "Executive Summary"
        ws1.views.sheetView[0].showGridLines = True

        # Styles
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

        # Title Block
        ws1["A1"] = f"IPO EVALUATION REPORT: {company_name.upper()}"
        ws1["A1"].font = title_font
        ws1["A2"] = f"Sector: {sector} | Issue Size: ₹{issue_size} Cr"
        ws1["A2"].font = Font(name="Calibri", size=11, italic=True)

        # Overview Metrics Table
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

        # Pillar Score Table
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

        # Sheet 2: Peer Valuation Matrix
        ws2 = wb.create_sheet(title="Peer Comparison")
        ws2.views.sheetView[0].showGridLines = True

        ws2["A1"] = "PEER VALUATION & RELATIVE METRICS"
        ws2["A1"].font = title_font

        for r in dataframe_to_rows(edited_peer_df, index=False, header=True):
            ws2.append(r)

        # Style Peer Table Header
        for cell in ws2[3]:
            cell.fill = header_fill
            cell.font = header_font

        # Auto-adjust column widths across sheets
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
