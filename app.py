"""
AuditAI - P2P Observation Drafter (MVP)

One user:       Internal auditor (associate / assistant manager)
One problem:    Writing audit observations for Procure-to-Pay exceptions takes too long
One workflow:   Upload PO sample -> rule checks find exceptions -> AI drafts observations
                -> auditor reviews/edits -> download draft report
One AI feature: Drafting audit observations from exceptions (Google Gemini)
"""

import io
import json
from datetime import date

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Pt

# ---------------------------------------------------------------------------
# Settings
# ---------------------------------------------------------------------------
st.set_page_config(page_title="AuditAI – P2P Observation Drafter", page_icon="📋", layout="wide")

REQUIRED_COLUMNS = ["PO_No", "PO_Date", "Vendor", "PR_No", "PR_Amount", "PO_Amount", "Approved_By_Role"]
DEFAULT_MODEL = "gemini-2.5-flash"

# The three P2P checks. "criteria" is the rule the auditor tests against.
CHECKS = {
    "PO_WITHOUT_PR": {
        "name": "Purchase order raised without a purchase requisition",
        "criteria": "Every purchase order must be supported by an approved purchase requisition.",
    },
    "PO_EXCEEDS_PR": {
        "name": "Purchase order value exceeds purchase requisition value",
        "criteria": "PO value must not exceed the approved PR value beyond the allowed tolerance.",
    },
    "APPROVAL_BEYOND_AUTHORITY": {
        "name": "Purchase order approved beyond delegated authority",
        "criteria": "POs must be approved by a role whose limit in the Delegation of Authority (DoA) "
                    "matrix covers the PO value.",
    },
}


def inr(amount: float) -> str:
    """Format a number in Indian style, e.g. 12,34,567."""
    amount = int(round(amount))
    sign = "-" if amount < 0 else ""
    s = str(abs(amount))
    if len(s) > 3:
        head, tail = s[:-3], s[-3:]
        groups = []
        while len(head) > 2:
            groups.insert(0, head[-2:])
            head = head[:-2]
        if head:
            groups.insert(0, head)
        s = ",".join(groups) + "," + tail
    return f"{sign}₹{s}"


# ---------------------------------------------------------------------------
# Step 2: rule-based checks (no AI here - these are deterministic)
# ---------------------------------------------------------------------------
def run_checks(df: pd.DataFrame, doa: dict, tolerance_pct: float) -> pd.DataFrame:
    """Return one row per exception found."""
    exceptions = []
    for _, row in df.iterrows():
        pr_no = str(row["PR_No"]).strip()
        has_pr = pr_no not in ("", "nan", "None")
        pr_amt, po_amt = row["PR_Amount"], row["PO_Amount"]
        role = str(row["Approved_By_Role"]).strip()

        if not has_pr:
            exceptions.append((row, "PO_WITHOUT_PR", "No PR number recorded"))
        elif pd.notna(pr_amt) and pr_amt > 0 and po_amt > pr_amt * (1 + tolerance_pct / 100):
            excess = (po_amt - pr_amt) / pr_amt * 100
            exceptions.append((row, "PO_EXCEEDS_PR",
                               f"PO exceeds PR by {inr(po_amt - pr_amt)} ({excess:.1f}%)"))

        limit = doa.get(role)
        if limit is None:
            exceptions.append((row, "APPROVAL_BEYOND_AUTHORITY", f"Role '{role}' not in DoA matrix"))
        elif po_amt > limit:
            exceptions.append((row, "APPROVAL_BEYOND_AUTHORITY",
                               f"{role} limit is {inr(limit)}; PO value is {inr(po_amt)}"))

    out = pd.DataFrame([{
        "Check": CHECKS[code]["name"],
        "check_code": code,
        "PO_No": r["PO_No"], "PO_Date": r["PO_Date"], "Vendor": r["Vendor"],
        "PR_No": r["PR_No"], "PR_Amount": r["PR_Amount"], "PO_Amount": r["PO_Amount"],
        "Approved_By_Role": r["Approved_By_Role"], "Detail": detail,
    } for r, code, detail in exceptions])
    return out


# ---------------------------------------------------------------------------
# Step 3: AI drafting of observations
# ---------------------------------------------------------------------------
def build_prompt(code: str, rows: pd.DataFrame, total_tested: int, context: dict) -> str:
    sample = rows[["PO_No", "Vendor", "PR_Amount", "PO_Amount", "Approved_By_Role", "Detail"]].head(10)
    return f"""You are assisting an internal auditor in India who is auditing the Procure-to-Pay process.
Draft ONE audit observation for the exceptions below, in the auditor's standard format.

Client: {context['client']} ({context['industry']})
Audit period: {context['period']}
Check performed: {CHECKS[code]['name']}
Criteria (policy): {CHECKS[code]['criteria']} {context['policy_ref']}
POs tested: {total_tested}
Exceptions found: {len(rows)}
Total PO value of exceptions: {inr(rows['PO_Amount'].sum())}
Exception details (first {len(sample)}):
{sample.to_csv(index=False)}

Rules:
- Use ONLY the facts given above. Do not invent names, dates, amounts or reasons.
- The cause is not known yet. Write likely causes as questions for the process owner,
  starting with "To be confirmed with management:".
- Write in clear, formal audit language. Keep each field to 2-4 sentences.
- Risk rating must be High, Medium or Low, based on value, frequency and control impact.

Return JSON with exactly these keys:
title, condition, criteria, cause, risk_implication, risk_rating, recommendation"""


def draft_with_ai(prompt: str, api_key: str, model: str) -> dict:
    from google import genai
    from google.genai import types

    client = genai.Client(api_key=api_key)
    response = client.models.generate_content(
        model=model,
        contents=prompt,
        config=types.GenerateContentConfig(response_mime_type="application/json", temperature=0.3),
    )
    data = json.loads(response.text)
    if isinstance(data, list):
        data = data[0]
    return data


def draft_without_ai(code: str, rows: pd.DataFrame, total_tested: int) -> dict:
    """Fallback template used when no API key is set, so the workflow can still be demonstrated."""
    return {
        "title": CHECKS[code]["name"],
        "condition": f"In {len(rows)} of {total_tested} POs tested (total value {inr(rows['PO_Amount'].sum())}), "
                     f"the following was noted: {CHECKS[code]['name'].lower()}.",
        "criteria": CHECKS[code]["criteria"],
        "cause": "To be confirmed with management.",
        "risk_implication": "[Auditor to complete]",
        "risk_rating": "Medium",
        "recommendation": "[Auditor to complete]",
    }


# ---------------------------------------------------------------------------
# Step 5: report export
# ---------------------------------------------------------------------------
FIELDS = [("condition", "Condition"), ("criteria", "Criteria"), ("cause", "Cause"),
          ("risk_implication", "Risk / Implication"), ("recommendation", "Recommendation")]


def build_docx(observations: list, context: dict, total_tested: int) -> bytes:
    doc = Document()
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(11)
    doc.add_heading(f"Draft Internal Audit Report – Procure-to-Pay", level=0)
    doc.add_paragraph(f"Client: {context['client']}    |    Period: {context['period']}    |    "
                      f"Prepared: {date.today().strftime('%d %b %Y')}")
    doc.add_paragraph(f"POs tested: {total_tested}. Status: DRAFT – for auditor review. "
                      "Observations were drafted with AI assistance and reviewed by the auditor.")

    doc.add_heading("Summary of observations", level=1)
    table = doc.add_table(rows=1, cols=3)
    table.style = "Light Grid Accent 1"
    for cell, text in zip(table.rows[0].cells, ["#", "Observation", "Risk rating"]):
        cell.text = text
    for i, obs in enumerate(observations, 1):
        cells = table.add_row().cells
        cells[0].text, cells[1].text, cells[2].text = str(i), obs["title"], obs["risk_rating"]

    for i, obs in enumerate(observations, 1):
        doc.add_heading(f"{i}. {obs['title']}  [{obs['risk_rating']}]", level=1)
        for key, label in FIELDS:
            p = doc.add_paragraph()
            p.add_run(f"{label}: ").bold = True
            p.add_run(obs[key])
        if obs.get("instances") is not None:
            doc.add_paragraph("Exception instances:").runs[0].bold = True
            inst = obs["instances"]
            t = doc.add_table(rows=1, cols=len(inst.columns))
            t.style = "Light Grid Accent 1"
            for cell, col in zip(t.rows[0].cells, inst.columns):
                cell.text = str(col)
            for _, r in inst.iterrows():
                for cell, val in zip(t.add_row().cells, r):
                    cell.text = str(val)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📋 AuditAI – P2P Observation Drafter")
st.caption("Upload your PO sample → find exceptions → AI drafts audit observations → you review → download the draft report.")

with st.sidebar:
    st.header("Engagement details")
    context = {
        "client": st.text_input("Client name", "Client A"),
        "industry": st.text_input("Industry", "Manufacturing"),
        "period": st.text_input("Audit period", "Apr–Sep 2026"),
        "policy_ref": st.text_input("Policy reference (optional)", "",
                                    help="e.g. 'Procurement Policy v3, section 4.2'"),
    }
    tolerance = st.number_input("Allowed PO over PR tolerance (%)", 0.0, 100.0, 5.0, 0.5)

    st.header("Delegation of Authority")
    st.caption("Edit to match the client's authority chart (limits in ₹).")
    doa_df = st.data_editor(
        pd.DataFrame({"Role": ["Purchase Manager", "Head – Procurement", "CFO"],
                      "Limit": [500000, 2500000, 10000000]}),
        num_rows="dynamic", hide_index=True, key="doa",
    )
    doa = {str(r.Role).strip(): float(r.Limit) for r in doa_df.itertuples() if pd.notna(r.Limit)}

    st.header("AI settings")
    def secret(name, default=""):
        try:
            return st.secrets.get(name, default)
        except Exception:  # no secrets file (e.g. running locally)
            return default

    api_key = secret("GEMINI_API_KEY")
    if not api_key:
        api_key = st.text_input("Gemini API key", type="password",
                                help="Leave empty to use a basic template instead of AI.")
    model = secret("GEMINI_MODEL", DEFAULT_MODEL)
    st.caption("✅ AI connected" if api_key else "⚠️ No API key – template mode")

# ---- Step 1: upload --------------------------------------------------------
st.subheader("1. Upload PO sample")
template_csv = pd.DataFrame(columns=REQUIRED_COLUMNS).to_csv(index=False).encode()
st.download_button("Download blank template (CSV)", template_csv, "po_sample_template.csv", "text/csv")
uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])

if not uploaded:
    st.info(f"Required columns: {', '.join(REQUIRED_COLUMNS)}")
    st.stop()

df = pd.read_csv(uploaded) if uploaded.name.endswith(".csv") else pd.read_excel(uploaded)
missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
if missing:
    st.error(f"Missing columns: {', '.join(missing)}. Please use the template.")
    st.stop()
for col in ["PR_Amount", "PO_Amount"]:
    df[col] = pd.to_numeric(df[col], errors="coerce")
df["PO_Amount"] = df["PO_Amount"].fillna(0)
st.write(f"{len(df)} POs loaded.")
with st.expander("View data"):
    st.dataframe(df, hide_index=True)

# ---- Step 2: checks --------------------------------------------------------
st.subheader("2. Exceptions found")
exc = run_checks(df, doa, tolerance)
if exc.empty:
    st.success("No exceptions found in this sample.")
    st.stop()

c1, c2, c3 = st.columns(3)
c1.metric("POs tested", len(df))
c2.metric("Exceptions", len(exc))
c3.metric("Value of exception POs", inr(exc.drop_duplicates("PO_No")["PO_Amount"].sum()))
st.dataframe(exc.drop(columns="check_code"), hide_index=True)

selected = st.multiselect("Checks to draft observations for",
                          options=list(exc["check_code"].unique()),
                          default=list(exc["check_code"].unique()),
                          format_func=lambda c: CHECKS[c]["name"])

# ---- Step 3: AI drafting ---------------------------------------------------
st.subheader("3. Draft observations")
if st.button("✨ Draft observations", type="primary", disabled=not selected):
    drafts = {}
    with st.spinner("Drafting observations..."):
        for code in selected:
            rows = exc[exc["check_code"] == code]
            try:
                if api_key:
                    d = draft_with_ai(build_prompt(code, rows, len(df), context), api_key, model)
                else:
                    d = draft_without_ai(code, rows, len(df))
            except Exception as e:
                st.warning(f"AI drafting failed for '{CHECKS[code]['name']}' ({e}). Using template instead.")
                d = draft_without_ai(code, rows, len(df))
            drafts[code] = {k: str(d.get(k, "")) for k in
                            ["title", "condition", "criteria", "cause", "risk_implication",
                             "risk_rating", "recommendation"]}
    st.session_state["drafts"] = drafts

if "drafts" not in st.session_state:
    st.stop()

# ---- Step 4: review --------------------------------------------------------
st.subheader("4. Review and edit")
st.caption("AI drafts can be wrong. Check every field against your working papers before accepting.")
accepted = []
for code, d in st.session_state["drafts"].items():
    if code not in exc["check_code"].values:
        continue
    with st.container(border=True):
        title = st.text_input("Title", d["title"], key=f"{code}_title")
        ratings = ["High", "Medium", "Low"]
        rating = st.selectbox("Risk rating", ratings,
                              index=ratings.index(d["risk_rating"]) if d["risk_rating"] in ratings else 1,
                              key=f"{code}_rating")
        obs = {"title": title, "risk_rating": rating}
        for key, label in FIELDS:
            obs[key] = st.text_area(label, d[key], key=f"{code}_{key}", height=90)
        rows = exc[exc["check_code"] == code]
        obs["instances"] = rows[["PO_No", "Vendor", "PR_Amount", "PO_Amount", "Approved_By_Role", "Detail"]]
        if st.checkbox("Reviewed – include in report", key=f"{code}_ok"):
            accepted.append(obs)

# ---- Step 5: export --------------------------------------------------------
st.subheader("5. Download draft report")
if not accepted:
    st.info("Tick 'Reviewed – include in report' on at least one observation.")
else:
    st.download_button("⬇️ Download report (.docx)", build_docx(accepted, context, len(df)),
                       f"AuditAI_P2P_draft_report_{date.today():%Y%m%d}.docx",
                       "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                       type="primary")
