# AuditAI – P2P Observation Drafter

DPM 2026 – Group 04 – MVP

## What it does
AuditAI helps an internal auditor write audit observations for the Procure-to-Pay (P2P) process.

1. **Upload** a sample of purchase orders (CSV or Excel).
2. **Rule checks** find exceptions (no AI):
   - PO raised without a purchase requisition (PR)
   - PO value higher than PR value beyond the allowed tolerance
   - PO approved by a role whose Delegation of Authority (DoA) limit is below the PO value
3. **AI drafting** (Google Gemini) turns each group of exceptions into a draft observation:
   condition, criteria, cause, risk/implication, risk rating, recommendation.
4. **Review**: the auditor edits each draft and ticks it as reviewed.
5. **Download** a draft report in Word (.docx).

## Input file columns
`PO_No, PO_Date, Vendor, PR_No, PR_Amount, PO_Amount, Approved_By_Role`
A blank template can be downloaded from the app.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```
Without an API key the app runs in template mode (no AI), so the workflow can still be shown.

## Deploy on Streamlit Community Cloud
1. Push this folder to a public GitHub repository.
2. On share.streamlit.io choose **New app** → select the repo → main file `app.py`.
3. In **Advanced settings → Secrets** add:
   ```toml
   GEMINI_API_KEY = "your-key-here"
   # optional: GEMINI_MODEL = "gemini-2.5-flash"
   ```
4. Deploy and copy the public URL into the presentation.

Get a free Gemini API key at https://aistudio.google.com/apikey. Never commit the key to GitHub.

## Data and privacy
- Uploaded files are processed in memory only and are not stored.
- Only the exception rows (PO number, vendor, amounts, approver role) are sent to the AI model.
  Use anonymised data for demos.
- AI output is always a draft. The auditor must review it before it goes into a report.

## AI use disclosure
The code was written with help from Claude (Anthropic) and reviewed and tested by the team.
The app uses Google Gemini to draft observations.
