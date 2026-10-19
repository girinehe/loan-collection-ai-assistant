"""
Step 7 (part 2): Bake dashboard_data.json into dashboard_template.html
to produce the final, self-contained dashboard.html.

The template has a placeholder, __DATA_JSON__, inside a
<script type="application/json"> block. This script replaces that
placeholder with the current contents of dashboard_data.json, so the
resulting dashboard.html works completely offline — just double-click
it to open in any browser, no server or internet connection needed.

Run order for the full pipeline, from a fresh Loan_Collection_Prototype.xlsx
containing only the original synthetic data (Loan_Contracts, Collection_Cases,
Payments, Bureau_Reporting_Log):

    python build_loan_data.py        # (re)generates the synthetic data sheets
    python rules_engine.py           # adds Engine_Recommendations + Audit_Log
    python ai_layer.py               # adds AI_Drafts
    python approval_layer.py         # adds Approval_Queue, extends Audit_Log
    python export_dashboard_data.py  # exports all sheets to dashboard_data.json
    python build_dashboard.py        # bakes dashboard_data.json into dashboard.html

Then open dashboard.html in a browser.
"""

TEMPLATE_PATH = "dashboard_template.html"
DATA_PATH = "dashboard_data.json"
OUTPUT_PATH = "dashboard.html"


def main():
    with open(TEMPLATE_PATH, "r", encoding="utf-8") as f:
        html = f.read()
    with open(DATA_PATH, "r", encoding="utf-8") as f:
        data_json = f.read()

    if "__DATA_JSON__" not in html:
        raise SystemExit(f"'__DATA_JSON__' placeholder not found in {TEMPLATE_PATH}")

    # Guard against a stray '</script>' inside the JSON breaking the page
    data_json_safe = data_json.replace("</script>", "<\\/script>")
    html = html.replace("__DATA_JSON__", data_json_safe)

    with open(OUTPUT_PATH, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Wrote {OUTPUT_PATH} ({len(html):,} bytes)")
    print("Open it directly in any browser — no server needed.")


if __name__ == "__main__":
    main()
