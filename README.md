# Loan Collection Assistant — AI Prototype

A working prototype of an AI-assisted loan collections workflow for a bank, built to demonstrate how agentic AI patterns — tool use, reflection & revision, and human-in-the-loop control — can be applied to a real BFSI process while keeping humans in charge of every consequential decision.

**This is a portfolio / learning project.** All customer names, loan numbers, and amounts are synthetic. It is not connected to any real bank system.

🔗  **[Open the live dashboard](https://girinehe.github.io/loan-collection-ai-assistant/dashboard.html)**

---

## What this is

Loan collections at a bank follows a fairly standard lifecycle: a loan goes from current to past-due, moves through escalating reminder stages, a collection agent engages the customer and tracks promises-to-pay, and — depending on how the customer responds — the case either resolves or escalates toward field visits, bureau reporting, legal notices, and eventually write-off.

This prototype builds a small, end-to-end system that:

1. **Tracks** 20 synthetic loan accounts through that lifecycle (`Loan_Contracts`, `Collection_Cases`, `Payments`, `Bureau_Reporting_Log`)
2. **Applies a rules engine** that buckets every loan by days-past-due and recommends a next action — critically, it enforces a **hard rule**: once a customer has broken a promise-to-pay beyond a policy threshold, the system forces escalation and does not allow an agent to quietly grant another extension
3. **Drafts** customer reminder messages and internal case summaries using an AI layer that **reflects on its own draft** and revises it — catching, for example, a reminder that drifted into premature "legal action" language
4. **Routes every consequential decision to a human.** The rules engine tags each loan Auto or Human-required; only Auto items execute without sign-off. Everything else sits in an approval queue until a named person approves, rejects, or edits the recommendation
5. **Logs everything.** A single audit trail records why the system flagged each case and what a human ultimately decided
6. **Displays it all** in a browser dashboard — click any loan to see its full decision trail, no spreadsheet required

## Why it's built this way

The interesting design problem in collections isn't "can AI send reminders" — it's **where exactly the line between automation and human judgment should sit**, and how to make that line enforceable in code rather than just written in a policy document that agents can quietly work around. This prototype's one hard rule (forced escalation after repeated broken promises) exists specifically to demonstrate that: a real, common failure mode in collections — staff re-extending difficult cases to avoid escalation paperwork — closed off by design, not by trust.

## Results on the synthetic dataset

| Metric | Value |
|---|---|
| Loans tracked | 20 |
| Handled automatically, no approval needed | 15 (75%) |
| Required a human decision | 5 (25%) |
| Audit log entries | 41 |
| AI drafts that needed revision after reflection | 4 of 13 |

## Project structure

```
Loan_Collection_Prototype/
├── README.md
├── Loan_Collection_Assistant.pptx     ← presentation deck
├── Loan_Collection_Prototype.xlsx     ← the "core banking" data + all results
├── dashboard.html                     ← ready-to-open dashboard (double-click it)
├── screenshots/
│   ├── overview.png
│   └── loan-detail.png
└── code/
    ├── build_loan_data.py             ← step 1: generates synthetic loan data
    ├── rules_engine.py                ← step 2: DPD buckets, escalation rules
    ├── ai_layer.py                    ← step 3: AI drafting with reflection/revision
    ├── approval_layer.py              ← step 4: human approval queue + audit log
    ├── export_dashboard_data.py       ← step 5: exports the workbook to JSON
    ├── build_dashboard.py             ← step 6: bakes the JSON into dashboard.html
    └── dashboard_template.html        ← the dashboard page source (with a data placeholder)
```

## Viewing the dashboard

No installation needed. Download or clone this repo, then double-click `dashboard.html` — it opens in any browser and works completely offline, since all the data is embedded directly in the file.

To view it as a hosted link instead of a local file: in this repo's **Settings → Pages**, set the source to the `main` branch and save. GitHub will publish `dashboard.html` at a public URL within a minute or two — paste that link at the top of this README.

## Re-running the pipeline

The whole thing is reproducible from scratch. Requires **Python 3** with `openpyxl` installed (`pip install openpyxl`):

```bash
cd code
python build_loan_data.py        # generates the synthetic loan/case/payment data
python rules_engine.py           # applies the rules engine, writes recommendations
python ai_layer.py               # drafts messages & summaries, with reflection/revision
python approval_layer.py         # simulates the human approval queue, extends the audit log
python export_dashboard_data.py  # exports the workbook to dashboard_data.json
python build_dashboard.py        # bakes the data into dashboard.html
```

Each script prints a short summary as it runs, and the final `dashboard.html` is written back into the `code/` folder — copy it up to the repo root (or adjust the path) to replace the shipped version.

> An optional real AI integration point exists in `ai_layer.py` (`_llm_call()`): if an `ANTHROPIC_API_KEY` environment variable is set, the script calls the live model instead of its template-based fallback drafter. Everything downstream — reflection, revision, the dashboard — works identically either way.

## Agentic AI patterns demonstrated

| Pattern | Where |
|---|---|
| **Planning** | The rules engine's DPD-bucket logic and escalation matrix |
| **Tool use** | Reading/writing the "core banking" data, triggering bureau-reporting flags |
| **Reflection + revision** | The AI layer drafts a message, checks it against a compliance/tone checklist, and rewrites it if it fails |
| **Human-in-the-loop** | Every Human-required recommendation is blocked from action until a named person logs a decision |

## Built with

Python (`openpyxl`), vanilla HTML/CSS/JavaScript (no frameworks — the dashboard is a single self-contained file), and PowerPoint generated via `pptxgenjs`.

## Author

**Girish Nehete** — built as part of self-directed study in AI for Financial Services.
