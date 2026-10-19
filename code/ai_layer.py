"""
Step 5: AI layer for the Loan Collection Assistant prototype.

This module adds two AI-generated outputs on top of the rules engine's
recommendations (step 4):

  1. Customer-facing messages — reminders / past-due notices, for loans
     the rules engine marked "Auto".
  2. Internal case summaries — for loans marked "Human-required", so an
     agent/officer reviewing the worklist gets a short brief instead of
     having to re-read every raw data field themselves.

AGENTIC PATTERN USED: Reflection + Revision
  Step A (DRAFT)   — generate a first-pass message/summary from the data.
  Step B (REFLECT) — check the draft against a short checklist (tone,
                      accuracy, required disclosures, length).
  Step C (REVISE)  — if the reflection step finds a problem, fix it and
                      re-check. Only the revised, passing version is kept.

This is the same loop a real production system would run, just using
simple deterministic checks here instead of a second LLM call — the
checklist in `reflect()` is exactly what you'd otherwise ask a model
to critique against.

HOW A REAL LLM WOULD PLUG IN
  In production, `draft_customer_message()` and `draft_case_summary()`
  would call an LLM, e.g.:

      from anthropic import Anthropic
      client = Anthropic()  # reads ANTHROPIC_API_KEY from environment

      response = client.messages.create(
          model="claude-sonnet-5",
          max_tokens=300,
          messages=[{"role": "user", "content": prompt}],
      )
      draft_text = response.content[0].text

  This sandbox has no API key configured, so this script runs a local
  "fallback drafter" instead — template logic that applies the SAME
  reflection/revision structure. The moment Girish adds a real API key,
  swapping the fallback for the real call above (see `USE_REAL_LLM`
  below) changes nothing else in the pipeline.
"""

import os
from datetime import date

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

WORKBOOK_PATH = "Loan_Collection_Prototype.xlsx"
TODAY = date(2026, 10, 16)

# Flip this on automatically if a real key is present; otherwise use the
# local fallback drafter. Either path runs the same reflect/revise loop.
USE_REAL_LLM = bool(os.environ.get("ANTHROPIC_API_KEY"))

if USE_REAL_LLM:
    from anthropic import Anthropic
    client = Anthropic()


# ---------------------------------------------------------------------------
# DRAFT step
# ---------------------------------------------------------------------------
def _llm_call(prompt: str) -> str:
    """Single point where a real model call happens, if a key is present."""
    response = client.messages.create(
        model="claude-sonnet-5",
        max_tokens=300,
        messages=[{"role": "user", "content": prompt}],
    )
    return response.content[0].text.strip()


def draft_customer_message(loan: dict) -> str:
    """First-pass customer-facing reminder, tone varying by severity."""
    name = loan["customer_name"].split()[0]  # first name, for a warmer tone
    status = loan["recommended_status"]
    emi = loan["emi_amount"] if "emi_amount" in loan else None
    dpd = loan["days_past_due"]

    if USE_REAL_LLM:
        prompt = (
            f"Write a short SMS (under 300 characters) to a bank loan customer "
            f"named {name}. Loan status: {status}, {dpd} days past due. "
            f"Tone must be polite, non-threatening, and RBI-compliant. "
            f"Do not mention legal action unless status is NAB or beyond."
        )
        return _llm_call(prompt)

    # --- Fallback drafter (no API key) ---
    if status == "PDO1":
        return (f"Dear {name}, your loan installment is now {dpd} day(s) overdue. "
                f"Please make the payment at your earliest convenience to avoid "
                f"further charges. If already paid, kindly ignore this message.")
    elif status == "PDO2":
        # Intentionally drafted a bit too strongly here, to show the
        # reflection step actually catching and softening it — a first
        # pass drifting toward "legal" language this early is a realistic
        # mistake for a template or a model prompt that isn't tuned yet.
        return (f"Dear {name}, your loan account is {dpd} days past due. "
                f"Please note that continued non-payment may lead to legal action. "
                f"We encourage you to settle the overdue amount or contact our "
                f"collection team to discuss available options.")
    elif status == "PDO3":
        return (f"Dear {name}, your loan account is significantly overdue ({dpd} days). "
                f"This requires urgent attention. Please contact our collections team "
                f"to discuss repayment or restructuring options before further action is taken.")
    else:
        return (f"Dear {name}, your loan account remains unresolved at {dpd} days past due. "
                f"Please contact us immediately to discuss your account.")


def draft_case_summary(loan: dict) -> str:
    """First-pass internal summary for a Human-required case."""
    name = loan["customer_name"]
    loan_id = loan["loan_id"]
    status = loan["recommended_status"]
    dpd = loan["days_past_due"]
    case_status = loan["case_status"]
    escalation = loan["escalation_level"]
    agent = loan["assigned_agent"]
    action = loan["recommended_action"]
    reason = loan["reason"]

    if USE_REAL_LLM:
        prompt = (
            f"Write a 3-sentence internal case summary for a bank collections officer. "
            f"Loan {loan_id}, customer {name}, status {status}, {dpd} days overdue, "
            f"case status {case_status}, escalation level {escalation}, assigned to {agent}. "
            f"System recommendation: {action}. Reason: {reason}. "
            f"Be factual and neutral, no speculation about the customer's intent."
        )
        return _llm_call(prompt)

    # --- Fallback drafter (no API key) ---
    return (
        f"Loan {loan_id} ({name}) is {dpd} days past due, currently at status {status}. "
        f"Case history shows status '{case_status}' with escalation level {escalation}, "
        f"assigned to {agent}. System recommendation: {action}. "
        f"Reason flagged for human review: {reason}."
    )


# ---------------------------------------------------------------------------
# REFLECT step — checklist a real reviewer (or a second LLM call) would apply
# ---------------------------------------------------------------------------
HARSH_WORDS = ["must pay immediately", "default", "seize", "blacklist", "warrant", "arrest"]

def reflect_customer_message(text: str, status: str) -> list:
    """Returns a list of problems found, empty if the draft passes."""
    issues = []
    lowered = text.lower()

    for word in HARSH_WORDS:
        if word in lowered:
            issues.append(f"Contains harsh/alarming language: '{word}'")

    if status in ("PDO1", "PDO2") and ("legal" in lowered or "court" in lowered):
        issues.append("Mentions legal action, but status is not yet NAB/escalated — too strong for this stage")

    if len(text) > 320:
        issues.append(f"Message too long for SMS ({len(text)} chars) — should be under ~320")

    if status in ("PDO2", "PDO3") and "contact" not in lowered and "collection" not in lowered:
        issues.append("Missing guidance on who the customer should contact")

    return issues


def reflect_case_summary(text: str) -> list:
    issues = []
    lowered = text.lower()

    # speculative/judgmental language a summary should avoid
    speculative_terms = ["probably", "seems like", "clearly trying to avoid", "irresponsible", "deliberately"]
    for term in speculative_terms:
        if term in lowered:
            issues.append(f"Contains speculative/judgmental language: '{term}'")

    if len(text) < 40:
        issues.append("Summary too short to be useful to a reviewing officer")

    return issues


# ---------------------------------------------------------------------------
# REVISE step — simple deterministic fixes for what reflection catches
# ---------------------------------------------------------------------------
def revise_customer_message(text: str, issues: list) -> str:
    revised = text
    for issue in issues:
        if "harsh" in issue.lower():
            for word in HARSH_WORDS:
                revised = revised.replace(word, "settle the overdue amount")
        if "legal action" in issue.lower():
            revised = revised.replace("legal action", "further steps")
        if "too long" in issue.lower():
            revised = revised[:300].rsplit(" ", 1)[0] + "."
    return revised


def revise_case_summary(text: str, issues: list) -> str:
    revised = text
    for issue in issues:
        if "speculative" in issue.lower():
            for term in ["probably", "seems like", "clearly trying to avoid", "irresponsible", "deliberately"]:
                revised = revised.replace(term, "")
    return revised


# ---------------------------------------------------------------------------
# Orchestration: draft -> reflect -> revise -> re-check
# ---------------------------------------------------------------------------
def generate_customer_message_with_reflection(loan: dict):
    status = loan["recommended_status"]
    draft = draft_customer_message(loan)
    issues = reflect_customer_message(draft, status)
    if not issues:
        return draft, "Passed on first draft — no issues found", False
    revised = revise_customer_message(draft, issues)
    remaining = reflect_customer_message(revised, status)
    note = "Issues found and fixed: " + "; ".join(issues)
    if remaining:
        note += " | STILL FLAGGED after revision: " + "; ".join(remaining)
    return revised, note, True


def generate_case_summary_with_reflection(loan: dict):
    draft = draft_case_summary(loan)
    issues = reflect_case_summary(draft)
    if not issues:
        return draft, "Passed on first draft — no issues found", False
    revised = revise_case_summary(draft, issues)
    remaining = reflect_case_summary(revised)
    note = "Issues found and fixed: " + "; ".join(issues)
    if remaining:
        note += " | STILL FLAGGED after revision: " + "; ".join(remaining)
    return revised, note, True


def load_recommendations(wb):
    ws = wb["Engine_Recommendations"]
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(dict(zip(headers, row)))
    return rows


def main():
    wb = openpyxl.load_workbook(WORKBOOK_PATH)
    recs = load_recommendations(wb)

    results = []
    for loan in recs:
        if loan["decision_type"] == "Auto" and loan["recommended_status"] != "Normal" and loan["recommended_status"] != "Written_Off":
            final_text, note, was_revised = generate_customer_message_with_reflection(loan)
            results.append({
                "loan_id": loan["loan_id"],
                "customer_name": loan["customer_name"],
                "draft_type": "Customer Message",
                "final_text": final_text,
                "revision_note": note,
                "was_revised": "Yes" if was_revised else "No",
            })
        elif loan["decision_type"] == "Human-required":
            final_text, note, was_revised = generate_case_summary_with_reflection(loan)
            results.append({
                "loan_id": loan["loan_id"],
                "customer_name": loan["customer_name"],
                "draft_type": "Case Summary (for officer)",
                "final_text": final_text,
                "revision_note": note,
                "was_revised": "Yes" if was_revised else "No",
            })

    # Write to a new sheet
    if "AI_Drafts" in wb.sheetnames:
        del wb["AI_Drafts"]

    HEADER_FILL = PatternFill(start_color="2E5C7A", end_color="2E5C7A", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True)
    REVISED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    cols = ["loan_id", "customer_name", "draft_type", "final_text", "revision_note", "was_revised"]
    ws = wb.create_sheet("AI_Drafts")
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for r in results:
        ws.append([r[c] for c in cols])
    for row in ws.iter_rows(min_row=2):
        row[0].alignment = Alignment(vertical="top")
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        if row[5].value == "Yes":
            for cell in row:
                cell.fill = REVISED_FILL
    widths = {"loan_id": 10, "customer_name": 18, "draft_type": 22, "final_text": 70, "revision_note": 55, "was_revised": 12}
    for i, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths[col]
    ws.freeze_panes = "A2"

    wb.save(WORKBOOK_PATH)

    revised_count = sum(1 for r in results if r["was_revised"] == "Yes")
    print(f"Generated {len(results)} AI drafts ({sum(1 for r in results if r['draft_type']=='Customer Message')} messages, "
          f"{sum(1 for r in results if 'Summary' in r['draft_type'])} case summaries)")
    print(f"Drafts revised after reflection: {revised_count}")
    print(f"Using real LLM: {USE_REAL_LLM}")
    print(f"Saved into: {WORKBOOK_PATH}")

    print("\n--- Sample output ---")
    for r in results[:3]:
        print(f"\n[{r['loan_id']}] {r['draft_type']}")
        print(f"  Text: {r['final_text']}")
        print(f"  Reflection: {r['revision_note']}")


if __name__ == "__main__":
    main()
