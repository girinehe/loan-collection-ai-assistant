"""
Step 6: Human approval + audit log for the Loan Collection Assistant prototype.

This is the control layer. Up to now:
  - Step 4 (rules engine) decided WHAT should happen and whether it's
    Auto or Human-required.
  - Step 5 (AI layer) drafted the actual text (customer message or
    case summary), with reflection/revision as a quality check.

Step 6 enforces the control matrix from step 2 as real behaviour:
  - AUTO items are logged as sent — the system acted on its own, and
    that fact is recorded, but no one had to approve it.
  - HUMAN-REQUIRED items go into an Approval_Queue and CANNOT be marked
    "actioned" until a named person records a decision: Approve,
    Reject, or Approve-with-edits. This is what makes the hard rules
    from step 4 (like the forced-escalation rule) actually meaningful —
    the system can flag a case, but a human still makes the final call.

In a real deployment, the approval decision would come from an actual
officer clicking a button in the dashboard (step 7). Here, since this
is an offline prototype, we simulate what a reviewing officer's
decision log would look like — clearly labelled as simulated — so the
Approval_Queue and Audit_Log have realistic content to demo and to
build the dashboard against.

OUTPUT:
  - Approval_Queue sheet: one row per Human-required item, with
    decision, decision-maker, timestamp, and any comment.
  - Audit_Log sheet (extended, not replaced): adds approval-decision
    rows alongside the rule-firing rows already logged in step 4, so
    there is ONE trail covering both "why the system flagged this"
    and "what the human decided."
"""

import random
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

WORKBOOK_PATH = "Loan_Collection_Prototype.xlsx"

random.seed(7)

# Synthetic officer names who would be reviewing the Human-required queue
OFFICERS = ["N. Fernandes (Collections Manager)", "P. Rao (Senior Collections Officer)",
            "A. D'Souza (Risk & Recovery)"]

# Simulated decision outcomes with realistic comments per case type
def simulate_decision(loan: dict) -> dict:
    """
    Returns a simulated officer decision for a Human-required loan.
    This stands in for a real person clicking Approve/Reject/Edit in
    the dashboard — clearly a placeholder for demo purposes, not a
    real approval.
    """
    action = loan["recommended_action"]
    officer = random.choice(OFFICERS)
    decided_at = datetime.now() - timedelta(hours=random.randint(1, 48))

    if "Force escalation" in action:
        decision = "Approved"
        comment = "Confirmed — repeated broken P2P, proceeding to escalation per policy. No further P2P extensions to be granted."
    elif "Proceed with escalation" in action:
        decision = "Approved"
        comment = "Field visit authorised. Agent to report back within 5 working days."
    elif "legal notice or moratorium" in action:
        decision = random.choice(["Approved", "Approved with edits"])
        if decision == "Approved with edits":
            comment = "Hold legal notice for now — offering moratorium discussion first, per customer's repayment history. Reclassify to NAB as recommended."
        else:
            comment = "Reclassify to NAB and proceed with legal notice preparation. Escalate to legal team."
    else:
        decision = "Approved"
        comment = "Reviewed, recommendation accepted."

    return {
        "decision": decision,
        "decided_by": officer,
        "decided_at": decided_at.strftime("%Y-%m-%d %H:%M"),
        "comment": comment,
    }


def load_sheet_as_dicts(ws):
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(dict(zip(headers, row)))
    return rows


def main():
    wb = openpyxl.load_workbook(WORKBOOK_PATH)
    recs = load_sheet_as_dicts(wb["Engine_Recommendations"])

    approval_rows = []
    new_audit_rows = []

    for loan in recs:
        loan_id = loan["loan_id"]
        decision_type = loan["decision_type"]
        action = loan["recommended_action"]

        if decision_type == "Auto":
            # System acted on its own — log it, nothing to approve.
            new_audit_rows.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "loan_id": loan_id,
                "rule_applied": "auto_execution",
                "outcome": f"Auto-executed without human approval: {action}",
            })
        else:
            # Human-required — goes into the approval queue.
            sim = simulate_decision(loan)
            approval_rows.append({
                "loan_id": loan_id,
                "customer_name": loan["customer_name"],
                "recommended_action": action,
                "reason_flagged": loan["reason"],
                "decision": sim["decision"],
                "decided_by": sim["decided_by"],
                "decided_at": sim["decided_at"],
                "comment": sim["comment"],
            })
            new_audit_rows.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "loan_id": loan_id,
                "rule_applied": "human_approval_decision",
                "outcome": f"{sim['decision']} by {sim['decided_by']} at {sim['decided_at']} — {sim['comment']}",
            })

    # -------------------------------------------------------------------
    # Write Approval_Queue sheet
    # -------------------------------------------------------------------
    if "Approval_Queue" in wb.sheetnames:
        del wb["Approval_Queue"]

    HEADER_FILL = PatternFill(start_color="4A3B6B", end_color="4A3B6B", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True)
    APPROVED_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    EDITED_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    cols = ["loan_id", "customer_name", "recommended_action", "reason_flagged",
            "decision", "decided_by", "decided_at", "comment"]
    ws = wb.create_sheet("Approval_Queue")
    ws.append(cols)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for r in approval_rows:
        ws.append([r[c] for c in cols])
    decision_col_idx = cols.index("decision") + 1
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(wrap_text=True, vertical="top")
        decision_val = row[decision_col_idx - 1].value
        fill = EDITED_FILL if decision_val == "Approved with edits" else APPROVED_FILL
        for cell in row:
            cell.fill = fill
    widths = {"loan_id": 10, "customer_name": 16, "recommended_action": 45, "reason_flagged": 45,
              "decision": 18, "decided_by": 28, "decided_at": 16, "comment": 55}
    for i, col in enumerate(cols, start=1):
        ws.column_dimensions[get_column_letter(i)].width = widths[col]
    ws.freeze_panes = "A2"

    # -------------------------------------------------------------------
    # Append to existing Audit_Log sheet (don't overwrite step 4's rows)
    # -------------------------------------------------------------------
    audit_ws = wb["Audit_Log"]
    for r in new_audit_rows:
        audit_ws.append([r["timestamp"], r["loan_id"], r["rule_applied"], r["outcome"]])
    for i, col_name in enumerate(["timestamp", "loan_id", "rule_applied", "outcome"], start=1):
        current_width = audit_ws.column_dimensions[get_column_letter(i)].width or 15
        max_len = max([len(str(c.value)) for c in audit_ws[get_column_letter(i)] if c.value] + [current_width])
        audit_ws.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 75)

    wb.save(WORKBOOK_PATH)

    approved = sum(1 for r in approval_rows if r["decision"] == "Approved")
    edited = sum(1 for r in approval_rows if r["decision"] == "Approved with edits")
    auto_count = sum(1 for loan in recs if loan["decision_type"] == "Auto")

    print(f"Auto-executed (no approval needed): {auto_count}")
    print(f"Sent to Approval_Queue: {len(approval_rows)}")
    print(f"  - Approved as-is: {approved}")
    print(f"  - Approved with edits: {edited}")
    print(f"Audit_Log now has {audit_ws.max_row - 1} total entries (rule-firing + approval decisions)")
    print(f"Saved into: {WORKBOOK_PATH}")

    print("\n--- Approval queue sample ---")
    for r in approval_rows:
        print(f"{r['loan_id']} | {r['decision']:20s} | {r['decided_by']} | {r['comment'][:60]}...")


if __name__ == "__main__":
    main()
