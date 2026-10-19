"""
Step 4: Rules Engine for the Loan Collection Assistant prototype.

Reads the synthetic "core banking + collection system" workbook built in
step 3, applies the policy logic Girish described (PDO bucket thresholds,
hard-rule escalation on repeated broken P2Ps, bureau-reporting triggers),
and produces a clear recommendation per loan: what should happen next,
and whether that action is Auto (system can do it) or Human-required
(per the control matrix from step 2).

Design choice: this engine does NOT silently overwrite the source sheets.
It reads Loan_Contracts + Collection_Cases + Bureau_Reporting_Log and
writes its output to two NEW sheets:
  - Engine_Recommendations  (one row per loan: what to do, auto/human, why)
  - Audit_Log               (one row per rule firing, for traceability)

This mirrors how a real collections engine would run as a nightly batch
job against the core system, producing a worklist for agents/officers
rather than mutating the book of record directly.
"""

from datetime import date, datetime

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SRC_PATH = "Loan_Collection_Prototype.xlsx"
OUT_PATH = "Loan_Collection_Prototype.xlsx"  # update same workbook, new sheets

TODAY = date(2026, 10, 16)

# ---------------------------------------------------------------------------
# Policy thresholds (step 2's control matrix, encoded)
# ---------------------------------------------------------------------------
def dpd_to_status(dpd: int) -> str:
    """Bucket a loan's days-past-due into its policy status."""
    if dpd == 0:
        return "Normal"
    elif dpd <= 30:
        return "PDO1"
    elif dpd <= 60:
        return "PDO2"
    elif dpd <= 90:
        return "PDO3"
    else:
        return "NAB"

# How many broken P2Ps are allowed before escalation becomes mandatory,
# regardless of what the agent wants to do. This is the hard rule we
# discussed: it stops a case from being quietly re-extended indefinitely.
MAX_BROKEN_P2P_BEFORE_FORCED_ESCALATION = 1  # i.e. the 2nd break forces it

AUTO = "Auto"
HUMAN = "Human-required"


def load_sheet_as_dicts(ws):
    headers = [c.value for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[0] is None:
            continue
        rows.append(dict(zip(headers, row)))
    return rows


def main():
    wb = openpyxl.load_workbook(SRC_PATH)
    loans = load_sheet_as_dicts(wb["Loan_Contracts"])
    cases = load_sheet_as_dicts(wb["Collection_Cases"])
    bureau_log = load_sheet_as_dicts(wb["Bureau_Reporting_Log"])

    cases_by_loan = {c["loan_id"]: c for c in cases}
    reported_loan_ids = {b["loan_id"] for b in bureau_log}

    recommendations = []
    audit_entries = []

    def log_audit(loan_id, rule, outcome):
        audit_entries.append({
            "timestamp": datetime.now().isoformat(timespec="seconds"),
            "loan_id": loan_id,
            "rule_applied": rule,
            "outcome": outcome,
        })

    for loan in loans:
        loan_id = loan["loan_id"]
        dpd = loan["days_past_due"]
        current_status = loan["loan_status"]
        case = cases_by_loan.get(loan_id)

        # --- Rule 1: status bucket recommendation ---------------------------
        if current_status == "Written_Off":
            # Already a human/committee decision; engine doesn't touch it.
            recommended_status = "Written_Off"
            status_match = True
        else:
            recommended_status = dpd_to_status(dpd)
            status_match = (recommended_status == current_status)
            log_audit(loan_id, "dpd_to_status_bucket",
                      f"dpd={dpd} -> {recommended_status} "
                      f"({'matches' if status_match else 'DIFFERS FROM'} recorded status {current_status})")

        # --- Rule 2: default action + decision type by status ---------------
        if recommended_status == "Normal":
            action, decision = "No action — loan current", AUTO
            reason = "Days past due = 0"

        elif recommended_status == "PDO1":
            if case is None:
                action, decision = "Open collection case; send past-due reminder", AUTO
                reason = "First entry into PDO1; no case exists yet"
                log_audit(loan_id, "auto_open_case", "Case opened on PDO1 entry")
            else:
                action, decision = "Continue reminders; monitor for P2P", AUTO
                reason = f"PDO1, existing case status = {case['case_status']}"

        elif recommended_status in ("PDO2", "PDO3"):
            if case is None:
                action, decision = "Escalate: open case urgently (overdue, no case on file)", HUMAN
                reason = "Loan is past PDO1 with no active case — process gap, needs review"
                log_audit(loan_id, "missing_case_flag", "Loan overdue with no case — flagged for human review")
            else:
                broken = case["case_status"] == "P2P_Broken"
                escalation_level = case.get("escalation_level") or 0
                if broken and escalation_level >= MAX_BROKEN_P2P_BEFORE_FORCED_ESCALATION:
                    action = "Force escalation — repeated broken P2P (hard rule, no agent override)"
                    decision = HUMAN
                    reason = f"escalation_level={escalation_level}, case_status=P2P_Broken — exceeds allowed cycles"
                    log_audit(loan_id, "forced_escalation_rule",
                              f"Broken P2P with escalation_level={escalation_level} -> forced escalation, agent override disabled")
                elif case["case_status"] == "P2P_Set":
                    action, decision = f"Monitor P2P due {case.get('p2p_date')}", AUTO
                    reason = "Customer has an open promise-to-pay, not yet due or broken"
                elif case["case_status"] == "P2P_Honoured":
                    action, decision = "Verify payment cleared; prepare to step down status", AUTO
                    reason = "P2P honoured per case log"
                elif case["case_status"] == "Escalated":
                    action, decision = "Proceed with escalation matrix (field visit / reference calls)", HUMAN
                    reason = "Case already escalated per prior cycle"
                else:
                    action, decision = "Agent to contact customer, establish new P2P", AUTO
                    reason = f"Case open ({case['case_status']}), no P2P on file yet"

        elif recommended_status == "NAB":
            actions = []
            if loan_id not in reported_loan_ids:
                actions.append("Report to bureau (CIBIL) — not yet reported")
                log_audit(loan_id, "bureau_reporting_trigger", "NAB status, not yet in bureau log -> flagged to report")
            actions.append("Stop interest accrual / reclassify to NAB in core system")
            actions.append("Evaluate for legal notice or moratorium offer")
            action = "; ".join(actions)
            decision = HUMAN  # legal/moratorium decisions always human; bureau feed below is auto-flagged separately
            reason = "Loan in non-accrual territory — multiple human-approval actions required"

        elif recommended_status == "Written_Off":
            action, decision = "No further collection action — write-off already recorded", AUTO
            reason = "Loan previously written off by risk/asset committee; engine takes no further action"

        else:
            action, decision = "Review manually", HUMAN
            reason = "Unrecognised status combination"

        recommendations.append({
            "loan_id": loan_id,
            "customer_name": loan["customer_name"],
            "days_past_due": dpd,
            "recorded_status": current_status,
            "recommended_status": recommended_status,
            "status_match": "Yes" if status_match else "REVIEW",
            "case_status": case["case_status"] if case else "No case",
            "escalation_level": case["escalation_level"] if case else "",
            "assigned_agent": case["assigned_agent"] if case else "",
            "recommended_action": action,
            "decision_type": decision,
            "reason": reason,
        })

    # -----------------------------------------------------------------------
    # Write Engine_Recommendations + Audit_Log into the workbook
    # -----------------------------------------------------------------------
    for sheet_name in ("Engine_Recommendations", "Audit_Log"):
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]

    HEADER_FILL = PatternFill(start_color="7A2E2E", end_color="7A2E2E", fill_type="solid")
    HEADER_FONT = Font(color="FFFFFF", bold=True)
    AUTO_FILL = PatternFill(start_color="E2F0D9", end_color="E2F0D9", fill_type="solid")
    HUMAN_FILL = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")

    rec_cols = ["loan_id", "customer_name", "days_past_due", "recorded_status",
                "recommended_status", "status_match", "case_status", "escalation_level",
                "assigned_agent", "recommended_action", "decision_type", "reason"]
    ws_rec = wb.create_sheet("Engine_Recommendations")
    ws_rec.append(rec_cols)
    for cell in ws_rec[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for r in recommendations:
        ws_rec.append([r[c] for c in rec_cols])
    # color decision_type column
    decision_col_idx = rec_cols.index("decision_type") + 1
    for row in ws_rec.iter_rows(min_row=2, min_col=decision_col_idx, max_col=decision_col_idx):
        for cell in row:
            cell.fill = AUTO_FILL if cell.value == AUTO else HUMAN_FILL
    for i, col in enumerate(rec_cols, start=1):
        max_len = max([len(str(col))] + [len(str(r[col])) for r in recommendations])
        ws_rec.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 45)
    ws_rec.freeze_panes = "A2"

    audit_cols = ["timestamp", "loan_id", "rule_applied", "outcome"]
    ws_audit = wb.create_sheet("Audit_Log")
    ws_audit.append(audit_cols)
    for cell in ws_audit[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for a in audit_entries:
        ws_audit.append([a[c] for c in audit_cols])
    for i, col in enumerate(audit_cols, start=1):
        max_len = max([len(str(col))] + [len(str(a[col])) for a in audit_entries]) if audit_entries else len(col)
        ws_audit.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 70)
    ws_audit.freeze_panes = "A2"

    wb.save(OUT_PATH)

    # -----------------------------------------------------------------------
    # Console summary
    # -----------------------------------------------------------------------
    auto_count = sum(1 for r in recommendations if r["decision_type"] == AUTO)
    human_count = sum(1 for r in recommendations if r["decision_type"] == HUMAN)
    mismatches = sum(1 for r in recommendations if r["status_match"] == "REVIEW")

    print(f"Processed {len(loans)} loans.")
    print(f"Auto-handled: {auto_count} | Human-required: {human_count}")
    print(f"Status mismatches flagged for review: {mismatches}")
    print(f"Audit entries logged: {len(audit_entries)}")
    print(f"Saved engine output into: {OUT_PATH}")

    print("\n--- Sample recommendations ---")
    for r in recommendations[:8]:
        print(f"{r['loan_id']} | {r['recommended_status']:10s} | {r['decision_type']:14s} | {r['recommended_action']}")


if __name__ == "__main__":
    main()
