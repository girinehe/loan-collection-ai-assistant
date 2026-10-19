"""
Builds synthetic data for the Loan Collection Assistant prototype.
Creates Loan_Collection_Prototype.xlsx with four sheets:
  - Loan_Contracts
  - Collection_Cases
  - Payments
  - Bureau_Reporting_Log

All data is fully synthetic (invented names, numbers) for demo purposes only.
"""

import random
from datetime import date, timedelta

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

random.seed(42)

TODAY = date(2026, 10, 16)

FIRST_NAMES = ["Ravi", "Sunita", "Arjun", "Meera", "Vikram", "Anjali", "Rahul",
               "Priya", "Sanjay", "Kavita", "Amit", "Neha", "Deepak", "Pooja",
               "Manoj", "Rekha", "Suresh", "Lata", "Vijay", "Shalini"]
LAST_NAMES = ["Sharma", "Patel", "Kumar", "Reddy", "Iyer", "Nair", "Gupta",
              "Joshi", "Mehta", "Rao", "Desai", "Pillai", "Kulkarni", "Singh",
              "Verma", "Bhat", "Chauhan", "Shetty", "Agarwal", "Menon"]

PRODUCT_TYPES = ["Personal Loan", "Auto Loan", "Home Loan", "Business Loan"]
AGENTS = ["R. Thakur (Internal)", "S. Naik (Internal)", "AgentPro Collections (Outsourced)",
          "QuickRecover Agency (Outsourced)", "M. Iyer (Internal)"]
CHANNELS = ["Call", "SMS", "WhatsApp", "Email"]

def synth_name():
    return f"{random.choice(FIRST_NAMES)} {random.choice(LAST_NAMES)}"

def rand_date(start, end):
    delta = (end - start).days
    return start + timedelta(days=random.randint(0, max(delta, 0)))

# ---------------------------------------------------------------------------
# Define 20 loans across a realistic spread of statuses
# ---------------------------------------------------------------------------
# status plan: 6 Normal, 4 PDO1, 4 PDO2, 3 PDO3, 2 NAB, 1 Written_Off
status_plan = (
    ["Normal"] * 6 +
    ["PDO1"] * 4 +
    ["PDO2"] * 4 +
    ["PDO3"] * 3 +
    ["NAB"] * 2 +
    ["Written_Off"] * 1
)

loans = []
for i, status in enumerate(status_plan, start=1):
    loan_id = f"LN{i:05d}"
    customer_id = f"CUST{i:04d}"
    customer_name = synth_name()
    product = random.choice(PRODUCT_TYPES)
    principal = random.choice([200000, 350000, 500000, 750000, 1000000, 1500000, 2500000])
    tenure = random.choice([12, 24, 36, 48, 60])
    rate = round(random.uniform(9.5, 15.5), 2)
    emi = round(principal / tenure * (1 + rate / 100 / 2), -1)  # rough synthetic EMI
    disb_date = rand_date(date(2022, 1, 1), date(2026, 6, 1))

    if status == "Normal":
        dpd = 0
        next_due = TODAY + timedelta(days=random.randint(1, 20))
        last_pay = TODAY - timedelta(days=random.randint(5, 25))
    elif status == "PDO1":
        dpd = random.randint(1, 30)
        next_due = TODAY - timedelta(days=dpd) + timedelta(days=30)
        last_pay = TODAY - timedelta(days=dpd + random.randint(20, 35))
    elif status == "PDO2":
        dpd = random.randint(31, 60)
        next_due = TODAY - timedelta(days=dpd) + timedelta(days=30)
        last_pay = TODAY - timedelta(days=dpd + random.randint(20, 35))
    elif status == "PDO3":
        dpd = random.randint(61, 90)
        next_due = TODAY - timedelta(days=dpd) + timedelta(days=30)
        last_pay = TODAY - timedelta(days=dpd + random.randint(20, 40))
    elif status == "NAB":
        dpd = random.randint(95, 150)
        next_due = TODAY - timedelta(days=dpd) + timedelta(days=30)
        last_pay = TODAY - timedelta(days=dpd + random.randint(30, 60))
    else:  # Written_Off
        dpd = random.randint(180, 365)
        next_due = TODAY - timedelta(days=dpd) + timedelta(days=30)
        last_pay = TODAY - timedelta(days=dpd + random.randint(60, 120))

    loans.append({
        "loan_id": loan_id,
        "customer_id": customer_id,
        "customer_name": customer_name,
        "product_type": product,
        "principal_amount": principal,
        "disbursement_date": disb_date,
        "emi_amount": emi,
        "tenure_months": tenure,
        "interest_rate": rate,
        "loan_status": status,
        "days_past_due": dpd,
        "next_due_date": next_due,
        "last_payment_date": last_pay,
    })

# ---------------------------------------------------------------------------
# Collection_Cases: only loans in PDO1 and beyond get a case
# ---------------------------------------------------------------------------
cases = []
case_counter = 1
for loan in loans:
    if loan["loan_status"] == "Normal":
        continue

    status = loan["loan_status"]
    opened_date = TODAY - timedelta(days=loan["days_past_due"] - random.randint(0, 3))

    # decide case_status and escalation based on severity
    if status == "PDO1":
        case_status = random.choice(["Open", "P2P_Set", "P2P_Honoured"])
        escalation = 0
    elif status == "PDO2":
        case_status = random.choice(["P2P_Set", "P2P_Broken", "Open"])
        escalation = random.choice([0, 1])
    elif status == "PDO3":
        case_status = random.choice(["P2P_Broken", "Escalated"])
        escalation = random.choice([1, 2])
    elif status == "NAB":
        case_status = "Escalated"
        escalation = random.choice([2, 3])
    else:  # Written_Off
        case_status = "Closed"
        escalation = 3

    p2p_date = None
    p2p_amount = None
    if case_status in ("P2P_Set", "P2P_Honoured", "P2P_Broken"):
        p2p_date = opened_date + timedelta(days=random.randint(3, 15))
        p2p_amount = loan["emi_amount"]

    chosen_agent = random.choice(AGENTS)
    cases.append({
        "case_id": f"CASE{case_counter:05d}",
        "loan_id": loan["loan_id"],
        "case_opened_date": opened_date,
        "assigned_agent": chosen_agent,
        "agent_type": "Outsourced" if "Outsourced" in chosen_agent else "Internal",
        "case_status": case_status,
        "p2p_date": p2p_date,
        "p2p_amount": p2p_amount,
        "contact_channel": random.choice(CHANNELS),
        "contact_count": random.randint(1, 8),
        "escalation_level": escalation,
        "notes": "" ,
    })
    case_counter += 1

# ---------------------------------------------------------------------------
# Payments: a handful of historical payments per loan
# ---------------------------------------------------------------------------
payments = []
pay_counter = 1
for loan in loans:
    n_payments = random.randint(2, 6)
    pay_date = loan["disbursement_date"] + timedelta(days=30)
    for _ in range(n_payments):
        if pay_date > TODAY:
            break
        payments.append({
            "payment_id": f"PAY{pay_counter:05d}",
            "loan_id": loan["loan_id"],
            "payment_date": pay_date,
            "amount": loan["emi_amount"] if random.random() > 0.15 else round(loan["emi_amount"] * random.uniform(0.4, 0.9), -1),
            "payment_type": random.choice(["EMI", "EMI", "EMI", "Manual", "Partial"]),
        })
        pay_counter += 1
        pay_date += timedelta(days=30)

# ---------------------------------------------------------------------------
# Bureau_Reporting_Log: loans at PDO2 and beyond get reported
# ---------------------------------------------------------------------------
bureau_log = []
log_counter = 1
for loan in loans:
    if loan["loan_status"] in ("PDO2", "PDO3", "NAB", "Written_Off"):
        bureau_log.append({
            "report_id": f"RPT{log_counter:05d}",
            "loan_id": loan["loan_id"],
            "report_date": TODAY - timedelta(days=random.randint(1, 15)),
            "reported_status": loan["loan_status"],
            "days_past_due_reported": loan["days_past_due"],
        })
        log_counter += 1

# ---------------------------------------------------------------------------
# Write to Excel with light formatting
# ---------------------------------------------------------------------------
wb = openpyxl.Workbook()
wb.remove(wb.active)

HEADER_FILL = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True)

def write_sheet(name, rows, columns):
    ws = wb.create_sheet(name)
    ws.append(columns)
    for cell in ws[1]:
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center")
    for row in rows:
        ws.append([row[c] for c in columns])
    # autosize columns roughly
    for i, col in enumerate(columns, start=1):
        max_len = max([len(str(col))] + [len(str(r[col])) for r in rows]) if rows else len(str(col))
        ws.column_dimensions[get_column_letter(i)].width = min(max_len + 3, 28)
    ws.freeze_panes = "A2"
    return ws

write_sheet("Loan_Contracts", loans, [
    "loan_id", "customer_id", "customer_name", "product_type", "principal_amount",
    "disbursement_date", "emi_amount", "tenure_months", "interest_rate",
    "loan_status", "days_past_due", "next_due_date", "last_payment_date"
])

write_sheet("Collection_Cases", cases, [
    "case_id", "loan_id", "case_opened_date", "assigned_agent", "agent_type",
    "case_status", "p2p_date", "p2p_amount", "contact_channel", "contact_count",
    "escalation_level", "notes"
])

write_sheet("Payments", payments, [
    "payment_id", "loan_id", "payment_date", "amount", "payment_type"
])

write_sheet("Bureau_Reporting_Log", bureau_log, [
    "report_id", "loan_id", "report_date", "reported_status", "days_past_due_reported"
])

# Add a README sheet explaining the synthetic dataset
readme = wb.create_sheet("README", 0)
readme["A1"] = "Loan Collection Assistant — Synthetic Prototype Data"
readme["A1"].font = Font(bold=True, size=14)
readme["A3"] = "This workbook is 100% synthetic data created for a portfolio prototype. No real customers, loans, or bank data are used."
readme["A5"] = "Sheets:"
readme["A6"] = "Loan_Contracts — mirrors the core banking/LMS loan master record"
readme["A7"] = "Collection_Cases — mirrors the loan collection system's case workflow (P2P tracking, escalation)"
readme["A8"] = "Payments — transaction history per loan"
readme["A9"] = "Bureau_Reporting_Log — simplified CIBIL/bureau reporting feed"
readme["A11"] = f"Status mix: {status_plan.count('Normal')} Normal, {status_plan.count('PDO1')} PDO1, {status_plan.count('PDO2')} PDO2, {status_plan.count('PDO3')} PDO3, {status_plan.count('NAB')} NAB, {status_plan.count('Written_Off')} Written_Off"
readme["A13"] = f"Generated as of (as-of date): {TODAY.isoformat()}"
for row in readme.iter_rows(min_row=1, max_row=13, min_col=1, max_col=1):
    for c in row:
        if c.value:
            c.alignment = Alignment(wrap_text=True)
readme.column_dimensions["A"].width = 100

OUT_PATH = "Loan_Collection_Prototype.xlsx"
wb.save(OUT_PATH)
print(f"Saved {OUT_PATH}")
print(f"Loans: {len(loans)} | Cases: {len(cases)} | Payments: {len(payments)} | Bureau log entries: {len(bureau_log)}")
