
import os
from io import BytesIO
from datetime import date
import pandas as pd
import streamlit as st
from dotenv import load_dotenv
from supabase import create_client
from reportlab.lib.pagesizes import landscape, letter
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet

load_dotenv()
st.set_page_config(
    page_title="Fabrication Business Manager",
    page_icon="🏗️",
    layout="wide",
)

TABLES = {
    "Customers": "fb_customers",
    "Suppliers": "fb_suppliers",
    "Workers": "fb_workers",
    "Projects / Sites": "fb_sites",
    "Transactions": "fb_transactions",
    "Attendance": "fb_attendance",
    "Worker Advances": "fb_worker_advances",
    "Salary Records": "fb_salary_records",
    "Site Visits": "fb_site_visits",
}


def make_client():
    url = os.getenv("SUPABASE_URL")
    key = os.getenv("SUPABASE_KEY")
    if not url or not key:
        st.error("Check SUPABASE_URL and SUPABASE_KEY in your .env file.")
        st.stop()
    return create_client(url, key)


def db():
    client = make_client()
    token = st.session_state.get("access_token")
    if token:
        client.postgrest.auth(token)
    return client


def get_rows(table, order="created_at"):
    query = db().table(table).select("*")
    try:
        return query.order(order, desc=True).execute().data or []
    except Exception:
        return query.execute().data or []


def save_row(table, values):
    return db().table(table).insert(values).execute()


def show_rows(records, heading):
    st.subheader(heading)
    if records:
        df = pd.DataFrame(records).drop(columns=["id"], errors="ignore")
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("No records yet.")


def export_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Report")
    return output.getvalue()


def export_pdf(df, title):
    output = BytesIO()
    doc = SimpleDocTemplate(output, pagesize=landscape(letter))
    data = [list(map(str, df.columns))]
    data += df.fillna("").astype(str).values.tolist()
    table = Table(data, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
        ("FONTSIZE", (0, 0), (-1, -1), 7),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
    ]))
    styles = getSampleStyleSheet()
    doc.build([Paragraph(title, styles["Title"]), Spacer(1, 10), table])
    return output.getvalue()


def export_buttons(df, filename):
    if df.empty:
        return
    a, b, c = st.columns(3)
    a.download_button(
        "Download Excel", export_excel(df), f"{filename}.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    b.download_button(
        "Download CSV", df.to_csv(index=False).encode("utf-8-sig"),
        f"{filename}.csv", "text/csv",
    )
    c.download_button(
        "Download PDF", export_pdf(df, filename.replace("_", " ").title()),
        f"{filename}.pdf", "application/pdf",
    )


def money(value):
    return f"₹{float(value or 0):,.2f}"


def loan_calculation(loan, repayments, as_of=None):
    """Calculate simple interest on the original principal."""
    as_of = as_of or date.today()
    start = date.fromisoformat(str(loan["loan_date"])[:10])
    days = max(0, (as_of - start).days)
    principal = float(loan["principal_amount"] or 0)
    interest_type = loan["interest_type"]

    if interest_type == "Fixed":
        interest = float(loan.get("fixed_interest_amount") or 0)
    elif interest_type == "Annual Simple":
        rate = float(loan.get("annual_interest_rate") or 0)
        interest = principal * (rate / 100) * days / 365
    else:
        interest = 0.0

    paid = sum(float(r["amount"] or 0) for r in repayments)
    total_due = principal + interest
    balance = max(0.0, total_due - paid)

    if balance <= 0.005:
        status = "Paid off"
    elif paid > 0:
        status = "Partially paid"
    else:
        status = "Unpaid"

    return {
        "principal": principal,
        "interest": interest,
        "total_due": total_due,
        "paid": paid,
        "balance": balance,
        "status": status,
        "days": days,
    }


def get_loan_repayments(loan_id):
    return (
        db().table("fb_loan_repayments")
        .select("*")
        .eq("loan_id", loan_id)
        .order("payment_date", desc=False)
        .execute().data or []
    )


st.title("🏗️ Fabrication Business Manager")
st.caption("Projects • Accounts • Workers • Loans • Reports")

if "user_email" not in st.session_state:
    st.session_state.user_email = None
if "access_token" not in st.session_state:
    st.session_state.access_token = None

# LOGIN
if not st.session_state.user_email:
    st.subheader("Secure Login")
    with st.form("login_form"):
        email = st.text_input("Approved email")
        password = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Log in", type="primary")
    if submitted:
        try:
            result = make_client().auth.sign_in_with_password({
                "email": email.strip(),
                "password": password,
            })
            st.session_state.user_email = result.user.email
            st.session_state.access_token = result.session.access_token
            st.rerun()
        except Exception:
            st.error("Login failed. Check your email and password.")
    st.info("Use an account already created in Supabase Authentication.")
    st.stop()

left, right = st.columns([4, 1])
left.success(f"Signed in as {st.session_state.user_email}")
if right.button("Log out"):
    st.session_state.user_email = None
    st.session_state.access_token = None
    st.rerun()

# CHECK DATABASE
try:
    db().table("fb_sites").select("id").limit(1).execute()
    db().table("fb_worker_loans").select("id").limit(1).execute()
    db().table("fb_loan_repayments").select("id").limit(1).execute()
except Exception as error:
    st.error("A required database table or permission is missing.")
    st.code(str(error))
    st.stop()

page = st.sidebar.radio(
    "MAIN MENU",
    [
        "Dashboard", "Projects / Sites", "Transactions", "Customers",
        "Suppliers", "Workers", "Attendance", "Worker Advances",
        "Salary Records", "Worker Loans", "Site Visits", "Reports",
    ],
)

# DASHBOARD
if page == "Dashboard":
    projects = get_rows("fb_sites")
    transactions = get_rows("fb_transactions", "transaction_date")
    customers = get_rows("fb_customers")
    workers = get_rows("fb_workers")

    income = sum(float(r.get("amount") or 0) for r in transactions
                 if r.get("direction") == "Income")
    expenses = sum(float(r.get("amount") or 0) for r in transactions
                   if r.get("direction") == "Expense")
    budgets = sum(float(r.get("agreed_budget") or 0) for r in projects)

    a, b, c, d = st.columns(4)
    a.metric("Projects", len(projects))
    b.metric("Customers", len(customers))
    c.metric("Money Received", money(income))
    d.metric("Recorded Expenses", money(expenses))

    a, b, c = st.columns(3)
    a.metric("Agreed Budgets", money(budgets))
    b.metric("Net Cash Flow", money(income - expenses))
    c.metric("Workers", len(workers))

    st.caption("Cash flow is not the same as profit. Loan balances are shown in Worker Loans.")
    show_rows(transactions[:10], "Recent Transactions")

# PROJECTS / SITES
elif page == "Projects / Sites":
    customers = get_rows("fb_customers")
    customer_map = {r["name"]: r["id"] for r in customers}

    with st.form("project_form"):
        name = st.text_input("Project / Site Name *")
        customer = st.selectbox("Customer", ["— None —"] + list(customer_map))
        budget = st.number_input("Agreed Budget (₹)", min_value=0.0, step=1000.0)
        location = st.text_input("Site Location")
        start = st.date_input("Start Date", value=date.today())
        status = st.selectbox("Status", ["Active", "On hold", "Completed", "Cancelled"])
        notes = st.text_area("Project Details")
        submitted = st.form_submit_button("Save Project", type="primary")

    if submitted:
        if not name.strip():
            st.error("Enter a project name.")
        else:
            try:
                save_row("fb_sites", {
                    "site_name": name.strip(),
                    "customer_id": customer_map.get(customer),
                    "agreed_budget": budget,
                    "location": location.strip() or None,
                    "start_date": str(start),
                    "status": status,
                    "notes": notes.strip() or None,
                })
                st.success("Project saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_sites")
    show_rows(records, "All Projects")
    export_buttons(pd.DataFrame(records), "projects")

# TRANSACTIONS
elif page == "Transactions":
    sites = get_rows("fb_sites")
    customers = get_rows("fb_customers")
    suppliers = get_rows("fb_suppliers")
    site_map = {r["site_name"]: r["id"] for r in sites}
    customer_map = {r["name"]: r["id"] for r in customers}
    supplier_map = {r["name"]: r["id"] for r in suppliers}

    with st.form("transaction_form"):
        day = st.date_input("Transaction Date", value=date.today())
        direction = st.selectbox("Transaction Type", ["Income", "Expense"])
        category = st.selectbox("Category", [
            "Customer Payment", "Materials / Steel", "Labour", "Transport",
            "Workshop Expense", "Electricity", "Repair", "Supplier Payment",
            "Equipment", "Other",
        ])
        amount = st.number_input("Amount (₹)", min_value=0.0, step=100.0)
        site = st.selectbox("Project / Site", ["— None —"] + list(site_map))
        customer = st.selectbox("Customer", ["— None —"] + list(customer_map))
        supplier = st.selectbox("Supplier", ["— None —"] + list(supplier_map))
        method = st.selectbox("Payment Method", [
            "Cash", "UPI", "Bank Transfer", "Cheque", "Other",
        ])
        description = st.text_input("Description / Reference")
        submitted = st.form_submit_button("Save Transaction", type="primary")

    if submitted:
        if amount <= 0:
            st.error("Amount must be greater than zero.")
        else:
            try:
                save_row("fb_transactions", {
                    "transaction_date": str(day),
                    "direction": direction,
                    "category": category,
                    "amount": amount,
                    "site_id": site_map.get(site),
                    "customer_id": customer_map.get(customer),
                    "supplier_id": supplier_map.get(supplier),
                    "payment_method": method,
                    "description": description.strip() or None,
                })
                st.success("Transaction saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_transactions", "transaction_date")
    show_rows(records, "All Transactions")
    export_buttons(pd.DataFrame(records), "transactions")

# CUSTOMERS, SUPPLIERS, WORKERS
elif page in ("Customers", "Suppliers", "Workers"):
    table = TABLES[page]
    with st.form("people_form"):
        name = st.text_input("Name *")
        phone = st.text_input("Phone")
        address = st.text_input("Address")
        role = st.text_input("Role / Supplier Type")
        daily_rate = 0.0
        monthly_salary = 0.0
        active = True
        if page == "Workers":
            daily_rate = st.number_input("Daily Rate (₹)", min_value=0.0, step=100.0)
            monthly_salary = st.number_input("Monthly Salary (₹)", min_value=0.0, step=1000.0)
            active = st.checkbox("Active Worker", value=True)
        notes = st.text_area("Notes")
        submitted = st.form_submit_button("Save Record", type="primary")

    if submitted:
        if not name.strip():
            st.error("Name is required.")
        else:
            data = {
                "name": name.strip(),
                "phone": phone.strip() or None,
                "notes": notes.strip() or None,
            }
            if page in ("Customers", "Suppliers"):
                data["address"] = address.strip() or None
            if page == "Suppliers":
                data["supplier_type"] = role.strip() or None
            if page == "Workers":
                data.update({
                    "role": role.strip() or None,
                    "daily_rate": daily_rate,
                    "monthly_salary": monthly_salary,
                    "active": active,
                })
            try:
                save_row(table, data)
                st.success("Record saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows(table)
    show_rows(records, f"All {page}")
    export_buttons(pd.DataFrame(records), page.lower().replace(" ", "_"))

# ATTENDANCE
elif page == "Attendance":
    workers = get_rows("fb_workers")
    sites = get_rows("fb_sites")
    worker_map = {r["name"]: r["id"] for r in workers}
    site_map = {r["site_name"]: r["id"] for r in sites}

    with st.form("attendance_form"):
        worker = st.selectbox("Worker", list(worker_map) or ["Add a worker first"])
        day = st.date_input("Date", value=date.today())
        status = st.selectbox("Attendance", ["Present", "Half day", "Absent", "Paid leave"])
        site = st.selectbox("Site", ["— None —"] + list(site_map))
        notes = st.text_input("Notes")
        submitted = st.form_submit_button("Save Attendance")

    if submitted:
        if not worker_map:
            st.error("Add a worker first.")
        else:
            try:
                save_row("fb_attendance", {
                    "worker_id": worker_map[worker],
                    "attendance_date": str(day),
                    "status": status,
                    "site_id": site_map.get(site),
                    "notes": notes.strip() or None,
                })
                st.success("Attendance saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_attendance", "attendance_date")
    show_rows(records, "Attendance History")
    export_buttons(pd.DataFrame(records), "attendance")

# WORKER ADVANCES
elif page == "Worker Advances":
    workers = get_rows("fb_workers")
    worker_map = {r["name"]: r["id"] for r in workers}
    with st.form("advance_form"):
        worker = st.selectbox("Worker", list(worker_map) or ["Add a worker first"])
        day = st.date_input("Advance Date", value=date.today())
        amount = st.number_input("Advance Given (₹)", min_value=0.0, step=100.0)
        recovered = st.number_input("Recovered So Far (₹)", min_value=0.0, step=100.0)
        notes = st.text_input("Notes")
        submitted = st.form_submit_button("Save Advance")

    if submitted:
        if not worker_map:
            st.error("Add a worker first.")
        elif amount <= 0 or recovered > amount:
            st.error("Enter a valid advance and recovery amount.")
        else:
            try:
                save_row("fb_worker_advances", {
                    "worker_id": worker_map[worker],
                    "advance_date": str(day),
                    "amount": amount,
                    "recovered_amount": recovered,
                    "notes": notes.strip() or None,
                })
                st.success("Advance saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_worker_advances", "advance_date")
    show_rows(records, "Worker Advances")
    export_buttons(pd.DataFrame(records), "worker_advances")

# SALARY RECORDS
elif page == "Salary Records":
    workers = get_rows("fb_workers")
    worker_map = {r["name"]: r["id"] for r in workers}
    with st.form("salary_form"):
        worker = st.selectbox("Worker", list(worker_map) or ["Add a worker first"])
        start = st.date_input("Period Start", value=date.today().replace(day=1))
        end = st.date_input("Period End", value=date.today())
        salary = st.number_input("Salary Earned (₹)", min_value=0.0, step=100.0)
        advance = st.number_input("Advance Deduction (₹)", min_value=0.0, step=100.0)
        other = st.number_input("Other Deduction (₹)", min_value=0.0, step=100.0)
        paid = st.number_input("Amount Actually Paid (₹)", min_value=0.0, step=100.0)
        paid_date = st.date_input("Payment Date", value=date.today())
        notes = st.text_input("Notes")
        submitted = st.form_submit_button("Save Salary Record")

    if submitted:
        if not worker_map:
            st.error("Add a worker first.")
        elif end < start or advance + other + paid > salary:
            st.error("Check dates and salary deductions/payment amounts.")
        else:
            try:
                save_row("fb_salary_records", {
                    "worker_id": worker_map[worker],
                    "period_start": str(start),
                    "period_end": str(end),
                    "salary_amount": salary,
                    "advance_deduction": advance,
                    "other_deduction": other,
                    "amount_paid": paid,
                    "paid_date": str(paid_date),
                    "notes": notes.strip() or None,
                })
                st.success("Salary record saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_salary_records", "period_end")
    if records:
        df = pd.DataFrame(records)
        df["balance_due"] = (
            df["salary_amount"].astype(float)
            - df["advance_deduction"].astype(float)
            - df["other_deduction"].astype(float)
            - df["amount_paid"].astype(float)
        )
        show_rows(df.to_dict("records"), "Salary and Outstanding Wages")
        export_buttons(df, "salary_records")
    else:
        st.info("No salary records yet.")

# WORKER LOANS
elif page == "Worker Loans":
    workers = get_rows("fb_workers")
    worker_map = {r["name"]: r["id"] for r in workers}
    tab1, tab2, tab3 = st.tabs([
        "Register Loan", "Record Repayment", "Loan Summary & History"
    ])

    with tab1:
        st.subheader("Register a New Worker Loan")
        with st.form("loan_form"):
            worker = st.selectbox("Worker", list(worker_map) or ["Add a worker first"])
            loan_date = st.date_input("Loan Date", value=date.today())
            principal = st.number_input("Original Loan Amount (₹)", min_value=0.0, step=500.0)
            interest_type = st.selectbox(
                "Interest Type", ["None", "Fixed", "Annual Simple"]
            )
            fixed_interest = 0.0
            annual_rate = 0.0
            if interest_type == "Fixed":
                fixed_interest = st.number_input(
                    "Total Fixed Interest (₹)", min_value=0.0, step=100.0
                )
            elif interest_type == "Annual Simple":
                annual_rate = st.number_input(
                    "Annual Simple Interest Rate (%)",
                    min_value=0.0, max_value=100.0, step=1.0
                )
                st.caption(
                    "Interest is calculated on the original principal "
                    "for elapsed days, using a 365-day year."
                )
            notes = st.text_area("Loan Notes / Agreement Reference")
            submitted = st.form_submit_button("Save Loan", type="primary")

        if submitted:
            if not worker_map:
                st.error("Add a worker first.")
            elif principal <= 0:
                st.error("Loan amount must be greater than zero.")
            else:
                try:
                    save_row("fb_worker_loans", {
                        "worker_id": worker_map[worker],
                        "loan_date": str(loan_date),
                        "principal_amount": principal,
                        "interest_type": interest_type,
                        "fixed_interest_amount": fixed_interest,
                        "annual_interest_rate": annual_rate,
                        "notes": notes.strip() or None,
                    })
                    st.success("Worker loan saved.")
                    st.rerun()
                except Exception as error:
                    st.error(str(error))

    with tab2:
        st.subheader("Record a Repayment")
        loans = db().table("fb_worker_loans").select("*").order(
            "loan_date", desc=True
        ).execute().data or []

        if not loans:
            st.info("Register a loan first.")
        else:
            loan_labels = {}
            for loan in loans:
                worker_name = next(
                    (w["name"] for w in workers if w["id"] == loan["worker_id"]),
                    "Unknown worker",
                )
                loan_labels[
                    f'{worker_name} — {loan["loan_date"]} — {money(loan["principal_amount"])}'
                ] = loan["id"]

            selected_label = st.selectbox("Select Loan", list(loan_labels))
            selected_id = loan_labels[selected_label]
            selected_loan = next(l for l in loans if l["id"] == selected_id)
            repayments = get_loan_repayments(selected_id)
            calc = loan_calculation(selected_loan, repayments)

            a, b, c = st.columns(3)
            a.metric("Principal", money(calc["principal"]))
            b.metric("Interest accrued", money(calc["interest"]))
            c.metric("Outstanding balance", money(calc["balance"]))
            st.caption(f"Status: {calc['status']}")

            if calc["balance"] <= 0.005:
                st.success("This loan is fully repaid. No further payment is due.")
            else:
                with st.form("repayment_form"):
                    payment_date = st.date_input("Payment Date", value=date.today())
                    amount = st.number_input(
                        "Amount Repaid (₹)", min_value=0.0,
                        max_value=float(calc["balance"]), step=100.0
                    )
                    method = st.selectbox(
                        "Payment Method", ["Cash", "UPI", "Bank Transfer", "Cheque", "Other"]
                    )
                    notes = st.text_input("Repayment Reference / Notes")
                    submitted = st.form_submit_button("Save Repayment", type="primary")

                if submitted:
                    if amount <= 0:
                        st.error("Repayment must be greater than zero.")
                    elif amount > calc["balance"] + 0.005:
                        st.error("Repayment cannot exceed the outstanding balance.")
                    else:
                        try:
                            save_row("fb_loan_repayments", {
                                "loan_id": selected_id,
                                "payment_date": str(payment_date),
                                "amount": amount,
                                "payment_method": method,
                                "notes": notes.strip() or None,
                            })
                            st.success("Repayment recorded.")
                            st.rerun()
                        except Exception as error:
                            st.error(str(error))

    with tab3:
        st.subheader("Loan Balances and Repayment History")
        loans = db().table("fb_worker_loans").select("*").order(
            "loan_date", desc=True
        ).execute().data or []

        summary = []
        for loan in loans:
            worker_name = next(
                (w["name"] for w in workers if w["id"] == loan["worker_id"]),
                "Unknown worker",
            )
            repayments = get_loan_repayments(loan["id"])
            calc = loan_calculation(loan, repayments)
            summary.append({
                "Worker": worker_name,
                "Loan Date": loan["loan_date"],
                "Interest Type": loan["interest_type"],
                "Principal (₹)": round(calc["principal"], 2),
                "Interest (₹)": round(calc["interest"], 2),
                "Total Due (₹)": round(calc["total_due"], 2),
                "Repaid (₹)": round(calc["paid"], 2),
                "Balance (₹)": round(calc["balance"], 2),
                "Status": calc["status"],
                "Loan ID": loan["id"],
            })

        if summary:
            summary_df = pd.DataFrame(summary)
            show_rows(summary_df.to_dict("records"), "All Worker Loans")
            export_buttons(summary_df, "worker_loan_summary")

            chosen = st.selectbox(
                "View Repayment History For",
                [f'{r["Worker"]} — {r["Loan Date"]} — {r["Loan ID"][:8]}'
                 for r in summary],
            )
            loan_id = chosen.split(" — ")[-1]
            loan_record = next(
                l for l in loans if l["id"].startswith(loan_id)
            )
            repayment_rows = get_loan_repayments(loan_record["id"])
            show_rows(repayment_rows, "Individual Repayment History")
            export_buttons(pd.DataFrame(repayment_rows), "repayment_history")
        else:
            st.info("No worker loans have been registered yet.")

        st.caption(
            "Interest rule: fixed interest is added once; annual simple interest "
            "uses the original principal and elapsed days. Repayments are subtracted "
            "from the calculated total due. This version does not recalculate interest "
            "on a declining principal balance or allocate each repayment separately "
            "between principal and interest."
        )

# SITE VISITS
elif page == "Site Visits":
    sites = get_rows("fb_sites")
    site_map = {r["site_name"]: r["id"] for r in sites}

    with st.form("visit_form"):
        site = st.selectbox("Project / Site", list(site_map) or ["Add a project first"])
        day = st.date_input("Visit Date", value=date.today())
        visited_by = st.text_input("Visited By")
        notes = st.text_area("Visit Notes")
        submitted = st.form_submit_button("Save Site Visit")

    if submitted:
        if not site_map:
            st.error("Add a project first.")
        else:
            try:
                save_row("fb_site_visits", {
                    "site_id": site_map[site],
                    "visit_date": str(day),
                    "visited_by": visited_by.strip() or None,
                    "notes": notes.strip() or None,
                })
                st.success("Site visit saved.")
                st.rerun()
            except Exception as error:
                st.error(str(error))

    records = get_rows("fb_site_visits", "visit_date")
    show_rows(records, "Site Visit History")
    export_buttons(pd.DataFrame(records), "site_visits")

# REPORTS
elif page == "Reports":
    choice = st.selectbox("Choose Records", list(TABLES.keys()))
    records = get_rows(TABLES[choice])
    df = pd.DataFrame(records)
    search = st.text_input("Search records")

    if not df.empty and search.strip():
        df = df[df.astype(str).apply(
            lambda col: col.str.contains(search, case=False, na=False)
        ).any(axis=1)]

    show_rows(df.to_dict("records") if not df.empty else [], choice)
    export_buttons(df, choice.lower().replace(" ", "_"))

st.sidebar.divider()
st.sidebar.caption("Keep regular backups and verify business accounts.")
