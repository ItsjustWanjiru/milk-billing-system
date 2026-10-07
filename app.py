import calendar
import hmac
import io
import re
import time
import zipfile
from datetime import date, datetime, timedelta
from zoneinfo import ZoneInfo

import openpyxl
import pandas as pd
import plotly.express as px
import requests
import streamlit as st
from fpdf import FPDF
from fpdf.enums import XPos, YPos
from openpyxl.worksheet.worksheet import Worksheet

# ---------------------------------------------------------------------------
# BUSINESS DETAILS (edit these to change what appears on invoices)
# ---------------------------------------------------------------------------
BUSINESS_NAME = "AMANI DAIRIES"
TAGLINE = "Reliable - Fresh - Local"
PAYMENT_METHOD = "M-PESA POCHI LA BIASHARA"
PAYMENT_NUMBER = "0722 686 720"
INVOICE_PREFIX = "AMD"
PAYMENT_TERMS_DAYS = 7
TIMEZONE = ZoneInfo("Africa/Nairobi")

COLOR_PRIMARY = (16, 43, 85)
COLOR_SECONDARY = (212, 175, 55)
COLOR_TEXT = (50, 50, 50)
COLOR_ALERT = (200, 0, 0)

# ---------------------------------------------------------------------------
# SPREADSHEET LAYOUT (matches the existing monthly sheets)
# ---------------------------------------------------------------------------
FIRST_CUSTOMER_COL = 10  # Column J
NAME_ROW = 2
RATE_ROW = 3
FIRST_DAY_ROW = 4  # Row 4 is day 1, row 34 is day 31
PREPAID_ROW = 37
END_OF_CUSTOMERS = ["Total", "Unaccounted", "Summary", "Fridge"]
SKIP_SHEETS = ["summary", "data", "client", "total", "template"]

# Used only if GOOGLE_SHEET_URL is not set in Streamlit secrets
DEFAULT_SHEET_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "14ykcy9qOUPu-wLp7Xzp6SJWqVmXUzGAT/export?format=xlsx"
)

MONTHS = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}


# ---------------------------------------------------------------------------
# SMALL HELPERS
# ---------------------------------------------------------------------------
def get_secret(key, default=None):
    """Read a Streamlit secret without crashing when no secrets are configured."""
    try:
        return st.secrets.get(key, default)
    except Exception:
        return default


def now_local():
    return datetime.now(TIMEZONE)


def clean_num(value):
    """Turn a cell into a number. Accepts 4000, -4000, (4000) and 'KES 4,000'."""
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    in_brackets = text.startswith("(") and text.endswith(")")
    text = re.sub(r"[^0-9.\-]", "", text)
    if text in ("", "-", "."):
        return 0.0
    try:
        number = float(text)
    except ValueError:
        return 0.0
    return -abs(number) if in_brackets else number


def parse_month(sheet_name):
    """Return (year, month) from a sheet name like 'Sep 2026'. Either may be None."""
    lower = sheet_name.lower()
    year_match = re.search(r"20\d{2}", sheet_name)
    year = int(year_match.group(0)) if year_match else None
    month = None
    for key, number in MONTHS.items():
        if re.search(rf"(?<![a-z]){key}", lower):
            month = number
            break
    return year, month


def sort_key(sheet_name):
    year, month = parse_month(sheet_name)
    return (year or 0, month or 0)


def days_in_sheet_month(sheet_name):
    year, month = parse_month(sheet_name)
    if not month:
        return 31
    return calendar.monthrange(year or now_local().year, month)[1]


def make_safe(text):
    """Make text printable with the built-in PDF fonts."""
    replacements = {
        "\u2022": "-", "\u2013": "-", "\u2014": "-", "\u2219": ".",
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
    }
    text = str(text)
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode("latin-1", "replace").decode("latin-1")


def safe_filename(text):
    cleaned = re.sub(r"[^\w\- ]+", "_", str(text)).strip()
    return cleaned or "customer"


def fmt(value):
    return f"{value:,.2f}"


def fmt_qty(value):
    return f"{value:g}"


def invoice_number(sheet_name, position):
    """Stable invoice number, e.g. AMD-202609-004 for the 4th customer in Sep 2026."""
    year, month = parse_month(sheet_name)
    if year and month:
        period = f"{year}{month:02d}"
    else:
        period = re.sub(r"[^A-Za-z0-9]", "", sheet_name).upper()[:8] or "NA"
    return f"{INVOICE_PREFIX}-{period}-{position:03d}"


def has_nothing_to_bill(customer):
    return (
        customer["billed_qty"] == 0
        and customer["spoilt_qty"] == 0
        and customer["prepaid"] == 0
        and customer["previous_balance"] == 0
    )


# ---------------------------------------------------------------------------
# READING THE WORKBOOK
# ---------------------------------------------------------------------------
def is_spoilt(cell):
    """Spoilt deliveries are marked with a red cell fill."""
    color = str(cell.fill.start_color.index).upper()
    return color == "2" or color.endswith("FF0000")


def get_month_data(ws):
    if not isinstance(ws, Worksheet):
        return []
    if not ws.cell(row=NAME_ROW, column=FIRST_CUSTOMER_COL).value:
        return []

    customers = []
    for col in range(FIRST_CUSTOMER_COL, ws.max_column + 1):
        name = ws.cell(row=NAME_ROW, column=col).value
        if not name or any(word in str(name) for word in END_OF_CUSTOMERS):
            break

        rate = clean_num(ws.cell(row=RATE_ROW, column=col).value)

        # Row 37: a positive amount is pre-paid, a negative amount such as -4000
        # or (4000) is an unpaid balance from before.
        row37_raw = ws.cell(row=PREPAID_ROW, column=col).value
        row37 = clean_num(row37_raw)
        prepaid = row37 if row37 > 0 else 0.0
        previous_balance = -row37 if row37 < 0 else 0.0
        row37_unreadable = (
            row37 == 0
            and row37_raw is not None
            and str(row37_raw).strip() not in ("", "-", "0")
        )
        billed_qty = 0.0
        spoilt_qty = 0.0
        daily = {}
        spoilt_days = []

        for day in range(1, 32):
            cell = ws.cell(row=FIRST_DAY_ROW + day - 1, column=col)
            qty = clean_num(cell.value)
            daily[day] = qty
            if qty > 0 and is_spoilt(cell):
                spoilt_qty += qty
                spoilt_days.append((day, qty))
            else:
                billed_qty += qty

        total_bill = billed_qty * rate
        customers.append({
            "position": len(customers) + 1,
            "name": str(name).strip(),
            "billed_qty": billed_qty,
            "spoilt_qty": spoilt_qty,
            "rate": rate,
            "total_bill": total_bill,
            "lost_revenue": spoilt_qty * rate,
            "prepaid": prepaid,
            "previous_balance": previous_balance,
            "balance": total_bill + previous_balance - prepaid,
            "row37_unreadable": row37_unreadable,
            "row37_raw": row37_raw,
            "daily_liters": daily,
            "spoilt_details": spoilt_days,
        })
    return customers


@st.cache_data(show_spinner=False)
def load_all_months(file_bytes):
    """Parse every billing sheet once. Cached, so clicking around the app stays fast."""
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    results = {}
    for sheet_name in wb.sheetnames:
        if any(word in sheet_name.lower() for word in SKIP_SHEETS):
            continue
        data = get_month_data(wb[sheet_name])
        if data:
            results[sheet_name] = data
    return results


def data_quality_issues(sheet_name, customers):
    """Things worth checking before invoices go out."""
    issues = []
    last_day = days_in_sheet_month(sheet_name)
    for c in customers:
        if c["row37_unreadable"]:
            issues.append(
                f"{c['name']}: row {PREPAID_ROW} says '{c['row37_raw']}', which is not a number, "
                f"so it was treated as 0."
            )
        if c["billed_qty"] > 0 and c["rate"] == 0:
            issues.append(f"{c['name']}: milk recorded but no rate in row {RATE_ROW}.")
        extra_days = [d for d in range(last_day + 1, 32) if c["daily_liters"].get(d, 0) > 0]
        if extra_days:
            days_text = ", ".join(str(d) for d in extra_days)
            issues.append(
                f"{c['name']}: litres entered on day {days_text}, "
                f"but this month only has {last_day} days."
            )
    return issues


def fetch_google_sheet(url):
    separator = "&" if "?" in url else "?"
    session = requests.Session()
    session.mount("https://", requests.adapters.HTTPAdapter(max_retries=3))
    response = session.get(f"{url}{separator}t={int(time.time())}", timeout=60)
    response.raise_for_status()
    # A real .xlsx file is a zip archive, which always starts with "PK".
    if not response.content.startswith(b"PK"):
        raise ValueError(
            "Google returned a web page instead of an Excel file. "
            "Check that the sheet's sharing settings allow the app to read it."
        )
    return response.content


# ---------------------------------------------------------------------------
# PDF INVOICE
# ---------------------------------------------------------------------------
class AmaniInvoice(FPDF):
    def header(self):
        self.set_font("Helvetica", "B", 22)
        self.set_text_color(*COLOR_PRIMARY)
        self.cell(100, 10, BUSINESS_NAME)
        self.set_font("Helvetica", "I", 10)
        self.set_text_color(100, 100, 100)
        self.cell(90, 10, TAGLINE, align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_draw_color(*COLOR_SECONDARY)
        self.set_line_width(0.8)
        self.line(10, 20, 200, 20)
        self.set_line_width(0.2)  # reset so table borders stay thin
        self.ln(6)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(130, 130, 130)
        self.cell(0, 5, f"Thank you for choosing {BUSINESS_NAME.title()}.", align="C")

    def draw_calendar_grid(self, customer, year, month):
        days_in_month = calendar.monthrange(year, month)[1] if (year and month) else 31
        label_w = 20
        day_w = (self.w - self.l_margin - self.r_margin - label_w) / 31
        h = 5
        spoilt_days = {d for d, _ in customer["spoilt_details"]}

        self.set_font("Helvetica", "B", 8)
        self.set_text_color(*COLOR_PRIMARY)
        self.cell(0, 6, "DAILY CONSUMPTION BREAKDOWN", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        self.set_draw_color(0, 0, 0)
        self.set_text_color(0, 0, 0)

        # Row 1: day numbers
        self.set_fill_color(245, 245, 245)
        self.cell(label_w, h, "Date", border=1, align="C", fill=True)
        for d in range(1, 32):
            self.cell(day_w, h, str(d) if d <= days_in_month else "", border=1, align="C", fill=True)
        self.ln(h)

        # Row 2: weekday names
        self.cell(label_w, h, "Day", border=1, align="C")
        self.set_font("Helvetica", "", 6)
        self.set_fill_color(225, 225, 225)
        for d in range(1, 32):
            if d > days_in_month:
                self.cell(day_w, h, "", border=1, fill=True)
                continue
            label = date(year, month, d).strftime("%a") if (year and month) else ""
            self.cell(day_w, h, label, border=1, align="C")
        self.ln(h)

        # Row 3: litres (spoilt days shown in red)
        self.set_font("Helvetica", "B", 7)
        self.cell(label_w, h, "Litres", border=1, align="C")
        for d in range(1, 32):
            if d > days_in_month:
                self.cell(day_w, h, "", border=1, fill=True)
                continue
            qty = customer["daily_liters"].get(d, 0)
            if d in spoilt_days:
                self.set_text_color(*COLOR_ALERT)
            self.cell(day_w, h, fmt_qty(qty) if qty > 0 else "-", border=1, align="C")
            self.set_text_color(0, 0, 0)
        self.ln(h + 6)


def create_branded_pdf(customer, sheet_name, invoice_no, issue_date):
    year, month = parse_month(sheet_name)
    if month and not year:
        year = issue_date.year
    due_date = issue_date + timedelta(days=PAYMENT_TERMS_DAYS)
    balance = customer["balance"]
    is_credit = balance < 0

    pdf = AmaniInvoice(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    top = pdf.get_y()

    # Bill to (left)
    pdf.set_xy(10, top)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(95, 5, "BILL TO:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "B", 13)
    pdf.set_text_color(*COLOR_TEXT)
    pdf.multi_cell(95, 7, make_safe(customer["name"]))

    # Invoice details (right)
    details = [
        ("INVOICE NO:", invoice_no),
        ("BILLING PERIOD:", make_safe(sheet_name).upper()),
        ("INVOICE DATE:", issue_date.strftime("%d %b %Y")),
        ("DUE DATE:", due_date.strftime("%d %b %Y")),
    ]
    for i, (label, value) in enumerate(details):
        pdf.set_xy(110, top + i * 5)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(*COLOR_PRIMARY)
        pdf.cell(40, 5, label, align="R")
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*COLOR_TEXT)
        pdf.cell(50, 5, value, align="R")

    # Balance box (now shows the amount)
    box_y = top + 23
    pdf.set_draw_color(0, 0, 0)
    pdf.set_fill_color(245, 245, 245)
    pdf.set_xy(140, box_y)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(60, 6, "CREDIT BALANCE" if is_credit else "BALANCE DUE", border="LTR", align="C", fill=True)
    pdf.set_xy(140, box_y + 6)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(*(COLOR_PRIMARY if is_credit else COLOR_ALERT))
    pdf.cell(60, 9, f"KES {fmt(abs(balance))}", border="LRB", align="C")

    # Daily grid
    pdf.set_xy(10, box_y + 20)
    pdf.draw_calendar_grid(customer, year, month)

    # Line items
    pdf.set_fill_color(*COLOR_PRIMARY)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", "B", 9)
    pdf.cell(80, 8, "  Description", border=1, align="L", fill=True)
    pdf.cell(30, 8, "Total Qty (L)", border=1, align="C", fill=True)
    pdf.cell(30, 8, "Rate (KES)", border=1, align="C", fill=True)
    pdf.cell(50, 8, "Total (KES)  ", border=1, align="R", fill=True, new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    pdf.set_text_color(*COLOR_TEXT)
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(80, 10, "  Fresh Milk Supplied", border=1)
    pdf.cell(30, 10, f"{customer['billed_qty']:.1f}", border=1, align="C")
    pdf.cell(30, 10, fmt(customer["rate"]), border=1, align="C")
    pdf.set_font("Helvetica", "B", 10)
    pdf.cell(50, 10, f"{fmt(customer['total_bill'])}  ", border=1, align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Totals
    pdf.ln(1)
    pdf.set_font("Helvetica", "", 9)
    total_rows = [("Sub-Total (this month):", fmt(customer["total_bill"]))]
    if customer["previous_balance"] > 0:
        total_rows.append(("Previous Unpaid Balance:", fmt(customer["previous_balance"])))
    else:
        total_rows.append(("Pre-Paid:", f"- {fmt(customer['prepaid'])}"))
    for label, value in total_rows:
        pdf.set_x(100)
        pdf.cell(60, 6, label, align="R")
        pdf.cell(40, 6, value, align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_x(100)
    pdf.set_font("Helvetica", "B", 11)
    pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(60, 10, "CREDIT BALANCE:" if is_credit else "TOTAL DUE:", border="T", align="R")
    pdf.cell(40, 10, f"KES {fmt(abs(balance))}", border="T", align="R", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Spoilt milk notice (wraps instead of running off the page)
    if customer["spoilt_qty"] > 0:
        pdf.ln(6)
        pdf.set_font("Helvetica", "B", 8)
        pdf.set_text_color(*COLOR_ALERT)
        pdf.cell(0, 5, "SPOILT MILK NOTICE (excluded from the bill):", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(*COLOR_TEXT)
        days_text = ", ".join(f"Day {d} ({fmt_qty(q)}L)" for d, q in customer["spoilt_details"])
        pdf.multi_cell(
            0, 4,
            f"These deliveries were recorded as spoilt and were not charged (shown in red above): "
            f"{days_text}. Total: {fmt_qty(customer['spoilt_qty'])}L.",
        )

    # Payment box
    pdf.ln(10)
    if pdf.get_y() + 30 > pdf.h - 15:
        pdf.add_page()
    y = pdf.get_y()
    pdf.set_fill_color(252, 248, 227)
    pdf.set_draw_color(*COLOR_SECONDARY)
    pdf.rect(10, y, 190, 25, style="DF")
    pdf.set_y(y + 3)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*COLOR_PRIMARY)
    pdf.cell(0, 4, "PAYMENT METHOD", align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 6, PAYMENT_METHOD, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 8, PAYMENT_NUMBER, align="C", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    return bytes(pdf.output())


@st.cache_data(show_spinner=False)
def build_invoice_zip(file_bytes, sheet_name, skip_empty, issue_date):
    customers = load_all_months(file_bytes)[sheet_name]
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for cust in customers:
            if skip_empty and has_nothing_to_bill(cust):
                continue
            inv = invoice_number(sheet_name, cust["position"])
            pdf_bytes = create_branded_pdf(cust, sheet_name, inv, issue_date)
            zf.writestr(f"{inv} - {safe_filename(cust['name'])}.pdf", pdf_bytes)
    return buffer.getvalue()


# ---------------------------------------------------------------------------
# APP
# ---------------------------------------------------------------------------
def check_password():
    """Only asks for a password if APP_PASSWORD is set in Streamlit secrets."""
    password = get_secret("APP_PASSWORD")
    if not password or st.session_state.get("authenticated"):
        return True
    entered = st.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered, str(password)):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def main():
    st.set_page_config(page_title="Amani Dairies Dashboard", page_icon="🥛", layout="wide")
    if not check_password():
        st.stop()

    st.title("🥛 Amani Dairies Performance Tracker")
    sheet_url = get_secret("GOOGLE_SHEET_URL", DEFAULT_SHEET_URL)

    # --- DATA SOURCE ---
    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔄 Sync with Google Sheet"):
            try:
                with st.spinner("Fetching data from Google Sheets..."):
                    st.session_state["file_bytes"] = fetch_google_sheet(sheet_url)
                    st.session_state["source"] = f"Google Sheet, synced at {now_local():%H:%M}"
                st.success("Sync successful.")
            except Exception as e:
                st.error(f"Sync failed: {e}")
    with c2:
        uploaded = st.file_uploader("Or upload an Excel file", type=["xlsx", "xlsm"])
        if uploaded is not None:
            # Only take the upload when it is a new file, so it does not overwrite a later sync
            signature = (uploaded.name, uploaded.size)
            if st.session_state.get("upload_signature") != signature:
                st.session_state["upload_signature"] = signature
                st.session_state["file_bytes"] = uploaded.getvalue()
                st.session_state["source"] = f"Uploaded file: {uploaded.name}"

    if "file_bytes" not in st.session_state:
        st.info("Sync with the Google Sheet or upload an Excel file to get started.")
        st.stop()

    st.caption(f"Data source: {st.session_state['source']}")
    file_bytes = st.session_state["file_bytes"]

    try:
        with st.spinner("Reading billing sheets..."):
            all_months = load_all_months(file_bytes)
    except Exception as e:
        st.error(f"Could not read the workbook: {e}")
        st.stop()

    if not all_months:
        st.warning(f"No valid billing sheets found. Customer names should start in column J, row {NAME_ROW}.")
        st.stop()

    months_desc = sorted(all_months, key=sort_key, reverse=True)
    months_asc = list(reversed(months_desc))

    # --- BUSINESS OVERVIEW ---
    st.header("📊 Cumulative Business Overview")
    total_rev = sum(c["total_bill"] for data in all_months.values() for c in data)
    total_loss = sum(c["lost_revenue"] for data in all_months.values() for c in data)
    total_litres = sum(c["billed_qty"] for data in all_months.values() for c in data)
    potential = total_rev + total_loss
    loss_pct = (total_loss / potential * 100) if potential else 0

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Lifetime Revenue", f"KES {total_rev:,.0f}")
    m2.metric("Revenue Lost (Spoilage)", f"KES {total_loss:,.0f}",
              delta=f"{loss_pct:.1f}% of potential", delta_color="inverse")
    m3.metric("Litres Supplied", f"{total_litres:,.1f} L")
    m4.metric("Billing Months", len(all_months))

    trend = pd.DataFrame([
        {
            "Month": m,
            "Revenue": sum(c["total_bill"] for c in all_months[m]),
            "Spoilage loss": sum(c["lost_revenue"] for c in all_months[m]),
        }
        for m in months_asc
    ])
    cust_totals = {}
    for data in all_months.values():
        for c in data:
            cust_totals[c["name"]] = cust_totals.get(c["name"], 0) + c["total_bill"]
    top_customers = (
        pd.DataFrame(list(cust_totals.items()), columns=["Customer", "Revenue"])
        .sort_values("Revenue", ascending=False)
        .head(10)
        .sort_values("Revenue")
    )

    left, right = st.columns(2)
    with left:
        st.plotly_chart(px.bar(
            trend, x="Month", y=["Revenue", "Spoilage loss"], barmode="group",
            title="Revenue vs. Spoilage Loss by Month",
            labels={"value": "KES", "variable": ""},
            color_discrete_map={"Revenue": "#102B55", "Spoilage loss": "#C80000"},
        ))
    with right:
        st.plotly_chart(px.bar(
            top_customers, x="Revenue", y="Customer", orientation="h",
            title="Top 10 Customers by Revenue", labels={"Revenue": "KES", "Customer": ""},
            color_discrete_sequence=["#102B55"],
        ))

    st.divider()

    # --- MONTHLY INVOICES ---
    st.header("🧾 Monthly Invoices")
    target = st.selectbox("Select month", months_desc)
    customers = all_months[target]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Billed this month", f"KES {sum(c['total_bill'] for c in customers):,.0f}")
    k2.metric("Pre-paid", f"KES {sum(c['prepaid'] for c in customers):,.0f}")
    k3.metric("Previous balances", f"KES {sum(c['previous_balance'] for c in customers):,.0f}")
    k4.metric("Total outstanding", f"KES {sum(max(c['balance'], 0) for c in customers):,.0f}")
    k5.metric("Customers", len(customers))

    issues = data_quality_issues(target, customers)
    if issues:
        with st.expander(f"⚠️ {len(issues)} item(s) to check before sending invoices", expanded=True):
            for issue in issues:
                st.write(f"- {issue}")

    table = pd.DataFrame(customers)[
        ["name", "billed_qty", "spoilt_qty", "rate", "total_bill",
         "previous_balance", "prepaid", "balance"]
    ].rename(columns={
        "name": "Customer", "billed_qty": "Billed (L)", "spoilt_qty": "Spoilt (L)",
        "rate": "Rate", "total_bill": "This Month (KES)",
        "previous_balance": "Previous Balance (KES)", "prepaid": "Pre-paid (KES)",
        "balance": "Total Due (KES)",
    }).round(2)
    table.insert(0, "Invoice No", [invoice_number(target, c["position"]) for c in customers])
    st.dataframe(table, hide_index=True)

    issue_date = now_local().date()

    st.subheader("Download invoices")
    skip_empty = st.checkbox("Skip customers with nothing to bill this month", value=True)
    with st.spinner("Preparing invoices..."):
        zip_bytes = build_invoice_zip(file_bytes, target, skip_empty, issue_date)
    st.download_button(
        f"📥 Download all invoices for {target} (ZIP)", zip_bytes,
        file_name=f"Amani_Invoices_{safe_filename(target)}.zip",
        mime="application/zip", type="primary",
    )

    idx = st.selectbox(
        "Or download one customer's invoice", range(len(customers)),
        format_func=lambda i: customers[i]["name"],
    )
    cust = customers[idx]
    inv = invoice_number(target, cust["position"])
    st.download_button(
        f"📄 Download invoice for {cust['name']}",
        create_branded_pdf(cust, target, inv, issue_date),
        file_name=f"{inv} - {safe_filename(cust['name'])}.pdf",
        mime="application/pdf",
    )


if __name__ == "__main__":
    main()
