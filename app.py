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
PAYMENT_METHOD = "M-PESA Pochi la Biashara"
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


def looks_like_total(value, other_values):
    """True when a cell holds the sum of the other days rather than one delivery."""
    filled = [v for v in other_values if v > 0]
    return value > 0 and len(filled) >= 2 and abs(value - sum(other_values)) < 0.01


def get_month_data(ws, last_day=31):
    """
    Read one monthly sheet. Only rows for days that exist in the month are billed
    (for example rows 4 to 33 in a 30-day month). Anything typed below that, such as
    a total in row 34, is set aside and reported instead of being added to the bill.
    """
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

        cells = {d: ws.cell(row=FIRST_DAY_ROW + d - 1, column=col) for d in range(1, 32)}
        values = {d: clean_num(cells[d].value) for d in range(1, 32)}
        month_values = [values[d] for d in range(1, last_day + 1)]
        ignored = []

        # Rows after the last day of the month are never billed
        for d in range(last_day + 1, 32):
            if values[d] > 0:
                ignored.append({
                    "day": d, "row": FIRST_DAY_ROW + d - 1, "qty": values[d],
                    "reason": "total" if looks_like_total(values[d], month_values) else "extra_day",
                })

        # A total typed into the last real day of the month (e.g. row 34 in August)
        last_value = values[last_day]
        if looks_like_total(last_value, month_values[:-1]):
            ignored.append({
                "day": last_day, "row": FIRST_DAY_ROW + last_day - 1, "qty": last_value,
                "reason": "total_in_month",
            })
            values[last_day] = 0.0

        billed_qty = 0.0
        spoilt_qty = 0.0
        daily = {}
        spoilt_days = []
        for d in range(1, last_day + 1):
            qty = values[d]
            daily[d] = qty
            if qty > 0 and is_spoilt(cells[d]):
                spoilt_qty += qty
                spoilt_days.append((d, qty))
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
            "ignored_entries": ignored,
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
        data = get_month_data(wb[sheet_name], days_in_sheet_month(sheet_name))
        if data:
            results[sheet_name] = data
    return results


def data_quality_issues(sheet_name, customers):
    """Things worth checking before invoices go out."""
    issues = []
    for c in customers:
        if c["row37_unreadable"]:
            issues.append(
                f"{c['name']}: row {PREPAID_ROW} says '{c['row37_raw']}', which is not a number, "
                f"so it was treated as 0."
            )
        if c["billed_qty"] > 0 and c["rate"] == 0:
            issues.append(f"{c['name']}: milk recorded but no rate in row {RATE_ROW}.")
    return issues


def ignored_entry_notes(sheet_name, customers):
    """Plain-English notes on cells that were left out of the bill."""
    last_day = days_in_sheet_month(sheet_name)
    notes = []
    for c in customers:
        for e in c["ignored_entries"]:
            qty = fmt_qty(e["qty"])
            if e["reason"] == "total_in_month":
                notes.append(
                    f"{c['name']}: row {e['row']} (day {e['day']}) has {qty}L, which equals the "
                    f"total of the other days, so it looks like a sum. Not billed."
                )
            elif e["reason"] == "total":
                notes.append(
                    f"{c['name']}: row {e['row']} has {qty}L, which matches the month's total. "
                    f"{sheet_name} only has {last_day} days, so this row was not billed."
                )
            else:
                notes.append(
                    f"{c['name']}: row {e['row']} has {qty}L, but {sheet_name} only has "
                    f"{last_day} days, so it was not billed. If this was a real delivery, "
                    f"move it to the correct day."
                )
    return notes


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
# Invoice palette
NAVY = COLOR_PRIMARY
GOLD = COLOR_SECONDARY
INK = (31, 41, 51)
MUTED = (110, 120, 135)
RULE = (222, 226, 233)
TINT = (236, 241, 248)
WEEKEND = (246, 247, 250)
OUTSIDE_MONTH = (233, 235, 239)
RED_TINT = (253, 236, 236)
CREAM = (252, 248, 234)
SOFT_WHITE = (200, 210, 228)
PAGE_MARGIN = 14


class AmaniInvoice(FPDF):
    def __init__(self, invoice_no=""):
        super().__init__(orientation="P", unit="mm", format="A4")
        self.invoice_no = invoice_no
        self.set_margins(PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN)
        self.set_auto_page_break(auto=True, margin=20)

    # ----- page furniture -----
    def header(self):
        self.set_fill_color(*NAVY)
        self.rect(0, 0, self.w, 36, style="F")
        self.set_fill_color(*GOLD)
        self.rect(0, 36, self.w, 1.2, style="F")

        self.set_xy(PAGE_MARGIN, 10)
        self.set_font("Helvetica", "B", 22)
        self.set_text_color(255, 255, 255)
        self.cell(110, 10, BUSINESS_NAME)
        self.set_xy(PAGE_MARGIN, 21)
        self.set_font("Helvetica", "", 8.5)
        self.set_text_color(*GOLD)
        self.set_char_spacing(1.2)
        self.cell(110, 5, TAGLINE.upper())
        self.set_char_spacing(0)

        self.set_xy(self.w - PAGE_MARGIN - 80, 9)
        self.set_font("Helvetica", "B", 24)
        self.set_text_color(255, 255, 255)
        self.cell(80, 11, "INVOICE", align="R")
        self.set_xy(self.w - PAGE_MARGIN - 80, 21)
        self.set_font("Helvetica", "", 10)
        self.set_text_color(*SOFT_WHITE)
        self.cell(80, 5, self.invoice_no, align="R")
        self.set_y(46)

    def footer(self):
        y = self.h - 14
        self.set_draw_color(*RULE)
        self.set_line_width(0.3)
        self.line(PAGE_MARGIN, y, self.w - PAGE_MARGIN, y)
        self.set_font("Helvetica", "", 8)
        self.set_text_color(*MUTED)
        self.set_xy(PAGE_MARGIN, y + 2)
        self.cell(100, 5, f"Thank you for choosing {BUSINESS_NAME.title()}.")
        self.set_xy(self.w - PAGE_MARGIN - 80, y + 2)
        self.cell(80, 5, f"{self.invoice_no}  |  Page {self.page_no()} of {{nb}}", align="R")

    # ----- small building blocks -----
    def small_label(self, x, y, text, color=MUTED, width=60):
        self.set_xy(x, y)
        self.set_font("Helvetica", "B", 7.5)
        self.set_text_color(*color)
        self.set_char_spacing(0.8)
        self.cell(width, 4, text.upper())
        self.set_char_spacing(0)

    def section_title(self, text):
        y = self.get_y()
        self.small_label(PAGE_MARGIN, y, text, color=NAVY, width=100)
        self.set_draw_color(*GOLD)
        self.set_line_width(0.7)
        self.line(PAGE_MARGIN, y + 5.5, PAGE_MARGIN + 12, y + 5.5)
        self.set_line_width(0.2)
        self.set_y(y + 8)

    def fit_font(self, text, style, max_size, max_width):
        """Use the largest font size (down to 9pt) that fits the width."""
        size = max_size
        while size > 9:
            self.set_font("Helvetica", style, size)
            if self.get_string_width(text) <= max_width:
                break
            size -= 0.5
        self.set_font("Helvetica", style, size)

    # ----- daily grid -----
    def draw_calendar_grid(self, customer, year, month):
        days_in_month = calendar.monthrange(year, month)[1] if (year and month) else 31
        label_w = 16
        day_w = (self.w - 2 * PAGE_MARGIN - label_w) / 31
        spoilt_days = {d for d, _ in customer["spoilt_details"]}
        rows = [("Date", 5.5), ("Day", 4.5), ("Litres", 7)]
        x0, y0 = PAGE_MARGIN, self.get_y()

        self.set_line_width(0.2)
        self.set_draw_color(*RULE)
        y = y0
        for r, (label, h) in enumerate(rows):
            self.set_xy(x0, y)
            self.set_font("Helvetica", "B", 6.5)
            self.set_text_color(*NAVY)
            self.set_fill_color(*TINT)
            self.cell(label_w, h, label.upper(), border=1, align="C", fill=True)

            for d in range(1, 32):
                in_month = d <= days_in_month
                known = bool(year and month and in_month)
                weekend = known and date(year, month, d).weekday() >= 5

                if not in_month:
                    fill = OUTSIDE_MONTH
                elif r == 2 and d in spoilt_days:
                    fill = RED_TINT
                elif r == 0:
                    fill = TINT
                elif weekend:
                    fill = WEEKEND
                else:
                    fill = None

                text, style, size, color = "", "", 6, MUTED
                if in_month and r == 0:
                    text, style, size, color = str(d), "B", 6.5, NAVY
                elif known and r == 1:
                    text, size = date(year, month, d).strftime("%a")[:2], 5.5
                elif in_month and r == 2:
                    qty = customer["daily_liters"].get(d, 0)
                    if qty > 0:
                        text, style, size = fmt_qty(qty), "B", 7
                        color = COLOR_ALERT if d in spoilt_days else INK
                    else:
                        text, size, color = "-", 7, (190, 195, 205)

                if fill:
                    self.set_fill_color(*fill)
                self.set_xy(x0 + label_w + (d - 1) * day_w, y)
                self.set_font("Helvetica", style, size)
                self.set_text_color(*color)
                self.cell(day_w, h, text, border=1, align="C", fill=bool(fill))
            y += h

        # Legend
        ly = y + 2
        self.set_font("Helvetica", "", 7)
        self.set_text_color(*MUTED)
        lx = x0
        legend = [(WEEKEND, "Weekend")]
        if spoilt_days:
            legend.append((RED_TINT, "Spoilt, not charged"))
        for color, label in legend:
            self.set_fill_color(*color)
            self.set_draw_color(*RULE)
            self.rect(lx, ly + 0.7, 3, 3, style="DF")
            self.set_xy(lx + 4, ly)
            w = self.get_string_width(label) + 1
            self.cell(w, 4.5, label)
            lx += 4 + w + 6
        self.set_y(ly + 10)


def create_branded_pdf(customer, sheet_name, invoice_no, issue_date):
    year, month = parse_month(sheet_name)
    if month and not year:
        year = issue_date.year
    due_date = issue_date + timedelta(days=PAYMENT_TERMS_DAYS)
    balance = customer["balance"]
    if balance >= 0.005:
        status, card_note, total_label = "Amount Due", f"Due by {due_date:%d %b %Y}", "TOTAL DUE"
    elif balance <= -0.005:
        status, card_note, total_label = "Credit Balance", "No payment needed", "CREDIT BALANCE"
    else:
        status, card_note, total_label = "Paid in Full", "Thank you", "BALANCE"

    pdf = AmaniInvoice(invoice_no)
    pdf.add_page()
    right_edge = pdf.w - PAGE_MARGIN
    top = pdf.get_y()

    # --- Billed to ---
    pdf.small_label(PAGE_MARGIN, top, "Billed to")
    pdf.set_xy(PAGE_MARGIN, top + 5)
    pdf.set_font("Helvetica", "B", 14)
    pdf.set_text_color(*NAVY)
    pdf.multi_cell(66, 6.5, make_safe(customer["name"]), align="L")

    # --- Dates ---
    details = [
        ("Billing period", make_safe(sheet_name)),
        ("Invoice date", issue_date.strftime("%d %b %Y")),
        ("Due date", due_date.strftime("%d %b %Y")),
    ]
    for i, (label, value) in enumerate(details):
        y = top + i * 10
        pdf.small_label(86, y, label, width=52)
        pdf.set_xy(86, y + 4)
        pdf.set_font("Helvetica", "B", 10)
        pdf.set_text_color(*INK)
        pdf.cell(52, 5, value)

    # --- Amount card ---
    card_x, card_w = right_edge - 52, 52
    pdf.set_fill_color(*NAVY)
    pdf.rect(card_x, top - 1, card_w, 30, style="F")
    pdf.set_fill_color(*GOLD)
    pdf.rect(card_x, top - 1, card_w, 1.2, style="F")
    pdf.small_label(card_x + 4, top + 3, status, color=GOLD, width=card_w - 8)
    amount_text = f"KES {fmt(abs(balance))}"
    pdf.fit_font(amount_text, "B", 16, card_w - 8)
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(card_x + 4, top + 9)
    pdf.cell(card_w - 8, 9, amount_text)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*SOFT_WHITE)
    pdf.set_xy(card_x + 4, top + 20)
    pdf.cell(card_w - 8, 5, card_note)

    # --- Daily deliveries ---
    pdf.set_y(top + 38)
    pdf.section_title("Daily deliveries")
    pdf.draw_calendar_grid(customer, year, month)

    # --- Line items ---
    widths = [86, 30, 30, right_edge - PAGE_MARGIN - 146]
    pdf.set_fill_color(*TINT)
    pdf.set_text_color(*NAVY)
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.set_char_spacing(0.6)
    for w, text, align in zip(widths, ["  DESCRIPTION", "QTY (L)", "RATE (KES)", "AMOUNT (KES)  "], "LRRR"):
        pdf.cell(w, 8, text, align=align, fill=True)
    pdf.set_char_spacing(0)
    pdf.ln(8)

    row_y = pdf.get_y()
    pdf.set_xy(PAGE_MARGIN + 2, row_y + 2)
    pdf.set_font("Helvetica", "B", 10)
    pdf.set_text_color(*INK)
    pdf.cell(80, 5, "Fresh milk supplied")
    pdf.set_xy(PAGE_MARGIN + 2, row_y + 7)
    pdf.set_font("Helvetica", "", 8)
    pdf.set_text_color(*MUTED)
    pdf.cell(80, 4, f"Deliveries for {make_safe(sheet_name)}")
    pdf.set_font("Helvetica", "", 10)
    pdf.set_text_color(*INK)
    x = PAGE_MARGIN + widths[0]
    for w, text in zip(widths[1:], [f"{customer['billed_qty']:.1f}", fmt(customer["rate"]),
                                     f"{fmt(customer['total_bill'])}  "]):
        pdf.set_xy(x, row_y)
        pdf.cell(w, 13, text, align="R")
        x += w
    pdf.set_draw_color(*RULE)
    pdf.set_line_width(0.3)
    pdf.line(PAGE_MARGIN, row_y + 13, right_edge, row_y + 13)

    # --- Totals (right) ---
    block_y = row_y + 18
    tx, tw = right_edge - 82, 82
    total_rows = [("Subtotal (this month)", fmt(customer["total_bill"]))]
    if customer["previous_balance"] > 0:
        total_rows.append(("Previous unpaid balance", fmt(customer["previous_balance"])))
    else:
        total_rows.append(("Pre-paid", f"- {fmt(customer['prepaid'])}"))
    y = block_y
    for label, value in total_rows:
        pdf.set_xy(tx, y)
        pdf.set_font("Helvetica", "", 9)
        pdf.set_text_color(*MUTED)
        pdf.cell(46, 7, label)
        pdf.set_font("Helvetica", "B", 9.5)
        pdf.set_text_color(*INK)
        pdf.cell(tw - 46, 7, value, align="R")
        y += 7
    y += 2
    pdf.set_fill_color(*NAVY)
    pdf.rect(tx, y, tw, 12, style="F")
    pdf.set_xy(tx + 3, y)
    pdf.set_font("Helvetica", "B", 9)
    pdf.set_text_color(*GOLD)
    pdf.set_char_spacing(0.8)
    pdf.cell(36, 12, total_label)
    pdf.set_char_spacing(0)
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(tw - 42, 12, f"KES {fmt(abs(balance))}", align="R")
    totals_end = y + 12

    # --- Spoilt milk note (left) ---
    note_end = block_y
    if customer["spoilt_qty"] > 0:
        nx, nw = PAGE_MARGIN, tx - PAGE_MARGIN - 8
        days_text = ", ".join(f"Day {d} ({fmt_qty(q)}L)" for d, q in customer["spoilt_details"])
        body = (
            f"These deliveries were recorded as spoilt and were not charged: {days_text}. "
            f"Total: {fmt_qty(customer['spoilt_qty'])}L."
        )
        pdf.set_font("Helvetica", "", 8)
        lines = pdf.multi_cell(nw - 9, 4, body, align="L", dry_run=True, output="LINES")
        box_h = 11 + 4 * len(lines)
        pdf.set_fill_color(*RED_TINT)
        pdf.rect(nx, block_y, nw, box_h, style="F")
        pdf.set_fill_color(*COLOR_ALERT)
        pdf.rect(nx, block_y, 1.2, box_h, style="F")
        pdf.small_label(nx + 5, block_y + 3, "Spoilt milk, not charged", color=COLOR_ALERT, width=nw - 9)
        pdf.set_xy(nx + 5, block_y + 8)
        pdf.set_font("Helvetica", "", 8)
        pdf.set_text_color(*INK)
        pdf.multi_cell(nw - 9, 4, body, align="L")
        note_end = block_y + box_h

    # --- How to pay ---
    py = max(totals_end, note_end) + 10
    if py + 32 > pdf.h - 20:
        pdf.add_page()
        py = pdf.get_y()
    pw = right_edge - PAGE_MARGIN
    pdf.set_fill_color(*CREAM)
    pdf.rect(PAGE_MARGIN, py, pw, 32, style="F")
    pdf.set_fill_color(*GOLD)
    pdf.rect(PAGE_MARGIN, py, 1.5, 32, style="F")

    pdf.small_label(PAGE_MARGIN + 7, py + 5, "How to pay", color=NAVY)
    pdf.set_xy(PAGE_MARGIN + 7, py + 11)
    pdf.set_font("Helvetica", "B", 10.5)
    pdf.set_text_color(*INK)
    pdf.cell(80, 5, PAYMENT_METHOD)
    pdf.set_xy(PAGE_MARGIN + 7, py + 18)
    pdf.set_font("Helvetica", "B", 20)
    pdf.set_text_color(*NAVY)
    pdf.cell(80, 9, PAYMENT_NUMBER)

    divider_x = PAGE_MARGIN + 92
    pdf.set_draw_color(*GOLD)
    pdf.set_line_width(0.3)
    pdf.line(divider_x, py + 6, divider_x, py + 26)
    steps = [
        "1.  Open M-PESA and select Lipa na M-PESA",
        "2.  Choose Pochi la Biashara",
        f"3.  Enter {PAYMENT_NUMBER}, the amount and your PIN",
    ]
    pdf.set_font("Helvetica", "", 8.5)
    pdf.set_text_color(*INK)
    for i, step in enumerate(steps):
        pdf.set_xy(divider_x + 6, py + 7 + i * 6.5)
        pdf.cell(right_edge - divider_x - 8, 5, step)

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


HERO_CSS = """
<style>
.amani-hero {
    background: linear-gradient(135deg, #102B55 0%, #1E4380 100%);
    border-bottom: 4px solid #D4AF37;
    border-radius: 14px;
    padding: 26px 30px;
    margin-bottom: 18px;
}
.amani-hero .brand {
    color: #D4AF37; font-size: 0.78rem; font-weight: 700;
    letter-spacing: 0.18em; text-transform: uppercase; margin: 0;
}
.amani-hero .title { color: #FFFFFF; font-size: 2rem; font-weight: 700; margin: 4px 0 6px 0; }
.amani-hero .sub { color: #D6DEEB; font-size: 1rem; margin: 0; }
.amani-step { color: #D4AF37; font-size: 1.6rem; font-weight: 700; line-height: 1; }
</style>
"""

CHART_BLUE = "#3D6BB3"   # readable on light and dark backgrounds
CHART_RED = "#E05252"
CHART_GOLD = "#D4AF37"


def render_hero(subtitle):
    st.markdown(HERO_CSS, unsafe_allow_html=True)
    st.markdown(
        f"""<div class="amani-hero">
            <p class="brand">Amani Dairies</p>
            <div class="title">🥛 Billing &amp; Performance Tracker</div>
            <p class="sub">{subtitle}</p>
        </div>""",
        unsafe_allow_html=True,
    )


def render_data_source(sheet_url):
    """The two ways to load data, shown as cards."""
    left, right = st.columns(2, gap="large")
    with left:
        with st.container(border=True):
            st.markdown("#### ☁️ Sync from Google Sheet")
            st.caption("Pulls the latest version of the billing sheet. Use this for normal monthly billing.")
            if st.button("Sync now", type="primary", width="stretch"):
                try:
                    with st.spinner("Fetching data from Google Sheets..."):
                        st.session_state["file_bytes"] = fetch_google_sheet(sheet_url)
                        st.session_state["source"] = f"Google Sheet, synced at {now_local():%H:%M}"
                    st.session_state["flash"] = "Synced with Google Sheet."
                    st.rerun()
                except Exception as e:
                    st.error(f"Sync failed: {e}")
    with right:
        with st.container(border=True):
            st.markdown("#### 📂 Upload an Excel file")
            st.caption("Use a saved copy if the sheet is unavailable or you want to bill from a backup.")
            uploaded = st.file_uploader(
                "Excel file", type=["xlsx", "xlsm"], label_visibility="collapsed"
            )
            if uploaded is not None:
                # Only take the upload when it is a new file, so it does not overwrite a later sync
                signature = (uploaded.name, uploaded.size)
                if st.session_state.get("upload_signature") != signature:
                    st.session_state["upload_signature"] = signature
                    st.session_state["file_bytes"] = uploaded.getvalue()
                    st.session_state["source"] = f"Uploaded file: {uploaded.name}"
                    st.session_state["flash"] = f"Loaded {uploaded.name}."
                    st.rerun()


def render_how_it_works():
    st.markdown("##### How it works")
    steps = [
        ("1", "Load the data", "Sync the Google Sheet or upload a copy. Every monthly sheet is read."),
        ("2", "Check the month", "Pick a month, review the totals and anything flagged for checking."),
        ("3", "Download invoices", "Get every invoice in one ZIP, or a single customer's invoice."),
    ]
    for col, (num, title, text) in zip(st.columns(3, gap="medium"), steps):
        with col:
            with st.container(border=True):
                st.markdown(f'<div class="amani-step">{num}</div>', unsafe_allow_html=True)
                st.markdown(f"**{title}**")
                st.caption(text)


def theme_tip():
    st.caption("Light or dark mode: open the ⋮ menu at the top right, choose Settings, then pick a theme.")


def main():
    st.set_page_config(page_title="Amani Dairies", page_icon="🥛", layout="wide")
    if not check_password():
        st.stop()

    sheet_url = get_secret("GOOGLE_SHEET_URL", DEFAULT_SHEET_URL)
    if "flash" in st.session_state:
        st.toast(st.session_state.pop("flash"), icon="✅")

    # --- START PAGE (no data yet) ---
    if "file_bytes" not in st.session_state:
        render_hero("Load this month's deliveries to see revenue, check the figures and download invoices.")
        render_data_source(sheet_url)
        st.write("")
        render_how_it_works()
        theme_tip()
        st.stop()

    # --- DASHBOARD ---
    render_hero(f"Data source: {st.session_state['source']}")
    with st.expander("🔄 Refresh or change data"):
        render_data_source(sheet_url)

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
    m1.metric("Lifetime Revenue", f"KES {total_rev:,.0f}", border=True)
    m2.metric("Revenue Lost (Spoilage)", f"KES {total_loss:,.0f}",
              delta=f"{loss_pct:.1f}% of potential", delta_color="inverse", border=True)
    m3.metric("Litres Supplied", f"{total_litres:,.1f} L", border=True)
    m4.metric("Billing Months", len(all_months), border=True)

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
            color_discrete_map={"Revenue": CHART_BLUE, "Spoilage loss": CHART_RED},
        ))
    with right:
        st.plotly_chart(px.bar(
            top_customers, x="Revenue", y="Customer", orientation="h",
            title="Top 10 Customers by Revenue", labels={"Revenue": "KES", "Customer": ""},
            color_discrete_sequence=[CHART_GOLD],
        ))

    st.divider()

    # --- MONTHLY INVOICES ---
    st.header("🧾 Monthly Invoices")
    target = st.selectbox("Select month", months_desc)
    customers = all_months[target]

    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("Billed this month", f"KES {sum(c['total_bill'] for c in customers):,.0f}", border=True)
    k2.metric("Pre-paid", f"KES {sum(c['prepaid'] for c in customers):,.0f}", border=True)
    k3.metric("Previous balances", f"KES {sum(c['previous_balance'] for c in customers):,.0f}", border=True)
    k4.metric("Total outstanding", f"KES {sum(max(c['balance'], 0) for c in customers):,.0f}", border=True)
    k5.metric("Customers", len(customers), border=True)

    issues = data_quality_issues(target, customers)
    if issues:
        with st.expander(f"⚠️ {len(issues)} item(s) to check before sending invoices", expanded=True):
            for issue in issues:
                st.write(f"- {issue}")

    notes = ignored_entry_notes(target, customers)
    if notes:
        with st.expander(f"🧹 {len(notes)} cell(s) left out of the bill", expanded=True):
            st.caption(
                f"{target} has {days_in_sheet_month(target)} days, so only rows "
                f"{FIRST_DAY_ROW} to {FIRST_DAY_ROW + days_in_sheet_month(target) - 1} are billed. "
                "Totals typed into the day rows are also left out."
            )
            for note in notes:
                st.write(f"- {note}")

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

    st.divider()
    theme_tip()


if __name__ == "__main__":
    main()
