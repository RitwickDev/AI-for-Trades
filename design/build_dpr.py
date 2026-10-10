"""Build the Mistri AI Detailed Project Report workbook.

Every total is an Excel formula. Every numeric rate is an input copied from a
named published source. Missing rates stay the text RATE NEEDED and the amount
formula returns a blank.
"""

from datetime import date
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.page import PageMargins

OUT = "/workspace/Mistri-AI-DPR.xlsx"
DIAGRAM = "/workspace/design/mistri-ai-system-context.png"

# Styles
TEAL = "0F766E"
INK = "1C1917"
WHITE = "FFFFFF"
YELLOW = "FFF3C4"
BLUE_FONT = "1D4ED8"
SECTION = "134E4A"
ZEBRA = "F8FAFC"
AMBER = "FEF3C7"
RED_SOFT = "FEE2E2"
GREEN_SOFT = "DCFCE7"
LINE = "D6D3D1"

thin = Border(
    left=Side(style="thin", color=LINE),
    right=Side(style="thin", color=LINE),
    top=Side(style="thin", color=LINE),
    bottom=Side(style="thin", color=LINE),
)
fill_teal = PatternFill("solid", fgColor=TEAL)
fill_section = PatternFill("solid", fgColor=SECTION)
fill_yellow = PatternFill("solid", fgColor=YELLOW)
fill_white = PatternFill("solid", fgColor=WHITE)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)
fill_amber = PatternFill("solid", fgColor=AMBER)
fill_green = PatternFill("solid", fgColor=GREEN_SOFT)
fill_red = PatternFill("solid", fgColor=RED_SOFT)
fill_head = PatternFill("solid", fgColor="115E59")

font_title = Font(name="Calibri", size=20, bold=True, color=INK)
font_h = Font(name="Calibri", size=14, bold=True, color=TEAL)
font_section = Font(name="Calibri", size=12, bold=True, color=WHITE)
font_header = Font(name="Calibri", size=10, bold=True, color=WHITE)
font_label = Font(name="Calibri", size=10, color=INK)
font_bold = Font(name="Calibri", size=10, bold=True, color=INK)
font_small = Font(name="Calibri", size=9, italic=True, color="57534E")
font_input = Font(name="Calibri", size=10, bold=True, color="92400E")
font_formula = Font(name="Calibri", size=10, color=BLUE_FONT)
font_white_bold = Font(name="Calibri", size=11, bold=True, color=WHITE)
wrap = Alignment(wrap_text=True, vertical="center")
left_align = Alignment(vertical="center", wrap_text=True)
center = Alignment(vertical="center", horizontal="center", wrap_text=True)

INR = '₹#,##0.00'
NUM = '#,##0.00'
NUM4 = '0.0000'
NUM8 = '0.00000000'
INT = '#,##0'
PCT = '0%'
DATE = 'DD MMM YYYY'


def apply_widths(ws, widths):
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def header_row(ws, row, labels, fills=None):
    for i, label in enumerate(labels, 1):
        cell = ws.cell(row, i, label)
        cell.font = font_header
        cell.fill = fill_teal
        cell.alignment = center
        cell.border = thin


def paint(cell, value, font, fill, fmt=None, align=None):
    cell.value = value
    cell.font = font
    cell.fill = fill
    cell.alignment = align or left_align
    cell.border = thin
    if fmt:
        cell.number_format = fmt
    return cell


def section(ws, row, text, cols):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=cols)
    cell = ws.cell(row, 1, text)
    cell.font = font_section
    cell.fill = fill_section
    cell.alignment = Alignment(vertical="center", indent=1)
    for i in range(1, cols + 1):
        ws.cell(row, i).fill = fill_section
        ws.cell(row, i).border = thin
    ws.row_dimensions[row].height = 22


def page_setup(ws, landscape=True, fit_width=1, fit_height=0):
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = fit_width
    ws.page_setup.fitToHeight = fit_height
    ws.page_setup.horizontalCentered = True
    ws.page_margins = PageMargins(left=0.5, right=0.5, top=0.6, bottom=0.6, header=0.25, footer=0.25)
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.oddHeader.left.text = "Mistri AI  ·  Detailed Project Report  ·  Delhi pilot"
    ws.oddFooter.left.text = "Yellow cells are inputs. Blue figures are formulas."
    ws.oddFooter.right.text = "Page &P of &N"
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 110
    ws.page_setup.scale = 100
    ws.print_options.horizontalCentered = True
    ws.sheet_format.defaultRowHeight = 18


def amount_formula(qty_cell, rate_cell):
    return f'=IF(OR({rate_cell}="",{rate_cell}="RATE NEEDED"),"",{qty_cell}*{rate_cell})'


def inr_amount_formula(qty_cell, rate_cell, currency_cell, fx_cell):
    return (
        f'=IF(OR({rate_cell}="",{rate_cell}="RATE NEEDED"),"",'
        f'IF({currency_cell}="USD",{qty_cell}*{rate_cell}*{fx_cell},{qty_cell}*{rate_cell}))'
    )


def main():
    wb = Workbook()
    wb.calculation = CalcProperties(calcMode="auto", fullCalcOnLoad=True, forceFullCalc=True)

    # ------------------------------------------------------------------ inputs
    ws_in = wb.active
    ws_in.title = "Inputs and Scope"
    page_setup(ws_in, landscape=True, fit_width=1, fit_height=0)
    apply_widths(ws_in, {"A": 38, "B": 22, "C": 18, "D": 88, "E": 28})
    ws_in.row_dimensions[1].height = 28
    ws_in["A1"] = "Mistri AI — inputs and scope"
    ws_in["A1"].font = font_title
    ws_in.merge_cells("A2:D2")
    ws_in["A2"] = (
        "Yellow cells with brown type are the only inputs. Change a yellow cell and the sizing, quantities, and totals recalculate. "
        "Do not type over a blue figure."
    )
    ws_in["A2"].font = font_small
    ws_in["A2"].alignment = wrap
    ws_in.row_dimensions[2].height = 32

    # key, label, value, unit, note, number format
    rows = [
        ("section", "CONFIRMED SCOPE", None, None, None, None),
        ("project", "Project name", "Mistri AI", "text", "Confirmed at intake", None),
        ("place", "Place", "Delhi, India", "text", "Confirmed at intake", None),
        ("workers", "Construction workers", 100, "accounts", "Pilot cap, confirmed", INT),
        ("contractors", "Contractors", 20, "accounts", "Pilot cap, confirmed", INT),
        ("vendors", "Vendors", 20, "accounts", "Pilot cap, confirmed", INT),
        ("months", "Pilot length", 12, "months", "Design, build, and one year of running, confirmed", INT),
        ("languages", "Languages", 2, "languages", "Hindi and English, confirmed", INT),
        ("channels", "Channels", 3, "channels", "Android, WhatsApp, and website, confirmed", INT),
        ("section", "VOLUME ASSUMPTIONS", None, None, None, None),
        ("jobs_per_contractor", "Jobs per contractor per month", 4, "jobs", "Assumption 1. A job is counted once.", INT),
        ("photos_per_job", "Description photos per job", 2, "photos", "Assumption 2", INT),
        ("audio_clips", "Audio clips per job", 1, "clips", "Assumption 2", INT),
        ("audio_minutes", "Minutes per audio clip", 1, "minutes", "Assumption 2. Used only to describe the clip.", NUM),
        ("photo_mb", "Average photo size", 2, "MB", "Assumption 9", NUM),
        ("audio_mb", "Average audio size", 1, "MB", "Assumption 9", NUM),
        ("site_updates", "Site updates per job", 12, "updates", "Assumption 3", INT),
        ("photos_per_update", "Photos per site update", 1, "photos", "Assumption 3", INT),
        ("bookings_per_job", "Bookings per job", 2, "bookings", "Assumption 4", INT),
        ("catalogue_edits", "Catalogue edits per vendor per month", 2, "edits", "Assumption 5", INT),
        ("logins_per_user", "Sign-ins per account per month", 8, "sign-ins", "Assumption 6. One SMS each.", INT),
        ("wa_job_share", "Share of jobs handled on WhatsApp", 0.5, "share", "Assumption 7", PCT),
        ("wa_msgs_per_job", "Business WhatsApp messages per WhatsApp job", 16, "messages", "Assumption 7. Estimate, match, booking, 12 site replies, invoice.", INT),
        ("vendor_wa_share", "Share of catalogue edits confirmed on WhatsApp", 0.5, "share", "Assumption 8", PCT),
        ("views_per_file", "Views of each media file during the year", 3, "views", "Assumption 11", INT),
        ("section", "PLATFORM ASSUMPTIONS", None, None, None, None),
        ("concurrency", "Peak concurrent share of accounts", 0.2, "share", "Assumption 12. Does not reduce the server count.", PCT),
        ("row_kb", "Database kilobytes per row", 5, "KB", "Assumption 13. Metadata only. Photos sit in object storage.", NUM),
        ("rds_min_gb", "Minimum database volume", 20, "GB", "Assumption 13. AWS RDS general-purpose minimum is 20 GiB.", INT),
        ("app_servers", "App servers running all month", 2, "servers", "Assumption 12. 2 vCPU and 4 GB each, so one can fail.", INT),
        ("hours_per_month", "Hours in a billed month", 730, "hours", "Assumption 21. AWS monthly price ÷ hourly price.", INT),
        ("ebs_gb", "Disk per app server", 30, "GB", "Assumption 14", INT),
        ("wa_free", "Free WhatsApp service messages per month", 1000, "messages", "Assumption 15. Set to 0 if messages are billed as utility templates.", INT),
        ("egress_free", "Free data transfer out per month", 100, "GB", "AWS free allowance, named on the BOM source.", INT),
        ("support_days", "Pilot support days per month", 2, "person-days", "Assumption 17", INT),
        ("section", "BUILD EFFORT (PLANNING DAYS)", None, None, None, None),
        ("d_MOD-AUTH", "Days — MOD-AUTH phone OTP sign-in", 8, "person-days", "Assumption 16", INT),
        ("d_MOD-PROF", "Days — MOD-PROF role and profile", 5, "person-days", "Assumption 16", INT),
        ("d_MOD-HOME", "Days — MOD-HOME home and my jobs", 5, "person-days", "Assumption 16", INT),
        ("d_MOD-JOB", "Days — MOD-JOB intake of text, audio, and photo", 12, "person-days", "Assumption 16", INT),
        ("d_MOD-EST", "Days — MOD-EST estimate engine", 15, "person-days", "Assumption 16", INT),
        ("d_JOB-NEW", "Days — JOB-NEW quantity rules", 6, "person-days", "Assumption 16", INT),
        ("d_JOB-EXT", "Days — JOB-EXT quantity rules", 6, "person-days", "Assumption 16", INT),
        ("d_JOB-FIT", "Days — JOB-FIT quantity rules", 6, "person-days", "Assumption 16", INT),
        ("d_MOD-CAT", "Days — MOD-CAT vendor catalogue", 8, "person-days", "Assumption 16", INT),
        ("d_MOD-MATCH", "Days — MOD-MATCH locality match", 8, "person-days", "Assumption 16", INT),
        ("d_MOD-BOOK", "Days — MOD-BOOK find and book", 8, "person-days", "Assumption 16", INT),
        ("d_MOD-SITE", "Days — MOD-SITE progress and attendance", 10, "person-days", "Assumption 16", INT),
        ("d_MOD-INV", "Days — MOD-INV internal invoice", 6, "person-days", "Assumption 16", INT),
        ("d_CH-AND", "Days — CH-AND Android app", 20, "person-days", "Assumption 16", INT),
        ("d_CH-WEB", "Days — CH-WEB website", 15, "person-days", "Assumption 16", INT),
        ("d_CH-WA", "Days — CH-WA WhatsApp flows", 12, "person-days", "Assumption 16", INT),
        ("d_PLT-LANG", "Days — PLT-LANG Hindi and English", 8, "person-days", "Assumption 16", INT),
        ("d_PLT-CLOUD", "Days — PLT-CLOUD environment setup", 6, "person-days", "Assumption 16", INT),
        ("d_DAT-ACC", "Days — DAT-ACC schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-JOB", "Days — DAT-JOB schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-EST", "Days — DAT-EST schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-CAT", "Days — DAT-CAT schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-BOOK", "Days — DAT-BOOK schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-SITE", "Days — DAT-SITE schema", 2, "person-days", "Assumption 16", INT),
        ("d_DAT-INV", "Days — DAT-INV schema", 2, "person-days", "Assumption 16", INT),
        ("d_EXT-OTP", "Days — EXT-OTP integration", 4, "person-days", "Assumption 16", INT),
        ("d_EXT-WABA", "Days — EXT-WABA integration", 8, "person-days", "Assumption 16", INT),
        ("d_USR-W", "Days — USR-W onboarding", 5, "person-days", "Assumption 16", INT),
        ("d_USR-C", "Days — USR-C onboarding", 4, "person-days", "Assumption 16", INT),
        ("d_USR-V", "Days — USR-V onboarding", 4, "person-days", "Assumption 16", INT),
        ("section", "MONEY INPUTS", None, None, None, None),
        ("day_rate", "Build day rate", "RATE NEEDED", "INR per person-day", "No day rate was given and no single published day rate is used. Type a number here to price every build line.", None),
        ("fx", "Rupees per US dollar", 96.6149, "INR per USD", "FBIL reference rate for 9 October 2026, republished from the RBI website (lawnut.com, retrieved 10 October 2026). The Hindu reported a provisional market close of 96.72 the same day. This cell uses the FBIL reference.", NUM4),
        ("contingency", "Contingency on priced subtotal", 0.10, "share", "Assumption 18. Planning allowance, not a market rate.", PCT),
        ("gst", "GST on priced subtotal plus contingency", 0.18, "share", "Assumption 19. SMS vendor pages retrieved 10 October 2026 still add GST at 18 percent. A tax adviser confirms which lines are taxable.", PCT),
        ("section", "PUBLISHED RATES AND OPEN RATES", None, None, None, None),
        ("ec2_usd", "App server hourly rate", 0.0448, "USD per hour", "AWS Price List, EC2 Linux on-demand, t3.medium, Asia Pacific (Mumbai), publication 2026-10-09. 2 vCPU, 4 GiB.", NUM4),
        ("rds_usd", "Database hourly rate", 0.053, "USD per hour", "AWS Price List, RDS for PostgreSQL, db.t3.small, Single-AZ, Asia Pacific (Mumbai), publication 2026-10-06. 2 vCPU, 2 GiB.", NUM4),
        ("rds_gp3", "Database disk rate", 0.131, "USD per GB-month", "AWS Price List, RDS general-purpose GP3, Single-AZ, Asia Pacific (Mumbai), publication 2026-10-06.", NUM4),
        ("s3_gb", "Object storage rate", 0.025, "USD per GB-month", "AWS Price List, S3 Standard, first 50 TB, Asia Pacific (Mumbai), publication 2026-09-28.", NUM4),
        ("s3_put", "Object PUT rate", 0.000005, "USD per request", "AWS Price List, S3 PUT/COPY/POST/LIST, Asia Pacific (Mumbai), publication 2026-09-28.", NUM8),
        ("s3_get", "Object GET rate", 0.0000004, "USD per request", "AWS Price List, S3 GET and other requests, Asia Pacific (Mumbai), publication 2026-09-28.", NUM8),
        ("egress_usd", "Data transfer out above the free allowance", 0.1093, "USD per GB", "National IT Service, AWS S3 India pricing, Mumbai tier after the free 100 GB, retrieved 10 October 2026. The free 100 GB is stated on https://aws.amazon.com/s3/pricing/ retrieved the same day.", NUM4),
        ("sms_inr", "Transactional SMS rate", 0.17, "INR per SMS", "SMSGatewayCenter published transactional slab, 10,000 units, INR 0.17 per SMS, GST extra, retrieved 10 October 2026.", NUM4),
        ("wa_inr", "WhatsApp business message rate", 0.115, "INR per message", "Meta WhatsApp Business Platform, India utility and service rate effective 1 October 2026, INR 0.1150, taxes extra. Rate card retrieved 10 October 2026.", NUM4),
        ("play_usd", "Android store registration", 25, "USD once", "Google Play Console Help, US$25 one-time registration fee, retrieved 10 October 2026.", NUM),
        ("ssl_inr", "TLS certificate", 0, "INR per year", "Let's Encrypt, certificates at no charge, https://letsencrypt.org/ retrieved 10 October 2026.", NUM),
        ("pg_inr", "PostgreSQL licence", 0, "INR", "PostgreSQL Licence, no fee. https://www.postgresql.org/about/licence/ The hosted instance is charged as the database hourly rate.", NUM),
        ("hardware_inr", "Purchased hardware rate", 0, "INR", "Confirmed scope: public cloud, and workers use their own phones. Nothing is bought.", NUM),
        ("domain_inr", "Domain name for one year", "RATE NEEDED", "INR per year", "No registrar rate was selected. Type a rupee figure from the registrar invoice.", None),
        ("ebs_usd", "App-server disk rate", "RATE NEEDED", "USD per GB-month", "Type the current AWS EBS gp3 rate for Asia Pacific (Mumbai). No EBS price file was retrieved for this issue.", None),
        ("alb_usd", "Load balancer monthly rate", "RATE NEEDED", "USD per month", "Type the current AWS Application Load Balancer rate for Asia Pacific (Mumbai), or delete the quantity if you will not use one.", None),
        ("backup_usd", "Database backup rate above the included allowance", "RATE NEEDED", "USD per GB-month", "Type the current RDS backup rate. AWS includes some backup storage; this line stays open so that allowance is not guessed.", None),
        ("dlt_inr", "DLT principal-entity registration", "RATE NEEDED", "INR", "Operator fees published in 2026 disagree on whether GST sits inside the figure. Type the operator invoice.", None),
        ("model_inr", "Estimate model fee per job", "RATE NEEDED", "INR per job", "Type a rate only if a separate India-region model is bought. The app servers are already in the cloud lines.", None),
        ("section", "TIMELINE INPUTS", None, None, None, None),
        ("start_date", "Project start", date(2026, 10, 13), "date", "Planning start, the Tuesday after this report. Edit the date.", DATE),
        ("build_weeks", "Build and integrate", 18, "weeks", "Assumption 24", INT),
        ("onboard_weeks", "Onboarding", 3, "weeks", "Assumption 24", INT),
        ("pilot_weeks", "Pilot operation", 52, "weeks", "Assumption 24. Twelve months counted as 52 weeks.", INT),
        ("close_weeks", "Close-out", 2, "weeks", "Assumption 24", INT),
    ]

    addr = {}
    r = 4
    for key, label, value, unit, note, fmt in rows:
        if key == "section":
            r += 1
            section(ws_in, r, label, 4)
            r += 1
            continue
        paint(ws_in.cell(r, 1), label, font_label, fill_white)
        value_cell = ws_in.cell(r, 2)
        paint(value_cell, value, font_input, fill_yellow, fmt)
        paint(ws_in.cell(r, 3), unit, font_label, fill_white, align=center)
        paint(ws_in.cell(r, 4), note, font_small, fill_white)
        ws_in.row_dimensions[r].height = 32
        addr[key] = f"='Inputs and Scope'!$B${r}"
        # bare ref for use inside larger formulas
        addr[key + "_ref"] = f"'Inputs and Scope'!$B${r}"
        r += 1

    scope_row = r + 2
    section(ws_in, scope_row, "SCOPE STATEMENT CONFIRMED 10 OCTOBER 2026", 4)
    scope_lines = [
        ("In scope", "Hindi and English tool for a Delhi pilot of up to the worker, contractor, and vendor caps above. Android, WhatsApp, and a website, on a public cloud in India. Phone number and one-time password. One role per account."),
        ("In scope", "Job types are a new building, an extension or extra floor, and an interior fit-out. The description is text, audio, or a photo. The system returns a cost estimate and a quantity list."),
        ("In scope", "Vendors type materials, prices, stock, and the Delhi locality they serve. Match uses that locality. Contractors and workers record site progress, attendance, photos, and a saved location, and they raise an internal invoice marked unpaid or paid."),
        ("Out of scope", "UPI and any money collected in the app. GST invoices and tax filing. Map services and distance. A starting price list or worker list. Aadhaar, bank accounts, and card numbers. A permission system beyond the one-time password, an audit log, and a formal privacy programme."),
        ("Out of scope", "Users above the pilot caps, work outside Delhi, and construction of the building itself. This report is the software only."),
    ]
    for i, (kind, text) in enumerate(scope_lines):
        rr = scope_row + 1 + i
        paint(ws_in.cell(rr, 1), kind, font_bold, fill_white)
        ws_in.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        paint(ws_in.cell(rr, 2), text, font_label, fill_white)
        ws_in.row_dimensions[rr].height = 48
    ws_in.auto_filter.ref = None
    ws_in.freeze_panes = "A5"
    ws_in.auto_filter.ref = "A4:D4"
    ws_in.print_title_rows = "1:2"
    ws_in.page_setup.fitToHeight = False
    ws_in.oddHeader.center.text = ""
    ws_in.sheet_view.zoomScale = 120

    # ------------------------------------------------------------------ sizing
    ws_s = wb.create_sheet("Sizing")
    page_setup(ws_s)
    apply_widths(ws_s, {"A": 16, "B": 62, "C": 78, "D": 22, "E": 20, "F": 16})
    ws_s["A1"] = "Mistri AI — sizing and quantities"
    ws_s["A1"].font = font_title
    ws_s.merge_cells("A2:F2")
    ws_s["A2"] = "Column D is a formula. Column C is the working in words. Yellow inputs live on Inputs and Scope."
    ws_s["A2"].font = font_small
    sizing = {}  # key -> cell ref D row

    def add_size(key, label, formula, unit, design_id, working, height=32):
        nonlocal srow
        paint(ws_s.cell(srow, 1), key, font_bold, fill_white, align=center)
        paint(ws_s.cell(srow, 2), label, font_label, fill_white)
        paint(ws_s.cell(srow, 3), working, font_small, fill_white)
        paint(ws_s.cell(srow, 4), formula, font_formula, fill_green, NUM)
        paint(ws_s.cell(srow, 5), unit, font_label, fill_white, align=center)
        paint(ws_s.cell(srow, 6), design_id, font_bold, fill_white, align=center)
        ws_s.row_dimensions[srow].height = height
        sizing[key] = f"'Sizing'!$D${srow}"
        srow += 1

    srow = 4
    section(ws_s, srow, "USERS", 6)
    srow += 1
    header_row(ws_s, srow, ["Key", "Measure", "Working", "Quantity", "Unit", "Design ID"])
    srow += 1
    I = lambda k: addr[k + "_ref"]
    add_size("S-ACC", "Registered accounts", f"={I('workers')}+{I('contractors')}+{I('vendors')}", "accounts", "USR-W", "Workers + contractors + vendors.")
    add_size("S-PEAK", "Peak concurrent users", f"={sizing['S-ACC']}*{I('concurrency')}", "users", "PLT-CLOUD", "Accounts × the peak share. This records load. Server count stays at the input.")

    section(ws_s, srow, "JOBS AND RECORDS", 6)
    srow += 1
    add_size("S-JOB-M", "Jobs per month", f"={I('contractors')}*{I('jobs_per_contractor')}", "jobs", "MOD-JOB", "Contractors × jobs per contractor. A worker-created job is inside this count so it is not added again.")
    add_size("S-JOB-Y", "Jobs in the pilot", f"={sizing['S-JOB-M']}*{I('months')}", "jobs", "MOD-JOB", "Jobs per month × months.")
    add_size("S-EST", "Estimates", f"={sizing['S-JOB-Y']}", "estimates", "MOD-EST", "One estimate per job.")
    add_size("S-MATCH", "Material matches", f"={sizing['S-JOB-Y']}", "matches", "MOD-MATCH", "One match per job.")
    add_size("S-BOOK", "Bookings", f"={sizing['S-JOB-Y']}*{I('bookings_per_job')}", "bookings", "MOD-BOOK", "Jobs × bookings per job.")
    add_size("S-INV", "Invoices", f"={sizing['S-JOB-Y']}", "invoices", "MOD-INV", "One internal invoice per job.")
    add_size("S-SITE", "Site updates", f"={sizing['S-JOB-Y']}*{I('site_updates')}", "updates", "MOD-SITE", "Jobs × site updates per job.")
    add_size("S-CAT", "Catalogue edits", f"={I('vendors')}*{I('catalogue_edits')}*{I('months')}", "edits", "MOD-CAT", "Vendors × edits per month × months.")
    add_size("S-LOGIN", "Sign-ins", f"={sizing['S-ACC']}*{I('logins_per_user')}*{I('months')}", "sign-ins", "MOD-AUTH", "Accounts × sign-ins per month × months.")
    add_size("S-SMS", "OTP text messages", f"={sizing['S-LOGIN']}", "SMS", "EXT-OTP", "One SMS per sign-in. WhatsApp does not send a second code.")

    section(ws_s, srow, "STORAGE", 6)
    srow += 1
    add_size("S-MB-JOB", "Megabytes of media per job", f"=({I('photos_per_job')}+{I('site_updates')}*{I('photos_per_update')})*{I('photo_mb')}+{I('audio_clips')}*{I('audio_mb')}", "MB", "DAT-JOB", "Description photos and site photos at the average photo size, plus audio clips at the average audio size.")
    add_size("S-MB-DESC", "Megabytes of description media per job", f"={I('photos_per_job')}*{I('photo_mb')}+{I('audio_clips')}*{I('audio_mb')}", "MB", "DAT-JOB", "Description photos and the audio clip only.")
    add_size("S-MB-SITE", "Megabytes of site media per job", f"={I('site_updates')}*{I('photos_per_update')}*{I('photo_mb')}", "MB", "DAT-SITE", "Site photos only.")
    add_size("S-GB-M", "New media per month", f"={sizing['S-JOB-M']}*{sizing['S-MB-JOB']}/1024", "GB", "PLT-CLOUD", "Jobs per month × MB per job ÷ 1024.")
    add_size("S-GB-END", "Media stored at the end of the pilot", f"={sizing['S-GB-M']}*{I('months')}", "GB", "PLT-CLOUD", "New media each month is kept for the whole pilot, so the last month holds months × the monthly addition.")
    add_size("S-GBMO", "Object storage billed", f"={sizing['S-GB-M']}*{I('months')}*({I('months')}+1)/2", "GB-months", "PLT-CLOUD", "Month 1 stores one month of media, month 2 stores two, and so on. The sum is monthly new × months × (months + 1) / 2.")
    add_size("S-GBMO-JOB", "Object storage for job descriptions", f"=IF({sizing['S-MB-JOB']}=0,0,{sizing['S-GBMO']}*{sizing['S-MB-DESC']}/{sizing['S-MB-JOB']})", "GB-months", "DAT-JOB", "Total object GB-months × the description share of each job's media.")
    add_size("S-GBMO-SITE", "Object storage for site photos", f"=IF({sizing['S-MB-JOB']}=0,0,{sizing['S-GBMO']}*{sizing['S-MB-SITE']}/{sizing['S-MB-JOB']})", "GB-months", "DAT-SITE", "Total object GB-months × the site-photo share. The two shares add back to the total.")

    section(ws_s, srow, "DATABASE ROWS AND THE PROVISIONED VOLUME", 6)
    srow += 1
    add_size("R-ACC", "DAT-ACC rows", f"={sizing['S-ACC']}+{sizing['S-LOGIN']}", "rows", "DAT-ACC", "One standing account row plus one row per sign-in.")
    add_size("R-JOB", "DAT-JOB rows", f"={sizing['S-JOB-Y']}", "rows", "DAT-JOB", "One metadata row per job. The file itself is in object storage.")
    add_size("R-EST", "DAT-EST rows", f"={sizing['S-EST']}", "rows", "DAT-EST", "One row per estimate.")
    add_size("R-CAT", "DAT-CAT rows", f"={sizing['S-CAT']}", "rows", "DAT-CAT", "One row per catalogue edit.")
    add_size("R-BOOK", "DAT-BOOK rows", f"={sizing['S-BOOK']}", "rows", "DAT-BOOK", "One row per booking.")
    add_size("R-SITE", "DAT-SITE rows", f"={sizing['S-SITE']}", "rows", "DAT-SITE", "One row per site update.")
    add_size("R-INV", "DAT-INV rows", f"={sizing['S-INV']}", "rows", "DAT-INV", "One row per invoice.")
    add_size(
        "R-ALL",
        "Database rows",
        f"={sizing['R-ACC']}+{sizing['R-JOB']}+{sizing['R-EST']}+{sizing['R-CAT']}+{sizing['R-BOOK']}+{sizing['R-SITE']}+{sizing['R-INV']}",
        "rows",
        "PLT-CLOUD",
        "Sum of the seven stores.",
    )
    add_size("S-DB-GB", "Calculated database size", f"={sizing['R-ALL']}*{I('row_kb')}/1024/1024", "GB", "PLT-CLOUD", "Rows × kilobytes per row ÷ 1024 ÷ 1024.")
    add_size("S-DB-PROV", "Provisioned database volume", f"=MAX({I('rds_min_gb')},{sizing['S-DB-GB']})", "GB", "PLT-CLOUD", "The greater of the calculated size and the 20 GB minimum AWS sells for this volume type.")
    for key, store in [
        ("R-ACC", "DAT-ACC"),
        ("R-JOB", "DAT-JOB"),
        ("R-EST", "DAT-EST"),
        ("R-CAT", "DAT-CAT"),
        ("R-BOOK", "DAT-BOOK"),
        ("R-SITE", "DAT-SITE"),
        ("R-INV", "DAT-INV"),
    ]:
        add_size(
            "GB-" + store,
            f"{store} share of the database bill",
            f"=IF({sizing['R-ALL']}=0,0,{sizing[key]}/{sizing['R-ALL']}*{sizing['S-DB-PROV']}*{I('months')})",
            "GB-months",
            store,
            "This store's share of rows × the provisioned volume × months. The shares add back to the full volume.",
        )

    section(ws_s, srow, "BANDWIDTH AND REQUESTS", 6)
    srow += 1
    add_size("S-FILES-JOB", "Description files", f"={sizing['S-JOB-Y']}*({I('photos_per_job')}+{I('audio_clips')})", "files", "DAT-JOB", "Jobs × (description photos + audio clips).")
    add_size("S-FILES-SITE", "Site photo files", f"={sizing['S-JOB-Y']}*{I('site_updates')}*{I('photos_per_update')}", "files", "DAT-SITE", "Jobs × site updates × photos per update.")
    add_size("S-PUT-JOB", "PUT requests for descriptions", f"={sizing['S-FILES-JOB']}", "requests", "DAT-JOB", "One PUT per description file.")
    add_size("S-PUT-SITE", "PUT requests for site photos", f"={sizing['S-FILES-SITE']}", "requests", "DAT-SITE", "One PUT per site photo.")
    add_size("S-GET-JOB", "GET requests for descriptions", f"={sizing['S-FILES-JOB']}*{I('views_per_file')}", "requests", "DAT-JOB", "Description files × views per file.")
    add_size("S-GET-SITE", "GET requests for site photos", f"={sizing['S-FILES-SITE']}*{I('views_per_file')}", "requests", "DAT-SITE", "Site files × views per file.")
    add_size("S-DOWN-Y", "Media downloaded in the year", f"={sizing['S-GB-END']}*{I('views_per_file')}", "GB", "PLT-CLOUD", "End-state media × views. Uploads are not charged as transfer out.")
    add_size("S-EGRESS-M", "Transfer out per month", f"=IF({I('months')}=0,0,{sizing['S-DOWN-Y']}/{I('months')})", "GB", "PLT-CLOUD", "Annual downloads ÷ months. Views are spread over the year.")
    add_size("S-EGRESS-BILL-M", "Billable transfer out per month", f"=MAX(0,{sizing['S-EGRESS-M']}-{I('egress_free')})", "GB", "PLT-CLOUD", "Monthly transfer minus the free allowance. The result is zero while the pilot stays inside the allowance.")
    add_size("S-EGRESS-BILL-Y", "Billable transfer out in the year", f"={sizing['S-EGRESS-BILL-M']}*{I('months')}", "GB", "PLT-CLOUD", "Billable GB per month × months.")

    section(ws_s, srow, "LICENCES AND MACHINE HOURS", 6)
    srow += 1
    add_size("S-EC2-H", "App-server hours", f"={I('app_servers')}*{I('hours_per_month')}*{I('months')}", "hours", "PLT-CLOUD", "Servers × 730 hours × months.")
    add_size("S-RDS-H", "Database hours", f"={I('hours_per_month')}*{I('months')}", "hours", "PLT-CLOUD", "One database × 730 hours × months.")
    add_size("S-EBS", "App-server disk", f"={I('app_servers')}*{I('ebs_gb')}*{I('months')}", "GB-months", "PLT-CLOUD", "Servers × GB per server × months.")
    add_size("S-ALB", "Load-balancer months", f"={I('months')}", "months", "PLT-CLOUD", "One load balancer for the whole pilot.")
    add_size("S-BACKUP", "Backup storage allowance", f"={sizing['S-DB-PROV']}*{I('months')}", "GB-months", "PLT-CLOUD", "Set equal to the provisioned database, one GB-month per GB per month. The rate stays open.")
    add_size("S-WA-M", "WhatsApp business messages per month", f"={sizing['S-JOB-M']}*{I('wa_job_share')}*{I('wa_msgs_per_job')}+{I('vendors')}*{I('vendor_wa_share')}*{I('catalogue_edits')}", "messages", "EXT-WABA", "WhatsApp jobs × messages per job, plus vendor catalogue confirmations.")
    add_size("S-WA-BILL-M", "Billable WhatsApp messages per month", f"=MAX(0,{sizing['S-WA-M']}-{I('wa_free')})", "messages", "EXT-WABA", "Monthly messages minus the free service allowance.")
    add_size("S-WA-BILL-Y", "Billable WhatsApp messages in the year", f"={sizing['S-WA-BILL-M']}*{I('months')}", "messages", "EXT-WABA", "Billable messages per month × months.")
    add_size("S-SUPPORT", "Pilot support days", f"={I('support_days')}*{I('months')}", "person-days", "PLT-CLOUD", "Support days per month × months.")

    section(ws_s, srow, "BUILD DAYS (COPIED FROM INPUTS)", 6)
    srow += 1
    day_keys = [
        "d_MOD-AUTH", "d_MOD-PROF", "d_MOD-HOME", "d_MOD-JOB", "d_MOD-EST",
        "d_JOB-NEW", "d_JOB-EXT", "d_JOB-FIT", "d_MOD-CAT", "d_MOD-MATCH",
        "d_MOD-BOOK", "d_MOD-SITE", "d_MOD-INV", "d_CH-AND", "d_CH-WEB",
        "d_CH-WA", "d_PLT-LANG", "d_PLT-CLOUD", "d_DAT-ACC", "d_DAT-JOB",
        "d_DAT-EST", "d_DAT-CAT", "d_DAT-BOOK", "d_DAT-SITE", "d_DAT-INV",
        "d_EXT-OTP", "d_EXT-WABA", "d_USR-W", "d_USR-C", "d_USR-V",
    ]
    for key in day_keys:
        add_size(key, key.replace("d_", "Build days "), f"={I(key)}", "person-days", key.replace("d_", ""), "Planning allowance. Edit it on Inputs and Scope.")
    add_size("S-DAYS", "Build days excluding pilot support", "=" + "+".join(sizing[k] for k in day_keys), "person-days", "PLT-CLOUD", "Sum of the planning allowances above. Pilot support is counted on its own line.")

    section(ws_s, srow, "CHECKS", 6)
    srow += 1
    add_size(
        "CHK-DB",
        "Database shares minus the full volume (must be 0)",
        "=" + "+".join(sizing["GB-" + s] for s in ["DAT-ACC", "DAT-JOB", "DAT-EST", "DAT-CAT", "DAT-BOOK", "DAT-SITE", "DAT-INV"]) + f"-{sizing['S-DB-PROV']}*{I('months')}",
        "GB-months",
        "PLT-CLOUD",
        "The seven shares must equal provisioned GB × months.",
    )
    add_size(
        "CHK-S3",
        "Object shares minus object total (must be 0)",
        f"={sizing['S-GBMO-JOB']}+{sizing['S-GBMO-SITE']}-{sizing['S-GBMO']}",
        "GB-months",
        "PLT-CLOUD",
        "Description media plus site media must equal all object storage.",
    )
    ws_s.freeze_panes = "A6"
    ws_s.auto_filter.ref = "A5:F5"
    ws_s.print_title_rows = "1:5"
    ws_s.auto_filter.ref = "A5:F200"
    ws_s.page_setup.fitToHeight = False

    # ------------------------------------------------------------------ design
    ws_d = wb.create_sheet("Design")
    page_setup(ws_d)
    apply_widths(ws_d, {"A": 16, "B": 28, "C": 55, "D": 22, "E": 36, "F": 28, "G": 22})
    ws_d["A1"] = "Mistri AI — design"
    ws_d["A1"].font = font_title
    ws_d.merge_cells("A2:G2")
    ws_d["A2"] = "Trace key = Yes means a BOQ or BOM line must cite this ID. Screens cite the parent module and are not priced a second time."
    ws_d["A2"].font = font_small
    ws_d.row_dimensions[2].height = 20

    components = [
        ("USR-W", "User", "Construction worker", "Yes", "Creates jobs, books, views own site updates and invoices"),
        ("USR-C", "User", "Contractor", "Yes", "Worker actions, plus site updates and marking invoices paid"),
        ("USR-V", "User", "Vendor", "Yes", "Catalogue, and invoices that name the vendor"),
        ("CH-AND", "Channel", "Android app", "Yes", "Hindi and English screens"),
        ("CH-WA", "Channel", "WhatsApp", "Yes", "The same tasks as a conversation"),
        ("CH-WEB", "Channel", "Website", "Yes", "The Android screens in a browser"),
        ("PLT-CLOUD", "Platform", "Public cloud in India", "Yes", "Hosts the application and the data for one year"),
        ("PLT-LANG", "Platform", "Hindi and English", "Yes", "Both languages on every channel"),
        ("EXT-OTP", "External", "SMS one-time-password gateway", "Yes", "Sends the code for sign-in"),
        ("EXT-WABA", "External", "WhatsApp Business API", "Yes", "Carries the WhatsApp channel"),
        ("MOD-AUTH", "Module", "Phone OTP sign-in", "Yes", "Phone number and one-time password"),
        ("MOD-PROF", "Module", "Role and profile", "Yes", "One role, trade, and locality"),
        ("MOD-HOME", "Module", "Home and my jobs", "Yes", "Entry point and the account's jobs"),
        ("MOD-JOB", "Module", "Job intake", "Yes", "Text, audio, photo, and job type"),
        ("MOD-EST", "Module", "Estimate and quantities", "Yes", "Cost estimate and quantity list"),
        ("MOD-CAT", "Module", "Vendor catalogue", "Yes", "Materials, price, stock, locality"),
        ("MOD-MATCH", "Module", "Material match", "Yes", "Match by the locality the vendor typed"),
        ("MOD-BOOK", "Module", "Find and book", "Yes", "Find by trade and record the booking"),
        ("MOD-SITE", "Module", "Site progress", "Yes", "Progress, attendance, photos, saved location"),
        ("MOD-INV", "Module", "Internal invoice", "Yes", "Unpaid or paid, no money collected"),
        ("DAT-ACC", "Store", "Accounts", "Yes", "Phone, role, trade, locality, sign-ins"),
        ("DAT-JOB", "Store", "Jobs and media", "Yes", "Text, audio, photos, job type"),
        ("DAT-EST", "Store", "Estimates", "Yes", "Estimate and quantity list"),
        ("DAT-CAT", "Store", "Catalogue", "Yes", "Price, stock, locality"),
        ("DAT-BOOK", "Store", "Bookings", "Yes", "Who was booked, on which job"),
        ("DAT-SITE", "Store", "Site records", "Yes", "Attendance, photos, saved location"),
        ("DAT-INV", "Store", "Invoices", "Yes", "Amount and paid or unpaid"),
        ("JOB-NEW", "Job type", "New building", "Yes", "Quantity rules inside the estimate"),
        ("JOB-EXT", "Job type", "Extension or extra floor", "Yes", "Quantity rules inside the estimate"),
        ("JOB-FIT", "Job type", "Interior fit-out", "Yes", "Quantity rules inside the estimate"),
    ]
    screens = [
        ("SCR-01", "Enter phone number", "MOD-AUTH", "Android, website", "All roles"),
        ("SCR-02", "Enter one-time password", "MOD-AUTH", "Android, website", "All roles"),
        ("SCR-03", "Role and profile", "MOD-PROF", "Android, website", "All roles"),
        ("SCR-04", "Home", "MOD-HOME", "Android, website", "All roles"),
        ("SCR-05", "My jobs", "MOD-HOME", "Android, website", "Worker, contractor"),
        ("SCR-06", "New job", "MOD-JOB", "Android, website", "Worker, contractor"),
        ("SCR-07", "Estimate and quantities", "MOD-EST", "Android, website", "Worker, contractor"),
        ("SCR-08", "Vendor catalogue", "MOD-CAT", "Android, website", "Vendor"),
        ("SCR-09", "Material match", "MOD-MATCH", "Android, website", "Worker, contractor"),
        ("SCR-10", "Find people by trade", "MOD-BOOK", "Android, website", "Worker, contractor"),
        ("SCR-11", "Booking", "MOD-BOOK", "Android, website", "Worker and contractor named on it"),
        ("SCR-12", "Site update", "MOD-SITE", "Android, website", "Contractor writes, worker views"),
        ("SCR-13", "Invoice", "MOD-INV", "Android, website", "Contractor writes, worker and vendor view their own"),
        ("SCR-W1", "Sign-in conversation", "MOD-AUTH", "WhatsApp", "All roles"),
        ("SCR-W2", "New job message", "MOD-JOB", "WhatsApp", "Worker, contractor"),
        ("SCR-W3", "Estimate reply", "MOD-EST", "WhatsApp", "Worker, contractor"),
        ("SCR-W4", "Catalogue replies", "MOD-CAT", "WhatsApp", "Vendor"),
        ("SCR-W5", "Match and book replies", "MOD-MATCH", "WhatsApp", "Worker, contractor"),
        ("SCR-W6", "Site update messages", "MOD-SITE", "WhatsApp", "Contractor writes, worker views"),
        ("SCR-W7", "Invoice status", "MOD-INV", "WhatsApp", "Each role sees their own"),
    ]
    section(ws_d, 4, "COMPONENTS — TRACE KEYS", 5)
    header_row(ws_d, 5, ["ID", "Kind", "Name", "Trace key", "Role in the pilot"])
    for i, rec in enumerate(components):
        for c, val in enumerate(rec, 1):
            paint(ws_d.cell(6 + i, c), val, font_bold if c in (1, 4) else font_label, fill_white if i % 2 == 0 else fill_zebra, align=center if c in (1, 4) else left_align)
        ws_d.row_dimensions[6 + i].height = 22
    comp_end = 5 + len(components)
    screen_header = comp_end + 2
    section(ws_d, screen_header, "SCREENS — EACH CITES ITS PARENT, WHICH IS THE TRACE KEY", 5)
    header_row(ws_d, screen_header + 1, ["ID", "Screen", "Parent", "Channel", "Who uses it"])
    for i, rec in enumerate(screens):
        rr = screen_header + 2 + i
        for c, val in enumerate(rec, 1):
            paint(ws_d.cell(rr, c), val, font_bold if c in (1, 3) else font_label, fill_white if i % 2 == 0 else fill_zebra, align=center if c in (1, 3) else left_align)
        ws_d.row_dimensions[rr].height = 20
    img_row = screen_header + 2 + len(screens) + 2
    section(ws_d, img_row, "SYSTEM CONTEXT", 5)
    ws_d.cell(img_row + 1, 1, "The diagram uses the same IDs as the table above.").font = font_small
    img = XLImage(DIAGRAM)
    img.width = 1100
    img.height = int(1100 * 1320 / 2400)
    ws_d.add_image(img, f"A{img_row + 2}")
    ws_d.row_dimensions[img_row + 2].height = 20
    ws_d.freeze_panes = "A6"
    ws_d.print_title_rows = "1:2"
    ws_d.page_setup.fitToHeight = False
    trace_ids = {rec[0] for rec in components}

    # ------------------------------------------------------------------ BOQ
    ws_b = wb.create_sheet("BOQ")
    page_setup(ws_b)
    apply_widths(ws_b, {"A": 12, "B": 12, "C": 18, "D": 14, "E": 62, "F": 16, "G": 16, "H": 18, "I": 20, "J": 78})
    ws_b["A1"] = "Mistri AI — bill of quantities"
    ws_b["A1"].font = font_title
    ws_b.merge_cells("A2:J2")
    ws_b["A2"] = "This sheet is the work. The rate of every person-day line reads the single day-rate input. While that input says RATE NEEDED, the amount stays blank."
    ws_b["A2"].font = font_small
    ws_b.row_dimensions[2].height = 28
    header_row(ws_b, 4, ["Line", "Phase", "Category", "Design ID", "Description", "Quantity", "Unit", "Rate", "Amount (INR)", "Source"])

    boq_items = [
        ("Build", "MOD-AUTH", "Build phone OTP sign-in on Android, WhatsApp, and the website", "d_MOD-AUTH", "person-days"),
        ("Build", "MOD-PROF", "Build role and profile, one role per phone number", "d_MOD-PROF", "person-days"),
        ("Build", "MOD-HOME", "Build home and my jobs", "d_MOD-HOME", "person-days"),
        ("Build", "MOD-JOB", "Build job intake for text, audio, and photo", "d_MOD-JOB", "person-days"),
        ("Build", "MOD-EST", "Build the estimate engine that turns a job into quantities and a cost", "d_MOD-EST", "person-days"),
        ("Build", "JOB-NEW", "Write quantity rules for a new building", "d_JOB-NEW", "person-days"),
        ("Build", "JOB-EXT", "Write quantity rules for an extension or extra floor", "d_JOB-EXT", "person-days"),
        ("Build", "JOB-FIT", "Write quantity rules for an interior fit-out", "d_JOB-FIT", "person-days"),
        ("Build", "MOD-CAT", "Build the vendor catalogue", "d_MOD-CAT", "person-days"),
        ("Build", "MOD-MATCH", "Build the locality match", "d_MOD-MATCH", "person-days"),
        ("Build", "MOD-BOOK", "Build find and book by trade", "d_MOD-BOOK", "person-days"),
        ("Build", "MOD-SITE", "Build site progress, attendance, photos, and saved location", "d_MOD-SITE", "person-days"),
        ("Build", "MOD-INV", "Build the internal invoice and the paid or unpaid mark", "d_MOD-INV", "person-days"),
        ("Build", "CH-AND", "Build the Android app shell", "d_CH-AND", "person-days"),
        ("Build", "CH-WEB", "Build the website", "d_CH-WEB", "person-days"),
        ("Build", "CH-WA", "Build the WhatsApp conversation", "d_CH-WA", "person-days"),
        ("Build", "PLT-LANG", "Build Hindi and English across every channel", "d_PLT-LANG", "person-days"),
        ("Build", "PLT-CLOUD", "Set up the India-region cloud environment", "d_PLT-CLOUD", "person-days"),
        ("Build", "DAT-ACC", "Create the account store", "d_DAT-ACC", "person-days"),
        ("Build", "DAT-JOB", "Create the job store", "d_DAT-JOB", "person-days"),
        ("Build", "DAT-EST", "Create the estimate store", "d_DAT-EST", "person-days"),
        ("Build", "DAT-CAT", "Create the catalogue store", "d_DAT-CAT", "person-days"),
        ("Build", "DAT-BOOK", "Create the booking store", "d_DAT-BOOK", "person-days"),
        ("Build", "DAT-SITE", "Create the site store", "d_DAT-SITE", "person-days"),
        ("Build", "DAT-INV", "Create the invoice store", "d_DAT-INV", "person-days"),
        ("Build", "EXT-OTP", "Connect the SMS one-time-password gateway", "d_EXT-OTP", "person-days"),
        ("Build", "EXT-WABA", "Connect the WhatsApp Business API", "d_EXT-WABA", "person-days"),
        ("Build", "USR-W", "Prepare worker onboarding for the pilot", "d_USR-W", "person-days"),
        ("Build", "USR-C", "Prepare contractor onboarding for the pilot", "d_USR-C", "person-days"),
        ("Build", "USR-V", "Prepare vendor onboarding for the pilot", "d_USR-V", "person-days"),
        ("Run", "PLT-CLOUD", "Support the pilot after go-live", "S-SUPPORT", "person-days"),
    ]
    boq_first = 5
    for i, (phase, did, desc, qty_key, unit) in enumerate(boq_items):
        rr = boq_first + i
        paint(ws_b.cell(rr, 1), f"BQ-{i+1:02d}", font_bold, fill_white, align=center)
        paint(ws_b.cell(rr, 2), phase, font_label, fill_white, align=center)
        paint(ws_b.cell(rr, 3), "Build labour", font_label, fill_white, align=center)
        paint(ws_b.cell(rr, 4), did, font_bold, fill_white, align=center)
        paint(ws_b.cell(rr, 5), desc, font_label, fill_white)
        paint(ws_b.cell(rr, 6), f"={sizing[qty_key]}", font_formula, fill_green, NUM)
        paint(ws_b.cell(rr, 7), unit, font_label, fill_white, align=center)
        paint(ws_b.cell(rr, 8), f"={I('day_rate')}", font_formula, fill_green)
        paint(ws_b.cell(rr, 9), amount_formula(f"F{rr}", f"H{rr}"), font_formula, fill_green, INR)
        paint(ws_b.cell(rr, 10), "Day rate is the yellow input on Inputs and Scope. No day rate was supplied and no published day rate is used, so the cell stays RATE NEEDED until you type a number.", font_small, fill_white)
        ws_b.row_dimensions[rr].height = 36
    boq_last = boq_first + len(boq_items) - 1
    tot = boq_last + 1
    paint(ws_b.cell(tot, 5), "BOQ total (blank amounts are ignored)", font_bold, fill_amber)
    paint(ws_b.cell(tot, 9), f"=SUM(I{boq_first}:I{boq_last})", font_formula, fill_amber, INR)
    for col in range(1, 11):
        if ws_b.cell(tot, col).value is None:
            ws_b.cell(tot, col).fill = fill_amber
            ws_b.cell(tot, col).border = thin
    ws_b.freeze_panes = "A5"
    ws_b.auto_filter.ref = f"A4:J{boq_last}"
    ws_b.print_title_rows = "1:4"
    ws_b.auto_filter.ref = f"A4:J{boq_last}"
    boq_amount_range = f"'BOQ'!$I${boq_first}:$I${boq_last}"
    boq_rate_range = f"'BOQ'!$H${boq_first}:$H${boq_last}"
    boq_id_range = f"'BOQ'!$D${boq_first}:$D${boq_last}"
    boq_phase_range = f"'BOQ'!$B${boq_first}:$B${boq_last}"
    boq_cat_range = f"'BOQ'!$C${boq_first}:$C${boq_last}"

    # ------------------------------------------------------------------ BOM
    ws_m = wb.create_sheet("BOM")
    page_setup(ws_m)
    bom_headers = ["Line", "Phase", "Section", "Category", "Design ID", "Item", "Specification", "Make or equivalent", "Quantity", "Unit", "Rate", "Currency", "Amount (INR)", "Source"]
    apply_widths(ws_m, {
        "A": 12, "B": 12, "C": 14, "D": 16, "E": 14, "F": 38, "G": 55, "H": 42,
        "I": 16, "J": 16, "K": 16, "L": 12, "M": 20, "N": 78,
    })
    ws_m["A1"] = "Mistri AI — bill of materials"
    ws_m["A1"].font = font_title
    ws_m.merge_cells("A2:N2")
    ws_m["A2"] = "Hardware, software, and services to buy or license. Quantity comes from Sizing. Rate comes from Inputs and Scope. Amount converts USD at the FBIL rate and stays blank when the rate is RATE NEEDED."
    ws_m["A2"].font = font_small
    ws_m.row_dimensions[2].height = 32
    header_row(ws_m, 4, bom_headers)

    # qty_key, rate_key, currency, phase, section, category, design, item, spec, make, unit, source
    bom_items = [
        ("literal0", "hardware_inr", "INR", "Build", "Hardware", "Hardware", "PLT-CLOUD",
         "Servers and phones", "No device is purchased. The pilot runs on the public cloud and on the users' own Android phones.",
         "None", "number", "Confirmed scope, 10 October 2026."),
        ("literal1", "play_usd", "USD", "Build", "Software", "Certificates and accounts", "CH-AND",
         "Android store account", "One Play Console developer account, once.",
         "Google Play or equivalent store", "account",
         "Google Play Console Help, US$25 one-time fee, retrieved 10 October 2026. https://support.google.com/googleplay/android-developer/answer/6112435"),
        ("literal1", "domain_inr", "INR", "Build", "Software", "Certificates and accounts", "CH-WEB",
         "Domain name", "One name for one year, .in or .com.",
         "Any ICANN registrar", "year",
         "RATE NEEDED. No registrar was chosen."),
        ("literal1", "ssl_inr", "INR", "Build", "Software", "Certificates and accounts", "PLT-CLOUD",
         "TLS certificate", "One certificate for the website, renewed through the pilot.",
         "Let's Encrypt or equivalent", "year",
         "Let's Encrypt, no charge, retrieved 10 October 2026. https://letsencrypt.org/"),
        ("literal1", "pg_inr", "INR", "Build", "Software", "Certificates and accounts", "PLT-CLOUD",
         "Database software licence", "PostgreSQL for all seven stores. The machine is priced on the database hourly line.",
         "PostgreSQL", "licence",
         "PostgreSQL Licence, no fee. https://www.postgresql.org/about/licence/"),
        ("S-EC2-H", "ec2_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "App servers", "Two machines, 2 vCPU and 4 GB memory, Linux, running all month in India, for the pilot length.",
         "Amazon EC2 t3.medium or equivalent", "hours",
         "AWS Price List, EC2 Linux on-demand t3.medium, Asia Pacific (Mumbai), USD 0.0448 per hour, publication 2026-10-09."),
        ("S-RDS-H", "rds_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "Database server", "One PostgreSQL instance, 2 vCPU and 2 GB memory, Single-AZ, India region.",
         "Amazon RDS db.t3.small or equivalent", "hours",
         "AWS Price List, RDS PostgreSQL db.t3.small Single-AZ, Asia Pacific (Mumbai), USD 0.053 per hour, publication 2026-10-06."),
        ("S-EBS", "ebs_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "App-server disk", "30 GB general-purpose disk on each app server, for the pilot length.",
         "Amazon EBS gp3 or equivalent", "GB-months",
         "RATE NEEDED. Enter the current Mumbai gp3 USD rate in the yellow input."),
        ("S-ALB", "alb_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "Load balancer", "One load balancer in front of the two app servers.",
         "Amazon ALB or equivalent", "months",
         "RATE NEEDED. Enter the current Mumbai monthly USD rate, or set the sizing quantity aside if you will not use one."),
        ("S-BACKUP", "backup_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "Database backup", "Backup held equal to the provisioned database for each month.",
         "Amazon RDS backup or equivalent", "GB-months",
         "RATE NEEDED. AWS includes some backup storage. The rate stays open so that allowance is not guessed."),
        ("S-EGRESS-BILL-Y", "egress_usd", "USD", "Run", "Services", "Cloud", "PLT-CLOUD",
         "Data transfer out", "Downloads of media above the free monthly allowance. Uploads are not charged.",
         "AWS data transfer or equivalent", "GB",
         "Free 100 GB per month stated on https://aws.amazon.com/s3/pricing/ retrieved 10 October 2026. Paid tier USD 0.1093 per GB from National IT Service, AWS S3 India pricing, retrieved the same day."),
        ("GB-DAT-ACC", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-ACC",
         "Account-store share of database disk", "This store's share of the one provisioned database volume.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "AWS Price List, RDS GP3 Single-AZ, Mumbai, USD 0.131 per GB-month, publication 2026-10-06. Share is a formula so the seven stores add up to the volume."),
        ("GB-DAT-JOB", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-JOB",
         "Job-store share of database disk", "Metadata only. The audio and photos are on the object-storage line.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("GB-DAT-EST", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-EST",
         "Estimate-store share of database disk", "This store's share of the one volume.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("GB-DAT-CAT", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-CAT",
         "Catalogue-store share of database disk", "This store's share of the one volume.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("GB-DAT-BOOK", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-BOOK",
         "Booking-store share of database disk", "This store's share of the one volume.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("GB-DAT-SITE", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-SITE",
         "Site-store share of database disk", "Metadata for attendance and location. Photos are on the object-storage line.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("GB-DAT-INV", "rds_gp3", "USD", "Run", "Services", "Cloud", "DAT-INV",
         "Invoice-store share of database disk", "This store's share of the one volume.",
         "Amazon RDS gp3 or equivalent", "GB-months",
         "Same RDS GP3 rate, publication 2026-10-06."),
        ("S-GBMO-JOB", "s3_gb", "USD", "Run", "Services", "Cloud", "DAT-JOB",
         "Description media storage", "Text stays in the database. Audio and description photos sit here, kept for the whole pilot.",
         "Amazon S3 Standard or equivalent", "GB-months",
         "AWS Price List, S3 Standard first 50 TB, Mumbai, USD 0.025 per GB-month, publication 2026-09-28."),
        ("S-GBMO-SITE", "s3_gb", "USD", "Run", "Services", "Cloud", "DAT-SITE",
         "Site-photo storage", "Site photos, kept for the whole pilot.",
         "Amazon S3 Standard or equivalent", "GB-months",
         "Same S3 Standard rate, publication 2026-09-28."),
        ("S-PUT-JOB", "s3_put", "USD", "Run", "Services", "Cloud", "DAT-JOB",
         "Description upload requests", "One PUT per description file.",
         "Amazon S3 or equivalent", "requests",
         "AWS Price List, S3 PUT, Mumbai, USD 0.000005 per request, publication 2026-09-28."),
        ("S-PUT-SITE", "s3_put", "USD", "Run", "Services", "Cloud", "DAT-SITE",
         "Site-photo upload requests", "One PUT per site photo.",
         "Amazon S3 or equivalent", "requests",
         "Same S3 PUT rate, publication 2026-09-28."),
        ("S-GET-JOB", "s3_get", "USD", "Run", "Services", "Cloud", "DAT-JOB",
         "Description download requests", "One GET per view of a description file.",
         "Amazon S3 or equivalent", "requests",
         "AWS Price List, S3 GET, Mumbai, USD 0.0000004 per request, publication 2026-09-28."),
        ("S-GET-SITE", "s3_get", "USD", "Run", "Services", "Cloud", "DAT-SITE",
         "Site-photo download requests", "One GET per view of a site photo.",
         "Amazon S3 or equivalent", "requests",
         "Same S3 GET rate, publication 2026-09-28."),
        ("S-SMS", "sms_inr", "INR", "Run", "Services", "Messaging", "EXT-OTP",
         "One-time-password SMS", "One transactional SMS per sign-in, India DLT route.",
         "SMSGatewayCenter or equivalent", "SMS",
         "SMSGatewayCenter published transactional slab, 10,000 units, INR 0.17 per SMS, GST extra, retrieved 10 October 2026."),
        ("literal1", "dlt_inr", "INR", "Build", "Services", "Messaging", "EXT-OTP",
         "DLT registration", "Principal entity and one sender header, required before OTP SMS will be delivered.",
         "Jio, Airtel, or Vi DLT", "registration",
         "RATE NEEDED. Published 2026 operator fees disagree on whether GST is inside the figure."),
        ("S-WA-BILL-Y", "wa_inr", "INR", "Run", "Services", "Messaging", "EXT-WABA",
         "WhatsApp business messages", "Billable service messages after the free monthly allowance. Taxes extra.",
         "Meta WhatsApp Business Platform", "messages",
         "Meta India utility and service rate INR 0.1150 effective 1 October 2026, taxes extra. Free allowance 1,000 service messages per phone number per month. Retrieved 10 October 2026."),
        ("S-JOB-Y", "model_inr", "INR", "Run", "Services", "Model", "MOD-EST",
         "Estimate model", "One generation per job from text, audio, or a photo, processed in India. Leave RATE NEEDED if this stays on the app servers.",
         "Any India-region model", "jobs",
         "RATE NEEDED. No India-region model price was selected."),
    ]

    bom_first = 5
    for i, item in enumerate(bom_items):
        qty_key, rate_key, currency, phase, section_name, category, did, item_name, spec, make, unit, source = item
        rr = bom_first + i
        if qty_key == "literal0":
            qty_formula = "=0"
        elif qty_key == "literal1":
            qty_formula = "=1"
        else:
            qty_formula = f"={sizing[qty_key]}"
        paint(ws_m.cell(rr, 1), f"BM-{i+1:02d}", font_bold, fill_white, align=center)
        paint(ws_m.cell(rr, 2), phase, font_label, fill_white, align=center)
        paint(ws_m.cell(rr, 3), section_name, font_label, fill_white, align=center)
        paint(ws_m.cell(rr, 4), category, font_label, fill_white, align=center)
        paint(ws_m.cell(rr, 5), did, font_bold, fill_white, align=center)
        paint(ws_m.cell(rr, 6), item_name, font_label, fill_white)
        paint(ws_m.cell(rr, 7), spec, font_small, fill_white)
        paint(ws_m.cell(rr, 8), make, font_label, fill_white)
        paint(ws_m.cell(rr, 9), qty_formula, font_formula, fill_green, NUM)
        paint(ws_m.cell(rr, 10), unit, font_label, fill_white, align=center)
        paint(ws_m.cell(rr, 11), f"={I(rate_key)}", font_formula, fill_green, NUM8 if rate_key in ("s3_put", "s3_get") else NUM4)
        paint(ws_m.cell(rr, 12), currency, font_label, fill_white, align=center)
        paint(ws_m.cell(rr, 13), inr_amount_formula(f"I{rr}", f"K{rr}", f"L{rr}", I("fx")), font_formula, fill_green, INR)
        paint(ws_m.cell(rr, 14), source, font_small, fill_white)
        ws_m.row_dimensions[rr].height = 48
    bom_last = bom_first + len(bom_items) - 1
    mtot = bom_last + 1
    paint(ws_m.cell(mtot, 6), "BOM total (blank amounts are ignored)", font_bold, fill_amber)
    paint(ws_m.cell(mtot, 13), f"=SUM(M{bom_first}:M{bom_last})", font_formula, fill_amber, INR)
    for col in range(1, 15):
        if ws_m.cell(mtot, col).value is None:
            ws_m.cell(mtot, col).fill = fill_amber
            ws_m.cell(mtot, col).border = thin
    ws_m.freeze_panes = "A5"
    ws_m.auto_filter.ref = f"A4:N{bom_last}"
    ws_m.print_title_rows = "1:4"
    ws_m.page_setup.fitToHeight = False
    bom_amount_range = f"'BOM'!$M${bom_first}:$M${bom_last}"
    bom_rate_range = f"'BOM'!$K${bom_first}:$K${bom_last}"
    bom_id_range = f"'BOM'!$E${bom_first}:$E${bom_last}"
    bom_phase_range = f"'BOM'!$B${bom_first}:$B${bom_last}"
    bom_cat_range = f"'BOM'!$D${bom_first}:$D${bom_last}"
    bom_section_range = f"'BOM'!$C${bom_first}:$C${bom_last}"

    # ------------------------------------------------------------------ timeline
    ws_t = wb.create_sheet("Timeline")
    page_setup(ws_t, landscape=True, fit_width=1, fit_height=1)
    apply_widths(ws_t, {"A": 28, "B": 55, "C": 18, "D": 28, "E": 20, "F": 20, "G": 36})
    ws_t["A1"] = "Mistri AI — timeline"
    ws_t["A1"].font = font_title
    ws_t.merge_cells("A2:G2")
    ws_t["A2"] = "Durations are the yellow week inputs. Dates are formulas from the start date. They are a plan, not a promise."
    ws_t["A2"].font = font_small
    header_row(ws_t, 4, ["Phase", "Milestone", "Depends on", "Duration (weeks)", "Start", "Finish", "Design IDs"])
    phases = [
        ("Design", "Design confirmed", "None", "0", "Confirmed scope and this design", "All trace keys"),
        ("Build", "Build and integrate", "Design", f"={I('build_weeks')}", "Android, WhatsApp, website, and cloud environment ready", "Modules, channels, stores, EXT-OTP, EXT-WABA"),
        ("Onboarding", "Pilot onboarding", "Build", f"={I('onboard_weeks')}", "Workers, contractors, and vendors can sign in", "USR-W, USR-C, USR-V"),
        ("Run", "Pilot operation", "Onboarding", f"={I('pilot_weeks')}", "One year of use inside the pilot caps", "PLT-CLOUD"),
        ("Close-out", "Pilot review", "Run", f"={I('close_weeks')}", "Quantities and open rates reviewed against actual use", "PLT-CLOUD"),
    ]
    for i, (phase, milestone, depends, duration, note, dids) in enumerate(phases):
        rr = 5 + i
        paint(ws_t.cell(rr, 1), phase, font_bold, fill_white)
        paint(ws_t.cell(rr, 2), milestone, font_label, fill_white)
        paint(ws_t.cell(rr, 3), depends, font_label, fill_white)
        if duration == "0":
            paint(ws_t.cell(rr, 4), "=0", font_formula, fill_green, INT, center)
        else:
            paint(ws_t.cell(rr, 4), duration, font_formula, fill_green, INT, center)
        if i == 0:
            paint(ws_t.cell(rr, 5), f"={I('start_date')}", font_formula, fill_green, DATE, center)
        else:
            paint(ws_t.cell(rr, 5), f"=F{rr-1}", font_formula, fill_green, DATE, center)
        paint(ws_t.cell(rr, 6), f"=E{rr}+D{rr}*7", font_formula, fill_green, DATE, center)
        paint(ws_t.cell(rr, 7), dids, font_small, fill_white)
        ws_t.row_dimensions[rr].height = 28
        # stash note under? The milestone column is the milestone. Add note in a second line by increasing description.
        ws_t.cell(rr, 2).value = f"{milestone}. {note}"
    ws_t["A11"] = "Pilot finish"
    ws_t["A11"].font = font_bold
    ws_t["B11"] = "=F9"
    ws_t["B11"].font = font_formula
    ws_t["B11"].number_format = DATE
    ws_t["C11"] = "Weeks from start"
    ws_t["D11"] = "=SUM(D5:D9)"
    ws_t["D11"].font = font_formula
    ws_t["D11"].number_format = INT
    timeline_finish = "'Timeline'!$B$11"
    timeline_start = "'Timeline'!$E$5"
    timeline_weeks = "'Timeline'!$D$11"
    ws_t.freeze_panes = "A5"

    # ------------------------------------------------------------------ cost
    ws_c = wb.create_sheet("Cost Summary")
    page_setup(ws_c, landscape=True, fit_width=1, fit_height=1)
    apply_widths(ws_c, {"A": 42, "B": 22, "C": 22, "D": 22, "E": 55})
    ws_c["A1"] = "Mistri AI — cost summary"
    ws_c["A1"].font = font_title
    ws_c.merge_cells("A2:E2")
    ws_c["A2"] = "Contingency and GST are separate from the line rates. Lines marked RATE NEEDED are counted and are not in the rupee total."
    ws_c["A2"].font = font_small
    header_row(ws_c, 4, ["Category", "Build (INR)", "Run (INR)", "Total (INR)", "What is included"])

    categories = [
        ("Build labour", "BOQ", "Build labour"),
        ("Cloud", "BOM", "Cloud"),
        ("Messaging", "BOM", "Messaging"),
        ("Certificates and accounts", "BOM", "Certificates and accounts"),
        ("Hardware", "BOM", "Hardware"),
        ("Model", "BOM", "Model"),
    ]
    for i, (name, sheet, cat) in enumerate(categories):
        rr = 5 + i
        paint(ws_c.cell(rr, 1), name, font_bold, fill_white)
        if sheet == "BOQ":
            build_f = (
                f'=IF(COUNTIFS({boq_phase_range},"Build",{boq_cat_range},A{rr},{boq_rate_range},">=0")=0,"",'
                f'SUMIFS({boq_amount_range},{boq_phase_range},"Build",{boq_cat_range},A{rr}))'
            )
            run_f = (
                f'=IF(COUNTIFS({boq_phase_range},"Run",{boq_cat_range},A{rr},{boq_rate_range},">=0")=0,"",'
                f'SUMIFS({boq_amount_range},{boq_phase_range},"Run",{boq_cat_range},A{rr}))'
            )
            note = "Person-days. Blank until the day rate is typed."
        else:
            build_f = (
                f'=IF(COUNTIFS({bom_phase_range},"Build",{bom_cat_range},A{rr},{bom_rate_range},">=0")=0,"",'
                f'SUMIFS({bom_amount_range},{bom_phase_range},"Build",{bom_cat_range},A{rr}))'
            )
            run_f = (
                f'=IF(COUNTIFS({bom_phase_range},"Run",{bom_cat_range},A{rr},{bom_rate_range},">=0")=0,"",'
                f'SUMIFS({bom_amount_range},{bom_phase_range},"Run",{bom_cat_range},A{rr}))'
            )
            note = "From the BOM. A blank category still has lines marked RATE NEEDED."
        paint(ws_c.cell(rr, 2), build_f, font_formula, fill_green, INR)
        paint(ws_c.cell(rr, 3), run_f, font_formula, fill_green, INR)
        paint(ws_c.cell(rr, 4), f'=IF(AND(B{rr}="",C{rr}=""),"",N(B{rr})+N(C{rr}))', font_formula, fill_green, INR)
        paint(ws_c.cell(rr, 5), note, font_small, fill_white)
        ws_c.row_dimensions[rr].height = 28
    cat_last = 4 + len(categories)
    sub = cat_last + 1
    paint(ws_c.cell(sub, 1), "Priced subtotal", font_bold, fill_amber)
    paint(ws_c.cell(sub, 2), f"=SUM(B5:B{cat_last})", font_formula, fill_amber, INR)
    paint(ws_c.cell(sub, 3), f"=SUM(C5:C{cat_last})", font_formula, fill_amber, INR)
    paint(ws_c.cell(sub, 4), f"=B{sub}+C{sub}", font_formula, fill_amber, INR)
    paint(ws_c.cell(sub, 5), "Sum of amounts that have a numeric rate.", font_small, fill_amber)
    for col in range(1, 6):
        ws_c.cell(sub, col).border = thin
        ws_c.cell(sub, col).fill = fill_amber

    # reconciliation block sits under, but contingency first
    cont_r = sub + 2
    paint(ws_c.cell(cont_r, 1), "Contingency", font_bold, fill_white)
    paint(ws_c.cell(cont_r, 2), f"={I('contingency')}", font_formula, fill_green, PCT)
    paint(ws_c.cell(cont_r, 4), f"=D{sub}*B{cont_r}", font_formula, fill_green, INR)
    paint(ws_c.cell(cont_r, 5), "Percent of the priced subtotal. It is not hidden inside the line rates.", font_small, fill_white)
    gst_r = cont_r + 1
    paint(ws_c.cell(gst_r, 1), "GST", font_bold, fill_white)
    paint(ws_c.cell(gst_r, 2), f"={I('gst')}", font_formula, fill_green, PCT)
    paint(ws_c.cell(gst_r, 4), f"=(D{sub}+D{cont_r})*B{gst_r}", font_formula, fill_green, INR)
    paint(ws_c.cell(gst_r, 5), "Percent of priced subtotal plus contingency. Confirm taxable lines with a tax adviser.", font_small, fill_white)
    grand = gst_r + 1
    paint(ws_c.cell(grand, 1), "Total with contingency and GST", font_white_bold, fill_teal)
    paint(ws_c.cell(grand, 4), f"=D{sub}+D{cont_r}+D{gst_r}", font_white_bold, fill_teal, INR)
    paint(ws_c.cell(grand, 5), "Excludes every line still marked RATE NEEDED.", font_white_bold, fill_teal)
    for col in range(1, 6):
        ws_c.cell(grand, col).fill = fill_teal
        ws_c.cell(grand, col).font = font_white_bold
        ws_c.cell(grand, col).border = thin
        ws_c.cell(grand, col).alignment = left_align
    ws_c.cell(grand, 4).number_format = INR
    ws_c.cell(grand, 4).font = font_white_bold

    open_r = grand + 2
    paint(ws_c.cell(open_r, 1), "BOQ lines still RATE NEEDED", font_bold, fill_red)
    paint(ws_c.cell(open_r, 4), f'=COUNTIF({boq_rate_range},"RATE NEEDED")', font_formula, fill_red, INT)
    paint(ws_c.cell(open_r + 1, 1), "BOM lines still RATE NEEDED", font_bold, fill_red)
    paint(ws_c.cell(open_r + 1, 4), f'=COUNTIF({bom_rate_range},"RATE NEEDED")', font_formula, fill_red, INT)
    for rr in (open_r, open_r + 1):
        for col in range(1, 6):
            ws_c.cell(rr, col).fill = fill_red
            ws_c.cell(rr, col).border = thin

    rec = open_r + 3
    section(ws_c, rec, "RECONCILIATION", 5)
    paint(ws_c.cell(rec + 1, 1), "BOQ total", font_label, fill_white)
    paint(ws_c.cell(rec + 1, 4), f"='BOQ'!I{tot}", font_formula, fill_green, INR)
    paint(ws_c.cell(rec + 2, 1), "BOM total", font_label, fill_white)
    paint(ws_c.cell(rec + 2, 4), f"='BOM'!M{mtot}", font_formula, fill_green, INR)
    paint(ws_c.cell(rec + 3, 1), "BOQ plus BOM", font_bold, fill_white)
    paint(ws_c.cell(rec + 3, 4), f"=D{rec+1}+D{rec+2}", font_formula, fill_green, INR)
    paint(ws_c.cell(rec + 4, 1), "Matches priced subtotal", font_bold, fill_white)
    paint(ws_c.cell(rec + 4, 4), f'=IF(ABS(D{rec+3}-D{sub})<0.01,"Yes","No")', font_formula, fill_green, align=center)
    for rr in range(rec + 1, rec + 5):
        for col in range(1, 6):
            if ws_c.cell(rr, col).value is None:
                ws_c.cell(rr, col).fill = fill_white
            ws_c.cell(rr, col).border = thin

    section_row = rec + 6
    section(ws_c, section_row, "BY CHANNEL OF BUY", 5)
    paint(ws_c.cell(section_row + 1, 1), "Hardware", font_label, fill_white)
    paint(ws_c.cell(section_row + 1, 4), f'=SUMIF({bom_section_range},"Hardware",{bom_amount_range})', font_formula, fill_green, INR)
    paint(ws_c.cell(section_row + 2, 1), "Software", font_label, fill_white)
    paint(ws_c.cell(section_row + 2, 4), f'=SUMIF({bom_section_range},"Software",{bom_amount_range})', font_formula, fill_green, INR)
    paint(ws_c.cell(section_row + 3, 1), "Services", font_label, fill_white)
    paint(ws_c.cell(section_row + 3, 4), f'=SUMIF({bom_section_range},"Services",{bom_amount_range})', font_formula, fill_green, INR)
    for rr in range(section_row + 1, section_row + 4):
        for col in range(1, 6):
            ws_c.cell(rr, col).border = thin
            if ws_c.cell(rr, col).value is None:
                ws_c.cell(rr, col).fill = fill_white

    cost_total_cell = f"'Cost Summary'!$D${grand}"
    cost_sub_cell = f"'Cost Summary'!$D${sub}"
    cost_cont_cell = f"'Cost Summary'!$D${cont_r}"
    cost_gst_cell = f"'Cost Summary'!$D${gst_r}"
    cost_open_boq = f"'Cost Summary'!$D${open_r}"
    cost_open_bom = f"'Cost Summary'!$D${open_r+1}"
    cost_match_cell = f"'Cost Summary'!$D${rec+4}"

    # ------------------------------------------------------------------ risks
    ws_r = wb.create_sheet("Risks and Compliance")
    page_setup(ws_r)
    apply_widths(ws_r, {"A": 10, "B": 28, "C": 55, "D": 42, "E": 14, "F": 22, "G": 40})
    ws_r["A1"] = "Mistri AI — risks and compliance"
    ws_r["A1"].font = font_title
    ws_r.merge_cells("A2:G2")
    ws_r["A2"] = "Effects that can be calculated point at the workbook. The rest are conditions to watch."
    ws_r["A2"].font = font_small
    header_row(ws_r, 4, ["ID", "Risk", "What it does to the pilot", "Effect on this report", "Level", "Design IDs", "Response"])
    risks = [
        ("R1", "Build day rate is open",
         "The person-day quantities are in the BOQ. They have no rupee amount until the day-rate input is filled. That labour is likely to be larger than the cloud bill.",
         f'="BOQ lines still open: "&{cost_open_boq}&". Priced total today excludes them."',
         "High", "All build IDs", "Type a day rate on Inputs and Scope. Every labour line then prices itself."),
        ("R2", "WhatsApp messages may all be billable",
         "The quantity subtracts 1,000 free service messages a month. If Meta treats these messages as utility templates, that allowance does not apply.",
         f'="Annual messages "&TEXT({sizing["S-WA-M"]}*{I("months")},"#,##0")&" × "&TEXT({I("wa_inr")},"₹0.0000")&" = "&TEXT({sizing["S-WA-M"]}*{I("months")}*{I("wa_inr")},"₹#,##0.00")&" if the free allowance is set to 0."',
         "Medium", "EXT-WABA, CH-WA", "Set the free-allowance input to 0 and the BOM recalculates. Keep marketing messages out of the pilot."),
        ("R3", "Worker data with OTP only",
         "The pilot stores name, phone, trade, photos, audio, and a saved location. The confirmed security bar is a one-time password, with no audit log and no formal privacy programme.",
         "A complaint or a DPDP notice can stop the pilot. It does not change a quantity until the scope changes.",
         "High", "DAT-ACC, DAT-JOB, DAT-SITE, MOD-AUTH", "Before go-live, add a short privacy notice and a way to delete an account, as a scope change."),
        ("R4", "SMS will not arrive without DLT",
         "India requires a registered principal entity and a template before transactional SMS is delivered. Sign-in depends on that SMS.",
         "The registration line is RATE NEEDED. Login quantity is already counted.",
         "High", "EXT-OTP, MOD-AUTH", "Start DLT registration during the build phase, not on the week of go-live."),
        ("R5", "Estimates use typed vendor prices",
         "There is no starting price list and no map. A match can only use what the vendor typed, including a wrong locality or a stale price.",
         "No cost line changes. Estimate quality is the product risk.",
         "Medium", "MOD-CAT, MOD-MATCH, MOD-EST", "Show the vendor's typed date on the match screen so the contractor sees how old the price is."),
        ("R6", "Single database zone",
         "The database price is a Single-AZ instance. A zone failure takes the pilot offline.",
         "A second zone is a different AWS rate. It is not in this issue.",
         "Medium", "PLT-CLOUD, DAT-ACC", "Accept the downtime for the pilot, or price a Multi-AZ instance as a change."),
        ("R7", "Dollar prices move",
         "Cloud lines are USD rates converted at the FBIL rate of 9 October 2026.",
         "Edit the rupees-per-dollar input. Every USD line recalculates.",
         "Low", "PLT-CLOUD", "Replace the FX input with the FBIL rate on the day you place the order."),
        ("R8", "Model fee is open",
         "Hindi speech and photo estimates may need a paid model. The app servers are already priced. A separate model must not be added twice.",
         "The model line is one per job and stays RATE NEEDED.",
         "Medium", "MOD-EST, MOD-JOB", "If the model runs on the two app servers, leave the rate as RATE NEEDED."),
    ]
    for i, rec in enumerate(risks):
        rr = 5 + i
        for c, val in enumerate(rec, 1):
            is_formula = isinstance(val, str) and val.startswith("=")
            paint(
                ws_r.cell(rr, c),
                val,
                font_formula if is_formula else (font_bold if c in (1, 2) else font_label),
                fill_green if is_formula else fill_white,
            )
        ws_r.row_dimensions[rr].height = 48
    comp_h = 5 + len(risks) + 2
    section(ws_r, comp_h, "COMPLIANCE", 7)
    compliance = [
        ("C1", "DPDP Act 2023", "Phone numbers, names, photos, audio, and location are personal data. The confirmed pilot does not include a privacy programme.", "Treat the risk R3 response as a gate before real workers are invited."),
        ("C2", "TRAI commercial-communication rules", "OTP SMS needs DLT registration, a header, and a template.", "Priced as a RATE NEEDED registration plus the SMS rate."),
        ("C3", "Meta WhatsApp Business policy", "Business-initiated messages follow Meta's categories and the rate card of 1 October 2026.", "The BOM uses the India service and utility rate and the published free allowance."),
        ("C4", "GST on this project", "The product does not raise GST invoices. The project's own purchases can still attract GST.", "GST is a separate line at the input rate, not inside the unit rates."),
        ("C5", "Payments", "The pilot does not collect money, so it does not need a payment-aggregator licence.", "UPI stays out of scope. The invoice is a record of paid or unpaid."),
        ("C6", "Maps and location", "A location saved with a site update is a record. It is not sent to a map service.", "No map licence is in the BOM."),
    ]
    header_row(ws_r, comp_h + 1, ["ID", "Rule", "How it touches this pilot", "Where it sits in the report", "", "", ""])
    for i, rec in enumerate(compliance):
        rr = comp_h + 2 + i
        for c, val in enumerate(rec, 1):
            paint(ws_r.cell(rr, c), val, font_bold if c == 1 else font_label, fill_white)
        ws_r.row_dimensions[rr].height = 36
    ws_r.freeze_panes = "A5"
    ws_r.print_title_rows = "1:4"
    risk1 = "'Risks and Compliance'!B5"
    risk2 = "'Risks and Compliance'!B6"
    risk3 = "'Risks and Compliance'!B7"

    # ------------------------------------------------------------------ assumptions
    ws_a = wb.create_sheet("Assumptions", 2)
    page_setup(ws_a)
    apply_widths(ws_a, {"A": 8, "B": 78, "C": 70, "D": 28})
    ws_a["A1"] = "Mistri AI — assumptions"
    ws_a["A1"].font = font_title
    ws_a.merge_cells("A2:D2")
    ws_a["A2"] = "Each assumption names the input it drives. Change the input, not this text, when you want the numbers to move."
    ws_a["A2"].font = font_small
    header_row(ws_a, 4, ["No.", "Assumption", "Effect on quantity or cost", "Input"])
    assumptions = [
        (1, "A job is counted once. Volume is the jobs-per-contractor input. A job a worker creates is inside that count.",
         f'="Jobs per month = "&TEXT({sizing["S-JOB-M"]},"#,##0")&", and "&TEXT({sizing["S-JOB-Y"]},"#,##0")&" in the pilot."',
         "Jobs per contractor per month"),
        (2, "Each job has one text description, the stated number of audio clips, and the stated number of description photos.",
         f'="Description media per job = "&TEXT({sizing["S-MB-DESC"]},"0.0")&" MB."',
         "Photos, audio clips, minutes"),
        (3, "Each job has the stated number of site updates, and each update has the stated number of photos plus attendance.",
         f'="Site updates = "&TEXT({sizing["S-SITE"]},"#,##0")&". Site media per job = "&TEXT({sizing["S-MB-SITE"]},"0.0")&" MB."',
         "Site updates, photos per update"),
        (4, "Each job produces one estimate, one material match, the stated number of bookings, and one invoice.",
         f'="Bookings = "&TEXT({sizing["S-BOOK"]},"#,##0")&". Invoices = "&TEXT({sizing["S-INV"]},"#,##0")&"."',
         "Bookings per job"),
        (5, "Vendors type every catalogue edit. There is no file to load on day one.",
         f'="Catalogue edits = "&TEXT({sizing["S-CAT"]},"#,##0")&"."',
         "Catalogue edits per vendor"),
        (6, "Each account signs in at the stated monthly rate. Each sign-in sends one SMS. WhatsApp does not send a second code.",
         f'="OTP SMS = "&TEXT({sizing["S-SMS"]},"#,##0")&"."',
         "Sign-ins per account"),
        (7, "The stated share of jobs runs on WhatsApp. Each of those jobs causes the stated number of business messages: estimate, match, booking, one reply per site update, and invoice.",
         "Sets the job part of the WhatsApp message count.",
         "WhatsApp job share and messages per job"),
        (8, "The stated share of catalogue edits is confirmed with one WhatsApp message.",
         f'="WhatsApp business messages per month = "&TEXT({sizing["S-WA-M"]},"#,##0.0")&"."',
         "Vendor WhatsApp share"),
        (9, "Photos and audio use the stated average sizes. Text is inside the database row, not in object storage.",
         f'="Media per job = "&TEXT({sizing["S-MB-JOB"]},"0.0")&" MB."',
         "Photo size, audio size"),
        (10, "Media is kept for the whole pilot. Object storage is billed on the growing total, not on the last month alone.",
         f'="Object storage = "&TEXT({sizing["S-GBMO"]},"#,##0.00")&" GB-months."',
         "Pilot length"),
        (11, "Each media file is viewed the stated number of times in the year. Uploads are free of transfer-out charges.",
         f'="Downloaded = "&TEXT({sizing["S-DOWN-Y"]},"#,##0.00")&" GB. Billable transfer = "&TEXT({sizing["S-EGRESS-BILL-Y"]},"#,##0.00")&" GB."',
         "Views per file"),
        (12, "Peak concurrency is the stated share of accounts. Two app servers still run all month so one can fail. Concurrency does not reduce the hours.",
         f'="Peak users = "&TEXT({sizing["S-PEAK"]},"0.0")&". App-server hours = "&TEXT({sizing["S-EC2-H"]},"#,##0")&"."',
         "Concurrent share, app servers"),
        (13, "Each database row is the stated size. The volume billed is the greater of that size and 20 GB, which is the smallest general-purpose RDS volume.",
         f'="Calculated "&TEXT({sizing["S-DB-GB"]},"0.000")&" GB. Provisioned "&TEXT({sizing["S-DB-PROV"]},"0")&" GB."',
         "Kilobytes per row, minimum volume"),
        (14, "Each app server has the stated disk. The GB-months are derived. The rupee rate is open.",
         f'="Disk = "&TEXT({sizing["S-EBS"]},"#,##0")&" GB-months, rate still RATE NEEDED."',
         "Disk per app server"),
        (15, "WhatsApp messages are treated as service messages, and the published free monthly allowance is subtracted. Set that input to 0 if they are utility templates.",
         f'="Billable WhatsApp messages = "&TEXT({sizing["S-WA-BILL-Y"]},"#,##0")&"."',
         "Free WhatsApp messages"),
        (16, "Build person-days are planning allowances, one per design item. They are not a measured take-off.",
         f'="Build days, excluding support = "&TEXT({sizing["S-DAYS"]},"#,##0")&". Amount stays blank until a day rate is typed."',
         "Days block"),
        (17, "After go-live the pilot is supported for the stated days each month.",
         f'="Support days = "&TEXT({sizing["S-SUPPORT"]},"#,##0")&"."',
         "Support days per month"),
        (18, "Contingency is the stated percent of the priced subtotal only. Open lines join it automatically once a rate is typed.",
         f'="Contingency = "&TEXT({cost_cont_cell},"₹#,##0.00")&"."',
         "Contingency"),
        (19, "GST is the stated percent of priced subtotal plus contingency, on its own line. A tax adviser confirms which purchases are taxable.",
         f'="GST = "&TEXT({cost_gst_cell},"₹#,##0.00")&"."',
         "GST"),
        (20, "US dollar prices convert at the FBIL reference rate on the inputs sheet.",
         f'="Rate = "&TEXT({I("fx")},"0.0000")&" rupees per dollar."',
         "Rupees per US dollar"),
        (21, "A billed month is 730 hours, which is the divisor AWS uses to show a monthly price from an hourly price.",
         "App-server hours and database hours use this figure.",
         "Hours in a billed month"),
        (22, "One load balancer is allowed for the two servers. Its rate is open.",
         f'="Load-balancer months = "&TEXT({sizing["S-ALB"]},"0")&"."',
         "Load balancer rate"),
        (23, "A separate estimate model, if bought, is one fee per job. It stays open so it is not double-counted with the app servers.",
         f'="Model units = "&TEXT({sizing["S-JOB-Y"]},"#,##0")&" jobs."',
         "Model fee per job"),
        (24, "The timeline weeks are planning allowances. They move the dates and do not by themselves move the cost.",
         f'="Plan length = "&TEXT({timeline_weeks},"0")&" weeks."',
         "Week inputs"),
    ]
    for i, (num, text, effect, inp) in enumerate(assumptions):
        rr = 5 + i
        paint(ws_a.cell(rr, 1), num, font_bold, fill_white, align=center)
        paint(ws_a.cell(rr, 2), text, font_label, fill_white)
        is_f = isinstance(effect, str) and effect.startswith("=")
        paint(ws_a.cell(rr, 3), effect, font_formula if is_f else font_label, fill_green if is_f else fill_white)
        paint(ws_a.cell(rr, 4), inp, font_small, fill_white)
        ws_a.row_dimensions[rr].height = 48
    ws_a.freeze_panes = "A5"
    ws_a.print_title_rows = "1:4"
    ws_a.page_setup.fitToHeight = False

    # ------------------------------------------------------------------ summary
    ws = wb.create_sheet("Summary", 0)
    page_setup(ws, landscape=True, fit_width=1, fit_height=1)
    apply_widths(ws, {"A": 28, "B": 88, "C": 28, "D": 36})
    ws.row_dimensions[1].height = 30
    ws["A1"] = "Mistri AI"
    ws["A1"].font = font_title
    ws.merge_cells("B1:D1")
    ws["B1"] = "Detailed Project Report — Delhi pilot"
    ws["B1"].font = Font(name="Calibri", size=16, color=TEAL, bold=True)
    ws.merge_cells("A2:D2")
    ws["A2"] = "First issue, 10 October 2026. Design confirmed with no change. Yellow cells on Inputs and Scope are the inputs."
    ws["A2"].font = font_small

    section(ws, 4, "PROJECT", 4)
    project_rows = [
        ("Name", f"={I('project')}"),
        ("Place", f"={I('place')}"),
        ("Kind", "IT system for construction workers, contractors, and vendors"),
        ("Accounts", f'={I("workers")}&" workers, "&{I("contractors")}&" contractors, "&{I("vendors")}&" vendors"'),
        ("Length", f'={I("months")}&" months, including build and one year of running"'),
        ("Channels", "Android app, WhatsApp, and a website, in Hindi and English"),
    ]
    for i, (label, val) in enumerate(project_rows):
        rr = 5 + i
        paint(ws.cell(rr, 1), label, font_bold, fill_white)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        is_f = isinstance(val, str) and val.startswith("=")
        paint(ws.cell(rr, 2), val, font_formula if is_f else font_label, fill_green if is_f else fill_white)
        for col in range(3, 5):
            ws.cell(rr, col).border = thin
            ws.cell(rr, col).fill = fill_green if is_f else fill_white

    section(ws, 12, "SCOPE IN FIVE LINES", 4)
    scope5 = [
        f'="1. Delhi pilot for "&{I("workers")}&" construction workers, "&{I("contractors")}&" contractors, and "&{I("vendors")}&" vendors."',
        "2. Android app, WhatsApp, and a website, in Hindi and English, on a public cloud in India.",
        "3. A job in text, audio, or a photo becomes a quantity list and a cost estimate, then a vendor match and a booking.",
        "4. Site progress, attendance, and photos are recorded, and an internal invoice is marked unpaid or paid.",
        "5. Sign-in is a phone one-time password. UPI, GST invoicing, and maps are outside this version.",
    ]
    for i, text in enumerate(scope5):
        rr = 13 + i
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
        is_f = text.startswith("=")
        paint(ws.cell(rr, 1), text, font_formula if is_f else font_label, fill_green if is_f else fill_white)
        for col in range(2, 5):
            ws.cell(rr, col).border = thin
            ws.cell(rr, col).fill = fill_green if is_f else fill_white
        ws.row_dimensions[rr].height = 22

    section(ws, 19, "TOTAL COST", 4)
    cost_lines = [
        ("Priced subtotal", f"={cost_sub_cell}", "Lines that have a sourced rate."),
        ("Contingency", f"={cost_cont_cell}", "Separate from the line rates."),
        ("GST", f"={cost_gst_cell}", "Separate from the line rates."),
        ("Total", f"={cost_total_cell}", "Priced subtotal plus contingency plus GST."),
        ("Still open", f'={cost_open_boq}&" build lines and "&{cost_open_bom}&" buy lines are marked RATE NEEDED. Their amounts are blank and are not in the total."', "Type a rate on Inputs and Scope and the total moves."),
    ]
    for i, (label, formula, note) in enumerate(cost_lines):
        rr = 20 + i
        paint(ws.cell(rr, 1), label, font_white_bold if label == "Total" else font_bold, fill_teal if label == "Total" else fill_white)
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
        paint(ws.cell(rr, 2), formula, font_white_bold if label == "Total" else font_formula, fill_teal if label == "Total" else fill_green, INR if label != "Still open" else None)
        paint(ws.cell(rr, 4), note, font_white_bold if label == "Total" else font_small, fill_teal if label == "Total" else fill_white)
        ws.cell(rr, 3).border = thin
        ws.cell(rr, 3).fill = fill_teal if label == "Total" else fill_green
        if label == "Total":
            ws.cell(rr, 2).number_format = INR
            ws.cell(rr, 2).font = font_white_bold
        ws.row_dimensions[rr].height = 28 if label == "Still open" else 20

    section(ws, 26, "TIMELINE", 4)
    paint(ws.cell(27, 1), "Start", font_bold, fill_white)
    paint(ws.cell(27, 2), f"={timeline_start}", font_formula, fill_green, DATE)
    paint(ws.cell(28, 1), "Finish", font_bold, fill_white)
    paint(ws.cell(28, 2), f"={timeline_finish}", font_formula, fill_green, DATE)
    paint(ws.cell(29, 1), "Duration", font_bold, fill_white)
    paint(ws.cell(29, 2), f'={timeline_weeks}&" weeks"', font_formula, fill_green)
    for rr in range(27, 30):
        ws.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=4)
        for col in range(3, 5):
            ws.cell(rr, col).border = thin
            ws.cell(rr, col).fill = fill_green

    section(ws, 31, "KEY RISKS", 4)
    for i, ref in enumerate([risk1, risk2, risk3]):
        rr = 32 + i
        ws.merge_cells(start_row=rr, start_column=1, end_row=rr, end_column=4)
        paint(ws.cell(rr, 1), f"={ref}", font_formula, fill_green)
        for col in range(2, 5):
            ws.cell(rr, col).border = thin
            ws.cell(rr, col).fill = fill_green
        ws.row_dimensions[rr].height = 22

    section(ws, 36, "HOW TO CHANGE A NUMBER", 4)
    ws.merge_cells("A37:D37")
    paint(ws.cell(37, 1), "Edit a yellow cell on Inputs and Scope. Rates that are still open are the yellow cells whose value is RATE NEEDED. Replace that word with a number. Do not type over a blue cell. Contingency and GST stay on this cost block and are not inside the unit rates.", font_label, fill_yellow)
    for col in range(2, 5):
        ws.cell(37, col).fill = fill_yellow
        ws.cell(37, col).border = thin
    ws.row_dimensions[37].height = 36
    ws.page_setup.fitToHeight = 1
    ws.page_setup.fitToWidth = 1
    ws.print_area = "A1:D37"
    ws.sheet_view.zoomScale = 120

    # ------------------------------------------------------------------ change log
    ws_l = wb.create_sheet("Change Log")
    page_setup(ws_l, landscape=True, fit_width=1, fit_height=1)
    apply_widths(ws_l, {"A": 12, "B": 20, "C": 28, "D": 78, "E": 40})
    ws_l["A1"] = "Mistri AI — change log"
    ws_l["A1"].font = font_title
    ws_l.merge_cells("A2:E2")
    ws_l["A2"] = "Corrections after your feedback are added here and applied everywhere they apply."
    ws_l["A2"].font = font_small
    header_row(ws_l, 4, ["No.", "Date", "Where", "What changed", "Carried into"])
    logs = [
        ("CL-01", date(2026, 10, 10), "Scope", "Scope confirmed with no change.", "Inputs and Scope"),
        ("CL-02", date(2026, 10, 10), "Design", "Design confirmed with no change to IDs, screens, or roles.", "Design, BOQ, BOM"),
        ("CL-03", date(2026, 10, 10), "Trace rule", "Screen IDs stay under the parent module and are not priced a second time. This was the rule in the confirmed design.", "BOQ, BOM"),
        ("CL-04", date(2026, 10, 10), "Rates", "No rate was invented. Published rates are named with their date. Every other rate is the input RATE NEEDED and its amount is blank.", "BOQ, BOM, Cost Summary"),
        ("CL-05", date(2026, 10, 10), "Self-check", "Checked before issue: every trace ID has a line, every line cites a trace ID, quantities are formulas from Sizing, totals reconcile, and every numeric rate has a source.", "Cost Summary reconciliation"),
        ("CL-06", date(2026, 10, 10), "Cost Summary", "A category with no numeric rate is left blank. It is not shown as zero, so the open day rate and the open model fee are not read as a free item.", "Summary total"),
    ]
    for i, rec in enumerate(logs):
        rr = 5 + i
        for c, val in enumerate(rec, 1):
            fmt = DATE if c == 2 else None
            paint(ws_l.cell(rr, c), val, font_bold if c == 1 else font_label, fill_white, fmt)
        ws_l.row_dimensions[rr].height = 36
    ws_l.freeze_panes = "A5"

    # ------------------------------------------------------------------ structural self-check
    boq_ids = [item[1] for item in boq_items]
    bom_ids = [item[6] for item in bom_items]
    used = set(boq_ids) | set(bom_ids)
    missing = sorted(trace_ids - used)
    unknown = sorted(used - trace_ids)
    if missing or unknown:
        raise SystemExit(f"Trace check failed. Missing {missing}. Unknown {unknown}.")

    # reorder sheets explicitly
    order = [
        "Summary",
        "Inputs and Scope",
        "Assumptions",
        "Design",
        "Sizing",
        "BOQ",
        "BOM",
        "Timeline",
        "Cost Summary",
        "Risks and Compliance",
        "Change Log",
    ]
    for i, name in enumerate(order):
        wb.move_sheet(name, offset=i - wb.sheetnames.index(name))

    wb.properties.title = "Mistri AI Detailed Project Report"
    wb.properties.subject = "Delhi pilot"
    wb.properties.creator = "Mistri AI DPR"
    wb.properties.description = "Scope, design, sizing, BOQ, and BOM for the Mistri AI Delhi pilot."

    wb.save(OUT)
    print("saved", OUT)
    print("boq lines", len(boq_items), "bom lines", len(bom_items))
    print("trace ids", len(trace_ids), "covered", len(used))
    print("sheets", wb.sheetnames)


if __name__ == "__main__":
    main()
