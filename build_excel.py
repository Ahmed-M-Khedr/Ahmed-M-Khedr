"""
build_excel.py
Generates Shoghly.xlsx – a fully-formatted Excel workbook that mirrors
the Shoghly MySQL database schema, pre-fills sample data in every table,
and exposes every query from the problem statement as a readable sheet.
"""

import openpyxl
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, GradientFill
)
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList

# ── colour palette ────────────────────────────────────────────────────────────
C = {
    "navy":     "1B2A4A",
    "blue":     "2563EB",
    "lblue":    "DBEAFE",
    "teal":     "0D9488",
    "lteal":    "CCFBF1",
    "green":    "16A34A",
    "lgreen":   "DCFCE7",
    "amber":    "D97706",
    "lamber":   "FEF3C7",
    "red":      "DC2626",
    "lred":     "FEE2E2",
    "purple":   "7C3AED",
    "lpurple":  "EDE9FE",
    "gray":     "6B7280",
    "lgray":    "F3F4F6",
    "white":    "FFFFFF",
    "black":    "111827",
    "gold":     "F59E0B",
    "lgold":    "FFFBEB",
}

# helper fills / fonts
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(hex_color="111827", bold=False, size=11, name="Calibri"):
    return Font(color=hex_color, bold=bold, size=size, name=name)

def center():
    return Alignment(horizontal="center", vertical="center",
                     wrap_text=True, readingOrder=2)

def right():
    return Alignment(horizontal="right", vertical="center",
                     wrap_text=True, readingOrder=2)

thin = Side(style="thin", color="D1D5DB")
med  = Side(style="medium", color="9CA3AF")
def border(style="thin"):
    s = thin if style == "thin" else med
    return Border(left=s, right=s, top=s, bottom=s)

# ── worksheet helpers ─────────────────────────────────────────────────────────
def set_col_widths(ws, widths):
    for col, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = w

def freeze(ws, cell="A2"):
    ws.freeze_panes = cell

def write_header_row(ws, row_num, headers, bg, fg="FFFFFF", size=11, bold=True):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row_num, column=col, value=h)
        c.fill    = fill(bg)
        c.font    = font(fg, bold=bold, size=size)
        c.alignment = center()
        c.border  = border("med")

def write_data_row(ws, row_num, values, bg=None, fg="111827", bold=False, align_fn=None):
    af = align_fn or center()
    for col, v in enumerate(values, 1):
        c = ws.cell(row=row_num, column=col, value=v)
        if bg:
            c.fill = fill(bg)
        c.font      = font(fg, bold=bold)
        c.alignment = af
        c.border    = border()

def row_height(ws, row, h):
    ws.row_dimensions[row].height = h

def merge_title(ws, row, col_start, col_end, text, bg, fg="FFFFFF", size=14):
    ws.merge_cells(start_row=row, start_column=col_start,
                   end_row=row, end_column=col_end)
    c = ws.cell(row=row, column=col_start, value=text)
    c.fill      = fill(bg)
    c.font      = font(fg, bold=True, size=size)
    c.alignment = center()
    c.border    = border("med")
    row_height(ws, row, 32)

# ─────────────────────────────────────────────────────────────────────────────
# SAMPLE DATA
# ─────────────────────────────────────────────────────────────────────────────
contractors = [
    (1, "أحمد محمد",   "01012345678", "2024-01-01"),
    (2, "محمود علي",   "01198765432", "2024-01-15"),
    (3, "سامي يوسف",   "01065432100", "2024-02-01"),
]

projects = [
    (1, "مشروع النيل الأزرق",  "القاهرة - المعادي",    1500000.00, 1),
    (2, "مشروع الفيوم الجديد", "الفيوم - مدينة الجديدة", 900000.00, 2),
    (3, "مشروع الساحل",        "الإسكندرية - الساحل",   2200000.00, 3),
]

buildings = [
    (1, 1, "عمارة 101"),
    (2, 1, "عمارة 102"),
    (3, 2, "عمارة 201"),
    (4, 3, "عمارة 301"),
    (5, 3, "عمارة 302"),
]

location_details = [
    (1,  1, "الدور الأول",  1, "Floor"),
    (2,  1, "الدور الثاني", 2, "Floor"),
    (3,  1, "السلم",        None, "Stairs"),
    (4,  1, "السطح",        None, "Roof"),
    (5,  1, "البدروم",      None, "Basement"),
    (6,  2, "الدور الأول",  1, "Floor"),
    (7,  2, "السلم",        None, "Stairs"),
    (8,  3, "الدور الأول",  1, "Floor"),
    (9,  4, "الدور الأول",  1, "Floor"),
    (10, 5, "الدور الأول",  1, "Floor"),
]

facades = [
    (1, 1, "Front"),
    (2, 1, "Back"),
    (3, 1, "Left Side"),
    (4, 2, "Front"),
    (5, 3, "Front"),
    (6, 4, "Front"),
    (7, 4, "Right Side"),
]

workers = [
    (1, "كريم سالم",   "01011112222", "محارة",   350.00, 1),
    (2, "طارق حسن",   "01033334444", "نقاشة",   300.00, 1),
    (3, "إبراهيم عبد الله", "01055556666", "سباكة", 400.00, 2),
    (4, "عمر خالد",   "01077778888", "كهرباء",  420.00, 2),
    (5, "يوسف أحمد",  "01099990000", "محارة",   350.00, 3),
    (6, "مصطفى علي",  "01022221111", "نجارة",   380.00, 3),
]

attendance = [
    (1,  1, 1,    None, None, "2024-03-01", "Daily",     0,     0,     "Full Day", 0,    350.00),
    (2,  2, 2,    None, None, "2024-03-01", "Daily",     0,     0,     "Half Day", 0,    150.00),
    (3,  3, None, 1,    None, "2024-03-01", "Per Meter", 12.5,  80.00, None,       0,   1000.00),
    (4,  4, 3,    None, None, "2024-03-02", "Daily",     0,     0,     "Full Day", 50,   470.00),
    (5,  5, 8,    None, None, "2024-03-02", "Daily",     0,     0,     "Full Day", 0,    350.00),
    (6,  6, None, 4,    None, "2024-03-02", "Per Meter", 8.0,   90.00, None,       0,    720.00),
    (7,  1, 1,    None, None, "2024-03-03", "Daily",     0,     0,     "Full Day", 0,    350.00),
    (8,  2, 2,    None, None, "2024-03-03", "Daily",     0,     0,     "Full Day", 100,  400.00),
    (9,  3, None, 2,    None, "2024-03-03", "Per Meter", 10.0,  80.00, None,       0,    800.00),
    (10, 4, 3,    None, None, "2024-03-04", "Daily",     0,     0,     "Full Day", 0,    420.00),
]

worker_payments = [
    (1, 1, 1, 500.00,  "2024-03-05", "Advance",         "سلفة أولى"),
    (2, 2, 1, 300.00,  "2024-03-05", "Private",         "مصاريف شخصية"),
    (3, 3, 2, 1000.00, "2024-03-10", "Full Salary",     "راتب كامل"),
    (4, 4, 2, 200.00,  "2024-03-10", "Family Transfer", "تحويل للأهل"),
    (5, 5, 3, 700.00,  "2024-03-12", "Advance",         "سلفة"),
    (6, 6, 3, 400.00,  "2024-03-12", "Private",         ""),
]

general_expenses = [
    (1,  1, 1, 1,    None, "أسمنت",         "Materials", 3500.00, "2024-03-01", ""),
    (2,  1, 1, None, 1,    "دهان خارجي",    "Materials", 1200.00, "2024-03-02", ""),
    (3,  1, 2, 6,    None, "رمل ومواد",     "Materials", 800.00,  "2024-03-02", ""),
    (4,  2, 3, 8,    None, "أدوات سباكة",   "Tools",     600.00,  "2024-03-05", ""),
    (5,  3, 4, 9,    None, "نقل معدات",     "Transport", 2000.00, "2024-03-06", ""),
    (6,  1, 1, None, 2,    "بويات واجهة",   "Materials", 1500.00, "2024-03-07", ""),
    (7,  1, 1, 3,    None, "مواد عازلة",    "Materials", 950.00,  "2024-03-08", ""),
    (8,  2, None, None, None, "مصاريف متنوعة", "Other",  400.00,  "2024-03-09", ""),
]

shared_meals = [
    (1, 1, "2024-03-01", 480.00, "غداء كشري"),
    (2, 1, "2024-03-02", 600.00, "غداء فول وطعمية"),
    (3, 2, "2024-03-05", 350.00, "غداء أرز وبطاطس"),
    (4, 3, "2024-03-06", 720.00, "غداء كبدة وسجق"),
]

meal_participants = [
    (1,1),(1,2),(1,3),(1,4),
    (2,1),(2,2),(2,5),(2,6),
    (3,3),(3,4),(3,5),
    (4,5),(4,6),(4,1),(4,2),(4,3),(4,4),
]

contractor_incomes = [
    (1, 1, 1, 1, None, None, 200000.00, "2024-03-15", "دفعة أولى محارة"),
    (2, 1, 1, 2, None, None, 150000.00, "2024-03-20", "دفعة ثانية"),
    (3, 1, 1, None, 1,   None, 80000.00, "2024-03-25", "واجهة أمامية"),
    (4, 2, 2, 3, None, None, 120000.00, "2024-04-01", "دفعة مشروع 2"),
    (5, 3, 3, 4, None, None, 300000.00, "2024-04-05", "دفعة مشروع 3"),
]

# ─────────────────────────────────────────────────────────────────────────────
# QUERY RESULTS (pre-computed from sample data)
# ─────────────────────────────────────────────────────────────────────────────

# Query 1 – workers present today (using sample date)
q_present = [
    ("كريم سالم",   "Full Day", "Daily",     350.00),
    ("طارق حسن",   "Half Day", "Daily",     150.00),
    ("إبراهيم عبد الله", None, "Per Meter", 1000.00),
]

# Query 2 – absent workers
q_absent = []   # none absent in sample

# Query 3 – total incomes
total_incomes = sum(r[6] for r in contractor_incomes)   # 850000

# Query 4 – total expenses
total_expenses = sum(r[7] for r in general_expenses)    # 11950

# Query 5 – net profit (incomes – expenses – attendance earned)
total_earned_att = sum(r[11] for r in attendance)       # 5010
net_profit = total_incomes - total_expenses - total_earned_att  # 833040

# Query 6 – most owed worker
worker_earnings = {}
for a in attendance:
    wid = a[1]; amt = a[11]
    worker_earnings[wid] = worker_earnings.get(wid, 0) + amt
most_owed_id = max(worker_earnings, key=worker_earnings.get)
most_owed_name = next(w[1] for w in workers if w[0] == most_owed_id)
most_owed_amt  = worker_earnings[most_owed_id]

# Query 7 – project summary
proj_incomes  = {1: 430000, 2: 120000, 3: 300000}
proj_expenses = {1: 7950,   2: 1000,   3: 2000}

# Query 8 – buildings per project
bld_per_proj = {}
for b in buildings:
    bld_per_proj[b[1]] = bld_per_proj.get(b[1], 0) + 1

# Query 9 – meal cost per worker (simplified)
meal_cost_per_worker = {
    1: round(480/4 + 600/4, 2),        # meals 1,2
    2: round(480/4 + 600/4, 2),
    3: round(480/4 + 350/3 + 720/6, 2),
    4: round(480/4 + 350/3 + 720/6, 2),
    5: round(600/4 + 350/3 + 720/6, 2),
    6: round(600/4 + 720/6, 2),
}

# Query 10 – daily workers full account
daily_workers_report = []
for w in workers:
    wid = w[0]
    att_rows = [a for a in attendance if a[1]==wid and a[6]=="Daily"]
    days     = len([a for a in att_rows if a[9]!="Absent"])
    extras   = sum(a[10] for a in [a for a in attendance if a[1]==wid])
    earned   = sum(a[11] for a in attendance if a[1]==wid)
    payments = sum(p[3] for p in worker_payments if p[1]==wid)
    meals    = meal_cost_per_worker.get(wid, 0)
    net      = round(earned - payments - meals, 2)
    if att_rows:
        daily_workers_report.append((w[1], days, w[4], extras, earned, payments, meals, net))

# Query 11 – per-meter workers full account
meter_workers_report = []
for w in workers:
    wid = w[0]
    meter_rows = [a for a in attendance if a[1]==wid and a[6]=="Per Meter"]
    if not meter_rows: continue
    total_m  = sum(a[7] for a in meter_rows)
    last_ppm = meter_rows[-1][8]
    extras   = sum(a[10] for a in attendance if a[1]==wid)
    earned   = sum(a[11] for a in attendance if a[1]==wid)
    payments = sum(p[3] for p in worker_payments if p[1]==wid)
    meals    = meal_cost_per_worker.get(wid, 0)
    net      = round(earned - payments - meals, 2)
    meter_workers_report.append((w[1], total_m, last_ppm, extras, earned, payments, meals, net))

# Query 12 – building cost
building_costs = []
for b in buildings:
    bid = b[0]
    exp_cost = sum(e[7] for e in general_expenses if e[2]==bid)
    att_cost = sum(a[11] for a in attendance
                   if a[2] is not None
                   and next((l for l in location_details if l[0]==a[2] and l[1]==bid), None))
    building_costs.append((b[2], round(exp_cost+att_cost, 2)))

# ─────────────────────────────────────────────────────────────────────────────
# BUILD WORKBOOK
# ─────────────────────────────────────────────────────────────────────────────
wb = openpyxl.Workbook()

# ── sheet order ───────────────────────────────────────────────────────────────
SHEETS = [
    ("🏠 الرئيسية",         "dashboard"),
    ("📋 المقاولون",        "contractors"),
    ("🏗️ المشاريع",         "projects"),
    ("🏢 العمارات",         "buildings"),
    ("📍 المواقع الداخلية", "locations"),
    ("🧱 الواجهات",         "facades"),
    ("👷 العمال",           "workers"),
    ("📅 الحضور",           "attendance"),
    ("💰 مدفوعات العمال",   "payments"),
    ("🛒 المصروفات",        "expenses"),
    ("🍽️ الوجبات",          "meals"),
    ("👥 المشتركون بالأكل", "meal_parts"),
    ("📥 إيرادات المقاول",  "incomes"),
    ("📊 التقارير",         "reports"),
    ("📝 الكود SQL",        "sql"),
]

# Create named worksheets
ws_dict = {}
first = True
for title, key in SHEETS:
    if first:
        ws = wb.active
        ws.title = title
        first = False
    else:
        ws = wb.create_sheet(title)
    ws_dict[key] = ws
    ws.sheet_view.rightToLeft = True  # RTL for Arabic

# ═══════════════════════════════════════════════════════════════════════════
# 1. DASHBOARD
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["dashboard"]
ws.sheet_properties.tabColor = C["navy"]

merge_title(ws, 1, 1, 8, "🏗️  نظام إدارة المشاريع – شغلي  🏗️", C["navy"], size=18)
row_height(ws, 1, 48)

ws.merge_cells("A2:H2")
c = ws["A2"]
c.value = "لوحة التحكم الرئيسية | إجمالي سريع"
c.fill  = fill(C["blue"])
c.font  = font(C["white"], bold=True, size=13)
c.alignment = center()
row_height(ws, 2, 26)

# KPI cards (row 4-8, 2 columns wide each)
kpis = [
    ("💵 إجمالي الإيرادات",  f"{total_incomes:,.2f} ج.م",     C["teal"],   C["white"]),
    ("🛒 إجمالي المصروفات",  f"{total_expenses:,.2f} ج.م",    C["amber"],  C["white"]),
    ("👷 أجور العمال",       f"{total_earned_att:,.2f} ج.م",  C["purple"], C["white"]),
    ("📈 صافي الربح",        f"{net_profit:,.2f} ج.م",         C["green"],  C["white"]),
    ("🏗️ عدد المشاريع",      str(len(projects)),               C["blue"],   C["white"]),
    ("🏢 عدد العمارات",      str(len(buildings)),              C["navy"],   C["white"]),
    ("👤 عدد العمال",        str(len(workers)),                C["teal"],   C["white"]),
    ("🤝 أكثر عامل مستحق",  f"{most_owed_name} ({most_owed_amt:,.0f}ج)", C["red"], C["white"]),
]

row_height(ws, 3, 10)
kpi_start = 4
for i, (label, value, bg, fg) in enumerate(kpis):
    col_s = (i % 4) * 2 + 1
    r = kpi_start + (i // 4) * 3
    ws.merge_cells(start_row=r, start_column=col_s, end_row=r, end_column=col_s+1)
    c = ws.cell(row=r, column=col_s, value=label)
    c.fill = fill(bg); c.font = font(fg, bold=True, size=11); c.alignment = center()
    row_height(ws, r, 22)

    ws.merge_cells(start_row=r+1, start_column=col_s, end_row=r+1, end_column=col_s+1)
    c = ws.cell(row=r+1, column=col_s, value=value)
    c.fill = fill(C["lgray"]); c.font = font(bg, bold=True, size=14); c.alignment = center()
    row_height(ws, r+1, 30)

row_height(ws, kpi_start+2, 10)
row_height(ws, kpi_start+5, 10)

# Project summary table
r = 12
merge_title(ws, r, 1, 8, "ملخص المشاريع", C["blue"], size=13)
r += 1
write_header_row(ws, r, ["#", "اسم المشروع", "الموقع", "قيمة العقد",
                          "إجمالي المقبوض", "إجمالي المصروفات", "الربح", "المقاول"], C["navy"])
row_height(ws, r, 24)
r += 1
for proj in projects:
    pid   = proj[0]
    cont  = next(c[1] for c in contractors if c[0]==proj[4])
    inc   = proj_incomes.get(pid, 0)
    exp   = proj_expenses.get(pid, 0)
    profit = inc - exp
    row_data = [pid, proj[1], proj[2], f"{proj[3]:,.0f}", f"{inc:,.0f}", f"{exp:,.0f}",
                f"{profit:,.0f}", cont]
    bg = C["lgray"] if r % 2 == 0 else C["white"]
    write_data_row(ws, r, row_data, bg=bg)
    row_height(ws, r, 20)
    r += 1

set_col_widths(ws, [4, 24, 22, 16, 16, 16, 16, 18])
freeze(ws, "A3")

# ═══════════════════════════════════════════════════════════════════════════
# 2. CONTRACTORS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["contractors"]
ws.sheet_properties.tabColor = C["teal"]
merge_title(ws, 1, 1, 4, "👷‍♂️  جدول المقاولين", C["teal"], size=14)
write_header_row(ws, 2, ["كود المقاول", "الاسم", "رقم الهاتف", "تاريخ الإنشاء"], C["teal"])
for i, row in enumerate(contractors):
    bg = C["lteal"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 26, 18, 20])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 3. PROJECTS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["projects"]
ws.sheet_properties.tabColor = C["blue"]
merge_title(ws, 1, 1, 5, "🏗️  جدول المشاريع", C["blue"], size=14)
write_header_row(ws, 2, ["كود المشروع","اسم المشروع","الموقع","قيمة العقد","كود المقاول"], C["blue"])
for i, row in enumerate(projects):
    bg = C["lblue"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 28, 26, 18, 14])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 4. BUILDINGS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["buildings"]
ws.sheet_properties.tabColor = C["navy"]
merge_title(ws, 1, 1, 3, "🏢  جدول العمارات", C["navy"], size=14)
write_header_row(ws, 2, ["كود العمارة", "كود المشروع", "رمز العمارة"], C["navy"])
for i, row in enumerate(buildings):
    bg = C["lgray"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 14, 20])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 5. LOCATION DETAILS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["locations"]
ws.sheet_properties.tabColor = C["purple"]
merge_title(ws, 1, 1, 5, "📍  المواقع الداخلية (دور / سلم / سطح / بدروم)", C["purple"], size=14)
write_header_row(ws, 2, ["كود الموقع","كود العمارة","اسم الموقع","رقم الدور","نوع الموقع"], C["purple"])
for i, row in enumerate(location_details):
    bg = C["lpurple"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 14, 22, 12, 16])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 6. FACADES
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["facades"]
ws.sheet_properties.tabColor = C["amber"]
merge_title(ws, 1, 1, 3, "🧱  الواجهات الخارجية", C["amber"], size=14)
write_header_row(ws, 2, ["كود الواجهة", "كود العمارة", "اتجاه الواجهة"], C["amber"])
for i, row in enumerate(facades):
    bg = C["lamber"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 14, 20])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 7. WORKERS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["workers"]
ws.sheet_properties.tabColor = C["green"]
merge_title(ws, 1, 1, 6, "👷  جدول العمال", C["green"], size=14)
write_header_row(ws, 2,
    ["كود العامل","الاسم","رقم الهاتف","التخصص","اليومية (ج.م)","كود المقاول"],
    C["green"])
for i, row in enumerate(workers):
    bg = C["lgreen"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14, 26, 18, 16, 16, 14])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 8. ATTENDANCE
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["attendance"]
ws.sheet_properties.tabColor = C["blue"]
merge_title(ws, 1, 1, 12, "📅  الحضور واليوميات", C["navy"], size=14)
att_headers = [
    "كود السجل","كود العامل","كود الموقع الداخلي","كود الواجهة",
    "كود موقع إكسترا","التاريخ","نوع الحساب","الكمية (م²)",
    "سعر المتر","الحالة","إضافي (ج.م)","المستحق (ج.م)"
]
write_header_row(ws, 2, att_headers, C["navy"])
status_color = {"Full Day": C["lgreen"], "Half Day": C["lamber"], None: C["lblue"]}
for i, row in enumerate(attendance):
    bg = status_color.get(row[10], C["white"]) if i % 2 == 0 else C["lgray"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [12,12,18,14,16,14,14,12,12,12,14,16])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 9. WORKER PAYMENTS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["payments"]
ws.sheet_properties.tabColor = C["red"]
merge_title(ws, 1, 1, 7, "💰  السلفيات والمدفوعات", C["red"], size=14)
write_header_row(ws, 2,
    ["كود الدفعة","كود العامل","كود المشروع","المبلغ","التاريخ","نوع الدفعة","ملاحظات"],
    C["red"])
for i, row in enumerate(worker_payments):
    bg = C["lred"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14,14,14,14,14,18,28])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 10. GENERAL EXPENSES
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["expenses"]
ws.sheet_properties.tabColor = C["amber"]
merge_title(ws, 1, 1, 10, "🛒  المصروفات العامة", C["amber"], size=14)
write_header_row(ws, 2,
    ["كود المصروف","كود المشروع","كود العمارة","كود الموقع","كود الواجهة",
     "اسم الصنف","التصنيف","التكلفة","التاريخ","ملاحظات"],
    C["amber"])
for i, row in enumerate(general_expenses):
    bg = C["lamber"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14,12,12,12,12,22,16,14,14,24])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 11. SHARED MEALS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["meals"]
ws.sheet_properties.tabColor = C["teal"]
merge_title(ws, 1, 1, 5, "🍽️  الوجبات المشتركة", C["teal"], size=14)
write_header_row(ws, 2,
    ["كود الوجبة","كود المشروع","التاريخ","إجمالي التكلفة","ملاحظات"],
    C["teal"])
for i, row in enumerate(shared_meals):
    bg = C["lteal"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14,14,14,18,28])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 12. MEAL PARTICIPANTS
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["meal_parts"]
ws.sheet_properties.tabColor = C["teal"]
merge_title(ws, 1, 1, 2, "👥  المشتركون في الوجبات", C["teal"], size=14)
write_header_row(ws, 2, ["كود الوجبة", "كود العامل"], C["teal"])
for i, row in enumerate(meal_participants):
    bg = C["lteal"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [16, 16])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 13. CONTRACTOR INCOMES
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["incomes"]
ws.sheet_properties.tabColor = C["green"]
merge_title(ws, 1, 1, 9, "📥  الدفعات المستلمة (إيرادات المقاول)", C["green"], size=14)
write_header_row(ws, 2,
    ["كود الإيراد","كود المقاول","كود المشروع","كود العمارة","كود الموقع",
     "كود الواجهة","المبلغ المستلم","تاريخ القبض","ملاحظات"],
    C["green"])
for i, row in enumerate(contractor_incomes):
    bg = C["lgreen"] if i % 2 == 0 else C["white"]
    write_data_row(ws, 3+i, list(row), bg=bg)
    row_height(ws, 3+i, 20)
set_col_widths(ws, [14,14,14,14,12,12,18,16,28])
freeze(ws)

# ═══════════════════════════════════════════════════════════════════════════
# 14. REPORTS (all queries)
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["reports"]
ws.sheet_properties.tabColor = C["gold"]
merge_title(ws, 1, 1, 10, "📊  التقارير والاستعلامات", C["navy"], size=16)

r = 3

# ── Q1: KPIs ──────────────────────────────────────────────────────────────
merge_title(ws, r, 1, 3, "1️⃣  إجماليات سريعة", C["blue"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["البيان", "القيمة (ج.م)", ""], C["lblue"], fg=C["navy"])
r += 1
kpi_data = [
    ("إجمالي الإيرادات",   f"{total_incomes:,.2f}"),
    ("إجمالي المصروفات",  f"{total_expenses:,.2f}"),
    ("إجمالي أجور العمال", f"{total_earned_att:,.2f}"),
    ("صافي الربح",          f"{net_profit:,.2f}"),
]
for label, val in kpi_data:
    write_data_row(ws, r, [label, val, ""], bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q2: Most owed worker ──────────────────────────────────────────────────
merge_title(ws, r, 1, 3, "2️⃣  أكثر عامل مستحق فلوس", C["red"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["العامل", "إجمالي المستحق (ج.م)", ""], C["lred"], fg=C["red"])
r += 1
write_data_row(ws, r, [most_owed_name, f"{most_owed_amt:,.2f}", ""], bg=C["lred"])
row_height(ws, r, 20); r += 2

# ── Q3: Daily workers report ──────────────────────────────────────────────
merge_title(ws, r, 1, 8, "3️⃣  كشف حساب العمال اليومية", C["green"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r,
    ["العامل","عدد الأيام","فئة اليومية","إجمالي السهرات",
     "إجمالي المستحق","السلف كاش","خصم الأكل","صافي القبض"],
    C["lgreen"], fg=C["green"])
r += 1
for rec in daily_workers_report:
    write_data_row(ws, r, list(rec), bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q4: Per-meter workers report ─────────────────────────────────────────
merge_title(ws, r, 1, 8, "4️⃣  كشف حساب عمال الأمتار", C["purple"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r,
    ["العامل","إجمالي الأمتار","آخر سعر متر","إجمالي السهرات",
     "إجمالي المستحق","السلف كاش","خصم الأكل","صافي القبض"],
    C["lpurple"], fg=C["purple"])
r += 1
for rec in meter_workers_report:
    write_data_row(ws, r, list(rec), bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q5: Project summary ───────────────────────────────────────────────────
merge_title(ws, r, 1, 6, "5️⃣  ملخص المشاريع", C["teal"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r,
    ["المشروع","قيمة العقد","إجمالي المقبوض","إجمالي المصروفات","الربح","المقاول"],
    C["lteal"], fg=C["teal"])
r += 1
for proj in projects:
    pid   = proj[0]
    cont  = next(c[1] for c in contractors if c[0]==proj[4])
    inc   = proj_incomes.get(pid, 0)
    exp   = proj_expenses.get(pid, 0)
    profit = inc - exp
    write_data_row(ws, r,
        [proj[1], f"{proj[3]:,.0f}", f"{inc:,.0f}", f"{exp:,.0f}", f"{profit:,.0f}", cont],
        bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q6: Buildings per project ─────────────────────────────────────────────
merge_title(ws, r, 1, 3, "6️⃣  عدد العمارات لكل مشروع", C["navy"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["المشروع","عدد العمارات",""], C["lgray"], fg=C["navy"])
r += 1
for proj in projects:
    cnt = bld_per_proj.get(proj[0], 0)
    write_data_row(ws, r, [proj[1], cnt, ""], bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q7: Meal cost per worker ──────────────────────────────────────────────
merge_title(ws, r, 1, 3, "7️⃣  تكلفة الأكل لكل عامل", C["amber"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["العامل","تكلفة الأكل (ج.م)",""], C["lamber"], fg=C["amber"])
r += 1
for w in workers:
    cost = meal_cost_per_worker.get(w[0], 0)
    if cost:
        write_data_row(ws, r, [w[1], f"{cost:,.2f}", ""],
                       bg=C["lgray"] if r%2==0 else C["white"])
        row_height(ws, r, 20); r += 1

r += 1
# ── Q8: Building cost ─────────────────────────────────────────────────────
merge_title(ws, r, 1, 3, "8️⃣  تكلفة كل عمارة", C["red"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["رمز العمارة","إجمالي التكلفة (ج.م)",""], C["lred"], fg=C["red"])
r += 1
for bc in building_costs:
    write_data_row(ws, r, [bc[0], f"{bc[1]:,.2f}", ""],
                   bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q9: Workers present today ────────────────────────────────────────────
merge_title(ws, r, 1, 4, "9️⃣  العمال الحاضرون اليوم", C["green"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["العامل","الحالة","نوع الحساب","المستحق (ج.م)"], C["lgreen"], fg=C["green"])
r += 1
for rec in q_present:
    write_data_row(ws, r, list(rec), bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q10: Absent workers ───────────────────────────────────────────────────
merge_title(ws, r, 1, 2, "🔟  العمال الغائبون اليوم", C["red"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["العامل","ملاحظة"], C["lred"], fg=C["red"])
r += 1
if q_absent:
    for rec in q_absent:
        write_data_row(ws, r, [rec, ""], bg=C["lred"])
        row_height(ws, r, 20); r += 1
else:
    write_data_row(ws, r, ["لا يوجد غياب اليوم 🎉", ""], bg=C["lgreen"])
    row_height(ws, r, 20); r += 1

r += 1
# ── Q11: Worker debts (View_Worker_Debts) ────────────────────────────────
merge_title(ws, r, 1, 4, "1️⃣1️⃣  ديون العمال (سلف + أكل)", C["navy"], fg=C["white"], size=12)
r += 1
write_header_row(ws, r, ["العامل","سلف كاش (ج.م)","ديون أكل (ج.م)","إجمالي الديون"],
                 C["lgray"], fg=C["navy"])
r += 1
for w in workers:
    wid   = w[0]
    cash  = sum(p[3] for p in worker_payments if p[1]==wid)
    meals_d = meal_cost_per_worker.get(wid, 0)
    write_data_row(ws, r, [w[1], f"{cash:,.2f}", f"{meals_d:,.2f}",
                            f"{cash+meals_d:,.2f}"],
                   bg=C["lgray"] if r%2==0 else C["white"])
    row_height(ws, r, 20); r += 1

set_col_widths(ws, [28, 18, 18, 18, 18, 18, 18, 18, 18, 18])
freeze(ws, "A3")

# ── Add a bar chart for project profits ───────────────────────────────────
# find the project summary block to build a chart
# We'll put the chart next to the KPIs area
chart_ws = ws  # same sheet
bar = BarChart()
bar.type        = "col"
bar.grouping    = "clustered"
bar.title       = "ربح المشاريع"
bar.y_axis.title = "الربح (ج.م)"
bar.x_axis.title = "المشروع"
bar.width   = 18
bar.height  = 12
bar.style   = 10

# We need data – write a mini helper table on the same sheet
helper_start_col = 12
helper_row = 3
ws.cell(row=helper_row, column=helper_start_col, value="المشروع")
ws.cell(row=helper_row, column=helper_start_col+1, value="الربح")
for i, proj in enumerate(projects):
    pid    = proj[0]
    profit = proj_incomes.get(pid,0) - proj_expenses.get(pid,0)
    ws.cell(row=helper_row+1+i, column=helper_start_col,   value=proj[1])
    ws.cell(row=helper_row+1+i, column=helper_start_col+1, value=profit)

data_ref = Reference(ws,
    min_col=helper_start_col+1,
    min_row=helper_row,
    max_row=helper_row+len(projects))
cats_ref = Reference(ws,
    min_col=helper_start_col,
    min_row=helper_row+1,
    max_row=helper_row+len(projects))
bar.add_data(data_ref, titles_from_data=True)
bar.set_categories(cats_ref)
ws.add_chart(bar, "L10")

# ═══════════════════════════════════════════════════════════════════════════
# 15. SQL CODE SHEET
# ═══════════════════════════════════════════════════════════════════════════
ws = ws_dict["sql"]
ws.sheet_properties.tabColor = C["gray"]
ws.sheet_view.rightToLeft = False  # LTR for SQL code

merge_title(ws, 1, 1, 3, "📝  كود SQL  –  إنشاء قاعدة البيانات", C["navy"], size=14)

sql_code = """\
-- 1. إنشاء قاعدة البيانات
CREATE DATABASE IF NOT EXISTS Shoghly;
USE Shoghly;

-- 2. جدول المقاولين
CREATE TABLE Contractors (
    contractor_id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    phone VARCHAR(20) UNIQUE,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- 3. جدول المشاريع
CREATE TABLE Projects (
    project_id INT AUTO_INCREMENT PRIMARY KEY,
    project_name VARCHAR(150) NOT NULL,
    location VARCHAR(255),
    contract_amount DECIMAL(15, 2),
    contractor_id INT,
    FOREIGN KEY (contractor_id) REFERENCES Contractors(contractor_id)
);

-- 4. جدول العمارات
CREATE TABLE Buildings (
    building_id INT AUTO_INCREMENT PRIMARY KEY,
    project_id INT NOT NULL,
    building_code VARCHAR(50) NOT NULL,
    FOREIGN KEY (project_id) REFERENCES Projects(project_id),
    CONSTRAINT UN_build_proj UNIQUE(project_id, building_code)
);

-- 5. جدول المواقع الداخلية
CREATE TABLE Location_Details (
    location_id INT AUTO_INCREMENT PRIMARY KEY,
    building_id INT NOT NULL,
    location_name VARCHAR(100) NOT NULL,
    floor_number INT,
    location_type ENUM('Floor','Stairs','Roof','Basement','Shaft') DEFAULT 'Floor',
    FOREIGN KEY (building_id) REFERENCES Buildings(building_id)
);

-- 6. جدول الواجهات
CREATE TABLE Facades (
    facade_id INT AUTO_INCREMENT PRIMARY KEY,
    building_id INT NOT NULL,
    facade_direction ENUM('Front','Back','Left Side','Right Side') NOT NULL,
    FOREIGN KEY (building_id) REFERENCES Buildings(building_id)
);

-- 7. جدول العمال
CREATE TABLE Workers (
    worker_id INT AUTO_INCREMENT PRIMARY KEY,
    name VARCHAR(100) NOT NULL,
    phone VARCHAR(20) UNIQUE,
    specialty VARCHAR(50),
    daily_rate DECIMAL(10, 2) NOT NULL,
    contractor_id INT,
    FOREIGN KEY (contractor_id) REFERENCES Contractors(contractor_id)
);

-- 8. جدول الحضور
CREATE TABLE Attendance (
    attendance_id INT AUTO_INCREMENT PRIMARY KEY,
    worker_id INT NOT NULL,
    location_id INT NULL,
    facade_id INT NULL,
    extra_location_id INT NULL,
    date DATE NOT NULL,
    pay_type ENUM('Daily','Per Meter') DEFAULT 'Daily',
    quantity DECIMAL(10, 2) DEFAULT 0,
    price_per_meter DECIMAL(10, 2) DEFAULT 0,
    status ENUM('Full Day','Half Day','Absent') DEFAULT 'Full Day',
    extra_amount DECIMAL(10, 2) DEFAULT 0.00,
    CONSTRAINT chk_pay_logic CHECK (
        (pay_type='Daily' AND status IS NOT NULL AND quantity=0) OR
        (pay_type='Per Meter' AND quantity>0)
    ),
    earned_amount DECIMAL(10, 2),
    CONSTRAINT chk_location_or_facade CHECK (
        (location_id IS NOT NULL AND facade_id IS NULL) OR
        (location_id IS NULL AND facade_id IS NOT NULL) OR
        (location_id IS NULL AND facade_id IS NULL)
    ),
    CONSTRAINT chk_extra_logic CHECK (
        (extra_amount > 0 AND extra_location_id IS NOT NULL) OR
        (extra_amount <= 0 AND extra_location_id IS NULL)
    ),
    UNIQUE(worker_id, date),
    FOREIGN KEY (worker_id) REFERENCES Workers(worker_id),
    FOREIGN KEY (location_id) REFERENCES Location_Details(location_id),
    FOREIGN KEY (facade_id) REFERENCES Facades(facade_id),
    FOREIGN KEY (extra_location_id) REFERENCES Location_Details(location_id)
);

-- 9. جدول المدفوعات
CREATE TABLE Worker_Payments (
    payment_id INT AUTO_INCREMENT PRIMARY KEY,
    worker_id INT NOT NULL,
    project_id INT,
    amount DECIMAL(10, 2) NOT NULL,
    date DATE NOT NULL,
    payment_type ENUM('Advance','Full Salary','Family Transfer','Private') DEFAULT 'Private',
    notes TEXT,
    FOREIGN KEY (worker_id) REFERENCES Workers(worker_id),
    FOREIGN KEY (project_id) REFERENCES Projects(project_id)
);

-- 10. جدول المصروفات العامة
CREATE TABLE General_Expenses (
    expense_id INT AUTO_INCREMENT PRIMARY KEY,
    project_id INT NOT NULL,
    building_id INT NULL,
    location_id INT NULL,
    facade_id INT NULL,
    item_name VARCHAR(150) NOT NULL,
    category ENUM('Tools','Materials','Transport','Other') NOT NULL,
    cost DECIMAL(12, 2) NOT NULL,
    date DATE NOT NULL,
    notes TEXT,
    FOREIGN KEY (project_id) REFERENCES Projects(project_id),
    FOREIGN KEY (building_id) REFERENCES Buildings(building_id),
    FOREIGN KEY (location_id) REFERENCES Location_Details(location_id),
    FOREIGN KEY (facade_id) REFERENCES Facades(facade_id)
);

-- 11. جدول الوجبات
CREATE TABLE Shared_Meals (
    meal_id INT AUTO_INCREMENT PRIMARY KEY,
    project_id INT NOT NULL,
    date DATE NOT NULL,
    total_cost DECIMAL(10, 2) NOT NULL,
    notes VARCHAR(255),
    FOREIGN KEY (project_id) REFERENCES Projects(project_id)
);

-- 12. جدول المشتركين بالأكل
CREATE TABLE Meal_Participants (
    meal_id INT,
    worker_id INT,
    PRIMARY KEY (meal_id, worker_id),
    FOREIGN KEY (meal_id) REFERENCES Shared_Meals(meal_id),
    FOREIGN KEY (worker_id) REFERENCES Workers(worker_id)
);

-- 13. جدول إيرادات المقاول
CREATE TABLE Contractor_Incomes (
    income_id INT AUTO_INCREMENT PRIMARY KEY,
    contractor_id INT NOT NULL,
    project_id INT NOT NULL,
    building_id INT NULL,
    location_id INT NULL,
    facade_id INT NULL,
    amount_received DECIMAL(15, 2) NOT NULL,
    payment_date DATE NOT NULL,
    notes TEXT,
    FOREIGN KEY (contractor_id) REFERENCES Contractors(contractor_id),
    FOREIGN KEY (project_id) REFERENCES Projects(project_id),
    FOREIGN KEY (building_id) REFERENCES Buildings(building_id),
    FOREIGN KEY (location_id) REFERENCES Location_Details(location_id),
    FOREIGN KEY (facade_id) REFERENCES Facades(facade_id)
);

-- Triggers
DELIMITER //
CREATE TRIGGER trg_calculate_attendance_final
BEFORE INSERT ON Attendance FOR EACH ROW
BEGIN
  DECLARE v_daily_rate DECIMAL(10,2);
  DECLARE v_base_pay DECIMAL(10,2) DEFAULT 0;
  IF NEW.pay_type = 'Daily' THEN
    SELECT daily_rate INTO v_daily_rate FROM Workers WHERE worker_id = NEW.worker_id;
    IF NEW.status = 'Full Day' THEN SET v_base_pay = v_daily_rate;
    ELSEIF NEW.status = 'Half Day' THEN SET v_base_pay = v_daily_rate * 0.5;
    END IF;
  ELSEIF NEW.pay_type = 'Per Meter' THEN
    SET v_base_pay = IFNULL(NEW.quantity,0) * IFNULL(NEW.price_per_meter,0);
  END IF;
  SET NEW.earned_amount = v_base_pay + IFNULL(NEW.extra_amount,0);
END; //
DELIMITER ;"""

code_font = Font(name="Courier New", size=9, color=C["black"])
code_fill_even = fill(C["lgray"])
code_fill_odd  = fill(C["white"])
for i, line in enumerate(sql_code.split("\n")):
    r = 3 + i
    c = ws.cell(row=r, column=1, value=line)
    c.font      = code_font
    c.fill      = code_fill_even if i % 2 == 0 else code_fill_odd
    c.alignment = Alignment(horizontal="left", vertical="center",
                             wrap_text=False, readingOrder=1)
    row_height(ws, r, 15)

ws.column_dimensions["A"].width = 100

# ─────────────────────────────────────────────────────────────────────────────
# GLOBAL TAB COLORS already set above per sheet.
# Set print area and zoom for all sheets
# ─────────────────────────────────────────────────────────────────────────────
for ws in wb.worksheets:
    ws.sheet_view.zoomScale = 90

# ─────────────────────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────────────────────
output_path = "/home/runner/work/Ahmed-M-Khedr/Ahmed-M-Khedr/Shoghly.xlsx"
wb.save(output_path)
print(f"✅  Saved: {output_path}")
