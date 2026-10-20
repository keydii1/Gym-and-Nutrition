import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import LineChart, BarChart, Reference, Series
import datetime

wb = openpyxl.load_workbook('Personal_Bulking_Tracker_2026.xlsx')

# Color & Style tokens
NAVY_HEADER = '0F172A'  # Slate 900
BLUE_ACCENT = '2563EB'  # Blue 600
LIGHT_BLUE_FILL = 'EFF6FF' # Blue 50
NEUTRAL_GRAY_FILL = 'F8FAFC' # Slate 50
BORDER_COLOR = 'CBD5E1' # Slate 300
EMERALD_GREEN = '10B981' # Green 500
AMBER_WARNING = 'F59E0B' # Amber 500
PURPLE_SAMPLE = '8B5CF6' # Purple 500

thin_border = Border(
    left=Side(style='thin', color=BORDER_COLOR),
    right=Side(style='thin', color=BORDER_COLOR),
    top=Side(style='thin', color=BORDER_COLOR),
    bottom=Side(style='thin', color=BORDER_COLOR)
)

header_font = Font(name='Segoe UI', size=11, bold=True, color='FFFFFF')
header_fill = PatternFill(fill_type='solid', start_color='1E293B', end_color='1E293B')
formula_fill = PatternFill(fill_type='solid', start_color='F1F5F9', end_color='F1F5F9')
input_fill = PatternFill(fill_type='solid', start_color='F0FDF4', end_color='F0FDF4')
bold_font = Font(name='Segoe UI', size=10, bold=True)
regular_font = Font(name='Segoe UI', size=10)

print("Starting fixes across workbook...")

# ==========================================
# 1. SHEET 01: PROFILE_GOALS
# ==========================================
ws01 = wb['01_PROFILE_GOALS']
# Fix BMI Baseline formula in C10
# C8 is height (175), C9 is weight (50.3)
ws01['C10'] = "=C9/((C8/100)^2)"
ws01['C10'].number_format = '0.0'

# Fix Current Weight formula in C38: dynamically get latest morning weight
ws01['C38'] = "=LOOKUP(9.99E+307, '10_BODY_WEIGHT_LOG'!$C$4:$C$171)"
ws01['C38'].number_format = '0.0'

# Fix formulas referencing C38 (was C37 which was a text header!)
ws01['C39'] = "=MAX(0, C38-50.3)"
ws01['C39'].number_format = '0.0'
ws01['C40'] = "=MAX(0, 55-C38)"
ws01['C40'].number_format = '0.0'
ws01['C41'] = "=MAX(0, 60-C38)"
ws01['C41'].number_format = '0.0'
ws01['C42'] = "=MAX(0, 65-C38)"
ws01['C42'].number_format = '0.0'
ws01['C43'] = "=IF(C38<=50.3, 0, (C38-50.3)/(65-50.3))"
ws01['C43'].number_format = '0.0%'

print("Sheet 01 fixed.")

# ==========================================
# 2. SHEET 02: NUTRITION_SETUP
# ==========================================
ws02 = wb['02_NUTRITION_SETUP']
# Fix Katch-McArdle formula in C7 referencing FFM from Sheet 11 (K4 is 46.4)
ws02['C7'] = "=370+(21.6*'11_INBODY_HISTORY'!K4)"
ws02['C7'].number_format = '#,##0'

# Fix Macros check formula in C24: Protein*4 + Fat*9 + Carb*4
# C21 is Protein (100), C22 is Fat (55), C23 is Carb (300)
ws02['C24'] = "=(C21*4)+(C22*9)+(C23*4)"
ws02['C24'].number_format = '#,##0'

print("Sheet 02 fixed.")

# ==========================================
# 3. SHEET 06: DAILY_NUTRITION_LOG (Expand Daily Summary to 24 Weeks = 168 Days)
# ==========================================
ws06 = wb['06_DAILY_NUTRITION_LOG']

# We need rows N5 to N172 for all 168 days (from 2026-10-19 to 2027-04-04)
start_date = datetime.date(2026, 10, 19)

for day_idx in range(168):
    row_num = 5 + day_idx
    curr_date = start_date + datetime.timedelta(days=day_idx)
    date_val = datetime.datetime(curr_date.year, curr_date.month, curr_date.day)
    
    # Col N: Ngày Tổng Hợp
    ws06.cell(row_num, 14, date_val)
    ws06.cell(row_num, 14).number_format = 'yyyy-mm-dd'
    ws06.cell(row_num, 14).alignment = Alignment(horizontal='center')
    ws06.cell(row_num, 14).border = thin_border
    
    # Col O: Thứ
    ws06.cell(row_num, 15, f'=TEXT(N{row_num}, "dddd")')
    ws06.cell(row_num, 15).alignment = Alignment(horizontal='center')
    ws06.cell(row_num, 15).border = thin_border
    
    # Col P: Tổng Kcal
    ws06.cell(row_num, 16, f'=IF(U{row_num}=0, "", SUMIFS($F$5:$F$1500, $A$5:$A$1500, N{row_num}))')
    ws06.cell(row_num, 16).number_format = '#,##0'
    ws06.cell(row_num, 16).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 16).fill = formula_fill
    ws06.cell(row_num, 16).border = thin_border
    
    # Col Q: Tổng Đạm (g)
    ws06.cell(row_num, 17, f'=IF(U{row_num}=0, "", SUMIFS($G$5:$G$1500, $A$5:$A$1500, N{row_num}))')
    ws06.cell(row_num, 17).number_format = '0.0'
    ws06.cell(row_num, 17).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 17).fill = formula_fill
    ws06.cell(row_num, 17).border = thin_border
    
    # Col R: Tổng Carb (g)
    ws06.cell(row_num, 18, f'=IF(U{row_num}=0, "", SUMIFS($H$5:$H$1500, $A$5:$A$1500, N{row_num}))')
    ws06.cell(row_num, 18).number_format = '0.0'
    ws06.cell(row_num, 18).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 18).fill = formula_fill
    ws06.cell(row_num, 18).border = thin_border
    
    # Col S: Tổng Fat (g)
    ws06.cell(row_num, 19, f'=IF(U{row_num}=0, "", SUMIFS($I$5:$I$1500, $A$5:$A$1500, N{row_num}))')
    ws06.cell(row_num, 19).number_format = '0.0'
    ws06.cell(row_num, 19).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 19).fill = formula_fill
    ws06.cell(row_num, 19).border = thin_border
    
    # Col T: Tổng Chi Phí
    ws06.cell(row_num, 20, f'=IF(U{row_num}=0, "", SUMIFS($J$5:$J$1500, $A$5:$A$1500, N{row_num}))')
    ws06.cell(row_num, 20).number_format = '#,##0'
    ws06.cell(row_num, 20).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 20).fill = formula_fill
    ws06.cell(row_num, 20).border = thin_border
    
    # Col U: Số Bữa
    ws06.cell(row_num, 21, f'=COUNTIFS($A$5:$A$1500, N{row_num})')
    ws06.cell(row_num, 21).number_format = '0'
    ws06.cell(row_num, 21).alignment = Alignment(horizontal='center')
    ws06.cell(row_num, 21).fill = formula_fill
    ws06.cell(row_num, 21).border = thin_border
    
    # Col V: Trạng Thái Ngày
    ws06.cell(row_num, 22, f'=IF(U{row_num}=0, "CHƯA NHẬP", IF(U{row_num}<3, "INCOMPLETE (Chưa đủ bữa)", "COMPLETE (Đầy đủ)"))')
    ws06.cell(row_num, 22).alignment = Alignment(horizontal='center')
    ws06.cell(row_num, 22).fill = formula_fill
    ws06.cell(row_num, 22).border = thin_border
    
    # Col W: % Đạt Kcal
    ws06.cell(row_num, 23, f'=IF(P{row_num}="", "", P{row_num}/\'02_NUTRITION_SETUP\'!$C$20)')
    ws06.cell(row_num, 23).number_format = '0.0%'
    ws06.cell(row_num, 23).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 23).fill = formula_fill
    ws06.cell(row_num, 23).border = thin_border
    
    # Col X: % Đạt Protein
    ws06.cell(row_num, 24, f'=IF(Q{row_num}="", "", Q{row_num}/\'02_NUTRITION_SETUP\'!$C$21)')
    ws06.cell(row_num, 24).number_format = '0.0%'
    ws06.cell(row_num, 24).alignment = Alignment(horizontal='right')
    ws06.cell(row_num, 24).fill = formula_fill
    ws06.cell(row_num, 24).border = thin_border

print(f"Sheet 06 Daily Summary expanded to row 172 (168 days).")

# Also ensure meal log formulas exist for rows 17 to 300 so user can enter new meals immediately
for r in range(17, 301):
    ws06.cell(r, 6, f'=IF(OR(C{r}="", D{r}=""), "", IFERROR(VLOOKUP(C{r}, \'03_FOOD_DATABASE\'!$B$5:$L$60, 6, FALSE)*D{r}/100, IFERROR(VLOOKUP(C{r}, \'04_RECIPE_DATABASE\'!$B$5:$M$20, 8, FALSE)*D{r}, 0)))')
    ws06.cell(r, 7, f'=IF(OR(C{r}="", D{r}=""), "", IFERROR(VLOOKUP(C{r}, \'03_FOOD_DATABASE\'!$B$5:$L$60, 7, FALSE)*D{r}/100, IFERROR(VLOOKUP(C{r}, \'04_RECIPE_DATABASE\'!$B$5:$M$20, 9, FALSE)*D{r}, 0)))')
    ws06.cell(r, 8, f'=IF(OR(C{r}="", D{r}=""), "", IFERROR(VLOOKUP(C{r}, \'03_FOOD_DATABASE\'!$B$5:$L$60, 8, FALSE)*D{r}/100, IFERROR(VLOOKUP(C{r}, \'04_RECIPE_DATABASE\'!$B$5:$M$20, 10, FALSE)*D{r}, 0)))')
    ws06.cell(r, 9, f'=IF(OR(C{r}="", D{r}=""), "", IFERROR(VLOOKUP(C{r}, \'03_FOOD_DATABASE\'!$B$5:$L$60, 9, FALSE)*D{r}/100, IFERROR(VLOOKUP(C{r}, \'04_RECIPE_DATABASE\'!$B$5:$M$20, 11, FALSE)*D{r}, 0)))')
    ws06.cell(r, 10, f'=IF(OR(C{r}="", D{r}=""), "", IFERROR(VLOOKUP(C{r}, \'03_FOOD_DATABASE\'!$B$5:$L$60, 11, FALSE)*D{r}/100, IFERROR(VLOOKUP(C{r}, \'04_RECIPE_DATABASE\'!$B$5:$M$20, 12, FALSE)*D{r}, 0)))')
    for c in range(1, 13):
        ws06.cell(r, c).border = thin_border
    for c in [6, 10]:
        ws06.cell(r, c).number_format = '#,##0'
        ws06.cell(r, c).fill = formula_fill
    for c in [7, 8, 9]:
        ws06.cell(r, c).number_format = '0.0'
        ws06.cell(r, c).fill = formula_fill

print("Sheet 06 meal log rows populated up to row 300.")

# ==========================================
# 4. SHEET 09: WORKOUT_LOG (Add Session Summary Table)
# ==========================================
ws09 = wb['09_WORKOUT_LOG']

# Add Session Summary Table headers in N3:R3
ws09['N3'] = "Ngày Tập"
ws09['O3'] = "Buổi Tập"
ws09['P3'] = "Số Sets Hoàn Thành"
ws09['Q3'] = "Tổng Volume (kg)"
ws09['R3'] = "Trạng Thái Buổi Tập"

for c in range(14, 19):
    cell = ws09.cell(3, c)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal='center', vertical='center')
    cell.border = thin_border

# Generate 96 session rows (24 weeks * 4 sessions/week)
session_names = ['Upper A', 'Lower A', 'Upper B', 'Lower B']
# Weekly workout days: Mon (0), Tue (1), Thu (3), Fri (4)
workout_day_offsets = [0, 1, 3, 4]

session_row = 4
for w in range(24):
    week_start = start_date + datetime.timedelta(days=w*7)
    for s_idx in range(4):
        s_date = week_start + datetime.timedelta(days=workout_day_offsets[s_idx])
        s_name = session_names[s_idx]
        d_val = datetime.datetime(s_date.year, s_date.month, s_date.day)
        
        ws09.cell(session_row, 14, d_val)
        ws09.cell(session_row, 14).number_format = 'yyyy-mm-dd'
        ws09.cell(session_row, 14).alignment = Alignment(horizontal='center')
        ws09.cell(session_row, 14).border = thin_border
        
        ws09.cell(session_row, 15, s_name)
        ws09.cell(session_row, 15).alignment = Alignment(horizontal='center')
        ws09.cell(session_row, 15).border = thin_border
        
        # P: Completed sets count
        ws09.cell(session_row, 16, f'=COUNTIFS($A$4:$A$1500, N{session_row}, $I$4:$I$1500, "Đạt")')
        ws09.cell(session_row, 16).number_format = '0'
        ws09.cell(session_row, 16).alignment = Alignment(horizontal='center')
        ws09.cell(session_row, 16).fill = formula_fill
        ws09.cell(session_row, 16).border = thin_border
        
        # Q: Total volume
        ws09.cell(session_row, 17, f'=SUMIFS($K$4:$K$1500, $A$4:$A$1500, N{session_row})')
        ws09.cell(session_row, 17).number_format = '#,##0'
        ws09.cell(session_row, 17).alignment = Alignment(horizontal='right')
        ws09.cell(session_row, 17).fill = formula_fill
        ws09.cell(session_row, 17).border = thin_border
        
        # R: Status
        ws09.cell(session_row, 18, f'=IF(P{session_row}=0, "CHƯA TẬP", IF(P{session_row}>=10, "Hoàn thành đầy đủ", "Hoàn thành 1 phần"))')
        ws09.cell(session_row, 18).alignment = Alignment(horizontal='center')
        ws09.cell(session_row, 18).fill = formula_fill
        ws09.cell(session_row, 18).border = thin_border
        
        session_row += 1

# Populate empty set formulas up to row 300 for user workout entry
for r in range(16, 301):
    ws09.cell(r, 11, f'=IF(OR(E{r}="", F{r}=""), "", E{r}*F{r})')
    ws09.cell(r, 11).number_format = '#,##0'
    ws09.cell(r, 11).fill = formula_fill
    for c in range(1, 13):
        ws09.cell(r, c).border = thin_border

ws09.column_dimensions['N'].width = 14
ws09.column_dimensions['O'].width = 14
ws09.column_dimensions['P'].width = 18
ws09.column_dimensions['Q'].width = 18
ws09.column_dimensions['R'].width = 20

print("Sheet 09 Session Summary table added.")

# ==========================================
# 5. SHEET 12: WEEKLY_REVIEW (Update references to 24 weeks of Daily Summary & Workouts)
# ==========================================
ws12 = wb['12_WEEKLY_REVIEW']

for r in range(4, 28): # Weeks 1 to 24
    # F: Ngày Ghi Đủ (look up in full 168 days $N$5:$N$172)
    ws12.cell(r, 6, f'=COUNTIFS(\'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, ">="&B{r}, \'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, "<="&C{r}, \'06_DAILY_NUTRITION_LOG\'!$V$5:$V$172, "COMPLETE*")')
    
    # G: Calo TB ngày đủ
    ws12.cell(r, 7, f'=IFERROR(AVERAGEIFS(\'06_DAILY_NUTRITION_LOG\'!$P$5:$P$172, \'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, ">="&B{r}, \'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, "<="&C{r}, \'06_DAILY_NUTRITION_LOG\'!$V$5:$V$172, "COMPLETE*"), "")')
    ws12.cell(r, 7).number_format = '#,##0'
    
    # H: Protein TB ngày đủ
    ws12.cell(r, 8, f'=IFERROR(AVERAGEIFS(\'06_DAILY_NUTRITION_LOG\'!$Q$5:$Q$172, \'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, ">="&B{r}, \'06_DAILY_NUTRITION_LOG\'!$N$5:$N$172, "<="&C{r}, \'06_DAILY_NUTRITION_LOG\'!$V$5:$V$172, "COMPLETE*"), "")')
    ws12.cell(r, 8).number_format = '0.0'
    
    # L: Buổi Hoàn Thành (look up in Session Summary table in Sheet 09: $N$4:$N$100)
    ws12.cell(r, 12, f'=COUNTIFS(\'09_WORKOUT_LOG\'!$N$4:$N$100, ">="&B{r}, \'09_WORKOUT_LOG\'!$N$4:$N$100, "<="&C{r}, \'09_WORKOUT_LOG\'!$R$4:$R$100, "Hoàn thành*")')
    ws12.cell(r, 12).number_format = '0'
    
    # M: % Hoàn Thành Tập
    ws12.cell(r, 13, f'=IF(K{r}="", "", L{r}/K{r})')
    ws12.cell(r, 13).number_format = '0.0%'

print("Sheet 12 formulas updated for all 24 weeks.")

# ==========================================
# 6. SHEET 13: DASHBOARD (Fix KPI Formulas & Add Charts)
# ==========================================
ws13 = wb['13_DASHBOARD']

# Fix KPI Top Row (Merged cells reference fix!)
# A5 (Current Weight)
ws13['A5'] = "=LOOKUP(9.99E+307, '10_BODY_WEIGHT_LOG'!$C$4:$C$171)"
ws13['A5'].number_format = '0.0'

# D5 (Goal Weight)
ws13['D5'] = 65
ws13['D5'].number_format = '0.0'

# G5 (Weight Gained): =MAX(0, A5-50.3)
ws13['G5'] = "=MAX(0, A5-50.3)"
ws13['G5'].number_format = '+0.0'

# J5 (Remaining Weight): =MAX(0, D5-A5)
ws13['J5'] = "=MAX(0, D5-A5)"
ws13['J5'].number_format = '0.0'

# M5 (7-day MA Weight): =IFERROR(LOOKUP(9.99E+307, '10_BODY_WEIGHT_LOG'!$L$4:$L$171), "CHƯA ĐỦ DỮ LIỆU")
ws13['M5'] = "=IFERROR(LOOKUP(9.99E+307, '10_BODY_WEIGHT_LOG'!$L$4:$L$171), \"CHƯA ĐỦ DỮ LIỆU\")"
ws13['M5'].number_format = '0.0'

# P5 (Progress %): =IF(A5<=50.3, 0, (A5-50.3)/(D5-50.3))
ws13['P5'] = "=IF(A5<=50.3, 0, (A5-50.3)/(D5-50.3))"
ws13['P5'].number_format = '0.0%'

# Row 9 (Nutrition & Habit KPIs)
# A9 (Calo TB tuần hiện tại)
ws13['A9'] = "=IFERROR('12_WEEKLY_REVIEW'!G4, \"Chờ dữ liệu\")"
ws13['A9'].number_format = '#,##0'

# D9 (Protein TB tuần hiện tại)
ws13['D9'] = "=IFERROR('12_WEEKLY_REVIEW'!H4, \"Chờ dữ liệu\")"
ws13['D9'].number_format = '0.0'

# G9 (Tuân thủ calo)
ws13['G9'] = "=IFERROR('12_WEEKLY_REVIEW'!I4, 0)"
ws13['G9'].number_format = '0.0%'

# J9 (Tuân thủ Protein)
ws13['J9'] = "=IFERROR('12_WEEKLY_REVIEW'!J4, 0)"
ws13['J9'].number_format = '0.0%'

# M9 (Hoàn thành tập luyện)
ws13['M9'] = "=IFERROR('12_WEEKLY_REVIEW'!M4, 0)"
ws13['M9'].number_format = '0.0%'

# P9 (Giờ ngủ TB)
ws13['P9'] = "=IFERROR('12_WEEKLY_REVIEW'!N4, 7.5)"
ws13['P9'].number_format = '0.0'

print("Sheet 13 KPI cards fixed.")

# Let's inspect and enhance charts on Dashboard
# Chart 1: Weight & 7-day MA Line Chart
# Let's add series for 7-day MA (column L) to Chart 1 if not already present
ch1 = ws13._charts[0]
ch1.title = "Cân Nặng Buổi Sáng & Đường Trung Bình 7 Ngày (kg)"
ch1.style = 13
ch1.width = 16
ch1.height = 9.5
ch1.y_axis.title = "Cân Nặng (kg)"
ch1.x_axis.title = "Ngày Theo Dõi"

# Add second series for 7-day MA
# Series 1 is C4:C24 (Daily Weight)
# Let's see if we can add L4:L24 as Series 2
ref_ma = Reference(wb['10_BODY_WEIGHT_LOG'], min_col=12, min_row=3, max_row=24) # Col L
ch1.add_data(ref_ma, titles_from_data=True)

# Chart 2: Milestone Target Comparison (BarChart)
ch2 = ws13._charts[1]
ch2.title = "Lộ Trình Các Cột Mốc Cân Nặng Mục Tiêu (kg)"
ch2.style = 10
ch2.width = 16
ch2.height = 9.5
ch2.y_axis.title = "Mục Tiêu (kg)"

# Chart 3: Weekly Planned vs Completed Workouts (BarChart)
ch3 = ws13._charts[2]
ch3.title = "Số Buổi Tập Dự Kiến vs Hoàn Thành Thực Tế (Theo Tuần)"
ch3.style = 11
ch3.width = 16
ch3.height = 9.5
ch3.y_axis.title = "Số Buổi Tập"
ch3.x_axis.title = "Tuần"

# Chart 4: Daily Energy Intake vs Target (LineChart)
ch4 = ws13._charts[3]
ch4.title = "Năng Lượng Ăn Vào Hằng Ngày vs Mục Tiêu (2.100 kcal)"
ch4.style = 13
ch4.width = 16
ch4.height = 9.5
ch4.y_axis.title = "Năng Lượng (kcal)"
ch4.x_axis.title = "Ngày"

# Add Chart 5: Daily Protein vs Target (100g)
ch5 = LineChart()
ch5.title = "Protein Ăn Vào Hằng Ngày vs Mục Tiêu (100.0 g/ngày)"
ch5.style = 13
ch5.width = 16
ch5.height = 9.5
ch5.y_axis.title = "Protein (g)"
ch5.x_axis.title = "Ngày"

ref_protein_data = Reference(ws06, min_col=17, min_row=4, max_row=18) # Col Q: Tổng Đạm
ref_dates = Reference(ws06, min_col=14, min_row=5, max_row=18) # Col N: Ngày
ch5.add_data(ref_protein_data, titles_from_data=True)
ch5.set_categories(ref_dates)
ws13.add_chart(ch5, "A39")

# Add Chart 6: Exercise Progressive Overload Trend (Incline DB Bench Press)
# Plot volume or load across recorded sessions
ch6 = BarChart()
ch6.title = "Khối Lượng Tập Luyện (Volume Load = Tạ x Reps) Buổi Tập"
ch6.style = 10
ch6.width = 16
ch6.height = 9.5
ch6.y_axis.title = "Volume (kg x reps)"
ch6.x_axis.title = "Buổi Tập"

ref_session_vol = Reference(ws09, min_col=17, min_row=3, max_row=11) # Col Q: Tổng Volume
ref_session_names = Reference(ws09, min_col=15, min_row=4, max_row=11) # Col O: Buổi Tập
ch6.add_data(ref_session_vol, titles_from_data=True)
ch6.set_categories(ref_session_names)
ws13.add_chart(ch6, "I39")

print("Added Chart 5 (Protein Tracker) and Chart 6 (Progressive Overload Volume) to Dashboard.")

# Save modified workbook
wb.save('Personal_Bulking_Tracker_2026.xlsx')
print("Personal_Bulking_Tracker_2026.xlsx successfully updated and saved!")
