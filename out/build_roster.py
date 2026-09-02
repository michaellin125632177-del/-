# -*- coding: utf-8 -*-
"""牙醫體系 統一出勤表範本 (班表 + 月結統計)"""
import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

F = "微軟正黑體"
OUT = "/home/user/-/out/牙醫體系_統一出勤表_範本.xlsx"

YEAR, MONTH = 2026, 10
DAYS_IN_MONTH = 31           # 2026-10
DAY_C0, DAY_C1 = 4, 34       # D..AH  (1..31)
ROW0, ROW1 = 6, 35           # 班表人員列
N_PEOPLE = ROW1 - ROW0 + 1

# ---------- 樣式 ----------
def font(sz=10, b=False, color="000000"):
    return Font(name=F, size=sz, bold=b, color=color)

TITLE_F   = font(16, True)
HDR_F     = font(10, True, "FFFFFF")
HDR_FILL  = PatternFill("solid", fgColor="1F4E5F")
SUB_FILL  = PatternFill("solid", fgColor="DCE6EC")
IN_FILL   = PatternFill("solid", fgColor="FFF2CC")   # 黃底 = 要手動填
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")   # 灰底 = 自動計算
WKND_FILL = PatternFill("solid", fgColor="DDEBF7")
LEAVE_FILL= PatternFill("solid", fgColor="D9D9D9")
SUPP_FILL = PatternFill("solid", fgColor="C6E0B4")
GAP_FILL  = PatternFill("solid", fgColor="FCE4E4")
ALERT_FILL= PatternFill("solid", fgColor="FF9999")

thin = Side(style="thin", color="AAAAAA")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR  = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

def put(ws, cell, value, f=None, fill=None, align=None, fmt=None, border=True):
    c = ws[cell]
    c.value = value
    c.font = f or font()
    if fill: c.fill = fill
    if align: c.alignment = align
    if fmt: c.number_format = fmt
    if border: c.border = BOX
    return c

def header_row(ws, row, labels, start_col=2):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX

wb = Workbook()

# ============================================================ 1. 說明
ws = wb.active
ws.title = "說明"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 100

put(ws, "B2", "牙醫體系 — 統一出勤表 使用說明", TITLE_F, border=False)
ws.merge_cells("B2:C2")

blocks = [
 ("這份檔案是什麼",
  "全體系共用的月班表範本。統一的是「欄位」與「班別代碼」,不是逼各院所用同一種排班方式。\n"
  "醫師用診次計、助理與醫護長用工時計,兩者共用同一套代碼,統計才能橫向比較。"),
 ("分頁怎麼用",
  "① 設定 — 班別代碼、院所清單、人員名冊。全體系只有總部可改,各院所不要動。\n"
  "② 班表 — 各院所每月填這一頁。只需填黃底格子。\n"
  "③ 月結統計 — 全自動,不要手動改(只有「加班時數」那一欄要手填)。"),
 ("一間院所一個檔",
  "請從本範本「另存新檔」給每間院所各一份,檔名例:2026-10_中山院_出勤表.xlsx。\n"
  "若要放在同一個檔案內,必須把「班表」與「月結統計」兩個分頁「一起選取後複製」,\n"
  "否則月結統計會算到別間院所的班表。"),
 ("填表順序",
  "1. 在「班表」B2 選院所,E2 填年、G2 填月。\n"
  "2. A 欄用下拉選單選員工編號 → 姓名與職類會自動帶出。\n"
  "3. D 欄起逐日選班別代碼(下拉選單,不要自己打字)。\n"
  "4. 檢查第 37~39 列「當日人力」有沒有變紅(助理不足 2 人會示警)。\n"
  "5. 到「月結統計」把打卡機的加班時數填進黃底欄,月結完成。"),
 ("顏色代表什麼",
  "黃底 = 你要填的格子     灰底 = 公式自動算,別動\n"
  "淺藍欄 = 週六/週日       灰色格 = 休假類代碼\n"
  "綠色格 = 支援他院        粉紅格 = 該日還沒排班(漏排提醒)\n"
  "紅色格 = 當日人力低於門檻"),
 ("交表時程(建議)",
  "每月 25 日前,各院所醫護長把下月班表交回總部,總部彙整後鎖表。\n"
  "臨時調班一律在「班表」上改並註記,不要只在 LINE 群講。"),
 ("⚠ 法規提醒",
  "本表是「預排班表 + 月結統計」。依勞動基準法第 30 條,雇主仍須另備「逐日、記載至分鐘」的\n"
  "實際出勤紀錄(打卡機或簽到紀錄),並保存 5 年。本表不能取代那份紀錄。\n"
  "受僱醫師是否適用勞基法,依其僱傭契約與主管機關公告認定,建議請人資或顧問確認後再定義工時欄。"),
 ("範例資料",
  "「班表」第 6~14 列與「設定」的院所、人員都是範例資料,上線前請整批刪除換成實際名單。\n"
  "「設定」的班別時間(如早診 09:00-12:00)也請改成貴體系實際的診次時間。"),
]
r = 4
for h, body in blocks:
    put(ws, f"B{r}", h, font(10, True), SUB_FILL, LEFT)
    put(ws, f"C{r}", body, font(10), None, WRAP)
    ws.row_dimensions[r].height = 15 * (body.count("\n") + 1) + 6
    r += 1

# ============================================================ 2. 設定
st = wb.create_sheet("設定")
st.sheet_view.showGridLines = False
for col, w in {"A":2,"B":10,"C":16,"D":10,"E":10,"F":12,"G":8,"H":10,"I":10,"J":2,"K":10,"L":2,"M":12}.items():
    st.column_dimensions[col].width = w

put(st, "B1", "設定表(總部維護,各院所請勿修改)", TITLE_F, border=False)

# --- 班別代碼表 rows 4..19 ---
put(st, "B3", "一、班別代碼表", font(11, True), border=False)
header_row(st, 4, ["代碼","名稱","開始","結束","排班工時","診次","計出勤","類別"])
CODES = [
    ("D1","早診","09:00","12:00",3.0,1,1,"上班"),
    ("D2","午診","14:00","17:00",3.0,1,1,"上班"),
    ("D3","晚診","18:00","21:00",3.0,1,1,"上班"),
    ("FD","全日診","09:00","21:00",6.0,2,1,"上班"),
    ("A","早班","09:00","18:00",8.0,0,1,"上班"),
    ("P","晚班","13:00","21:30",8.0,0,1,"上班"),
    ("ADM","行政班","09:00","18:00",8.0,0,1,"上班"),
    ("支","支援他院","—","—",8.0,2,1,"上班"),
    ("訓","教育訓練/學會","—","—",8.0,0,1,"上班"),
    ("OFF","排休","—","—",0.0,0,0,"休假"),
    ("特","特休","—","—",0.0,0,0,"休假"),
    ("病","病假","—","—",0.0,0,0,"休假"),
    ("事","事假","—","—",0.0,0,0,"休假"),
    ("公","公假","—","—",0.0,0,0,"休假"),
    ("國","國定假日","—","—",0.0,0,0,"休假"),
    ("休","休診","—","—",0.0,0,0,"休假"),
]
CODE_R0, CODE_R1 = 5, 5 + len(CODES) - 1          # 5..20
WORK_R1 = 5 + 8                                    # 上班類最後一列 = 13
for i, row in enumerate(CODES):
    r = CODE_R0 + i
    for j, v in enumerate(row):
        c = st.cell(row=r, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if j == 1: c.alignment = LEFT
        if j == 4: c.number_format = "0.0"
put(st, f"B{CODE_R1+1}", "※ 特休/公假為有薪假,以「天數」計,故排班工時記 0,薪資另依請假天數計算。",
    font(9, color="808080"), None, LEFT, border=False)

# --- 星期對照 K5:K11 ---
put(st, "K4", "星期對照", font(9, True), SUB_FILL, CTR)
for i, w in enumerate(["日","一","二","三","四","五","六"]):
    put(st, f"K{5+i}", w, font(), CALC_FILL, CTR)
put(st, "M4", "職類", font(9, True), SUB_FILL, CTR)
for i, v in enumerate(["醫師","醫護長","助理"]):
    put(st, f"M{5+i}", v, font(), CALC_FILL, CTR)

# --- 院所清單 rows 25..30 ---
CL_R0 = 25
put(st, f"B{CL_R0-1}", "二、院所清單", font(11, True), border=False)
header_row(st, CL_R0, ["院所代碼","院所名稱","負責人","電話","備註"])
CLINICS = [
    ("Z01","中山院","陳彥廷","02-2500-0000","總院"),
    ("Z02","板橋院","林佳蓉","02-2960-0000",""),
    ("Z03","竹北院","黃冠宇","03-558-0000",""),
    ("Z04","台中院","(待補)","04-2200-0000","2026Q4 開幕"),
]
for i, row in enumerate(CLINICS):
    for j, v in enumerate(row):
        c = st.cell(row=CL_R0 + 1 + i, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
CL_L0, CL_L1 = CL_R0 + 1, CL_R0 + 6   # 留 6 列空間

# --- 人員名冊 rows 34..63 ---
PP_R0 = 34
put(st, f"B{PP_R0-1}", "三、人員名冊", font(11, True), border=False)
header_row(st, PP_R0, ["員工編號","姓名","職類","所屬院所","到職日","備註"])
PEOPLE = [
    ("D001","陳彥廷","醫師","中山院","2019-03-01","院長"),
    ("D002","林佳蓉","醫師","中山院","2021-07-15",""),
    ("D003","黃冠宇","醫師","中山院","2023-02-01","週三支援板橋院"),
    ("N001","吳淑芬","醫護長","中山院","2018-05-06",""),
    ("A001","張怡君","助理","中山院","2022-09-01",""),
    ("A002","李宗翰","助理","中山院","2023-04-10",""),
    ("A003","王思婷","助理","中山院","2024-01-08",""),
    ("A004","蔡孟蓉","助理","中山院","2024-11-01",""),
    ("A005","許家豪","助理","中山院","2025-06-16","兼職"),
]
PP_L0 = PP_R0 + 1
PP_L1 = PP_R0 + 30
for i in range(30):
    r = PP_L0 + i
    vals = PEOPLE[i] if i < len(PEOPLE) else ("",)*6
    for j, v in enumerate(vals):
        c = st.cell(row=r, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if i < len(PEOPLE): c.fill = IN_FILL
        if j == 5: c.alignment = LEFT

CODE_RNG  = f"設定!$B${CODE_R0}:$B${CODE_R1}"
WORK_RNG  = f"設定!$B${CODE_R0}:$B${WORK_R1}"
HRS_RNG   = f"設定!$F${CODE_R0}:$F${CODE_R1}"
SES_RNG   = f"設定!$G${CODE_R0}:$G${CODE_R1}"
ATT_RNG   = f"設定!$H${CODE_R0}:$H${CODE_R1}"
LEAVE_RNG = f"設定!$B${WORK_R1+1}:$B${CODE_R1}"

# ============================================================ 3. 班表
sc = wb.create_sheet("班表")
sc.sheet_view.showGridLines = False
sc.freeze_panes = "D6"
sc.column_dimensions["A"].width = 11
sc.column_dimensions["B"].width = 11
sc.column_dimensions["C"].width = 9
for c in range(DAY_C0, DAY_C1 + 1):
    sc.column_dimensions[get_column_letter(c)].width = 4.6

put(sc, "A1", "牙醫體系 — 月班表", TITLE_F, border=False)
sc.merge_cells(start_row=1, start_column=1, end_row=1, end_column=DAY_C1)

put(sc, "A2", "院所", font(10, True), SUB_FILL, CTR)
put(sc, "B2", "中山院", font(10, True), IN_FILL, CTR)
put(sc, "D2", "年", font(10, True), SUB_FILL, CTR)
put(sc, "E2", YEAR, font(10, True), IN_FILL, CTR, "0")
put(sc, "F2", "月", font(10, True), SUB_FILL, CTR)
put(sc, "G2", MONTH, font(10, True), IN_FILL, CTR, "0")
put(sc, "I2", "← 黃底= 手動填。日期與星期會自動更新。", font(9, color="808080"), None, LEFT, border=False)

for cell, lab in (("A4","員工編號"),("B4","姓名"),("C4","職類")):
    put(sc, cell, lab, HDR_F, HDR_FILL, CTR)
    sc.merge_cells(f"{cell[0]}4:{cell[0]}5")

for c in range(DAY_C0, DAY_C1 + 1):
    d = c - DAY_C0 + 1
    cell = sc.cell(row=4, column=c)
    cell.value = f'=IF({d}<=DAY(EOMONTH(DATE($E$2,$G$2,1),0)),{d},"")'
    cell.font, cell.fill, cell.alignment, cell.border = HDR_F, HDR_FILL, CTR, BOX
    L = get_column_letter(c)
    w = sc.cell(row=5, column=c)
    w.value = (f'=IF({L}$4="","",INDEX(設定!$K$5:$K$11,'
               f'WEEKDAY(DATE($E$2,$G$2,{L}$4),1)))')
    w.font, w.fill, w.alignment, w.border = font(9, True), SUB_FILL, CTR, BOX

# 範例班表資料
def build_sample():
    sched = {}
    for i, (eid, name, role, *_ ) in enumerate(PEOPLE):
        row = []
        for d in range(1, DAYS_IN_MONTH + 1):
            wd = dt.date(YEAR, MONTH, d).weekday()   # 0=Mon .. 6=Sun
            if role == "醫師":
                if wd == 6: code = "休"
                elif wd == i: code = "OFF"
                elif eid == "D003" and wd == 2: code = "支"
                else: code = ["FD","D1","D2","D3","FD","D2"][(d + i) % 6]
            elif role == "醫護長":
                if wd == 6: code = "OFF"
                elif wd == 3: code = "ADM"
                else: code = "A"
            else:
                k = i - 4
                if wd == 6: code = "OFF"
                elif wd == (k % 6): code = "OFF"
                else: code = "A" if (d + k) % 2 == 0 else "P"
        # 插幾筆假別讓統計看得出來
            row.append(code)
        sched[eid] = row
    sched["D002"][9]  = "訓"
    sched["A001"][14] = "特"
    sched["A003"][21] = "病"
    sched["N001"][6]  = "特"
    return sched

SAMPLE = build_sample()

for i in range(N_PEOPLE):
    r = ROW0 + i
    sc.row_dimensions[r].height = 18
    eid = PEOPLE[i][0] if i < len(PEOPLE) else None
    a = sc.cell(row=r, column=1, value=eid)
    a.font, a.fill, a.border, a.alignment = font(), IN_FILL, BOX, CTR
    b = sc.cell(row=r, column=2)
    b.value = f'=IFERROR(INDEX(設定!$C${PP_L0}:$C${PP_L1},MATCH($A{r},設定!$B${PP_L0}:$B${PP_L1},0)),"")'
    b.font, b.fill, b.border, b.alignment = font(), CALC_FILL, BOX, CTR
    cc = sc.cell(row=r, column=3)
    cc.value = f'=IFERROR(INDEX(設定!$D${PP_L0}:$D${PP_L1},MATCH($A{r},設定!$B${PP_L0}:$B${PP_L1},0)),"")'
    cc.font, cc.fill, cc.border, cc.alignment = font(), CALC_FILL, BOX, CTR
    for c in range(DAY_C0, DAY_C1 + 1):
        d = c - DAY_C0 + 1
        cell = sc.cell(row=r, column=c)
        if eid and d <= DAYS_IN_MONTH:
            cell.value = SAMPLE[eid][d - 1]
        cell.font, cell.border, cell.alignment = font(9), BOX, CTR

# 當日人力檢核
LBL = {37: "當日醫師人數", 38: "當日助理/醫護長人數", 39: "當日總人力"}
put(sc, "A36", "人力檢核(自動計算)", font(10, True), SUB_FILL, LEFT)
sc.merge_cells("A36:C36")
for r, lab in LBL.items():
    put(sc, f"A{r}", lab, font(10, True), SUB_FILL, LEFT)
    sc.merge_cells(f"A{r}:C{r}")
    for c in range(DAY_C0, DAY_C1 + 1):
        L = get_column_letter(c)
        cell = sc.cell(row=r, column=c)
        on_duty = f'(COUNTIF({WORK_RNG},{L}${ROW0}:{L}${ROW1})>0)'
        if r == 37:
            fx = f'SUMPRODUCT(($C${ROW0}:$C${ROW1}="醫師")*{on_duty})'
        elif r == 38:
            fx = (f'SUMPRODUCT((($C${ROW0}:$C${ROW1}="助理")+($C${ROW0}:$C${ROW1}="醫護長"))'
                  f'*{on_duty})')
        else:
            fx = f'{L}37+{L}38'
        cell.value = f'=IF({L}$4="","",{fx})'
        cell.font, cell.fill, cell.border, cell.alignment = font(9, True), CALC_FILL, BOX, CTR

put(sc, "A41",
    "※ 門檻設定:有看診的日子至少 1 位醫師、2 位助理/醫護長,低於門檻自動變紅;"
    "全院休診日不會示警。要調門檻請改第 37~38 列的條件式格式。",
    font(9, color="808080"), None, LEFT, border=False)

# 資料驗證
dv_code = DataValidation(type="list", formula1=CODE_RNG, allow_blank=True,
                         showErrorMessage=True, errorTitle="班別代碼無效",
                         error="請從下拉選單選擇「設定」分頁定義的班別代碼。")
sc.add_data_validation(dv_code)
dv_code.add(f"{get_column_letter(DAY_C0)}{ROW0}:{get_column_letter(DAY_C1)}{ROW1}")

dv_emp = DataValidation(type="list", formula1=f"設定!$B${PP_L0}:$B${PP_L1}", allow_blank=True)
sc.add_data_validation(dv_emp)
dv_emp.add(f"A{ROW0}:A{ROW1}")

dv_cli = DataValidation(type="list", formula1=f"設定!$C${CL_L0}:$C${CL_L1}", allow_blank=True)
sc.add_data_validation(dv_cli)
dv_cli.add("B2")

# 條件式格式
GRID = f"{get_column_letter(DAY_C0)}{ROW0}:{get_column_letter(DAY_C1)}{ROW1}"
D0 = get_column_letter(DAY_C0)
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}{ROW0}="")'], fill=GAP_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'{D0}{ROW0}="支"'], fill=SUPP_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'COUNTIF({LEAVE_RNG},{D0}{ROW0})>0'], fill=LEAVE_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'OR({D0}$5="六",{D0}$5="日")'], fill=WKND_FILL))
sc.conditional_formatting.add(f"{D0}4:{get_column_letter(DAY_C1)}5", FormulaRule(
    formula=[f'OR({D0}$5="六",{D0}$5="日")'], fill=WKND_FILL))
# 全院休診日(當日無人上班)不視為人力不足,避免週日誤報
sc.conditional_formatting.add(f"{D0}37:{get_column_letter(DAY_C1)}37", FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}39>0,{D0}37<1)'], fill=ALERT_FILL))
sc.conditional_formatting.add(f"{D0}38:{get_column_letter(DAY_C1)}38", FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}37>0,{D0}38<2)'], fill=ALERT_FILL))

# ============================================================ 4. 月結統計
ms = wb.create_sheet("月結統計")
ms.sheet_view.showGridLines = False
ms.freeze_panes = "D5"
widths = {"A":11,"B":11,"C":9,"D":9,"E":9,"F":8,"G":7,"H":7,"I":7,"J":7,"K":7,"L":9,"M":9,"N":10,"O":10,"P":14}
for col, w in widths.items(): ms.column_dimensions[col].width = w

put(ms, "A1", "月結統計(自動計算,僅「加班時數」需手填)", TITLE_F, border=False)
ms.merge_cells("A1:P1")
put(ms, "A2", "院所", font(10, True), SUB_FILL, CTR)
put(ms, "B2", "=班表!B2", font(10, True), CALC_FILL, CTR)
put(ms, "C2", "期間", font(10, True), SUB_FILL, CTR)
put(ms, "D2", '=班表!E2&"年"&班表!G2&"月"', font(10, True), CALC_FILL, CTR)

HEADS = ["員工編號","姓名","職類","出勤天數","排班工時","總診次","排休","特休","病假","事假",
         "公假","國定假日","支援他院","加班時數","實際工時","排班完整度檢核"]
for i, h in enumerate(HEADS):
    c = ms.cell(row=4, column=1 + i, value=h)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
ms.row_dimensions[4].height = 30

MS_R0 = 5
DAYS_FX = "DAY(EOMONTH(DATE(班表!$E$2,班表!$G$2,1),0))"
for i in range(N_PEOPLE):
    r = MS_R0 + i
    sr = ROW0 + i
    rng = f"班表!$D{sr}:$AH{sr}"
    guard = f'IF($A{r}="","",'
    rows = {
      "A": f'=IF(班表!$A{sr}="","",班表!$A{sr})',
      "B": f'={guard}班表!$B{sr})',
      "C": f'={guard}班表!$C{sr})',
      "D": f'={guard}SUMPRODUCT(COUNTIF({rng},{CODE_RNG}),{ATT_RNG}))',
      "E": f'={guard}SUMPRODUCT(COUNTIF({rng},{CODE_RNG}),{HRS_RNG}))',
      "F": f'={guard}SUMPRODUCT(COUNTIF({rng},{CODE_RNG}),{SES_RNG}))',
      "G": f'={guard}COUNTIF({rng},"OFF"))',
      "H": f'={guard}COUNTIF({rng},"特"))',
      "I": f'={guard}COUNTIF({rng},"病"))',
      "J": f'={guard}COUNTIF({rng},"事"))',
      "K": f'={guard}COUNTIF({rng},"公"))',
      "L": f'={guard}COUNTIF({rng},"國"))',
      "M": f'={guard}COUNTIF({rng},"支"))',
      "O": f'={guard}$E{r}+$N{r})',
      "P": (f'={guard}IF(COUNTA({rng})={DAYS_FX},"OK",'
            f'"⚠ 缺"&({DAYS_FX}-COUNTA({rng}))&"天"))'),
    }
    for col, fx in rows.items():
        c = ms[f"{col}{r}"]
        c.value, c.font, c.border, c.alignment = fx, font(), BOX, CTR
        c.fill = CALC_FILL
        if col in ("E", "O"): c.number_format = "0.0"
        if col == "P": c.alignment = CTR
    n = ms[f"N{r}"]
    n.value = 0
    n.font, n.fill, n.border, n.alignment, n.number_format = font(), IN_FILL, BOX, CTR, "0.0"

TOT = MS_R0 + N_PEOPLE
put(ms, f"A{TOT}", "合計", font(10, True), SUB_FILL, CTR)
ms.merge_cells(f"A{TOT}:C{TOT}")
for col in "DEFGHIJKLMNO":
    c = ms[f"{col}{TOT}"]
    c.value = f"=SUM({col}{MS_R0}:{col}{TOT-1})"
    c.font, c.fill, c.border, c.alignment = font(10, True), SUB_FILL, BOX, CTR
    if col in ("E","O"): c.number_format = "0.0"
put(ms, f"P{TOT}", "", font(), SUB_FILL, CTR)

ms.conditional_formatting.add(f"P{MS_R0}:P{TOT-1}", FormulaRule(
    formula=[f'AND($A{MS_R0}<>"",P{MS_R0}<>"OK")'], fill=ALERT_FILL))

notes = [
 "※ 出勤天數/排班工時/總診次 = 依「設定」分頁的班別代碼權重自動加總,改代碼權重這裡就會跟著變。",
 "※ 加班時數(N 欄)請由打卡機或簽到紀錄填入,本表不會自己算 —— 預排班表不等於實際出勤。",
 "※ 檢核欄顯示「⚠ 缺 N 天」代表該員當月還有 N 天沒排班,交表前要補完。",
]
for i, t in enumerate(notes):
    put(ms, f"A{TOT+2+i}", t, font(9, color="808080"), None, LEFT, border=False)

wb.save(OUT)
print("saved:", OUT)
