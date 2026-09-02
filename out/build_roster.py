# -*- coding: utf-8 -*-
"""牙醫體系 統一出勤表範本
分頁:說明 / 設定 / 班表 / 打卡匯入 / 出勤紀錄 / 月結統計
"""
import datetime as dt
import random
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

F = "微軟正黑體"
OUT = "/home/user/-/out/牙醫體系_統一出勤表_範本.xlsx"

YEAR, MONTH = 2026, 10
DAYS_IN_MONTH = 31
DAY_C0, DAY_C1 = 4, 34          # 班表 D..AH = 1..31 日
ROW0, ROW1 = 6, 35              # 班表人員列
N_PEOPLE = ROW1 - ROW0 + 1      # 30

PUNCH_R0, PUNCH_R1 = 3, 1002    # 打卡匯入資料列
ATT_R0 = 3                      # 出勤紀錄資料起始列
ATT_R1 = ATT_R0 + N_PEOPLE * DAYS_IN_MONTH - 1   # 932

# ---------------------------------------------------------------- 樣式
def font(sz=10, b=False, color="000000"):
    return Font(name=F, size=sz, bold=b, color=color)

TITLE_F   = font(16, True)
HDR_F     = font(10, True, "FFFFFF")
HDR_FILL  = PatternFill("solid", fgColor="1F4E5F")
SUB_FILL  = PatternFill("solid", fgColor="DCE6EC")
IN_FILL   = PatternFill("solid", fgColor="FFF2CC")
CALC_FILL = PatternFill("solid", fgColor="F2F2F2")
WKND_FILL = PatternFill("solid", fgColor="DDEBF7")
LEAVE_FILL= PatternFill("solid", fgColor="D9D9D9")
SUPP_FILL = PatternFill("solid", fgColor="C6E0B4")
GAP_FILL  = PatternFill("solid", fgColor="FCE4E4")
ALERT_FILL= PatternFill("solid", fgColor="FF9999")
OT_FILL   = PatternFill("solid", fgColor="FCE4D6")

thin = Side(style="thin", color="AAAAAA")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)
CTR  = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

FMT_TIME = "hh:mm"
FMT_HR   = '0.00;-0.00;"–"'
FMT_MIN  = '0;-0;"–"'
FMT_DATE = "yyyy/m/d"

def put(ws, cell, value, f=None, fill=None, align=None, fmt=None, border=True):
    c = ws[cell]; c.value = value
    c.font = f or font()
    if fill: c.fill = fill
    if align: c.alignment = align
    if fmt: c.number_format = fmt
    if border: c.border = BOX
    return c

def header_row(ws, row, labels, start_col=2, height=None):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start_col + i, value=lab)
        c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    if height: ws.row_dimensions[row].height = height

T = lambda h, m: dt.time(h, m)

wb = Workbook()

# ================================================================ 設定資料
# 代碼, 名稱, 應到, 應退, 休息分鐘, 排班工時, 診次, 計出勤, 類別
CODES = [
    ("D1", "早診",          T(9,0),  T(12,0),   0, 3.0, 1, 1, "上班"),
    ("D2", "午診",          T(14,0), T(17,0),   0, 3.0, 1, 1, "上班"),
    ("D3", "晚診",          T(18,0), T(21,0),   0, 3.0, 1, 1, "上班"),
    ("FD", "全日診",        T(9,0),  T(17,0), 120, 6.0, 2, 1, "上班"),
    ("A",  "早班",          T(9,0),  T(18,0),  60, 8.0, 0, 1, "上班"),
    ("P",  "晚班",          T(12,30),T(21,30), 60, 8.0, 0, 1, "上班"),
    ("ADM","行政班",        T(9,0),  T(18,0),  60, 8.0, 0, 1, "上班"),
    ("支", "支援他院",      None,    None,     60, 8.0, 2, 1, "上班"),
    ("訓", "教育訓練/學會", None,    None,     60, 8.0, 0, 1, "上班"),
    ("OFF","排休",          None,    None,      0, 0.0, 0, 0, "休假"),
    ("特", "特休",          None,    None,      0, 0.0, 0, 0, "休假"),
    ("病", "病假",          None,    None,      0, 0.0, 0, 0, "休假"),
    ("事", "事假",          None,    None,      0, 0.0, 0, 0, "休假"),
    ("公", "公假",          None,    None,      0, 0.0, 0, 0, "休假"),
    ("國", "國定假日",      None,    None,      0, 0.0, 0, 0, "休假"),
    ("休", "休診",          None,    None,      0, 0.0, 0, 0, "休假"),
]
# 自我檢查:有固定時段的代碼,應退 - 應到 - 休息 必須等於排班工時
for c in CODES:
    if c[2] and c[3]:
        span = (dt.datetime.combine(dt.date.min, c[3]) -
                dt.datetime.combine(dt.date.min, c[2])).seconds / 3600
        assert abs(span - c[4] / 60 - c[5]) < 1e-9, f"代碼 {c[0]} 工時定義不一致"

CODE_R0 = 5
CODE_R1 = CODE_R0 + len(CODES) - 1      # 20
WORK_R1 = CODE_R0 + 8                   # 13,上班類最後一列
CODE_BY = {c[0]: c for c in CODES}

CLINICS = [
    ("Z01", "中山院", "陳彥廷", "02-2500-0000", "總院"),
    ("Z02", "板橋院", "林佳蓉", "02-2960-0000", ""),
    ("Z03", "竹北院", "黃冠宇", "03-558-0000", ""),
    ("Z04", "台中院", "(待補)", "04-2200-0000", "2026Q4 開幕"),
]
PEOPLE = [
    ("D001", "陳彥廷", "醫師",   "中山院", "2019-03-01", "院長"),
    ("D002", "林佳蓉", "醫師",   "中山院", "2021-07-15", ""),
    ("D003", "黃冠宇", "醫師",   "中山院", "2023-02-01", "週三支援板橋院"),
    ("N001", "吳淑芬", "醫護長", "中山院", "2018-05-06", ""),
    ("A001", "張怡君", "助理",   "中山院", "2022-09-01", ""),
    ("A002", "李宗翰", "助理",   "中山院", "2023-04-10", ""),
    ("A003", "王思婷", "助理",   "中山院", "2024-01-08", ""),
    ("A004", "蔡孟蓉", "助理",   "中山院", "2024-11-01", ""),
    ("A005", "許家豪", "助理",   "中山院", "2025-06-16", "兼職"),
]
PARAM_R0 = 25                       # 計算參數表頭列,參數值在 26~28 列
P_OT_MIN  = "設定!$C$26"            # 加班認定門檻(分)
P_OT_UNIT = "設定!$C$27"            # 加班計算單位(分)
P_GRACE   = "設定!$C$28"            # 遲到早退寬限(分)
CL_R0 = 31; CL_L0, CL_L1 = 32, 37
PP_R0 = 40; PP_L0, PP_L1 = 41, 70

# 常用範圍字串
R_CODE  = f"設定!$B${CODE_R0}:$B${CODE_R1}"
R_WORK  = f"設定!$B${CODE_R0}:$B${WORK_R1}"
R_LEAVE = f"設定!$B${WORK_R1+1}:$B${CODE_R1}"
R_IN    = f"設定!$D${CODE_R0}:$D${CODE_R1}"
R_OUT   = f"設定!$E${CODE_R0}:$E${CODE_R1}"
R_REST  = f"設定!$F${CODE_R0}:$F${CODE_R1}"
R_HRS   = f"設定!$G${CODE_R0}:$G${CODE_R1}"
R_SES   = f"設定!$H${CODE_R0}:$H${CODE_R1}"
R_ATT   = f"設定!$I${CODE_R0}:$I${CODE_R1}"
R_NAME  = f"設定!$C${PP_L0}:$C${PP_L1}"
R_ROLE  = f"設定!$D${PP_L0}:$D${PP_L1}"
R_EID   = f"設定!$B${PP_L0}:$B${PP_L1}"
DAYS_FX = "DAY(EOMONTH(DATE(班表!$E$2,班表!$G$2,1),0))"

# ================================================================ 1. 說明
ws = wb.active; ws.title = "說明"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 104
put(ws, "B2", "牙醫體系 — 統一出勤表 使用說明", TITLE_F, border=False)
ws.merge_cells("B2:C2")

BLOCKS = [
 ("這份檔案是什麼",
  "全體系共用的月出勤表。統一的是「欄位」與「班別代碼」,不是逼各院所用同一種排班方式。\n"
  "醫師用診次計、助理與醫護長用工時計,兩者共用同一套代碼,統計才能橫向比較。\n"
  "檔案同時涵蓋「預排班表」與「實際出勤紀錄」兩件事——法規上這兩份缺一不可。"),
 ("六個分頁怎麼分工",
  "① 說明 — 就是這一頁。\n"
  "② 設定 — 班別代碼、院所清單、人員名冊。總部維護,各院所勿動。\n"
  "③ 班表 — 下個月誰哪天上什麼班。各院所醫護長每月填。\n"
  "④ 打卡匯入 — 把打卡機匯出的資料整段貼進來。只有四欄。\n"
  "⑤ 出勤紀錄 — 逐人逐日的法定出勤紀錄,自動比對班表與打卡。\n"
  "⑥ 月結統計 — 全自動,不需手填。"),
 ("每月作業流程",
  "上月底:在「班表」排好下月班 → 檢查底下的人力檢核列 → 交總部。\n"
  "當月底:把打卡機匯出檔貼進「打卡匯入」 → 到「出勤紀錄」看異常欄 → 逐筆處理 →\n"
  "「月結統計」直接就是結果,送人資。"),
 ("一間院所一個檔",
  "請從本範本另存新檔給每間院所各一份,檔名例:2026-10_中山院_出勤表.xlsx。\n"
  "不建議把多間院所塞進同一個檔——「出勤紀錄」與「月結統計」的公式都綁定同一份「班表」。"),
 ("打卡匯入要準備什麼",
  "只需四欄:員工編號、日期、上班時間、下班時間,一天一列。\n"
  "多數打卡系統(如快打卡、Femas、NUEiP)都能匯出這種格式;欄位順序不同就先在 Excel 調整。\n"
  "日期要是真正的日期格式、時間要是真正的時間格式,不能是文字,否則比對會失敗。\n"
  "貼上後看最右邊的檢核欄:會抓出「查無員編」與「同一天重複打卡」兩種問題。"),
 ("出勤紀錄怎麼看",
  "一人一天一列,共 930 列(30 人 × 31 天),沒用到的人員列會自動留白。\n"
  "「出勤異常」欄自動判定五種狀態:正常、遲到、早退、未打卡、假日出勤。\n"
  "紅底 = 需要處理的異常;橘底 = 當天有加班。用上方篩選鈕只看異常最快。\n"
  "沒有打卡機的院所,可以直接把「實際上班/實際下班」兩欄的公式刪掉改成手填,其餘照算。"),
 ("顏色代表什麼",
  "黃底 = 你要填的格子       灰底 = 公式自動算,別動\n"
  "淺藍欄 = 週六/週日        灰色格 = 休假類代碼\n"
  "綠色格 = 支援他院          粉紅格 = 該日還沒排班\n"
  "紅色 = 人力不足或出勤異常  橘色 = 有加班"),
 ("交表時程(建議)",
  "每月 25 日前,各院所醫護長把下月班表交回總部,總部彙整後鎖表。\n"
  "每月 5 日前,完成上月的打卡匯入與異常處理,月結送人資。\n"
  "臨時調班一律在「班表」上改並註記,不要只在 LINE 群講。"),
 ("⚠ 法規提醒",
  "依勞動基準法第 30 條,雇主應置備勞工出勤紀錄,逐日記載出勤情形至分鐘為止,並保存 5 年。\n"
  "「出勤紀錄」分頁就是為此設計的;請每月結束後另存一份唯讀封存檔,不要只留最新版覆蓋。\n"
  "加班時數依實際工時扣除排班工時計算,未含加班費率換算——加班費請依勞基法第 24 條與\n"
  "貴體系薪資辦法另行計算。受僱醫師是否適用勞基法依僱傭契約與主管機關認定,建議請人資確認。"),
 ("四個內建假設(請確認後調整)",
  "一、休息時間依「設定」分頁各代碼的固定值扣除,不是依實際打卡判斷。\n"
  "二、加班要超過門檻(預設 30 分)才認定,並無條件捨去到計算單位(預設 30 分)。\n"
  "   否則每天晚幾分鐘打卡都會被算成加班,一個月會憑空多出好幾小時。\n"
  "三、打卡與應到/應退時間差在寬限內(預設 5 分)不判定遲到早退。\n"
  "   以上三項都在「設定」分頁的計算參數區,改一次全檔生效。\n"
  "四、加班別自動判定為:週日=例假、週六=休息日、代碼「國」=國定假日、其餘=平日。\n"
  "   若貴體系的例假日不是週日(輪班制常見),請改「出勤紀錄」Q 欄的公式。"),
 ("範例資料",
  "「班表」「打卡匯入」與「設定」的院所、人員都是範例(中山院 2026 年 10 月,9 位同仁),\n"
  "上線前請整批刪除換成實際名單。班別時間也請改成貴體系實際的診次時段。\n"
  "範例刻意放了遲到、早退、未打卡、假日出勤、加班各一筆,方便你看異常判定怎麼跑。"),
]
r = 4
for h, body in BLOCKS:
    put(ws, f"B{r}", h, font(10, True), SUB_FILL, LEFT)
    put(ws, f"C{r}", body, font(10), None, WRAP)
    ws.row_dimensions[r].height = 15 * (body.count("\n") + 1) + 8
    r += 1

# ================================================================ 2. 設定
st = wb.create_sheet("設定")
st.sheet_view.showGridLines = False
for col, w in {"A":2,"B":10,"C":16,"D":9,"E":9,"F":10,"G":10,"H":7,"I":8,"J":8,
               "K":2,"L":10,"M":2,"N":12}.items():
    st.column_dimensions[col].width = w
put(st, "B1", "設定表(總部維護,各院所請勿修改)", TITLE_F, border=False)

put(st, "B3", "一、班別代碼表", font(11, True), border=False)
header_row(st, 4, ["代碼","名稱","應到","應退","休息(分)","排班工時","診次","計出勤","類別"])
for i, row in enumerate(CODES):
    r = CODE_R0 + i
    for j, v in enumerate(row):
        c = st.cell(row=r, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if j == 1: c.alignment = LEFT
        if j in (2, 3): c.number_format = FMT_TIME
        if j == 5: c.number_format = "0.0"
put(st, f"B{CODE_R1+1}",
    "※ 有固定時段的代碼必須滿足「應退 − 應到 − 休息 = 排班工時」,否則出勤紀錄會算出假加班。",
    font(9, color="808080"), None, LEFT, border=False)
put(st, f"B{CODE_R1+2}",
    "※ 支援他院與教育訓練沒有固定時段,遲到早退不判定,工時依實際打卡計。",
    font(9, color="808080"), None, LEFT, border=False)
put(st, f"B{CODE_R1+3}",
    "※ 特休與公假為有薪假,以天數計,故排班工時記 0,薪資另依請假天數計算。",
    font(9, color="808080"), None, LEFT, border=False)

put(st, "L4", "星期對照", font(9, True), SUB_FILL, CTR)
for i, w in enumerate(["日","一","二","三","四","五","六"]):
    put(st, f"L{5+i}", w, font(), CALC_FILL, CTR)
put(st, "N4", "職類", font(9, True), SUB_FILL, CTR)
for i, v in enumerate(["醫師","醫護長","助理"]):
    put(st, f"N{5+i}", v, font(), CALC_FILL, CTR)

put(st, f"B{PARAM_R0-1}", "二、計算參數(全體系一致,改這裡會影響所有統計)",
    font(11, True), border=False)
header_row(st, PARAM_R0, ["參數", "值", "說明"])
PARAMS = [
    ("加班認定門檻(分)", 30,
     "超過排班工時多少分鐘才認定為加班。低於門檻視為正常收尾,不計加班。"),
    ("加班計算單位(分)", 30,
     "認定為加班後,無條件捨去到此單位。填 1 表示逐分鐘計。"),
    ("遲到早退寬限(分)", 5,
     "打卡與應到/應退時間差在此範圍內不判定為遲到或早退。"),
]
for i, (nm, val, desc) in enumerate(PARAMS):
    r = PARAM_R0 + 1 + i
    put(st, f"B{r}", nm, font(), SUB_FILL, LEFT)
    put(st, f"C{r}", val, font(10, True), IN_FILL, CTR, "0")
    put(st, f"D{r}", desc, font(9), None, LEFT)
    st.merge_cells(f"D{r}:J{r}")
put(st, f"B{PARAM_R0+4}",
    "※ 這三個參數是勞資雙方的認定慣例,不是法律規定的數字。"
    "上線前請與人資或勞務顧問確認,並讓同仁知道規則。",
    font(9, color="808080"), None, LEFT, border=False)

put(st, f"B{CL_R0-1}", "三、院所清單", font(11, True), border=False)
header_row(st, CL_R0, ["院所代碼","院所名稱","負責人","電話","備註"])
for i, row in enumerate(CLINICS):
    for j, v in enumerate(row):
        c = st.cell(row=CL_L0 + i, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR

put(st, f"B{PP_R0-1}", "四、人員名冊", font(11, True), border=False)
header_row(st, PP_R0, ["員工編號","姓名","職類","所屬院所","到職日","備註"])
for i in range(30):
    r = PP_L0 + i
    vals = PEOPLE[i] if i < len(PEOPLE) else ("",) * 6
    for j, v in enumerate(vals):
        c = st.cell(row=r, column=2 + j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if i < len(PEOPLE): c.fill = IN_FILL
        if j == 5: c.alignment = LEFT

# ================================================================ 3. 班表
sc = wb.create_sheet("班表")
sc.sheet_view.showGridLines = False
sc.freeze_panes = "D6"
sc.column_dimensions["A"].width = 11
sc.column_dimensions["B"].width = 11
sc.column_dimensions["C"].width = 9
for c in range(DAY_C0, DAY_C1 + 1):
    sc.column_dimensions[get_column_letter(c)].width = 4.6

put(sc, "A1", "牙醫體系 — 月班表(預排)", TITLE_F, border=False)
sc.merge_cells(start_row=1, start_column=1, end_row=1, end_column=DAY_C1)
put(sc, "A2", "院所", font(10, True), SUB_FILL, CTR)
put(sc, "B2", "中山院", font(10, True), IN_FILL, CTR)
put(sc, "D2", "年", font(10, True), SUB_FILL, CTR)
put(sc, "E2", YEAR, font(10, True), IN_FILL, CTR, "0")
put(sc, "F2", "月", font(10, True), SUB_FILL, CTR)
put(sc, "G2", MONTH, font(10, True), IN_FILL, CTR, "0")
put(sc, "I2", "← 黃底 = 手動填。日期與星期會自動更新,整份檔案的年月都以這裡為準。",
    font(9, color="808080"), None, LEFT, border=False)

for cell, lab in (("A4","員工編號"), ("B4","姓名"), ("C4","職類")):
    put(sc, cell, lab, HDR_F, HDR_FILL, CTR)
    sc.merge_cells(f"{cell[0]}4:{cell[0]}5")
for c in range(DAY_C0, DAY_C1 + 1):
    d = c - DAY_C0 + 1
    L = get_column_letter(c)
    cell = sc.cell(row=4, column=c)
    cell.value = f'=IF({d}<={DAYS_FX},{d},"")'
    cell.font, cell.fill, cell.alignment, cell.border = HDR_F, HDR_FILL, CTR, BOX
    w = sc.cell(row=5, column=c)
    w.value = f'=IF({L}$4="","",INDEX(設定!$L$5:$L$11,WEEKDAY(DATE($E$2,$G$2,{L}$4),1)))'
    w.font, w.fill, w.alignment, w.border = font(9, True), SUB_FILL, CTR, BOX

DOC_OFF = {"D001": 0, "D002": 1, "D003": 4}     # 各醫師固定排休日(週一/二/五)

def build_sample():
    sched = {}
    for i, (eid, name, role, *_) in enumerate(PEOPLE):
        row = []
        for d in range(1, DAYS_IN_MONTH + 1):
            wd = dt.date(YEAR, MONTH, d).weekday()      # 0=Mon .. 6=Sun
            if role == "醫師":
                if wd == 6: code = "休"
                elif eid == "D003" and wd == 2: code = "支"
                elif wd == DOC_OFF[eid]: code = "OFF"
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
            row.append(code)
        sched[eid] = row
    sched["D002"][12] = "訓"          # 10/13 學會
    sched["A001"][14] = "特"
    sched["A002"][20] = "事"
    sched["A003"][21] = "病"
    sched["N001"][6]  = "特"
    for eid in sched:                 # 10/10 國慶日,全院休診
        sched[eid][9] = "國"
    return sched

SAMPLE = build_sample()

for i in range(N_PEOPLE):
    r = ROW0 + i
    sc.row_dimensions[r].height = 18
    eid = PEOPLE[i][0] if i < len(PEOPLE) else None
    a = sc.cell(row=r, column=1, value=eid)
    a.font, a.fill, a.border, a.alignment = font(), IN_FILL, BOX, CTR
    for col, src in ((2, R_NAME), (3, R_ROLE)):
        c = sc.cell(row=r, column=col)
        c.value = f'=IFERROR(INDEX({src},MATCH($A{r},{R_EID},0)),"")'
        c.font, c.fill, c.border, c.alignment = font(), CALC_FILL, BOX, CTR
    for c in range(DAY_C0, DAY_C1 + 1):
        d = c - DAY_C0 + 1
        cell = sc.cell(row=r, column=c)
        if eid and d <= DAYS_IN_MONTH:
            cell.value = SAMPLE[eid][d - 1]
        cell.font, cell.border, cell.alignment = font(9), BOX, CTR

LBL = {37: "當日醫師人數", 38: "當日助理/醫護長人數", 39: "當日總人力"}
put(sc, "A36", "人力檢核(自動計算)", font(10, True), SUB_FILL, LEFT)
sc.merge_cells("A36:C36")
for r, lab in LBL.items():
    put(sc, f"A{r}", lab, font(10, True), SUB_FILL, LEFT)
    sc.merge_cells(f"A{r}:C{r}")
    for c in range(DAY_C0, DAY_C1 + 1):
        L = get_column_letter(c)
        on_duty = f'(COUNTIF({R_WORK},{L}${ROW0}:{L}${ROW1})>0)'
        if r == 37:   fx = f'SUMPRODUCT(($C${ROW0}:$C${ROW1}="醫師")*{on_duty})'
        elif r == 38: fx = (f'SUMPRODUCT((($C${ROW0}:$C${ROW1}="助理")+'
                            f'($C${ROW0}:$C${ROW1}="醫護長"))*{on_duty})')
        else:         fx = f'{L}37+{L}38'
        cell = sc.cell(row=r, column=c)
        cell.value = f'=IF({L}$4="","",{fx})'
        cell.font, cell.fill, cell.border, cell.alignment = font(9, True), CALC_FILL, BOX, CTR
put(sc, "A41",
    "※ 門檻:有看診的日子至少 1 位醫師、2 位助理/醫護長,低於門檻自動變紅;"
    "全院休診日不示警。要調門檻請改第 37~38 列的條件式格式。",
    font(9, color="808080"), None, LEFT, border=False)

dv_code = DataValidation(type="list", formula1=R_CODE, allow_blank=True,
                         showErrorMessage=True, errorTitle="班別代碼無效",
                         error="請從下拉選單選擇「設定」分頁定義的班別代碼。")
sc.add_data_validation(dv_code)
dv_code.add(f"{get_column_letter(DAY_C0)}{ROW0}:{get_column_letter(DAY_C1)}{ROW1}")
dv_emp = DataValidation(type="list", formula1=R_EID, allow_blank=True)
sc.add_data_validation(dv_emp); dv_emp.add(f"A{ROW0}:A{ROW1}")
dv_cli = DataValidation(type="list", formula1=f"設定!$C${CL_L0}:$C${CL_L1}", allow_blank=True)
sc.add_data_validation(dv_cli); dv_cli.add("B2")

D0 = get_column_letter(DAY_C0); D1L = get_column_letter(DAY_C1)
GRID = f"{D0}{ROW0}:{D1L}{ROW1}"
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}{ROW0}="")'], fill=GAP_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'{D0}{ROW0}="支"'], fill=SUPP_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'COUNTIF({R_LEAVE},{D0}{ROW0})>0'], fill=LEAVE_FILL, stopIfTrue=True))
sc.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'OR({D0}$5="六",{D0}$5="日")'], fill=WKND_FILL))
sc.conditional_formatting.add(f"{D0}4:{D1L}5", FormulaRule(
    formula=[f'OR({D0}$5="六",{D0}$5="日")'], fill=WKND_FILL))
sc.conditional_formatting.add(f"{D0}37:{D1L}37", FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}39>0,{D0}37<1)'], fill=ALERT_FILL))
sc.conditional_formatting.add(f"{D0}38:{D1L}38", FormulaRule(
    formula=[f'AND({D0}$4<>"",{D0}37>0,{D0}38<2)'], fill=ALERT_FILL))

# ================================================================ 4. 打卡匯入
pc = wb.create_sheet("打卡匯入")
pc.sheet_view.showGridLines = False
pc.freeze_panes = "A3"
for col, w in {"A":13,"B":13,"C":12,"D":12,"E":16,"F":16,"G":2,"H":72}.items():
    pc.column_dimensions[col].width = w
put(pc, "A1", "打卡匯入(把打卡機匯出的資料貼在下面四欄)", TITLE_F, border=False)
pc.merge_cells("A1:F1")
put(pc, "H1",
    "貼上規則:一天一列。日期要是真正的日期格式、時間要是真正的時間格式,"
    "不能是文字。E、F 欄是公式,不要覆蓋。",
    font(9, color="808080"), None, WRAP, border=False)
header_row(pc, 2, ["員工編號","日期","上班時間","下班時間","對照鍵(自動)","檢核(自動)"],
           start_col=1, height=24)

def build_punches():
    """依範例班表產生對應的打卡資料,並刻意放入五種異常。"""
    rnd = random.Random(20261001)
    rows = []
    for eid, *_ in PEOPLE:
        for d in range(1, DAYS_IN_MONTH + 1):
            code = SAMPLE[eid][d - 1]
            spec = CODE_BY[code]
            if spec[8] != "上班":
                continue
            date = dt.date(YEAR, MONTH, d)
            if spec[2] and spec[3]:
                start = (dt.datetime.combine(date, spec[2])
                         - dt.timedelta(minutes=rnd.randint(0, 7)))
                end = (dt.datetime.combine(date, spec[3])
                       + dt.timedelta(minutes=rnd.randint(0, 18)))
            else:                                   # 支援他院 / 教育訓練
                start = dt.datetime.combine(date, T(9, 0))
                end = dt.datetime.combine(date, T(18, 0))
            rows.append([eid, date, start.time(), end.time()])
    idx = {(r[0], r[1].day): i for i, r in enumerate(rows)}

    def first_day(eid, code, skip=0):
        hits = [d for d in range(1, DAYS_IN_MONTH + 1) if SAMPLE[eid][d - 1] == code]
        return hits[skip]

    # 遲到 18 分
    d = first_day("A001", "A"); rows[idx[("A001", d)]][2] = T(9, 18)
    # 早退 50 分
    d = first_day("A002", "P"); rows[idx[("A002", d)]][3] = T(20, 40)
    # 未打卡:整列刪除
    d = first_day("A003", "A", 2); rows.pop(idx[("A003", d)])
    idx = {(r[0], r[1].day): i for i, r in enumerate(rows)}
    # 加班:早班做到 20:30
    d = first_day("A004", "A", 1); rows[idx[("A004", d)]][3] = T(20, 30)
    # 假日出勤:醫師在休診日進診所
    d = [x for x in range(1, DAYS_IN_MONTH + 1) if SAMPLE["D001"][x - 1] == "休"][1]
    rows.append(["D001", dt.date(YEAR, MONTH, d), T(9, 0), T(13, 0)])
    rows.sort(key=lambda r: (r[1], r[0]))
    return rows

PUNCHES = build_punches()
for i in range(PUNCH_R0, PUNCH_R1 + 1):
    k = i - PUNCH_R0
    vals = PUNCHES[k] if k < len(PUNCHES) else [None] * 4
    for j, v in enumerate(vals):
        c = pc.cell(row=i, column=1 + j, value=v)
        c.font, c.border, c.alignment, c.fill = font(9), BOX, CTR, IN_FILL
        if j == 1: c.number_format = FMT_DATE
        if j in (2, 3): c.number_format = FMT_TIME
    e = pc.cell(row=i, column=5)
    e.value = f'=IF($A{i}="","",$A{i}&"|"&DAY($B{i}))'
    e.font, e.border, e.alignment, e.fill = font(9), BOX, CTR, CALC_FILL
    f = pc.cell(row=i, column=6)
    f.value = (f'=IF($A{i}="","",'
               f'IF(COUNTIF({R_EID},$A{i})=0,"⚠ 查無員編",'
               f'IF(COUNTIF($E${PUNCH_R0}:$E${PUNCH_R1},$E{i})>1,"⚠ 重複打卡","OK")))')
    f.font, f.border, f.alignment, f.fill = font(9), BOX, CTR, CALC_FILL
pc.auto_filter.ref = f"A2:F{PUNCH_R1}"
pc.conditional_formatting.add(f"A{PUNCH_R0}:F{PUNCH_R1}", FormulaRule(
    formula=[f'AND($A{PUNCH_R0}<>"",LEFT($F{PUNCH_R0},1)="⚠")'], fill=GAP_FILL))

# ================================================================ 5. 出勤紀錄
at = wb.create_sheet("出勤紀錄")
at.sheet_view.showGridLines = False
at.freeze_panes = "D3"
AT_COLS = [
    ("A", "日期",       11), ("B", "星期",  6), ("C", "員工編號", 11),
    ("D", "姓名",       11), ("E", "職類",  9), ("F", "排班代碼",  9),
    ("G", "應到",        8), ("H", "應退",  8), ("I", "實際上班",  10),
    ("J", "實際下班",   10), ("K", "休息(分)", 9), ("L", "實際工時", 10),
    ("M", "出勤異常",   11), ("N", "遲到(分)", 9), ("O", "早退(分)", 9),
    ("P", "加班時數",   10), ("Q", "加班別",  10), ("R", "備註",     22),
]
for col, _, w in AT_COLS:
    at.column_dimensions[col].width = w
put(at, "A1", "出勤紀錄(法定紀錄:逐日記載至分鐘,保存 5 年)", TITLE_F, border=False)
at.merge_cells("A1:R1")
header_row(at, 2, [lab for _, lab, _ in AT_COLS], start_col=1, height=28)

for p in range(N_PEOPLE):
    srow = ROW0 + p
    for d in range(1, DAYS_IN_MONTH + 1):
        r = ATT_R0 + p * DAYS_IN_MONTH + (d - 1)
        dcol = get_column_letter(DAY_C0 + d - 1)
        DUE_HR = f'IFERROR(INDEX({R_HRS},MATCH($F{r},{R_CODE},0)),0)'
        fx = {
"A": f'=IF(OR(班表!$A{srow}="",{d}>{DAYS_FX}),"",DATE(班表!$E$2,班表!$G$2,{d}))',
"B": f'=IF($A{r}="","",INDEX(設定!$L$5:$L$11,WEEKDAY($A{r},1)))',
"C": f'=IF($A{r}="","",班表!$A{srow})',
"D": f'=IF($C{r}="","",班表!$B{srow})',
"E": f'=IF($C{r}="","",班表!$C{srow})',
"F": f'=IF($C{r}="","",班表!{dcol}{srow})',
"G": f'=IF($F{r}="","",IFERROR(INDEX({R_IN},MATCH($F{r},{R_CODE},0)),""))',
"H": f'=IF($F{r}="","",IFERROR(INDEX({R_OUT},MATCH($F{r},{R_CODE},0)),""))',
"I": (f'=IFERROR(INDEX(打卡匯入!$C${PUNCH_R0}:$C${PUNCH_R1},'
      f'MATCH($C{r}&"|"&DAY($A{r}),打卡匯入!$E${PUNCH_R0}:$E${PUNCH_R1},0)),"")'),
"J": (f'=IFERROR(INDEX(打卡匯入!$D${PUNCH_R0}:$D${PUNCH_R1},'
      f'MATCH($C{r}&"|"&DAY($A{r}),打卡匯入!$E${PUNCH_R0}:$E${PUNCH_R1},0)),"")'),
"K": f'=IF(OR($I{r}="",$J{r}=""),"",IFERROR(INDEX({R_REST},MATCH($F{r},{R_CODE},0)),0))',
"L": f'=IF(OR($I{r}="",$J{r}=""),"",ROUND(($J{r}-$I{r})*24-$K{r}/60,2))',
"N": (f'=IF(OR($C{r}="",$G{r}="",$I{r}=""),0,'
      f'IF(ROUND(($I{r}-$G{r})*1440,0)>{P_GRACE},'
      f'ROUND(($I{r}-$G{r})*1440,0),0))'),
"O": (f'=IF(OR($C{r}="",$H{r}="",$J{r}=""),0,'
      f'IF(ROUND(($H{r}-$J{r})*1440,0)>{P_GRACE},'
      f'ROUND(($H{r}-$J{r})*1440,0),0))'),
"M": (f'=IF($C{r}="","",'
      f'IF(AND(COUNTIF({R_LEAVE},$F{r})>0,$I{r}<>""),"假日出勤",'
      f'IF(AND(COUNTIF({R_WORK},$F{r})>0,$I{r}=""),"未打卡",'
      f'IF($I{r}="","",'
      f'IF($N{r}>0,"遲到",IF($O{r}>0,"早退","正常"))))))'),
"P": (f'=IF($L{r}="",0,'
      f'IF(($L{r}-{DUE_HR})*60>={P_OT_MIN},'
      f'ROUND(FLOOR(($L{r}-{DUE_HR})*60,{P_OT_UNIT})/60,2),0))'),
"Q": (f'=IF(OR($C{r}="",$P{r}=0),"",'
      f'IF($F{r}="國","國定假日",'
      f'IF($B{r}="日","例假",IF($B{r}="六","休息日","平日"))))'),
"R": None,
        }
        for col, _, _w in AT_COLS:
            c = at[f"{col}{r}"]
            v = fx.get(col)
            if v is not None:
                c.value = v
                c.fill = CALC_FILL
            else:
                c.fill = IN_FILL
            c.font, c.border, c.alignment = font(9), BOX, CTR
            if col == "A": c.number_format = FMT_DATE
            elif col in ("G","H","I","J"): c.number_format = FMT_TIME
            elif col in ("L","P"): c.number_format = FMT_HR
            elif col in ("K","N","O"): c.number_format = FMT_MIN
            elif col == "R": c.alignment = LEFT
at.auto_filter.ref = f"A2:R{ATT_R1}"
AT_RANGE = f"A{ATT_R0}:R{ATT_R1}"
at.conditional_formatting.add(AT_RANGE, FormulaRule(
    formula=[f'AND($C{ATT_R0}<>"",$M{ATT_R0}<>"",$M{ATT_R0}<>"正常")'],
    fill=GAP_FILL, stopIfTrue=True))
at.conditional_formatting.add(AT_RANGE, FormulaRule(
    formula=[f'AND($C{ATT_R0}<>"",$P{ATT_R0}>0)'], fill=OT_FILL))
at.conditional_formatting.add(f"M{ATT_R0}:M{ATT_R1}", FormulaRule(
    formula=[f'AND($M{ATT_R0}<>"",$M{ATT_R0}<>"正常")'], fill=ALERT_FILL))

# ================================================================ 6. 月結統計
ms = wb.create_sheet("月結統計")
ms.sheet_view.showGridLines = False
ms.freeze_panes = "D5"
MS_COLS = [
    ("A","員工編號",11), ("B","姓名",11), ("C","職類",9), ("D","出勤天數",9),
    ("E","排班工時",9), ("F","總診次",8), ("G","排休",7), ("H","特休",7),
    ("I","病假",7), ("J","事假",7), ("K","公假",7), ("L","國定假日",9),
    ("M","支援他院",9), ("N","加班時數",10), ("O","實際工時",10),
    ("P","遲到次數",9), ("Q","早退次數",9), ("R","未打卡",8),
    ("S","排班完整度檢核",15),
]
for col, _, w in MS_COLS:
    ms.column_dimensions[col].width = w
put(ms, "A1", "月結統計(全自動,不需手填)", TITLE_F, border=False)
ms.merge_cells("A1:S1")
put(ms, "A2", "院所", font(10, True), SUB_FILL, CTR)
put(ms, "B2", "=班表!B2", font(10, True), CALC_FILL, CTR)
put(ms, "C2", "期間", font(10, True), SUB_FILL, CTR)
put(ms, "D2", '=班表!E2&"年"&班表!G2&"月"', font(10, True), CALC_FILL, CTR)
header_row(ms, 4, [lab for _, lab, _ in MS_COLS], start_col=1, height=30)

MS_R0 = 5
AT_C = f"出勤紀錄!$C${ATT_R0}:$C${ATT_R1}"
AT_M = f"出勤紀錄!$M${ATT_R0}:$M${ATT_R1}"
AT_P = f"出勤紀錄!$P${ATT_R0}:$P${ATT_R1}"
for i in range(N_PEOPLE):
    r = MS_R0 + i
    srow = ROW0 + i
    rng = f"班表!$D{srow}:$AH{srow}"
    g = f'IF($A{r}="","",'
    rows = {
"A": f'=IF(班表!$A{srow}="","",班表!$A{srow})',
"B": f'={g}班表!$B{srow})',
"C": f'={g}班表!$C{srow})',
"D": f'={g}SUMPRODUCT(COUNTIF({rng},{R_CODE}),{R_ATT}))',
"E": f'={g}SUMPRODUCT(COUNTIF({rng},{R_CODE}),{R_HRS}))',
"F": f'={g}SUMPRODUCT(COUNTIF({rng},{R_CODE}),{R_SES}))',
"G": f'={g}COUNTIF({rng},"OFF"))',
"H": f'={g}COUNTIF({rng},"特"))',
"I": f'={g}COUNTIF({rng},"病"))',
"J": f'={g}COUNTIF({rng},"事"))',
"K": f'={g}COUNTIF({rng},"公"))',
"L": f'={g}COUNTIF({rng},"國"))',
"M": f'={g}COUNTIF({rng},"支"))',
"N": f'={g}SUMIF({AT_C},$A{r},{AT_P}))',
"O": f'={g}$E{r}+$N{r})',
"P": f'={g}COUNTIFS({AT_C},$A{r},{AT_M},"遲到"))',
"Q": f'={g}COUNTIFS({AT_C},$A{r},{AT_M},"早退"))',
"R": f'={g}COUNTIFS({AT_C},$A{r},{AT_M},"未打卡"))',
"S": (f'={g}IF(COUNTA({rng})={DAYS_FX},"OK",'
      f'"⚠ 缺"&({DAYS_FX}-COUNTA({rng}))&"天"))'),
    }
    for col, _, _w in MS_COLS:
        c = ms[f"{col}{r}"]
        c.value, c.font, c.border, c.alignment, c.fill = (
            rows[col], font(), BOX, CTR, CALC_FILL)
        if col in ("E", "N", "O"): c.number_format = "0.0"

TOT = MS_R0 + N_PEOPLE
put(ms, f"A{TOT}", "合計", font(10, True), SUB_FILL, CTR)
ms.merge_cells(f"A{TOT}:C{TOT}")
for col in "DEFGHIJKLMNOPQR":
    c = ms[f"{col}{TOT}"]
    c.value = f"=SUM({col}{MS_R0}:{col}{TOT-1})"
    c.font, c.fill, c.border, c.alignment = font(10, True), SUB_FILL, BOX, CTR
    if col in ("E", "N", "O"): c.number_format = "0.0"
put(ms, f"S{TOT}", "", font(), SUB_FILL, CTR)
ms.conditional_formatting.add(f"S{MS_R0}:S{TOT-1}", FormulaRule(
    formula=[f'AND($A{MS_R0}<>"",S{MS_R0}<>"OK")'], fill=ALERT_FILL))
for col in ("P", "Q", "R"):
    ms.conditional_formatting.add(f"{col}{MS_R0}:{col}{TOT-1}", FormulaRule(
        formula=[f'AND($A{MS_R0}<>"",{col}{MS_R0}>0)'], fill=GAP_FILL))

NOTES = [
 "※ 出勤天數/排班工時/總診次:依「設定」分頁的班別代碼權重自動加總,改權重這裡就會跟著變。",
 "※ 加班時數:由「出勤紀錄」P 欄加總而來。不使用打卡匯入的院所,可直接把 N 欄公式覆蓋為手動數字。",
 "※ 加班時數為「實際工時 − 排班工時」,尚未依勞基法第 24 條換算加班費率,薪資請另行計算。",
 "※ 遲到/早退/未打卡次數:由「出勤紀錄」的出勤異常欄計次,有值會標色,月結前應逐筆處理完。",
 "※ 排班完整度檢核顯示「⚠ 缺 N 天」代表該員當月還有 N 天沒排班,交表前要補完。",
]
for i, t in enumerate(NOTES):
    put(ms, f"A{TOT+2+i}", t, font(9, color="808080"), None, LEFT, border=False)

wb.save(OUT)
print("saved:", OUT)
print(f"出勤紀錄列數: {ATT_R1 - ATT_R0 + 1}  打卡範例筆數: {len(PUNCHES)}")
