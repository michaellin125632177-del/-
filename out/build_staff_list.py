# -*- coding: utf-8 -*-
"""日三牙醫體系 員工名冊。

資料來源是 build_roster.py 的 DOCTORS / NURSES / CLINICS,不另外維護一份,
免得兩邊對不起來。改名冊請改 build_roster.py,再重跑這支。

    python3 build_staff_list.py

主表全部寫實值——這份是要給人看的,不能在手機預覽或雲端硬碟預覽裡整欄空白。
統計格用公式(加人之後會自己更新),但另外把算好的數字注入 xlsx 的快取值,
所以不重算公式的檢視器也看得到數字,Excel 開啟時仍會重算一次。
"""
import importlib.util, sys, io, contextlib, re, shutil, zipfile, datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

OUT = "/home/user/-/out/日三牙醫體系_員工名冊.xlsx"
F = "微軟正黑體"

# ---------------------------------------------------------------- 讀資料
_spec = importlib.util.spec_from_file_location("br", "/home/user/-/out/build_roster.py")
_br = importlib.util.module_from_spec(_spec)
_argv, sys.argv = sys.argv, ["build_roster.py"]
with contextlib.redirect_stdout(io.StringIO()):
    _spec.loader.exec_module(_br)
sys.argv = _argv
DOCTORS, NURSES, CLINICS = _br.DOCTORS, _br.NURSES, _br.CLINICS

def font(sz=10, b=False, color="000000"):
    return Font(name=F, size=sz, bold=b, color=color)

def fill(hex_):
    # fgColor 與 bgColor 都寫、alpha 補 FF:一般底色看 fgColor,
    # 條件式格式的差異格式看 bgColor,兩邊都填才不會有一邊是白的。
    argb = "FF" + hex_ if len(hex_) == 6 else hex_
    return PatternFill("solid", start_color=argb, end_color=argb)

HDR_FILL, SUB_FILL, IN_FILL = fill("1F4E5F"), fill("DCE6EC"), fill("FFF2CC")
ALT_FILL = fill("FAFAF8")
CLINIC_FILL = {"晶悅": fill("FBE3EC"), "晶睿": fill("D9E7F5"), "晶匯": fill("DCEEDC"),
               "晶曜": fill("E8DFF2"), "寶貝牙": fill("FAE8A0")}
_t = Side(style="thin", color="C8CCC8")
BOX = Border(left=_t, right=_t, top=_t, bottom=_t)
CTR = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

# ---------------------------------------------------------------- 整理成一張表
# 醫師的「主要院所」不拿名單第一間充數——那只是排序。直接數門診表:
# 每週診次最多的那間才是主要院所,院長則以所轄院所為準。
import weekly as _wk
SHORT = {c[0]: c[1] for c in CLINICS}          # 悅 → 晶悅
SESSIONS = {}                                  # 姓名 → {院所: 每週診次}
for _code, _tbl in _wk.W.items():
    for _slot, _people in _tbl.items():
        for _nm, _flag in _people:
            SESSIONS.setdefault(_nm, {})
            SESSIONS[_nm][SHORT[_code]] = SESSIONS[_nm].get(SHORT[_code], 0) + 1

ROWS = []          # 員工編號,姓名,英文名,職類,職務,專科,主要院所,服務院所,備註
for eid, nm, eng, spec, duty, teams in DOCTORS:
    if duty.endswith("院長"):
        title, home = "院長", duty[:-2]
    else:
        title = ""
        by = SESSIONS.get(nm, {})
        # 診次最多的那間;同票就照服務院所的排序決定,結果才會穩定
        home = max(teams, key=lambda t: (by.get(t, 0), -teams.index(t))) if teams else ""
    note = "官網醫療團隊頁未收錄,資料待補" if not eng and not spec else ""
    ROWS.append([eid, nm, eng, "醫師", title, spec, home, " · ".join(teams), note])
for eid, nm, kind, home, _hire, note in NURSES:
    clinics = home.split("·")
    ROWS.append([eid, nm, "", kind, "", "", clinics[0], " · ".join(clinics), note])

R0 = 4                                   # 資料第一列
R1 = R0 + len(ROWS) - 1
COLS = [("序號", 6), ("員工編號", 11), ("姓名", 11), ("英文名", 19), ("職類", 9),
        ("職務", 8), ("專科", 22), ("主要院所", 11), ("服務院所", 27),
        ("跨院數", 8), ("到職日", 13), ("備註", 30)]
NC = len(COLS)
LASTC = get_column_letter(NC)

wb = Workbook()
ws = wb.active
ws.title = "員工名冊"
ws.sheet_view.showGridLines = False
ws.freeze_panes = f"A{R0}"
for i, (lab, w) in enumerate(COLS, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

ws["A1"] = f"日三牙醫體系 員工名冊({dt.date.today():%Y/%m/%d} 製)"
ws["A1"].font, ws["A1"].alignment = font(16, True), LEFT
ws.merge_cells(f"A1:{LASTC}1")
ws.row_dimensions[1].height = 30

# 第 2 列:人數統計。用公式,加人之後會自己更新。
SUMMARY = [
    ("B2", "總人數",  f'=COUNTA($B${R0}:$B${R1})'),
    ("D2", "醫師",    f'=COUNTIF($D${R0}:$D${R1},"醫師")'),
    ("F2", "醫護長",  f'=COUNTIF($D${R0}:$D${R1},"醫護長")'),
    ("H2", "管理部",  f'=COUNTIF($D${R0}:$D${R1},"管理部")'),
    ("J2", "院長",    f'=COUNTIF($E${R0}:$E${R1},"院長")'),
]
ws["A2"] = "人數"
ws["A2"].font, ws["A2"].fill, ws["A2"].alignment = font(10, True), SUB_FILL, CTR
for anchor, lab, fx in SUMMARY:
    col = anchor[0]
    lc = ws[f"{col}2"]
    lc.value, lc.font, lc.fill, lc.alignment = lab, font(9, color="4A5A5C"), SUB_FILL, CTR
    vcol = get_column_letter(ws[f"{col}2"].column + 1)
    vc = ws[f"{vcol}2"]
    vc.value, vc.font, vc.fill, vc.alignment = fx, font(11, True), SUB_FILL, CTR
    vc.number_format = "0"
for i in range(1, NC + 1):                       # 統計列整列鋪底
    ws.cell(row=2, column=i).fill = SUB_FILL
ws.row_dimensions[2].height = 22

for i, (lab, _w) in enumerate(COLS, start=1):
    c = ws.cell(row=3, column=i, value=lab)
    c.font, c.fill, c.alignment, c.border = font(10, True, "FFFFFF"), HDR_FILL, CTR, BOX
ws.row_dimensions[3].height = 24

for k, row in enumerate(ROWS):
    r = R0 + k
    eid, nm, eng, kind, title, spec, home, teams, note = row
    # 跨院數不寫死:數「服務院所」欄裡有幾個間隔點,改了院所會跟著變
    cross = (f'=IF($I{r}="","",LEN($I{r})-LEN(SUBSTITUTE($I{r},"·",""))+1)')
    vals = [k + 1, eid, nm, eng, kind, title, spec, home, teams, cross, None, note]
    for i, v in enumerate(vals, start=1):
        c = ws.cell(row=r, column=i, value=v)
        c.font, c.border, c.alignment = font(9.5), BOX, CTR
        if i in (4, 7, 9, 12): c.alignment = LEFT
        if k % 2 == 1: c.fill = ALT_FILL
    ws.cell(row=r, column=8).fill = CLINIC_FILL.get(home, ALT_FILL if k % 2 else PatternFill())
    hire = ws.cell(row=r, column=11)             # 到職日:要人資填的那一欄
    hire.fill, hire.number_format = IN_FILL, "yyyy/m/d"
    ws.cell(row=r, column=10).number_format = "0"

NOTE0 = R1 + 2
NOTES = [
    "【怎麼用】黃底的「到職日」是唯一要手填的欄位,請填真正的日期格式(例：2020/3/15),"
    "不要填成文字。填好之後特休年資才算得出來——目前 34 人全部空白。",
    "【自動欄位】第 2 列的人數與「跨院數」是公式,加人或改「服務院所」會自己更新,不必手改。",
    "【資料來源】醫師的姓名、英文名、專科、服務院所取自 gemray.tw 官網醫療團隊頁與五間院所門診表;"
    "醫護長與管理部由小編提供。",
    "【主要院所】醫師是依門診表算出每週診次最多的那一間,不是名單排序;"
    "院長則以所轄院所為準(陳昺元晶悅與晶匯各 5 診次,依院長職務歸晶匯)。"
    "跨院醫師在「服務院所」欄可以看到全部院所。",
    "【待確認】D025 陳爵安、D026 吳冠廷、D027 楊孟庭三位只出現在門診表,官網醫療團隊頁沒有,"
    "英文名與專科待補。",
    "【員工編號】D／N／M 三碼為本表暫編,非人事系統編號。若人事或打卡系統另有編號,應以那套為準。",
    "【個資】本表含員工個人資料,請依個資法限定使用目的並控管取用權限。",
]
for i, t in enumerate(NOTES):
    c = ws.cell(row=NOTE0 + i, column=1, value=t)
    c.font, c.alignment = font(9, color="6E7F81"), WRAP
    ws.merge_cells(start_row=NOTE0 + i, start_column=1,
                   end_row=NOTE0 + i, end_column=NC)
    ws.row_dimensions[NOTE0 + i].height = 15
ws.auto_filter.ref = f"A3:{LASTC}{R1}"
ws.print_title_rows = "3:3"

# ---------------------------------------------------------------- 院所對照
cs = wb.create_sheet("院所對照")
cs.sheet_view.showGridLines = False
CCOLS = [("代碼", 6), ("簡稱", 10), ("全名", 20), ("院長", 10), ("電話", 14),
         ("地址", 34), ("醫師", 7), ("醫護長", 8), ("管理部", 8), ("合計", 7)]
for i, (lab, w) in enumerate(CCOLS, start=1):
    cs.column_dimensions[get_column_letter(i)].width = w
cs["A1"] = "院所對照與人數"
cs["A1"].font, cs["A1"].alignment = font(14, True), LEFT
cs.merge_cells(f"A1:{get_column_letter(len(CCOLS))}1")
cs.row_dimensions[1].height = 26
for i, (lab, _w) in enumerate(CCOLS, start=1):
    c = cs.cell(row=2, column=i, value=lab)
    c.font, c.fill, c.alignment, c.border = font(10, True, "FFFFFF"), HDR_FILL, CTR, BOX
RNG_TEAM = f"員工名冊!$I${R0}:$I${R1}"
RNG_KIND = f"員工名冊!$D${R0}:$D${R1}"
for k, (code, short, full, head, tel, addr) in enumerate(CLINICS):
    r = 3 + k
    # 一個人可能服務多間院所(醫師跨院、文君兼管兩間),所以用包含比對
    def cnt(kind):
        return (f'=SUMPRODUCT(ISNUMBER(SEARCH("{short}",{RNG_TEAM}))'
                f'*({RNG_KIND}="{kind}"))')
    vals = [code, short, full, head, tel, addr,
            cnt("醫師"), cnt("醫護長"), cnt("管理部"),
            f"=SUM(G{r}:I{r})"]
    for i, v in enumerate(vals, start=1):
        c = cs.cell(row=r, column=i, value=v)
        c.font, c.border, c.alignment = font(9.5), BOX, CTR
        if i in (3, 6): c.alignment = LEFT
        if i >= 7: c.number_format = "0"
    cs.cell(row=r, column=2).fill = CLINIC_FILL[short]
n0 = 3 + len(CLINICS) + 1
for i, t in enumerate([
    "※ 人數是用「服務院所」欄做包含比對算的,所以跨院醫師會在每一間都被計入一次,"
    "各院所合計會大於總人數 34。要看不重複人數請以「員工名冊」第 2 列為準。",
    "※ 文君兼管晶睿與晶曜,是同一位、一個員工編號,兩間都會算到她。",
]):
    c = cs.cell(row=n0 + i, column=1, value=t)
    c.font, c.alignment = font(9, color="6E7F81"), WRAP
    cs.merge_cells(start_row=n0 + i, start_column=1,
                   end_row=n0 + i, end_column=len(CCOLS))
wb.save(OUT)

# ---------------------------------------------------------------- 注入快取值
# openpyxl 寫出來的公式沒有快取值,不重算公式的檢視器(手機預覽、雲端硬碟預覽)
# 會顯示空白。這裡把每個公式算好的結果寫進 xlsx 的 <v>,顯示就正常了;
# Excel 開檔仍會依 fullCalcOnLoad 重算一次,所以不會蓋掉真正的計算結果。
def cross_count(teams):
    return teams.count("·") + 1 if teams else 0

CACHE = {"員工名冊": {}, "院所對照": {}}
CACHE["員工名冊"]["C2"] = len(ROWS)
CACHE["員工名冊"]["E2"] = sum(1 for r in ROWS if r[3] == "醫師")
CACHE["員工名冊"]["G2"] = sum(1 for r in ROWS if r[3] == "醫護長")
CACHE["員工名冊"]["I2"] = sum(1 for r in ROWS if r[3] == "管理部")
CACHE["員工名冊"]["K2"] = sum(1 for r in ROWS if r[4] == "院長")
for k, row in enumerate(ROWS):
    CACHE["員工名冊"][f"J{R0+k}"] = cross_count(row[7])
for k, (code, short, *_r) in enumerate(CLINICS):
    r = 3 + k
    per = {}
    for kind in ("醫師", "醫護長", "管理部"):
        per[kind] = sum(1 for row in ROWS if row[3] == kind and short in row[7])
    CACHE["院所對照"][f"G{r}"] = per["醫師"]
    CACHE["院所對照"][f"H{r}"] = per["醫護長"]
    CACHE["院所對照"][f"I{r}"] = per["管理部"]
    CACHE["院所對照"][f"J{r}"] = sum(per.values())

sheet_of = {"員工名冊": "xl/worksheets/sheet1.xml",
            "院所對照": "xl/worksheets/sheet2.xml"}
src = OUT + ".tmp"
shutil.move(OUT, src)
zin = zipfile.ZipFile(src)
with zipfile.ZipFile(OUT, "w", zipfile.ZIP_DEFLATED) as zout:
    for item in zin.infolist():
        data = zin.read(item.filename)
        rev = {v: k for k, v in sheet_of.items()}
        if item.filename in rev:
            x = data.decode("utf-8")
            for ref, val in CACHE[rev[item.filename]].items():
                # openpyxl 會自己補一個空的 <v />,要連那個一起換掉
                pat = re.compile(
                    rf'(<c r="{ref}"[^>]*>)(<f>.*?</f>)(?:<v\s*/>|<v>.*?</v>)?(</c>)')
                m = pat.search(x)
                assert m, f"{rev[item.filename]}!{ref} 找不到公式格"
                x = pat.sub(rf'\1\2<v>{val}</v>\3', x, count=1)
            data = x.encode("utf-8")
        zout.writestr(item, data)
zin.close()
import os; os.remove(src)

print(f"saved: {OUT}")
print(f"員工名冊:{len(ROWS)} 人"
      f"(醫師 {sum(1 for r in ROWS if r[3]=='醫師')}、"
      f"醫護長 {sum(1 for r in ROWS if r[3]=='醫護長')}、"
      f"管理部 {sum(1 for r in ROWS if r[3]=='管理部')}、"
      f"其中院長 {sum(1 for r in ROWS if r[4]=='院長')} 位)")
print(f"注入快取值:{sum(len(v) for v in CACHE.values())} 格")
