# -*- coding: utf-8 -*-
"""日三牙醫體系 統一出勤表
分頁:說明 / 設定 / 醫師週班表 / 醫師班表 / 醫師月結 / 醫護長與管理部班表 / 打卡匯入 / 出勤紀錄 / 月結統計

架構重點:全體系一個檔。
  醫師走「診次制」——一位醫師一列,橫向 31 天 × 早/午/晚,格內填院所代碼。
  一格只容得下一間院所,同一診次被排到兩間院所在結構上就不可能發生。
  醫護長與管理部走「工時制」——班表 + 打卡 + 逐日出勤紀錄。本表不含助理。
"""
import datetime as dt
import sys
import pathlib as _pathlib, os
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from weekly import (W as WEEKLY_RAW, SESSION_TIME, NOTES, SPEC as WEB_SPEC,
                    HOLIDAYS_2026)
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

F = "微軟正黑體"
# 加上 --no-assistant 參數會產出「醫師版」:拿掉工時制那四個分頁。
# 醫護長與管理部班表被出勤紀錄與月結統計讀取,打卡匯入只被出勤紀錄讀取,
# 四頁互相牽連,只拔其中一頁會讓其餘變成 #REF!,所以要拔就是整組拔。
WITH_ASSISTANT = "--no-assistant" not in sys.argv
STAFF_SHEETS = ["醫護長與管理部班表", "打卡匯入", "出勤紀錄", "月結統計"]
# 輸出放在腳本自己所在的資料夾,不寫死絕對路徑——這支要在別人的電腦上跑。
HERE = _pathlib.Path(__file__).resolve().parent
OUT = str(HERE / ("日三牙醫體系_統一出勤表.xlsx" if WITH_ASSISTANT
                  else "日三牙醫體系_統一出勤表_醫師版.xlsx"))

# 本期年月。換月不必改程式:
#     GEMRAY_PERIOD=2026-11 python3 build_roster.py
# 沒給就用下面的預設值。天數一律由月份推算——寫死 31 會讓 11 月多出一天。
import os as _os, calendar as _cal
_p = _os.environ.get("GEMRAY_PERIOD", "2026-10").strip()
try:
    YEAR, MONTH = (int(x) for x in _p.replace("/", "-").split("-")[:2])
except ValueError:
    raise SystemExit(f"GEMRAY_PERIOD 格式應為 YYYY-MM,收到:{_p!r}")
assert 1 <= MONTH <= 12, f"月份不合法:{MONTH}"
DAYS_IN_MONTH = _cal.monthrange(YEAR, MONTH)[1]
SESSIONS = ["早", "午", "晚"]

N_DOC   = 40      # 醫師班表列數(24 位 + 保留)
N_ASST_SLOTS = 12                 # 工時制人員列數(醫護長 4 + 管理部 3 + 5 保留)
N_ASST  = N_ASST_SLOTS
PUNCH_N = 500     # 打卡匯入資料列數

# ---------------------------------------------------------------- 樣式
def font(sz=10, b=False, color="000000"):
    return Font(name=F, size=sz, bold=b, color=color)

TITLE_F   = font(16, True)
HDR_F     = font(10, True, "FFFFFF")
def fill(hex_):
    """實心底色。fgColor 與 bgColor 兩個都填,是因為 Excel 讀這兩者的規則不一樣:

    一般儲存格的實心底色看 fgColor,但條件式格式用的是「差異格式」(dxf),
    Excel 在 dxf 裡讀的是 bgColor。openpyxl 只寫 fgColor,所以同一個 PatternFill
    當一般底色會上色、當條件式格式卻是白的——本檔所有條件式格式的顏色
    先前都沒有顯示,原因就在這裡。兩個都寫,兩種情況都吃得到。
    """
    # 另外補上 FF 的 alpha:openpyxl 收到 6 碼色碼會寫成 "00RRGGBB",
    # 也就是完全透明。一般底色 Excel 會忽略這個 alpha,dxf 不見得。
    argb = "FF" + hex_ if len(hex_) == 6 else hex_
    return PatternFill("solid", start_color=argb, end_color=argb)

HDR_FILL  = fill("1F4E5F")
SUB_FILL  = fill("DCE6EC")
IN_FILL   = fill("FFF2CC")
CALC_FILL = fill("F2F2F2")
WKND_FILL = fill("DDEBF7")
LEAVE_FILL= fill("D9D9D9")
ALERT_FILL= fill("FF9999")
GAP_FILL  = fill("FCE4E4")
OT_FILL   = fill("FCE4D6")
SUPP_FILL = fill("C6E0B4")   # 醫護長與管理部班表的「支援他院」專用綠
# 五間院所各自的底色,讓醫師整月動線一眼看得出來
CLINIC_FILL = {
    "悅": fill("FBE3EC"),   # 粉
    "睿": fill("D9E7F5"),   # 藍
    "匯": fill("DCEEDC"),   # 綠
    "曜": fill("E8DFF2"),   # 紫
    # 寶貝牙的黃刻意比 IN_FILL(FFF2CC,「你要填的格子」)飽和,兩者不會混淆
    "寶": fill("FAE8A0"),   # 黃
}

thin = Side(style="thin", color="AAAAAA")
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)
DAYSEP = Border(left=Side(style="medium", color="7F7F7F"),
                right=thin, top=thin, bottom=thin)
CTR  = Alignment(horizontal="center", vertical="center")
LEFT = Alignment(horizontal="left", vertical="center")
WRAP = Alignment(horizontal="left", vertical="top", wrap_text=True)

FMT_TIME = "hh:mm"; FMT_HR = '0.00;-0.00;"–"'
FMT_MIN  = '0;-0;"–"'; FMT_DATE = "yyyy/m/d"; FMT_CNT = '0;-0;"–"'

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

# ================================================================ 主資料
# 院所:代碼, 簡稱, 全名, 院長, 電話, 地址
CLINICS = [
    ("悅", "晶悅",   "晶悅牙醫診所",   "林智俊", "03-4257710", "320桃園市中壢區新生路251號"),
    ("睿", "晶睿",   "晶睿牙醫診所",   "温健成", "03-4522360", "桃園市中壢區福州二街502號"),
    ("匯", "晶匯",   "晶匯牙醫診所",   "陳昺元", "03-4550719", "桃園市中壢區成章一街172號"),
    ("曜", "晶曜",   "晶曜牙醫診所",   "劉立德", "03-3150666", "桃園市桃園區中正路1398-1號"),
    ("寶", "寶貝牙", "寶貝牙牙醫診所", "黃育亭", "03-2871688", "桃園市大園區新生路四段273號2樓"),
]
CLINIC_CODES = [c[0] for c in CLINICS]
SHORT2CODE   = {c[1]: c[0] for c in CLINICS}

# 醫師:編號, 姓名, 英文名, 專科, 院長職務, 服務院所(簡稱)
DOCTORS = [
 ("D001","林智俊","JASON LIN","牙周植牙專科","晶悅院長",["晶悅","晶睿","晶匯","晶曜"]),
 ("D002","劉立德","LEADER LIU","植牙專科 家庭牙醫專科","晶曜院長",["晶睿","晶曜"]),
 ("D003","王泳泉","RYAN WANG","矯正專科","",["晶悅","晶睿"]),
 ("D004","劉玠旻","CHIEH-MIN LIU","兒童專科","",["晶悅","晶睿","晶曜"]),
 ("D005","陳志僑","CHIH-CHIAO CHEN","口外專科","",["晶悅","晶睿"]),
 ("D006","黃榆鈞","YU-CHUN HUANG","矯正專科","",["晶悅","寶貝牙"]),
 ("D007","温翊君","YI-CHUN WEN","根管治療專科","",["晶悅"]),
 ("D008","林紫亭","TZU-TING LIN","","",["晶悅"]),
 ("D009","陳昺元","PING-YUAN CHEN","美學植牙專科","晶匯院長",["晶悅","晶匯","寶貝牙"]),
 ("D010","蔡季軒","CHI-HSUAN TSAI","","",["晶悅"]),
 ("D011","温健成","JIAN-CHENG WEN","","晶睿院長",["晶睿","晶匯"]),
 ("D012","陳立軒","LI-HSUAN CHEN","","",["晶睿","晶匯","晶曜"]),
 ("D013","朱柏非","PO-FEI CHU","","",["晶悅","晶匯","寶貝牙"]),
 ("D014","鄭漢賢","HON-YIN CHENG","","",["晶悅","晶匯"]),
 ("D015","徐平","PING HSU","","",["晶悅","晶曜"]),
 ("D016","賴敏傑","MIN-CHIEH LAI","","",["晶睿"]),
 ("D017","吳柏賢","BO-XIAN WU","","",["晶睿","晶曜"]),
 ("D018","許博諺","PO-YEN HSU","","",["晶睿","晶曜"]),
 ("D019","翁崎紘","CHI-HONG WENG","","",["晶悅"]),
 ("D020","余柏萱","PO-HSUAN YU","","",["晶睿","寶貝牙"]),
 ("D021","鄭博文","PO-WEN CHENG","","",["晶悅"]),
 ("D022","黃育亭","YU-TING HUANG","","寶貝牙院長",["寶貝牙"]),
 ("D023","黃婷愉","TING-YU HUANG","","",["寶貝牙"]),
 ("D024","林紓安","SHU-AN LIN","","",["晶匯","寶貝牙"]),
 # 以下三位只出現在門診表,官網醫療團隊頁的 24 位名單中沒有,英文名與專科待補
 ("D025","陳爵安","","","",["晶睿"]),
 ("D026","吳冠廷","","","",["晶匯"]),
 ("D027","楊孟庭","","","",["晶匯"]),
]
# 隔週輪替的基準日:從這個星期一起連續數週,單數週為 A 組、雙數週為 B 組。
# 用連續週序而不是「當月第幾個星期幾」,否則跨月時輪替會斷掉
# (例:2026/10/31 與 11/7 會連續兩週派到同一間)。
# ⚠ 哪一組先目前無從得知。院所確認後,把基準日往前或往後挪一週即可整體對齊。
ALT_EPOCH = dt.date(2026, 9, 28)          # 2026 年第 40 週的星期一

# 隔週互換的四組:同一位醫師同一時段在兩間院所輪替,列在這裡的排在 B 組。
ALT_B = {("劉立德",3,"曜"), ("朱柏非",6,"匯"), ("王泳泉",1,"睿"), ("陳昺元",6,"匯")}

# 官網門診表的 icon 有標專科,醫療團隊頁沒列到的補進來
DOCTORS = [
    (eid, nm, eng, spec or WEB_SPEC.get(nm, ""), duty, teams)
    for eid, nm, eng, spec, duty, teams in DOCTORS
]

# 國定假日:全院休診,班表填「國」
HOLIDAY_SET = {dt.date(YEAR, m, d): name for m, d, name, _mk in HOLIDAYS_2026
               if m and d}

# 工時制人員:醫護長 + 管理部。助理已移出本表範圍。
# 文君兼管晶睿與晶曜——同一位、一個員工編號、一份打卡紀錄,不是兩個人。
# 管理部三位都在晶匯上班,院所欄就寫晶匯。但底下的各院所在班人數檢核只數醫護長:
# 那一列問的是「這間院所今天有沒有醫護長」,管理部在不在都不能代替。
NURSES = [
    ("N001", "ivy",  "醫護長", "晶悅",      "", ""),
    ("N002", "文君", "醫護長", "晶睿·晶曜", "", "兼管晶睿與晶曜"),
    ("N003", "小玲", "醫護長", "晶匯",      "", ""),
    ("N004", "娜娜", "醫護長", "寶貝牙",    "", ""),
    ("M001", "孟諭", "管理部", "晶匯",      "", ""),
    ("M002", "怡雯", "管理部", "晶匯",      "", ""),
    ("M003", "小華", "管理部", "晶匯",      "", ""),
]
ASSISTANTS = NURSES

# 班別代碼。醫護長與管理部共用同一套代碼與同一套時段。
#   代碼,名稱,計出勤,類別
STAFF_TYPES = ["醫護長", "管理部"]
WORK_CODES = [
    ("早",   "早班",          1, "上班"),
    ("午",   "午班",          1, "上班"),
    ("晚",   "晚班",          1, "上班"),
    ("早午", "早班+午班",     1, "上班"),
    ("午晚", "午班+晚班",     1, "上班"),
    ("早晚", "早班+晚班",     1, "上班"),
    ("全日", "早+午+晚",      1, "上班"),
    ("支",   "支援他院",      1, "上班"),
    ("訓",   "教育訓練/會議",  1, "上班"),
    ("OFF", "排休",          0, "休假"),
    ("特",   "特休",          0, "休假"),
    ("病",   "病假",          0, "休假"),
    ("事",   "事假",          0, "休假"),
    ("公",   "公假",          0, "休假"),
    ("國",   "國定假日",      0, "休假"),
    ("休",   "休診",          0, "休假"),
]
N_WORK_CODE = 9                   # 上班類代碼數(早…訓),其餘為休假類

# 代碼 → (應到, 應退, 休息分, 排班工時)。醫護長與管理部共用同一套時段。
# 「休息(分)」= 診次之間不算工時的空檔,滿足「應退 − 應到 − 休息 = 排班工時」。
# 支/訓 沒有固定時段(依實際),但仍給排班工時,否則加班會從 0 起算。
# 日後若某個職類要拆出自己的時段,把這張表改成 {職類: {代碼: ...}},
# 設定分頁多一欄「職類」與一欄比對鍵、出勤紀錄的 MATCH 改用「職類|代碼」即可。
SHIFT_TIME = {
    "早":   (T(9,0),  T(12,0),   0, 3.0),
    "午":   (T(14,0), T(17,0),   0, 3.0),
    "晚":   (T(18,0), T(21,0),   0, 3.0),
    "早午": (T(9,0),  T(17,0),  120, 6.0),
    "午晚": (T(14,0), T(21,0),   60, 6.0),
    "早晚": (T(9,0),  T(21,0),  360, 6.0),
    "全日": (T(9,0),  T(21,0),  180, 9.0),
    "支":   (None, None, 0, 6.0),
    "訓":   (None, None, 0, 6.0),
}

# 設定分頁的班別代碼表:一個代碼一列。
WC_TABLE = []
for code, name, att, kind in WORK_CODES:
    tin, tout, rest, hrs = SHIFT_TIME.get(code, (None, None, 0, 0.0))
    WC_TABLE.append([code, name, tin, tout, rest, hrs, att, kind])

for code, name, tin, tout, rest, hrs, att, kind in WC_TABLE:
    if tin and tout:                      # 工時定義一致性檢查
        span = (dt.datetime.combine(dt.date.min, tout)
                - dt.datetime.combine(dt.date.min, tin)).seconds / 3600
        assert abs(span - rest/60 - hrs) < 1e-9, f"代碼「{code}」工時定義不一致"
    if kind == "休假":
        assert not tin and not tout and hrs == 0, f"代碼「{code}」是休假卻有工時"

DOC_CODES = ([(c[0], f"{c[1]}牙醫") for c in CLINICS] +
             [("訓", "教育訓練 / 學會"), ("OFF", "排休"), ("特", "特休"),
              ("病", "病假"), ("事", "事假"), ("公", "公假"),
              ("國", "國定假日"), ("休", "休診")])
DOC_LEAVE = [c[0] for c in DOC_CODES if c[0] not in CLINIC_CODES and c[0] != "訓"]

def wk_label(hit):
    """週班表格子的顯示寫法。hit = [(院所代碼, 標記), ...]

    悅        每週固定
    悅(隔)    隔週看診
    悅/睿     兩院輪替(斜線前 = 單數週)
    寶(不定)  官網註明非每週固定
    悅(註)    有附加條件,詳見附註欄
    """
    if len(hit) > 1:
        return "/".join(cl for cl, _ in hit)
    cl, fg = hit[0]
    if "~" in fg: return f"{cl}(隔)"
    if "*" in fg: return f"{cl}(不定)"
    if "!" in fg: return f"{cl}(註)"
    return cl

# ================================================================ 設定分頁座標
SET_Y, SET_M = "設定!$C$2", "設定!$D$2"
# 座標全部由上一區塊推算,增減代碼不會再壓到下面的區塊
WC_R0    = 6                                    # 一、班別代碼:表頭 5、資料 6 起
WC_R1    = WC_R0 + len(WC_TABLE) - 1
WC_WORK1 = WC_R0 + N_WORK_CODE - 1              # 上班類最後一列
DC_R0    = WC_R1 + 7                            # 二、醫師診次代碼:標題 -2、表頭 -1
                                                # (+1~+3 是班別時段表底下的三行註記)
DC_R1    = DC_R0 + len(DOC_CODES) - 1
PM_R0    = DC_R1 + 5                            # 三、計算參數:標題 -1、表頭 PM_R0、值 +1 起
CL_R0    = PM_R0 + 8                            # 四、院所清單:標題 -1、表頭 CL_R0
CL_L0    = CL_R0 + 1
CL_L1    = CL_L0 + 6
PP_R0    = CL_L1 + 3                            # 五、人員名冊
PP_L0    = PP_R0 + 1
PP_L1    = PP_L0 + 139

# 班別代碼表欄位:B 代碼 C 名稱 D 應到 E 應退 F 休息 G 排班工時 H 計出勤 I 類別
R_WC     = f"設定!$B${WC_R0}:$B${WC_R1}"
R_WC_W   = f"設定!$B${WC_R0}:$B${WC_WORK1}"
R_WC_L   = f"設定!$B${WC_WORK1+1}:$B${WC_R1}"
R_WC_IN  = f"設定!$D${WC_R0}:$D${WC_R1}"
R_WC_OUT = f"設定!$E${WC_R0}:$E${WC_R1}"
R_WC_RST = f"設定!$F${WC_R0}:$F${WC_R1}"
R_WC_HRS = f"設定!$G${WC_R0}:$G${WC_R1}"
R_WC_ATT = f"設定!$H${WC_R0}:$H${WC_R1}"
R_DC     = f"設定!$B${DC_R0}:$B${DC_R1}"
P_OT_MIN  = f"設定!$C${PM_R0+1}"
P_OT_UNIT = f"設定!$C${PM_R0+2}"
P_GRACE   = f"設定!$C${PM_R0+3}"
P_EPOCH   = f"設定!$C${PM_R0+4}"                 # 隔週基準日
R_EID  = f"設定!$B${PP_L0}:$B${PP_L1}"
R_NAME = f"設定!$C${PP_L0}:$C${PP_L1}"
R_ROLE = f"設定!$D${PP_L0}:$D${PP_L1}"
R_SPEC = f"設定!$E${PP_L0}:$E${PP_L1}"
R_DUTY = f"設定!$F${PP_L0}:$F${PP_L1}"
R_HOME = f"設定!$G${PP_L0}:$G${PP_L1}"
R_WEEK = "設定!$O$6:$O$12"
R_HOL  = "設定!$W$6:$W$40"                       # 國定假日清單(W 欄,與上列區塊不重疊)
R_OPEN = f"設定!$U${CL_L0}:$AO${CL_L0+4}"        # 五間院所的開診時段旗標(含週日)
DAYS_FX = f"DAY(EOMONTH(DATE({SET_Y},{SET_M},1),0))"

# ── 門診表 → 每位醫師的每週時段 ────────────────────────────
DOC_WEEK = {}                       # 姓名 -> {(星期, 診次): (院所代碼, 標記)}
for _cl, _cells in WEEKLY_RAW.items():
    for (_wd, _ss), _lst in _cells.items():
        for _nm, _flag in _lst:
            DOC_WEEK.setdefault(_nm, {})[(_wd, _ss, _cl)] = _flag
NAME2EID = {d[1]: d[0] for d in DOCTORS}
# 各院所有排診的時段 = 有開診。週日五院全休;週六晚診五院全休;寶貝牙平日早診休診。
OPEN = {c[0]: {(w, t) for (w, t) in WEEKLY_RAW.get(c[0], {})} for c in CLINICS}
ALL_CLOSED = {(w, t) for w in range(1, 8) for t in range(3)
              if not any((w, t) in OPEN[c] for c in OPEN)}
_unknown = [n for n in DOC_WEEK if n not in NAME2EID]
assert not _unknown, f"門診表出現名冊沒有的醫師:{_unknown}"

def expand_month(name, year, month, ndays):
    """把每週固定門診表展開成當月 31 天 × 3 診次。"""
    out = []
    slots = DOC_WEEK.get(name, {})
    for d in range(1, ndays + 1):
        date = dt.date(year, month, d)
        wd = date.weekday() + 1                      # 1=一 … 7=日
        wk = (date - ALT_EPOCH).days // 7            # 自基準日起的連續週序
        for ss in range(3):
            if date in HOLIDAY_SET:         # 國定假日,全院休診
                out.append("國"); continue
            if (wd, ss) in ALL_CLOSED:      # 五間院所該時段全休
                out.append("休"); continue
            hit = [(cl, fg) for (w, t, cl), fg in slots.items() if w == wd and t == ss]
            if not hit:
                out.append(""); continue
            val = ""
            for cl, fg in hit:
                if "~" in fg:                        # 隔週
                    group_b = (name, wd, cl) in ALT_B
                    if (wk % 2 == 1) != group_b:
                        continue
                val = cl
                break
            out.append(val)
    return out

# 名冊是靜態資料,直接寫實值而非公式——公式沒有快取值,不重算的檢視器會顯示空白
ROSTER_BY_EID = {}
for _eid, _nm, _eng, _spec, _duty, _teams in DOCTORS:
    _home = SHORT2CODE[_teams[0]] if not _duty else SHORT2CODE[_duty.replace("院長", "")]
    ROSTER_BY_EID[_eid] = (_nm, "醫師", _spec, _duty,
                           next(c[1] for c in CLINICS if c[0] == _home))
for _a in ASSISTANTS:
    ROSTER_BY_EID[_a[0]] = (_a[1], _a[2], "", "", _a[3])

WK_CH = "一二三四五六日"
def wd_of(d):     # 該日星期(中文)
    return WK_CH[dt.date(YEAR, MONTH, d).weekday()]
PERIOD = f"{YEAR} 年 {MONTH} 月"

# 名單未匯入時在標題直接講明,避免有人以為分頁壞掉
PENDING = not ASSISTANTS
PEND_TAG = "　　⚠ 人員名單尚未匯入,本分頁目前是空的" if PENDING else ""

def title_of(text):
    return text + PEND_TAG

def lookup(eid, idx, fallback_formula):
    """已知人員寫實值,空白列留公式(之後填編號會自動帶出)。"""
    return ROSTER_BY_EID[eid][idx] if eid in ROSTER_BY_EID else fallback_formula

wb = Workbook()

# ================================================================ 1. 說明
ws = wb.active; ws.title = "說明"
ws.sheet_view.showGridLines = False
ws.column_dimensions["A"].width = 2
ws.column_dimensions["B"].width = 20
ws.column_dimensions["C"].width = 106
put(ws, "B2", "日三牙醫體系 — 統一出勤表 使用說明", TITLE_F, border=False)
ws.merge_cells("B2:C2")

BLOCKS = [
 ("為什麼是一個檔",
  "體系 27 位醫師裡有 18 位跨院看診,林智俊院長一個人就跑四間。\n"
  "若一間院所一個檔,排他一個月要開四個檔,而且沒有任何機制擋得住「同一個早診\n"
  "同時被排在晶悅和晶睿」。所以全體系共用這一個檔。"),
 ("醫師與工時制人員為什麼分開排",
  "醫師的一天是「幾個診次」,醫護長與管理部的一天是「幾個小時」,兩者不能塞進同一種格子。\n"
  "醫師 → 醫師班表:一人一列,橫向 31 天 × 早/午/晚,格內填院所代碼(悅睿匯曜寶)。\n"
  "   一格只容得下一間院所,衝堂在結構上就不可能發生。\n"
  "醫護長 → 醫護長與管理部班表:一人一列,橫向 31 天,格內填班別代碼(早/午/晚/早午/OFF/特…)。\n"
  "本表目前不含助理。"),
 ("九個分頁怎麼分工",
  "① 說明 ② 設定 — 總部維護,各院所勿動。\n"
  "③ 醫師週班表 — 官網門診表的固定週輪值,是排月班的母表。\n"
  "④ 醫師班表 ⑤ 醫師月結 — 本月實際診次,由週班表自動展開。\n"
  "⑥ 醫護長與管理部班表 — 各院所醫護長填自己那一列。\n"
  "⑦ 打卡匯入 ⑧ 出勤紀錄 ⑨ 月結統計 — 醫護長與管理部的法定出勤與月結。"),
 ("週班表與月班表的關係",
  "門診表是固定的每週輪值,所以月班表不必一格一格填——「醫師班表」已經照\n"
  "「醫師週班表」把整個月展開好了,直接改例外即可(請假就把該格改成假別代碼)。\n"
  "⚠ 改「設定」的本期年月只會更新日期與星期列,格子裡的班不會重排。換月份時\n"
  "   請告知重新產生,或自行把上個月的格子整段調整。"),
 ("各院所診次時間不一樣",
  "晶悅/晶睿/晶匯/晶曜:早 09:00~11:30、午 14:00~16:30、晚 18:00~20:30。\n"
  "寶貝牙:早 09:00~14:00、午 14:00~17:00、晚 17:30~21:00,且週一至週五早診休診。\n"
  "五間院所週日全休、週六晚診全休。醫師班表的早/午/晚只是時段代號,\n"
  "實際幾點到幾點依「設定」院所清單裡該院所的定義。開診時段表也在那裡。"),
 ("每月作業流程",
  "上月底:總部排醫師班表 → 看底下五列「各院所每診次醫師數」有沒有 0 → 發給各院所。\n"
  "     各院所醫護長與管理部排自己那一列 → 檢查人力檢核列 → 交回。\n"
  "當月底:打卡機匯出檔貼進「打卡匯入」→ 到「出勤紀錄」處理異常 → 月結送人資。"),
 ("醫師班表怎麼填",
  "每一天有三格:早診、午診、晚診。格內填院所代碼,空白代表該診次沒排班。\n"
  "  悅 = 晶悅   睿 = 晶睿   匯 = 晶匯   曜 = 晶曜   寶 = 寶貝牙\n"
  "請假也是填在診次格裡(特/病/事/公/國/OFF/休),因為醫師請假常常只請半天。\n"
  "所以醫師的月結一律以「診次」為單位,不是天數——請假 3 個診次就是請一天。\n"
  "五種院所代碼各有底色,一位醫師整月在五間之間怎麼跑,橫著看一列就知道。"),
 ("改班表哪裡會跟著變",
  "資料是這樣流的,兩邊改都會立刻反映到月結:\n"
  "  醫師週班表 →(公式)→ 醫師班表 →(公式)→ 醫師月結\n"
  "\n"
  "① 改「醫師週班表」= 改固定門診(某位醫師從此改成週三晚診)。\n"
  "   醫師班表整個月的對應格子會立刻跟著變,月結也立刻變。\n"
  "② 直接改「醫師班表」的格子 = 這個月的單次例外(某天請假、臨時調班)。\n"
  "   月結一樣立刻變。\n"
  "\n"
  "⚠ 但②有個副作用:你打字覆蓋的那一格,原本的公式就沒了,\n"
  "   從此不再跟著週班表走。那種格子會自動加上紅色外框提醒你。\n"
  "   要讓它恢復連動,從左右鄰近沒被改過的格子複製一格貼回來即可。\n"
  "   (旁邊的格子公式一樣,貼過來會自動對應到正確的日期與診次。)"),
 ("醫師週班表的標記怎麼看",
  "週班表是母表,要表示「這一週在 A 院、下一週在 B 院」,所以比醫師班表多幾種寫法:\n"
  "  悅        每週固定在晶悅\n"
  "  悅(隔)    隔週在晶悅(單數週有、雙數週沒有)\n"
  "  悅/睿     兩間輪替——斜線前是單數週、斜線後是雙數週,全表共四組\n"
  "  寶(不定)  官網註明非每週固定,需向院所確認\n"
  "  悅(註)    有附加條件,寫在該列最右邊的附註欄\n"
  "完整對照表在「設定」分頁的「二之二、醫師週班表的標記」。\n"
  "醫師班表(月班表)不用這些標記,一格就是一個單純的代碼。"),
 ("出勤紀錄怎麼看",
  "一人一天一列,自動把醫護長與管理部班表與打卡對起來,判定五種狀態:\n"
  "正常、遲到、早退、未打卡、假日出勤。紅底 = 需處理;橘底 = 當天有加班。\n"
  "用上方篩選鈕只看非「正常」的最快。沒有打卡機的院所,可把「實際上班/下班」\n"
  "兩欄公式刪掉改成手填,其餘照算。"),
 ("加班怎麼認定",
  "超過排班工時的分鐘數要跨過門檻(預設 30 分)才算加班,並無條件捨去到計算\n"
  "單位(預設 30 分)。否則每天晚幾分鐘打卡都會被算成加班,一個月會憑空多出\n"
  "好幾小時。遲到早退另有寬限(預設 5 分)。三個參數在「設定」的計算參數區。\n"
  "這三個數字是勞資慣例、不是法規,上線前請人資確認,並讓同仁事先知道規則。"),
 ("兩種下拉箭頭不一樣",
  "標題列上的箭頭是「篩選鈕」,用來隱藏或顯示列,不是拿來選資料的。\n"
  "要選員工編號或班別代碼,請點「資料列」的格子——游標移上去右邊才會出現選單箭頭。\n"
  "沒排班或沒打卡的格子本來就是空的,篩選鈕點開看到「(空格)」是正常的。"),
 ("如果有些格子是空白的",
  "姓名、專科、日期、星期都是實際文字,任何檢視器都看得到。\n"
  "但統計類的格子(醫師月結的診次數、班表底下的人力檢核、出勤紀錄、月結統計)是公式,\n"
  "在不會重算公式的檢視器裡會顯示空白——手機預覽、雲端硬碟預覽、部分看圖工具都是這樣。\n"
  "用 Excel 或 Google 試算表正式開啟就會算出來;若仍空白,按 Ctrl+Alt+F9 強制重算。\n"
  "這些欄位不能寫死成固定數字,否則你改了班表,統計不會跟著變。"),
 ("顏色代表什麼",
  "黃底 = 你要填的格子        灰底 = 公式自動算,別動\n"
  "淺藍欄 = 週六/週日          灰色格 = 休假類代碼\n"
  "五種院所底色 = 醫師該診次在哪一間\n"
  "紅色 = 人力不足或出勤異常    橘色 = 有加班"),
 ("⚠ 法規提醒",
  "依勞動基準法第 30 條,雇主應置備勞工出勤紀錄,逐日記載至分鐘為止,保存 5 年。\n"
  "「出勤紀錄」分頁為此設計,但必須每月另存唯讀封存檔才算數,不要只留一份覆蓋。\n"
  "加班時數為「實際工時 − 排班工時」,未依第 24 條換算費率,薪資須另行計算。\n"
  "受僱醫師是否適用勞基法依僱傭契約與主管機關認定,建議請人資確認後再定義醫師工時。"),
 ("⚠ 需要核對的三件事",
  "一、門診表是從截圖轉錄的,姓名與隔週標記請逐格核對「醫師週班表」。\n"
  "   若能提供五個門診表頁面的網頁存檔,可以完全免除轉錄誤差。\n"
  "二、陳爵安、吳冠廷、楊孟庭三位只出現在門診表,官網醫療團隊頁的名單沒有,\n"
  "   已暫編 D025~D027,英文名與專科待補。\n"
  "三、隔週互換有四組(劉立德週三、朱柏非週六、王泳泉週一、陳昺元週六),\n"
  "   同一時段在兩間院所輪替。目前假設第 1、3、5 個該星期幾在前一間,\n"
  "   第 2、4 個在後一間——哪一週在哪一間無從得知,務必確認。"),
 ("目前待補的資料",
  "一、醫護長與管理部的到職日尚未提供——特休天數依年資給,沒有到職日算不出特休餘額。\n"
  "二、本表目前不含助理。日後要納入,在「設定」人員名冊加人、並把醫護長與管理部班表的\n"
  "   列數調大即可,班別代碼與公式都不用改。\n"
  "三、員工編號目前為暫編 D001~D027,若人事或打卡系統另有編號應以那套為準。\n"
  "四、官網門診表有星號註記者(並非每週固定)需逐一向院所確認。"),
]
if not WITH_ASSISTANT:
    DROP = {"出勤紀錄怎麼看", "加班怎麼認定"}
    SWAP = {
      "九個分頁怎麼分工": ("五個分頁怎麼分工",
        "① 說明 ② 設定 — 總部維護,各院所勿動。\n"
        "③ 醫師週班表 — 官網門診表的固定週輪值,是排月班的母表。\n"
        "④ 醫師班表 ⑤ 醫師月結 — 本月實際診次,由週班表自動展開。"),
      "醫師與工時制人員為什麼分開排": ("這一版只有醫師",
        "這是醫師版,工時制的四個分頁(醫護長與管理部班表、打卡匯入、出勤紀錄、月結統計)\n"
        "已經拿掉,需要時用完整版。那四頁彼此相依——出勤紀錄要讀醫護長與管理部班表與\n"
        "打卡匯入、月結統計要讀出勤紀錄——所以是整組拿掉,不是只拔一頁。\n"
        "「設定」分頁裡的班別時段表保留不動,完整版會用到。"),
      "每月作業流程": ("每月作業流程",
        "上月底:排醫師班表 → 看底下五列「各院所每診次醫師數」有沒有 0 → 發給各院所。\n"
        "要改固定門診就改「醫師週班表」,整個月會立刻跟著變;\n"
        "只是這個月的例外(請假、調班)就直接改「醫師班表」的格子。"),
      "如果有些格子是空白的": ("如果有些格子是空白的",
        "姓名、專科、日期、星期都是實際文字,任何檢視器都看得到。\n"
        "但班別格與統計格(醫師月結的診次數、班表底下的人力檢核)是公式,\n"
        "在不會重算公式的檢視器裡會顯示空白——手機預覽、雲端硬碟預覽都是這樣。\n"
        "用 Excel 或 Google 試算表正式開啟就會算出來;若仍空白,按 Ctrl+Alt+F9 強制重算。"),
      "⚠ 法規提醒": ("⚠ 法規提醒",
        "依勞動基準法第 30 條,雇主應置備勞工出勤紀錄,逐日記載至分鐘為止,保存 5 年。\n"
        "這一版只有醫師的診次統計,沒有法定出勤紀錄——那在完整版的「出勤紀錄」分頁,\n"
        "涵蓋醫護長與管理部。這一版不能單獨拿來因應勞檢。\n"
        "受僱醫師是否適用勞基法、工時如何認定,依僱傭契約與主管機關認定,\n"
        "建議請人資確認後再決定醫師要不要納入工時管理。"),
      "目前待補的資料": ("目前待補的資料",
        "一、醫護長與管理部的出勤在完整版,那邊有醫護長與管理部班表、打卡匯入、出勤紀錄、\n"
        "   月結統計四個分頁。本表目前不含助理。\n"
        "二、員工編號目前為暫編 D001~D027,若人事或打卡系統另有編號應以那套為準。\n"
        "三、官網門診表有星號註記者(並非每週固定)需逐一向院所確認。"),
      "兩種下拉箭頭不一樣": ("兩種下拉箭頭不一樣",
        "標題列上的箭頭是「篩選鈕」,用來隱藏或顯示列,不是拿來選資料的。\n"
        "要選員工編號或診次代碼,請點「資料列」的格子——游標移上去右邊才會出現選單箭頭。"),
    }
    BLOCKS = [(SWAP.get(h, (h, b))[0], SWAP.get(h, (h, b))[1])
              for h, b in BLOCKS if h not in DROP]
    put(ws, "B2", "日三牙醫體系 — 統一出勤表(醫師版)使用說明", TITLE_F, border=False)

r = 4
for h, body in BLOCKS:
    put(ws, f"B{r}", h, font(10, True), SUB_FILL, LEFT)
    put(ws, f"C{r}", body, font(10), None, WRAP)
    ws.row_dimensions[r].height = 15 * (body.count("\n") + 1) + 8
    r += 1

# ================================================================ 2. 設定
st = wb.create_sheet("設定")
st.sheet_view.showGridLines = False
for col, w in {"A":2,"B":11,"C":18,"D":9,"E":9,"F":10,"G":11,"H":8,"I":9,"J":9,
               "K":2,"L":12,"M":2,"N":9,"O":8,"P":2,"Q":13,"R":2,"S":16,"T":2,"W":14,"X":26,"Y":8}.items():
    st.column_dimensions[col].width = w
put(st, "B1", "設定表(總部維護,各院所請勿修改)", TITLE_F, border=False)
put(st, "B2", "本期年月", font(10, True), SUB_FILL, CTR)
put(st, "C2", YEAR, font(10, True), IN_FILL, CTR, "0")
put(st, "D2", MONTH, font(10, True), IN_FILL, CTR, "0")
put(st, "E2", "← 整份檔案的日期、星期都以這裡為準", font(9, color="808080"),
    None, LEFT, border=False)

put(st, f"B{WC_R0-2}", "一、班別代碼(工時制:醫護長、管理部共用)", font(11, True),
    border=False)
header_row(st, WC_R0-1, ["代碼","名稱","應到","應退","休息(分)","排班工時",
                         "計出勤","類別"])
for i, row in enumerate(WC_TABLE):
    code, name, tin, tout, rest, hrs, att, kind = row
    for j, v in enumerate(row):
        c = st.cell(row=WC_R0+i, column=2+j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if j == 1: c.alignment = LEFT
        if j in (2, 3): c.number_format = FMT_TIME
        if j == 5: c.number_format = "0.0"
    if kind == "休假":
        for j in range(8):
            st.cell(row=WC_R0+i, column=2+j).fill = LEAVE_FILL
put(st, f"B{WC_R1+1}",
    "※ 醫護長與管理部用同一套代碼、同一套時段:早 09:00–12:00、午 14:00–17:00、"
    "晚 18:00–21:00。填表的人只要填代碼。",
    font(9, color="808080"), None, LEFT, border=False)
put(st, f"B{WC_R1+2}",
    "※ 有固定時段的代碼必須滿足「應退 − 應到 − 休息 = 排班工時」,否則出勤紀錄會算出不存在的加班。"
    "建置腳本會自動檢查。", font(9, color="808080"), None, LEFT, border=False)
put(st, f"B{WC_R1+3}",
    "※ 注意:勞基法第 30 條的正常工時是每日 8 小時。「全日」排 9 小時,多出來的"
    "1 小時本質上就是延長工時,但本表的加班是用「實際 − 排班」算的,排進去就不會"
    "被抓出來。要按法規認定,請不要把「全日」當成常態班別。",
    font(9, color="A8433C"), None, LEFT, border=False)

put(st, f"B{DC_R0-2}", "二、醫師診次代碼(醫師班表用:一格填一個代碼)",
    font(11, True), border=False)
header_row(st, DC_R0-1, ["代碼","意義"])
for i, (code, mean) in enumerate(DOC_CODES):
    put(st, f"B{DC_R0+i}", code, font(10, True),
        CLINIC_FILL.get(code, CALC_FILL), CTR)
    put(st, f"C{DC_R0+i}", mean, font(), None, LEFT)
put(st, f"B{DC_R1+1}",
    "※ 醫師班表一格 = 一個診次。填院所代碼表示該診次在哪間看診,填休假代碼表示該診次請假。",
    font(9, color="808080"), None, LEFT, border=False)
put(st, f"B{DC_R1+2}",
    "※ 「醫師週班表」除了上面這些代碼,還會用到右邊那幾種標記——那是母表,"
    "要表示隔週與輪替,所以下拉選項比較多。",
    font(9, color="808080"), None, LEFT, border=False)

# ── 週班表的標記對照(放在診次代碼表右邊)──────────────────
put(st, f"E{DC_R0-2}", "二之二、醫師週班表的標記(母表專用)", font(11, True), border=False)
MARK_R0 = DC_R0 - 1
for i, lab in enumerate(["寫法", "意思", "舉例說明"]):
    c = st.cell(row=MARK_R0, column=5 + i, value=lab)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
MARKS = [
    ("悅",       "每週固定",      "每一週的這個診次都在晶悅"),
    ("悅(隔)",   "隔週看診",      "單數週在晶悅,雙數週這個診次不排"),
    ("悅/睿",    "兩間院所輪替",  "斜線前 = 單數週、斜線後 = 雙數週(全表共四組)"),
    ("寶(不定)", "非每週固定",    "官網註明「詳情請聯繫院所」,需逐一向院所確認"),
    ("悅(註)",   "有附加條件",    "條件寫在該列最右邊的附註欄"),
    ("OFF 特 病", "休假類",       "整個診次請假,寫法與醫師班表相同"),
]
for i, (w, mean, ex) in enumerate(MARKS):
    r = MARK_R0 + 1 + i
    put(st, f"E{r}", w, font(10, True), CALC_FILL, CTR)
    st.column_dimensions["E"].width = 11
    put(st, f"F{r}", mean, font(9), None, CTR)
    put(st, f"G{r}", ex, font(9), None, LEFT)
    st.merge_cells(f"G{r}:J{r}")
put(st, f"E{MARK_R0 + len(MARKS) + 1}",
    "※ 單數/雙數週以計算參數區的「隔週基準日」起算,跨月跨年連續不會斷。",
    font(9, color="808080"), None, LEFT, border=False)
put(st, f"E{MARK_R0 + len(MARKS) + 2}",
    "※ 週班表的資料格有下拉選單,但不強制——輸入清單外的寫法只會跳提醒,按「是」照樣存入。",
    font(9, color="808080"), None, LEFT, border=False)

put(st, "W4", f"六、{YEAR} 年國定假日(全院休診,班表自動填「國」)",
    font(10, True), SUB_FILL, LEFT)
st.merge_cells("W4:Y4")
for i, lab in enumerate(["日期", "名稱", "類別"]):
    c = st.cell(row=5, column=23 + i, value=lab)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
WK_CH = "一二三四五六日"
for i in range(35):
    r = 6 + i
    if i < len(HOLIDAYS_2026):
        m, d, name, mk = HOLIDAYS_2026[i]
        date = dt.date(YEAR, m, d)
        put(st, f"W{r}", date, font(9), IN_FILL, CTR, FMT_DATE)
        put(st, f"X{r}", f"{name}(週{WK_CH[date.weekday()]})", font(9), IN_FILL, LEFT)
        put(st, f"Y{r}", "補假" if mk else "", font(9), IN_FILL, CTR)
    else:
        for cc in range(23, 26):
            c = st.cell(row=r, column=cc)
            c.font, c.border, c.alignment = font(9), BOX, CTR
put(st, "W{}".format(6 + 36),
    "※ ⚠ 這份是草稿,務必與行政院人事行政總處公告的「政府行政機關辦公日曆表」核對。"
    "法定紀念日日期可信度高,但補假與調整放假(彈性放假)每年公告,以官方版本為準。",
    font(9, color="A8433C"), None, LEFT, border=False)
put(st, "W{}".format(6 + 37),
    "※ 改這裡不會自動改醫師班表——月班表是產生出來的固定值,改完要重新產生。",
    font(9, color="808080"), None, LEFT, border=False)

put(st, "S5", "工時制人員編號(下拉用)", font(9, True), SUB_FILL, CTR)
AST_EIDS = [a[0] for a in ASSISTANTS]
for i in range(N_ASST_SLOTS):
    put(st, f"S{6+i}", AST_EIDS[i] if i < len(AST_EIDS) else "",
        font(9), CALC_FILL if i < len(AST_EIDS) else None, CTR)

put(st, "Q5", "醫師編號(下拉用)", font(9, True), SUB_FILL, CTR)
DOC_EIDS = [d[0] for d in DOCTORS]
for i in range(40):
    put(st, f"Q{6+i}", DOC_EIDS[i] if i < len(DOC_EIDS) else "",
        font(9), CALC_FILL if i < len(DOC_EIDS) else None, CTR)

put(st, "N5", "星期對照", font(9, True), SUB_FILL, CTR)
put(st, "O5", "", font(9, True), SUB_FILL, CTR)
for i, w in enumerate(["日","一","二","三","四","五","六"]):
    put(st, f"N{6+i}", i+1, font(9), CALC_FILL, CTR)
    put(st, f"O{6+i}", w, font(9), CALC_FILL, CTR)

# 標題要在表頭的上一列,否則會被 header_row 覆蓋掉
put(st, f"B{PM_R0-1}", "三、計算參數(全體系一致)", font(11, True), border=False)
header_row(st, PM_R0, ["參數","值","說明"], start_col=2)
PARAMS = [("加班認定門檻(分)", 30, "超過排班工時多少分鐘才認定為加班。低於門檻視為正常收尾。"),
          ("加班計算單位(分)", 30, "認定為加班後,無條件捨去到此單位。填 1 表示逐分鐘計。"),
          ("遲到早退寬限(分)", 5,  "打卡與應到/應退時間差在此範圍內不判定為遲到或早退。"),
          ("隔週基準日", ALT_EPOCH,
           "從這個星期一起連續數週。單數週 = 週班表前面那個代碼,雙數週 = 後面那個。"
           "若實際輪替對調了,把這個日期往前或往後挪 7 天即可整體對齊。")]
for i, (nm, val, desc) in enumerate(PARAMS):
    r = PM_R0 + 1 + i
    put(st, f"B{r}", nm, font(), SUB_FILL, LEFT)
    put(st, f"C{r}", val, font(10, True), IN_FILL, CTR,
        FMT_DATE if isinstance(val, dt.date) else "0")
    put(st, f"D{r}", desc, font(9), None, LEFT); st.merge_cells(f"D{r}:J{r}")
put(st, f"B{PM_R0+5}",
    "※ 前三個是勞資雙方的認定慣例,不是法律規定的數字。上線前請與人資或勞務顧問確認,"
    "並讓同仁事先知道規則。", font(9, color="808080"), None, LEFT, border=False)

put(st, f"B{CL_R0-1}", "四、院所清單", font(11, True), border=False)
header_row(st, CL_R0, ["代碼","簡稱","全名","院長","電話","地址",
                       "早診起","早診訖","午診起","午診訖","晚診起","晚診訖"])
CLINIC_ROWS = [tuple(c) + tuple(t for pair in SESSION_TIME[c[0]] for t in pair)
               for c in CLINICS]
for i in range(CL_L1 - CL_L0 + 1):
    vals = CLINIC_ROWS[i] if i < len(CLINIC_ROWS) else ("",)*12
    for j, v in enumerate(vals):
        c = st.cell(row=CL_L0+i, column=2+j, value=v)
        c.font, c.border, c.alignment = font(), BOX, CTR
        if j in (2,5): c.alignment = LEFT
        if j >= 6: c.font = font(9)
        if j == 0 and v: c.fill = CLINIC_FILL[v]

put(st, "U{}".format(CL_R0), "開診時段(1 = 有開診,0 = 休診;順序為週一早…週日晚)",
    font(9, True), SUB_FILL, LEFT)
st.merge_cells(start_row=CL_R0, start_column=21, end_row=CL_R0, end_column=41)
for i in range(CL_L1 - CL_L0 + 1):
    code = CLINICS[i][0] if i < len(CLINICS) else None
    for w in range(7):                       # 含週日(五院皆休,全 0),
        for t in range(3):                   # 否則公式在週日會 INDEX 超界變 #REF!
            c = st.cell(row=CL_L0 + i, column=21 + w * 3 + t)
            c.value = (1 if code and (w + 1, t) in OPEN[code] else 0) if code else None
            c.font, c.border, c.alignment = font(8), BOX, CTR
            if code and (w + 1, t) not in OPEN[code]: c.fill = LEAVE_FILL
for w in range(7):
    for t in range(3):
        c = st.cell(row=CL_R0 - 1, column=21 + w * 3 + t,
                    value=f"{'一二三四五六日'[w]}{SESSIONS[t]}")
        c.font, c.fill, c.alignment, c.border = font(8, True, "FFFFFF"), HDR_FILL, CTR, BOX

put(st, f"B{CL_L1+2}",
    "※ 寶貝牙的診次時段與其他四院不同,且週一至週五早診休診、週六晚診休診。"
    "醫師班表的早/午/晚只是時段代號,實際幾點到幾點依上表該院所的定義。",
    font(9, color="808080"), None, LEFT, border=False)

put(st, f"B{PP_R0-1}", "五、人員名冊", font(11, True), border=False)
header_row(st, PP_R0, ["員工編號","姓名","職類","專科","職務","主要院所",
                       "服務院所(醫師)","英文名","到職日","備註"])
ROSTER = []
for eid, nm, eng, spec, duty, teams in DOCTORS:
    home = SHORT2CODE[teams[0]] if not duty else SHORT2CODE[duty.replace("院長","")]
    ROSTER.append([eid, nm, "醫師", spec, duty,
                   next(c[1] for c in CLINICS if c[0] == home),
                   " · ".join(teams), eng, "", ""])
# 醫師版沒有工時制的分頁,名冊就不列那些人,免得看到名字卻找不到班表。
if WITH_ASSISTANT:
    for a in ASSISTANTS:
        ROSTER.append([a[0], a[1], a[2], "", "", a[3], "", "", a[4], a[5]])
for i in range(PP_L1 - PP_L0 + 1):
    vals = ROSTER[i] if i < len(ROSTER) else [""]*10
    for j, v in enumerate(vals):
        c = st.cell(row=PP_L0+i, column=2+j, value=v)
        c.font, c.border, c.alignment = font(9), BOX, CTR
        if i < len(ROSTER): c.fill = IN_FILL
        if j in (3, 6, 9): c.alignment = LEFT
_n_nur = sum(1 for a in ASSISTANTS if a[2] == "醫護長")
_n_mgt = sum(1 for a in ASSISTANTS if a[2] == "管理部")
print(f"人員名冊:{len(ROSTER)} 人(醫師 {len(DOCTORS)}、"
      f"醫護長 {_n_nur if WITH_ASSISTANT else 0}、管理部 {_n_mgt if WITH_ASSISTANT else 0})")

# ================================================================ 3. 醫師週班表
wk = wb.create_sheet("醫師週班表")
wk.sheet_view.showGridLines = False
WK_ROW0 = 5
WK_ROW1 = WK_ROW0 + N_DOC - 1
WK_C0 = 3                                   # C 欄起,18 格
wk.freeze_panes = "C5"
wk.column_dimensions["A"].width = 9
wk.column_dimensions["B"].width = 10
for i in range(18):
    wk.column_dimensions[get_column_letter(WK_C0 + i)].width = 7.2
wk.column_dimensions[get_column_letter(WK_C0 + 18)].width = 3
wk.column_dimensions[get_column_letter(WK_C0 + 19)].width = 40
put(wk, "A1", "醫師週班表(固定門診表 · 這是排月班的母表)", TITLE_F, border=False)
wk.merge_cells(start_row=1, start_column=1, end_row=1, end_column=WK_C0 + 17)
put(wk, "A2",
    "資料來源:官網五間院所門診表頁面,由原始碼直接解析產生(231 筆逐格比對一致)。"
    "格內寫法見「設定」分頁的「二之二、醫師週班表的標記」。"
    "改這裡等於改固定門診:「醫師班表」整個月的對應格子會立刻跟著變,月結也會。",
    font(9, color="808080"), None, LEFT, border=False)
for lab, col in (("員工編號", 1), ("姓名", 2)):
    c = wk.cell(row=3, column=col, value=lab)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    wk.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col)
WD_LAB = ["一", "二", "三", "四", "五", "六"]
for w in range(6):
    a = WK_C0 + w * 3
    c = wk.cell(row=3, column=a, value=f"週{WD_LAB[w]}")
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    wk.merge_cells(start_row=3, start_column=a, end_row=3, end_column=a + 2)
    for sidx in range(3):
        cc = wk.cell(row=4, column=a + sidx, value=SESSIONS[sidx])
        cc.font, cc.fill, cc.alignment = font(8), SUB_FILL, CTR
        cc.border = DAYSEP if sidx == 0 else BOX
put(wk, f"{get_column_letter(WK_C0+19)}3", "附註", HDR_F, HDR_FILL, CTR)
wk.merge_cells(start_row=3, start_column=WK_C0+19, end_row=4, end_column=WK_C0+19)

for i in range(N_DOC):
    r = WK_ROW0 + i
    wk.row_dimensions[r].height = 17
    eid = DOCTORS[i][0] if i < len(DOCTORS) else None
    nm = DOCTORS[i][1] if i < len(DOCTORS) else None
    a = wk.cell(row=r, column=1, value=eid)
    a.font, a.fill, a.border, a.alignment = font(9), IN_FILL, BOX, CTR
    b = wk.cell(row=r, column=2)
    b.value = lookup(eid, 0, f'=IFERROR(INDEX({R_NAME},MATCH($A{r},{R_EID},0)),"")')
    b.font, b.fill, b.border, b.alignment = font(9), CALC_FILL, BOX, CTR
    slots = DOC_WEEK.get(nm, {}) if nm else {}
    notes = []
    for w in range(6):
        for sidx in range(3):
            cc = wk.cell(row=r, column=WK_C0 + w * 3 + sidx)
            hit = [(cl, fg) for (ww, tt, cl), fg in slots.items()
                   if ww == w + 1 and tt == sidx]
            if hit:
                # 單數週那組在前、雙數週那組在後,「斜線前 = 單數週」才是通則
                hit = sorted(hit, key=lambda x: ((nm, w + 1, x[0]) in ALT_B, x[0]))
                cc.value = wk_label(hit)
                cc.fill = CLINIC_FILL.get(hit[0][0], CALC_FILL)
            cc.font, cc.alignment = font(8, True), CTR
            cc.border = DAYSEP if sidx == 0 else BOX
    for (ww, tt, cl), fg in sorted(slots.items()):
        if "!" in fg and (cl, ww, tt) not in ():
            key = (cl, ww, tt)
        if "!" in fg:
            notes.append(NOTES.get((cl, ww, tt), "有附加條件"))
    n = wk.cell(row=r, column=WK_C0 + 19, value=" / ".join(dict.fromkeys(notes)))
    n.font, n.border, n.alignment = font(8), BOX, LEFT
# 週班表的儲存格值比醫師班表複雜:除了單一代碼,還有隔週(悅~)、
# 兩院輪替(悅~睿~)、官網星號(寶*)與附註(悅!)。列成下拉方便選,
# 但不強制——showErrorMessage=False,任何值都還是打得進去。
WK_USED = sorted({wk.cell(row=r, column=c).value
                  for r in range(WK_ROW0, WK_ROW1 + 1)
                  for c in range(WK_C0, WK_C0 + 18)
                  if wk.cell(row=r, column=c).value})
_base = [c for c in CLINIC_CODES] + [f"{c}(隔)" for c in CLINIC_CODES]
WK_CHOICES = (_base + [v for v in WK_USED
                       if v not in _base and v != "訓" and v not in DOC_LEAVE]
              + ["訓"] + DOC_LEAVE)
WK_CHOICES = list(dict.fromkeys(WK_CHOICES))
WKC_R0 = 6
put(st, f"AP{WKC_R0-1}", "醫師週班表下拉選單來源(自動產生,勿刪)",
    font(9, True), SUB_FILL, CTR)
st.merge_cells(f"AP{WKC_R0-1}:AR{WKC_R0-1}")
st.column_dimensions["AP"].width = 16
for i, v in enumerate(WK_CHOICES):
    put(st, f"AP{WKC_R0+i}", v, font(9), CALC_FILL, CTR)
# ⚠ showErrorMessage=False 會讓 Excel 整條驗證失效(下拉不會出現)。
# 要「有下拉但不強制」,正確做法是 errorStyle="warning":輸入清單外的值只跳
# 提醒,按「是」仍可存入。清單直接內嵌,不依賴跨分頁參照,最不容易出問題。
_inline = ",".join(WK_CHOICES)
assert len(_inline) <= 255, f"內嵌清單過長({len(_inline)} 字元)"
dv_wkcell = DataValidation(
    type="list", formula1=f'"{_inline}"', allow_blank=True,
    showErrorMessage=True, errorStyle="warning",
    errorTitle="不在常用清單中",
    error="這個寫法不在常用清單裡。若是特殊組合(例如三間輪替),按「是」即可照樣存入。")
wk.add_data_validation(dv_wkcell)
dv_wkcell.add(f"{get_column_letter(WK_C0)}{WK_ROW0}:"
              f"{get_column_letter(WK_C0+17)}{WK_ROW1}")
print(f"週班表下拉選項 {len(WK_CHOICES)} 項:{'、'.join(WK_CHOICES)}")

NOTE0 = WK_ROW1 + 2
for i, t in enumerate([
  "※ 資料由官網五個門診表頁面直接解析產生,非人工轉錄。寫法:悅(隔) = 隔週看診、"
  "悅/睿 = 兩院輪替(斜線前為單數週)、寶(不定) = 官網註明非每週固定、悅(註) = 見附註欄。",
  "※ 資料格有下拉選單(常用值),但不強制——特殊組合直接打字也可以,不會被擋。",
  "※ 隔週的四組互換(劉立德週三、朱柏非週六、王泳泉週一、陳昺元週六)以固定基準日(2026/9/28 起)"
  "連續數週判定單雙週,跨月不會斷。哪一組先需要院所確認——確認後把基準日挪一週即可整體對齊。",
  "※ 週日五間院所皆休診。寶貝牙另有週一至週五早診休診、週六晚診休診。",
]):
    put(wk, f"A{NOTE0+i}", t, font(9, color="808080"), None, LEFT, border=False)
dv_wk = DataValidation(type="list", formula1="設定!$Q$6:$Q$45", allow_blank=True,
                       showErrorMessage=True, errorStyle="warning",
                       errorTitle="不在醫師名冊中",
                       error="這個編號不在醫師名冊裡。確定要用請按「是」。")
wk.add_data_validation(dv_wk); dv_wk.add(f"A{WK_ROW0}:A{WK_ROW1}")

# ================================================================ 4. 醫師班表
ds = wb.create_sheet("醫師班表")
ds.sheet_view.showGridLines = False
DS_C0 = 5                                   # E 欄起
DS_H0 = 57                                  # 隱藏輔助列 57/58/59
DS_HCOL = 99                                # 隱藏輔助欄 CU
DS_HCOL_L = get_column_letter(DS_HCOL)
DS_ROW0 = 6
DS_ROW1 = DS_ROW0 + N_DOC - 1               # 45
ds.freeze_panes = "E6"
for col, w in {"A":9,"B":10,"C":15,"D":11}.items():
    ds.column_dimensions[col].width = w
def dcol(d, sidx):                           # 第 d 日、第 sidx 診次的欄號
    return DS_C0 + (d-1)*3 + sidx
for d in range(1, DAYS_IN_MONTH+1):
    for sidx in range(3):
        # 4.0 是為了「OFF」。3.4 剛好卡在邊界上(實測需 23px、可用 23px),
        # 院所代碼是一個字沒問題,但 OFF 會被隔壁切掉。
        ds.column_dimensions[get_column_letter(dcol(d, sidx))].width = 4.0
LAST_C = dcol(DAYS_IN_MONTH, 2)

put(ds, "A1", "醫師班表(診次制 · 一格 = 一個診次 · 已依週班表填好本月)", TITLE_F, border=False)
ds.merge_cells(start_row=1, start_column=1, end_row=1, end_column=min(LAST_C, 40))
put(ds, "A2", "期間", font(10, True), SUB_FILL, CTR)
ds["B2"] = PERIOD
ds["B2"].font, ds["B2"].fill, ds["B2"].alignment, ds["B2"].border = (
    font(10, True), CALC_FILL, CTR, BOX)
put(ds, "D2", "圖例:", font(9, True), None, LEFT, border=False)
for i, (code, short, *_rest) in enumerate(CLINICS):
    c = ds.cell(row=2, column=DS_C0 + i*4)
    c.value = f"{code} {short}"
    c.font, c.fill, c.alignment, c.border = font(9), CLINIC_FILL[code], CTR, BOX
    ds.merge_cells(start_row=2, start_column=DS_C0+i*4, end_row=2, end_column=DS_C0+i*4+3)

for lab, col in (("員工編號",1), ("姓名",2), ("專科",3), ("職務",4)):
    c = ds.cell(row=3, column=col, value=lab)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    ds.merge_cells(start_row=3, start_column=col, end_row=5, end_column=col)

for d in range(1, DAYS_IN_MONTH+1):
    a = dcol(d, 0); L = get_column_letter(a)
    c = ds.cell(row=3, column=a, value=d)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    ds.merge_cells(start_row=3, start_column=a, end_row=3, end_column=a+2)
    for sidx in range(3):
        cc = dcol(d, sidx); CL = get_column_letter(cc)
        w = ds.cell(row=4, column=cc)
        w.value = wd_of(d)
        w.font, w.fill, w.alignment = font(8, True), SUB_FILL, CTR
        w.border = DAYSEP if sidx == 0 else BOX
        s = ds.cell(row=5, column=cc, value=SESSIONS[sidx])
        s.font, s.fill, s.alignment = font(8), SUB_FILL, CTR
        s.border = DAYSEP if sidx == 0 else BOX

for i in range(N_DOC):
    r = DS_ROW0 + i
    ds.row_dimensions[r].height = 17
    eid = DOCTORS[i][0] if i < len(DOCTORS) else None
    a = ds.cell(row=r, column=1, value=eid)
    a.font, a.fill, a.border, a.alignment = font(9), IN_FILL, BOX, CTR
    for col, src, ix in ((2, R_NAME, 0), (3, R_SPEC, 2), (4, R_DUTY, 3)):
        c = ds.cell(row=r, column=col)
        c.value = lookup(eid, ix, f'=IFERROR(INDEX({src},MATCH($A{r},{R_EID},0)),"")')
        c.font, c.fill, c.border, c.alignment = font(9), CALC_FILL, BOX, CTR
        if col == 3: c.alignment = LEFT
    # 輔助欄:這位醫師在「醫師週班表」的第幾列(用員工編號比對,不靠列位置硬綁)
    hcell = ds.cell(row=r, column=DS_HCOL)
    hcell.value = (f'=IFERROR(MATCH($A{r},醫師週班表!$A${WK_ROW0}:$A${WK_ROW1},0),0)')
    hcell.font = font(8, color="BBBBBB")
    for d in range(1, DAYS_IN_MONTH+1):
        for sidx in range(3):
            cc = ds.cell(row=r, column=dcol(d, sidx))
            L = get_column_letter(dcol(d, sidx))
            W = (f'IFERROR(INDEX(醫師週班表!$C${WK_ROW0}:$T${WK_ROW1},'
                 f'${DS_HCOL_L}{r},{L}${DS_H0}),"")')
            # ⚠ INDEX 指到空白儲存格會回傳數字 0(不是空字串),
            # 沒擋掉的話沒排班的格子會顯示 0。文字值與 0 永不相等,所以用 =0 判斷。
            # 解析順序:悅/睿(輪替)→ 悅(隔)→ 其他帶括號的 → 原樣輸出
            P = f'{L}${DS_H0+1}'
            cc.value = (
                f'=IF({L}${DS_H0+2}="X","",'
                f'IF({L}${DS_H0+2}<>"",{L}${DS_H0+2},'
                f'IF({W}=0,"",'
                f'IF(MID({W},2,1)="/",MID({W},1+{P}*2,1),'
                f'IF(MID({W},3,1)="隔",IF({P}=0,LEFT({W},1),""),'
                f'IF(MID({W},2,1)="(",LEFT({W},1),'
                f'{W}))))))')
            cc.font, cc.alignment = font(9, True), CTR
            cc.border = DAYSEP if sidx == 0 else BOX

# ── 隱藏的運算輔助列/欄 ───────────────────────────────
# 第 57 列 = 週班表的第幾格(星期×診次)、58 列 = 單雙週、59 列 = 當日狀態
for d in range(1, DAYS_IN_MONTH + 1):
    for sidx in range(3):
        cc = dcol(d, sidx); L = get_column_letter(cc)
        ds.cell(row=DS_H0, column=cc).value = (
            f'=IF({d}>{DAYS_FX},1,(WEEKDAY(DATE({SET_Y},{SET_M},{d}),2)-1)*3+{sidx+1})')
        ds.cell(row=DS_H0+1, column=cc).value = (
            f'=IF({d}>{DAYS_FX},0,'
            f'MOD(INT((DATE({SET_Y},{SET_M},{d})-{P_EPOCH})/7),2))')
        ds.cell(row=DS_H0+2, column=cc).value = (
            f'=IF({d}>{DAYS_FX},"X",'
            f'IF(COUNTIF({R_HOL},DATE({SET_Y},{SET_M},{d}))>0,"國",'
            f'IF(SUMPRODUCT(INDEX({R_OPEN},0,{L}${DS_H0}))=0,"休","")))')
for hr in range(DS_H0, DS_H0 + 3):
    put(ds, f"A{hr}", "↓運算用,請勿刪", font(8, color="BBBBBB"), None, LEFT, border=False)
    ds.row_dimensions[hr].hidden = True
ds.column_dimensions[DS_HCOL_L].hidden = True

TALLY0 = DS_ROW1 + 2                         # 47
put(ds, f"A{TALLY0-1}", "各院所每診次醫師數(自動計算,0 表示該診次沒有醫師)",
    font(10, True), SUB_FILL, LEFT)
ds.merge_cells(start_row=TALLY0-1, start_column=1, end_row=TALLY0-1, end_column=4)
for k, (code, short, *_r) in enumerate(CLINICS):
    r = TALLY0 + k
    put(ds, f"A{r}", f"{code} {short}", font(9, True), CLINIC_FILL[code], CTR)
    ds.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    for d in range(1, DAYS_IN_MONTH+1):
        for sidx in range(3):
            cc = dcol(d, sidx); CL = get_column_letter(cc)
            c = ds.cell(row=r, column=cc)
            orow = CL_L0 + k
            c.value = (f'=IF({CL}$4="","",'
                       f'IF(COUNTIF({R_HOL},DATE({SET_Y},{SET_M},{d}))>0,"假",'
                       f'IF(INDEX(設定!$U${orow}:$AO${orow},'
                       f'(WEEKDAY(DATE({SET_Y},{SET_M},{d}),2)-1)*3+{sidx+1})=0,"休",'
                       f'COUNTIF({CL}${DS_ROW0}:{CL}${DS_ROW1},"{code}"))))')
            c.font, c.fill, c.alignment = font(8, True), CALC_FILL, CTR
            c.border = DAYSEP if sidx == 0 else BOX
            c.number_format = FMT_CNT
for _i, _t in enumerate([
  "※ 一格只容得下一間院所,所以同一位醫師同一診次被排到兩間院所在結構上不可能發生。"
  "上面五列是反向檢查:某間院所某個診次掛 0,表示那個時段沒有醫師。",
  "※ 本表已依「醫師週班表」把固定門診展開成本月的值,可以直接改。請假就把該格改成假別代碼。",
  "※ 姓名、專科、日期、星期都是實際文字不是公式,任何檢視器都看得到;"
  "底下的人力檢核列是公式,不重算的檢視器會空白。",
  "※ 格子是公式,從「醫師週班表」自動展開。改週班表、改設定的本期年月,這裡都會立刻跟著變。",
  "※ 要改單次例外(請假、臨時調班)直接在格子上打字即可,月結一樣立刻反映——"
  "但那一格的公式會被蓋掉、不再跟著週班表走,所以會自動加上紅色外框提醒。"
  "要恢復連動,從旁邊沒被改過的格子複製一格貼回來。",
], ):
    put(ds, f"A{TALLY0+len(CLINICS)+1+_i}", _t, font(9, color="808080"), None, LEFT, border=False)

dv_doc = DataValidation(type="list", formula1=R_DC, allow_blank=True,
                        showErrorMessage=True, errorTitle="診次代碼無效",
                        error="請填院所代碼(悅睿匯曜寶)或休假代碼。")
ds.add_data_validation(dv_doc)
dv_doc.add(f"{get_column_letter(DS_C0)}{DS_ROW0}:{get_column_letter(LAST_C)}{DS_ROW1}")
dv_deid = DataValidation(type="list", formula1="設定!$Q$6:$Q$45", allow_blank=True,
                         showErrorMessage=True, errorStyle="warning",
                         errorTitle="不在醫師名冊中",
                         error="這個編號不在醫師名冊裡。確定要用請按「是」,"
                               "但建議先到「設定」分頁把人加進名冊。")
ds.add_data_validation(dv_deid); dv_deid.add(f"A{DS_ROW0}:A{DS_ROW1}")

E0 = get_column_letter(DS_C0); EL = get_column_letter(LAST_C)
GRID = f"{E0}{DS_ROW0}:{EL}{DS_ROW1}"
for code in CLINIC_CODES:
    ds.conditional_formatting.add(GRID, FormulaRule(
        formula=[f'{E0}{DS_ROW0}="{code}"'], fill=CLINIC_FILL[code], stopIfTrue=True))
ds.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'COUNTIF({R_DC},{E0}{DS_ROW0})>0'], fill=LEAVE_FILL, stopIfTrue=True))
ds.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'OR({E0}$4="六",{E0}$4="日")'], fill=WKND_FILL))
ds.conditional_formatting.add(f"{E0}4:{EL}5", FormulaRule(
    formula=[f'OR({E0}$4="六",{E0}$4="日")'], fill=WKND_FILL))
# 手動覆蓋掉公式的格子加邊框標示,提醒該格已與週班表脫鉤
ds.conditional_formatting.add(GRID, FormulaRule(
    formula=[f'AND({E0}$4<>"",NOT(_xlfn.ISFORMULA({E0}{DS_ROW0})))'],
    border=Border(left=Side(style="medium", color="C00000"),
                  right=Side(style="medium", color="C00000"),
                  top=Side(style="medium", color="C00000"),
                  bottom=Side(style="medium", color="C00000"))))

for k, code in enumerate(CLINIC_CODES):
    r = TALLY0 + k
    ds.conditional_formatting.add(f"{E0}{r}:{EL}{r}", FormulaRule(
        formula=[f'AND({E0}$4<>"",ISNUMBER({E0}{r}),{E0}{r}=0)'], fill=ALERT_FILL))

# ================================================================ 4. 醫師月結
dm = wb.create_sheet("醫師月結")
dm.sheet_view.showGridLines = False
dm.freeze_panes = "C5"
DM_COLS = ([("A","員工編號",11), ("B","姓名",11), ("C","專科",15), ("D","職務",11)]
           + [(get_column_letter(5+i), f"{c[1]}診次", 10) for i, c in enumerate(CLINICS)]
           + [("J","總診次",9), ("K","服務院所數",11), ("L","教育訓練",9),
              ("M","排休",7), ("N","特休",7), ("O","病假",7), ("P","事假",7),
              ("Q","公假",7), ("R","國定假日",9), ("S","休診",7),
              ("T","休假/休診合計",12), ("U","未排診次",10)])
for col, _, w in DM_COLS: dm.column_dimensions[col].width = w
put(dm, "A1", "醫師月結(全自動 · 以診次為單位)", TITLE_F, border=False)
dm.merge_cells("A1:U1")
put(dm, "A2", "期間", font(10, True), SUB_FILL, CTR)
dm["A2"].alignment = CTR
dm["B2"] = PERIOD
dm["B2"].font, dm["B2"].fill, dm["B2"].alignment, dm["B2"].border = (
    font(10, True), CALC_FILL, CTR, BOX)
header_row(dm, 4, [lab for _, lab, _ in DM_COLS], start_col=1, height=30)
DM_R0 = 5
for i in range(N_DOC):
    r = DM_R0 + i; sr = DS_ROW0 + i
    rng = f"醫師班表!${E0}{sr}:${EL}{sr}"
    g = f'IF($A{r}="","",'
    vals = {
        "A": (DOCTORS[i][0] if i < len(DOCTORS)
              else f'=IF(醫師班表!$A{sr}="","",醫師班表!$A{sr})'),
        "B": lookup(DOCTORS[i][0] if i < len(DOCTORS) else None, 0, f'={g}醫師班表!$B{sr})'),
        "C": lookup(DOCTORS[i][0] if i < len(DOCTORS) else None, 2, f'={g}醫師班表!$C{sr})'),
        "D": lookup(DOCTORS[i][0] if i < len(DOCTORS) else None, 3, f'={g}醫師班表!$D{sr})'),
        "J": f'={g}SUM($E{r}:$I{r}))',
        "K": f'={g}SUMPRODUCT(($E{r}:$I{r}>0)*1))',
        "L": f'={g}COUNTIF({rng},"訓"))',
        "T": f'={g}SUM($M{r}:$S{r}))',
        "U": f'={g}{DAYS_FX}*3-COUNTIF({rng},"?*"))',
    }
    for k, c in enumerate(CLINIC_CODES):
        vals[get_column_letter(5+k)] = f'={g}COUNTIF({rng},"{c}"))'
    for k, code in enumerate(DOC_LEAVE):
        vals[get_column_letter(13+k)] = f'={g}COUNTIF({rng},"{code}"))'
    for col, _, _w in DM_COLS:
        c = dm[f"{col}{r}"]
        c.value, c.font, c.border, c.alignment, c.fill = (
            vals[col], font(9), BOX, CTR, CALC_FILL)
        if col == "C": c.alignment = LEFT
        if col in [get_column_letter(5+k) for k in range(5)]:
            c.fill = CLINIC_FILL[CLINIC_CODES[[get_column_letter(5+k)
                     for k in range(5)].index(col)]]
DM_TOT = DM_R0 + N_DOC
put(dm, f"A{DM_TOT}", "合計", font(10, True), SUB_FILL, CTR)
dm.merge_cells(f"A{DM_TOT}:D{DM_TOT}")
for k in range(4, 21):
    col = get_column_letter(k+1)
    c = dm[f"{col}{DM_TOT}"]
    c.value = f"=SUM({col}{DM_R0}:{col}{DM_TOT-1})"
    c.font, c.fill, c.border, c.alignment = font(10, True), SUB_FILL, BOX, CTR
for i, t in enumerate([
  "※ 醫師以「診次」為單位,不是天數。請假 3 個診次等於請一天。",
  "※ 本頁全部由「醫師班表」計算而來。改週班表或改醫師班表,這裡都會立刻跟著變。",
  "※ 服務院所數 = 本月實際有排診的院所家數,可用來看跨院負荷。",
  "※ 未排診次 = 當月總診次格數(天數 × 3)扣掉已填格數,不代表應該排滿。",
]):
    put(dm, f"A{DM_TOT+2+i}", t, font(9, color="808080"), None, LEFT, border=False)


# ================================================================ 5. 醫護長與管理部班表
AS_C0, AS_C1 = 5, 5 + DAYS_IN_MONTH - 1      # E..AI
AS_ROW0 = 5
AS_ROW1 = AS_ROW0 + N_ASST - 1               # 64
A0 = get_column_letter(AS_C0); A1 = get_column_letter(AS_C1)

asx = wb.create_sheet("醫護長與管理部班表")
asx.sheet_view.showGridLines = False
asx.freeze_panes = "E5"
for col, w in {"A":9,"B":10,"C":9,"D":10}.items():
    asx.column_dimensions[col].width = w
for c in range(AS_C0, AS_C1 + 1):
    # 5.6 是為了「早午」「午晚」「全日」這類兩個字的代碼。4.4 只夠一個字,
    # 兩字會被右邊的格子切掉(隔壁有內容就不會溢出顯示)。
    asx.column_dimensions[get_column_letter(c)].width = 5.6
put(asx, "A1", title_of("醫護長與管理部班表(工時制)"), TITLE_F,
    IN_FILL if PENDING else None, LEFT, border=False)
asx.merge_cells(start_row=1, start_column=1, end_row=1, end_column=AS_C1)
put(asx, "A2", "期間", font(10, True), SUB_FILL, CTR)
asx["B2"] = PERIOD
asx["B2"].font, asx["B2"].fill, asx["B2"].alignment, asx["B2"].border = (
    font(10, True), CALC_FILL, CTR, BOX)
put(asx, "E2",
    "← 各院所醫護長與管理部各填自己那一列。院所欄由人員名冊自動帶出。"
    "要選員工編號請點「資料列」的格子(第 5 列以下),標題列上的箭頭是篩選鈕不是選單。",
    font(9, color="808080"), None, LEFT, border=False)

for lab, col in (("員工編號",1), ("姓名",2), ("職類",3), ("院所",4)):
    c = asx.cell(row=3, column=col, value=lab)
    c.font, c.fill, c.alignment, c.border = HDR_F, HDR_FILL, CTR, BOX
    asx.merge_cells(start_row=3, start_column=col, end_row=4, end_column=col)
for c in range(AS_C0, AS_C1 + 1):
    d = c - AS_C0 + 1; L = get_column_letter(c)
    h = asx.cell(row=3, column=c, value=d)
    h.font, h.fill, h.alignment, h.border = HDR_F, HDR_FILL, CTR, BOX
    w = asx.cell(row=4, column=c)
    w.value = wd_of(d)
    w.font, w.fill, w.alignment, w.border = font(9, True), SUB_FILL, CTR, BOX

for i in range(N_ASST):
    r = AS_ROW0 + i
    asx.row_dimensions[r].height = 17
    eid = ASSISTANTS[i][0] if i < len(ASSISTANTS) else None
    a = asx.cell(row=r, column=1, value=eid)
    a.font, a.fill, a.border, a.alignment = font(9), IN_FILL, BOX, CTR
    for col, src, ix in ((2, R_NAME, 0), (3, R_ROLE, 1), (4, R_HOME, 4)):
        c = asx.cell(row=r, column=col)
        c.value = lookup(eid, ix, f'=IFERROR(INDEX({src},MATCH($A{r},{R_EID},0)),"")')
        c.font, c.fill, c.border, c.alignment = font(9), CALC_FILL, BOX, CTR
    for c in range(AS_C0, AS_C1 + 1):
        cc = asx.cell(row=r, column=c)
        cc.font, cc.border, cc.alignment = font(9), BOX, CTR

AS_TALLY = AS_ROW1 + 2
# 五間院所各一列,再加「管理部」一列。院所那五列只數醫護長:那一列問的是
# 「這間院所今天有沒有醫護長」,管理部同樣在晶匯上班,但不能代替醫護長,
# 混在一起數會讓「小玲請假、三位管理部在」看起來像有人顧,那正是要抓的狀況。
# (short, 顯示標籤, 底色, 職類, 是否比對院所)
TALLY_ROWS = [(short, f"{code} {short}", CLINIC_FILL[code], "醫護長", True)
              for code, short, *_r in CLINICS]
TALLY_ROWS.append(("", "管 管理部", SUB_FILL, "管理部", False))
put(asx, f"A{AS_TALLY-1}", "當日在班人數(自動計算)", font(10, True), SUB_FILL, LEFT)
asx.merge_cells(start_row=AS_TALLY-1, start_column=1, end_row=AS_TALLY-1, end_column=4)
for k, (short, label, fill_, jobtype, by_clinic) in enumerate(TALLY_ROWS):
    r = AS_TALLY + k
    put(asx, f"A{r}", label, font(9, True), fill_, CTR)
    asx.merge_cells(start_row=r, start_column=1, end_row=r, end_column=4)
    for c in range(AS_C0, AS_C1 + 1):
        L = get_column_letter(c)
        cell = asx.cell(row=r, column=c)
        d_ = c - AS_C0 + 1
        # 院所欄可能寫成「晶睿·晶曜」(兼管兩間),所以用包含比對而非完全相等
        clinic_term = (f'ISNUMBER(SEARCH("{short}",$D${AS_ROW0}:$D${AS_ROW1}))*'
                       if by_clinic else "")
        cell.value = (f'=IF({L}$3="","",'
                      f'IF(COUNTIF({R_HOL},DATE({SET_Y},{SET_M},{d_}))>0,"假",'
                      f'IF({L}$4="日","休",'
                      f'SUMPRODUCT({clinic_term}'
                      f'($C${AS_ROW0}:$C${AS_ROW1}="{jobtype}")'
                      f'*(COUNTIF({R_WC_W},{L}${AS_ROW0}:{L}${AS_ROW1})>0)))))')
        cell.font, cell.fill, cell.alignment, cell.border = (
            font(9, True), CALC_FILL, CTR, BOX)
        cell.number_format = FMT_CNT
put(asx, f"A{AS_TALLY+len(TALLY_ROWS)+1}",
    "※ 一間院所一位醫護長,所以檢核的是「當天有沒有醫護長在」:掛 0 會變紅。"
    "國定假日顯示「假」、週日顯示「休」,都不示警。文君兼管晶睿與晶曜,兩間都會算到她。"
    "管理部三位也在晶匯,但不列入晶匯那一列——他們代替不了醫護長,"
    "混在一起數會讓「醫護長請假、管理部在」看起來像有人顧。他們自己在最後一列。"
    "門檻要調請改這幾列的條件式格式。", font(9, color="808080"), None, LEFT, border=False)

# 用內嵌清單而非跨分頁範圍:選單少一層相依,代碼表挪位置也不會壞
_wc_inline = ",".join(c[0] for c in WORK_CODES)
assert len(_wc_inline) <= 255, f"班別代碼清單過長({len(_wc_inline)} 字元)"
dv_wc = DataValidation(type="list", formula1=f'"{_wc_inline}"', allow_blank=True,
                       showErrorMessage=True, errorStyle="warning",
                       errorTitle="班別代碼無效",
                       error="這不是「設定」分頁定義的班別代碼。確定要用請按「是」。")
asx.add_data_validation(dv_wc)
dv_wc.add(f"{A0}{AS_ROW0}:{A1}{AS_ROW1}")
dv_aeid = DataValidation(type="list", formula1=f"設定!$S$6:$S${5+N_ASST_SLOTS}",
                         allow_blank=True, showErrorMessage=True, errorStyle="warning",
                         errorTitle="不在人員名冊中",
                         error="這個編號不在工時制人員名冊裡。確定要用請按「是」。")
asx.add_data_validation(dv_aeid); dv_aeid.add(f"A{AS_ROW0}:A{AS_ROW1}")

AGRID = f"{A0}{AS_ROW0}:{A1}{AS_ROW1}"
asx.conditional_formatting.add(AGRID, FormulaRule(
    formula=[f'AND({A0}$3<>"",$A{AS_ROW0}<>"",{A0}{AS_ROW0}="")'],
    fill=GAP_FILL, stopIfTrue=True))
asx.conditional_formatting.add(AGRID, FormulaRule(
    formula=[f'{A0}{AS_ROW0}="支"'], fill=SUPP_FILL, stopIfTrue=True))
asx.conditional_formatting.add(AGRID, FormulaRule(
    formula=[f'COUNTIF({R_WC_L},{A0}{AS_ROW0})>0'], fill=LEAVE_FILL, stopIfTrue=True))
asx.conditional_formatting.add(AGRID, FormulaRule(
    formula=[f'OR({A0}$4="六",{A0}$4="日")'], fill=WKND_FILL))
asx.conditional_formatting.add(f"{A0}3:{A1}4", FormulaRule(
    formula=[f'OR({A0}$4="六",{A0}$4="日")'], fill=WKND_FILL))
for k in range(len(TALLY_ROWS)):
    r = AS_TALLY + k
    asx.conditional_formatting.add(f"{A0}{r}:{A1}{r}", FormulaRule(
        formula=[f'AND({A0}$3<>"",ISNUMBER({A0}{r}),{A0}{r}=0)'], fill=ALERT_FILL))

# ================================================================ 6. 打卡匯入
PUNCH_R0 = 3
PUNCH_R1 = PUNCH_R0 + PUNCH_N - 1
pc = wb.create_sheet("打卡匯入")
pc.sheet_view.showGridLines = False
pc.freeze_panes = "A3"
# B 是日期欄,留 14 才放得下「2026/10/10」這種兩位數月日
for col, w in {"A":13,"B":14,"C":12,"D":12,"E":16,"F":16,"G":2,"H":76}.items():
    pc.column_dimensions[col].width = w
put(pc, "A1", title_of("打卡匯入(把打卡機匯出的資料貼在下面四欄)"), TITLE_F,
    IN_FILL if PENDING else None, LEFT, border=False)
pc.merge_cells("A1:F1")
put(pc, "H1", "貼上規則:一天一列。日期要是真正的日期格式、時間要是真正的時間格式,"
    "不能是文字。E、F 欄是公式,不要覆蓋。", font(9, color="808080"), None, WRAP, border=False)
header_row(pc, 2, ["員工編號","日期","上班時間","下班時間","對照鍵(自動)","檢核(自動)"],
           start_col=1, height=24)
for i in range(PUNCH_R0, PUNCH_R1 + 1):
    for j in range(4):
        c = pc.cell(row=i, column=1 + j)
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

# ================================================================ 7. 出勤紀錄
ATT_R0 = 3
ATT_R1 = ATT_R0 + N_ASST * DAYS_IN_MONTH - 1
at = wb.create_sheet("出勤紀錄")
at.sheet_view.showGridLines = False
at.freeze_panes = "E3"
AT_COLS = [("A","日期",14), ("B","星期",6), ("C","員工編號",11), ("D","姓名",11),
           ("E","院所",10), ("F","職類",9), ("G","排班代碼",9), ("H","應到",8),
           ("I","應退",8), ("J","實際上班",10), ("K","實際下班",10),
           ("L","休息(分)",9), ("M","實際工時",10), ("N","出勤異常",11),
           ("O","遲到(分)",9), ("P","早退(分)",9), ("Q","加班時數",10),
           ("R","加班別",10), ("S","備註",22)]
for col, _, w in AT_COLS: at.column_dimensions[col].width = w
put(at, "A1", title_of("出勤紀錄(法定紀錄:逐日記載至分鐘,保存 5 年)"), TITLE_F,
    IN_FILL if PENDING else None, LEFT, border=False)
at.merge_cells("A1:S1")
# 表頭每一欄都會有一個小箭頭,那是篩選鈕不是下拉選單——這一頁除了備註欄之外
# 全是公式,沒有東西可選。會被誤認成選單,所以在旁邊直接講明。
at.column_dimensions["U"].width = 2
put(at, "V1", "↑ 表頭上的箭頭是「篩選」鈕,不是下拉選單。"
    "這一頁只有最右邊的備註欄要手填,其餘都是公式自動算的。"
    "處理異常請點「出勤異常」欄的箭頭,把「正常」取消勾選。",
    font(9, color="808080"), None, LEFT, border=False)
header_row(at, 2, [lab for _, lab, _ in AT_COLS], start_col=1, height=28)
for pi in range(N_ASST):
    srow = AS_ROW0 + pi
    for d in range(1, DAYS_IN_MONTH + 1):
        r = ATT_R0 + pi * DAYS_IN_MONTH + (d - 1)
        scol = get_column_letter(AS_C0 + d - 1)
        DUE = f'IFERROR(INDEX({R_WC_HRS},MATCH($G{r},{R_WC},0)),0)'
        fx = {
"A": f'=IF(OR(醫護長與管理部班表!$A{srow}="",{d}>{DAYS_FX}),"",DATE({SET_Y},{SET_M},{d}))',
"B": f'=IF($A{r}="","",INDEX({R_WEEK},WEEKDAY($A{r},1)))',
"C": f'=IF($A{r}="","",醫護長與管理部班表!$A{srow})',
"D": f'=IF($C{r}="","",醫護長與管理部班表!$B{srow})',
"E": f'=IF($C{r}="","",醫護長與管理部班表!$D{srow})',
"F": f'=IF($C{r}="","",醫護長與管理部班表!$C{srow})',
"G": f'=IF($C{r}="","",醫護長與管理部班表!{scol}{srow})',
"H": (f'=IF($G{r}="","",IFERROR(IF(INDEX({R_WC_IN},MATCH($G{r},{R_WC},0))=0,"",'
      f'INDEX({R_WC_IN},MATCH($G{r},{R_WC},0))),""))'),
"I": (f'=IF($G{r}="","",IFERROR(IF(INDEX({R_WC_OUT},MATCH($G{r},{R_WC},0))=0,"",'
      f'INDEX({R_WC_OUT},MATCH($G{r},{R_WC},0))),""))'),
"J": (f'=IFERROR(INDEX(打卡匯入!$C${PUNCH_R0}:$C${PUNCH_R1},'
      f'MATCH($C{r}&"|"&DAY($A{r}),打卡匯入!$E${PUNCH_R0}:$E${PUNCH_R1},0)),"")'),
"K": (f'=IFERROR(INDEX(打卡匯入!$D${PUNCH_R0}:$D${PUNCH_R1},'
      f'MATCH($C{r}&"|"&DAY($A{r}),打卡匯入!$E${PUNCH_R0}:$E${PUNCH_R1},0)),"")'),
"L": f'=IF(OR($J{r}="",$K{r}=""),"",IFERROR(INDEX({R_WC_RST},MATCH($G{r},{R_WC},0)),0))',
"M": f'=IF(OR($J{r}="",$K{r}=""),"",ROUND(($K{r}-$J{r})*24-$L{r}/60,2))',
"O": (f'=IF(OR($C{r}="",$H{r}="",$J{r}=""),0,'
      f'IF(ROUND(($J{r}-$H{r})*1440,0)>{P_GRACE},ROUND(($J{r}-$H{r})*1440,0),0))'),
"P": (f'=IF(OR($C{r}="",$I{r}="",$K{r}=""),0,'
      f'IF(ROUND(($I{r}-$K{r})*1440,0)>{P_GRACE},ROUND(($I{r}-$K{r})*1440,0),0))'),
"N": (f'=IF($C{r}="","",'
      f'IF(AND(COUNTIF({R_WC_L},$G{r})>0,$J{r}<>""),"假日出勤",'
      f'IF(AND(COUNTIF({R_WC_W},$G{r})>0,$J{r}=""),"未打卡",'
      f'IF($J{r}="","",IF($O{r}>0,"遲到",IF($P{r}>0,"早退","正常"))))))'),
"Q": (f'=IF($M{r}="",0,IF(($M{r}-{DUE})*60>={P_OT_MIN},'
      f'ROUND(FLOOR(($M{r}-{DUE})*60,{P_OT_UNIT})/60,2),0))'),
"R": (f'=IF(OR($C{r}="",$Q{r}=0),"",IF($G{r}="國","國定假日",'
      f'IF($B{r}="日","例假",IF($B{r}="六","休息日","平日"))))'),
"S": None,
        }
        for col, _, _w in AT_COLS:
            c = at[f"{col}{r}"]
            v = fx.get(col)
            if v is not None: c.value = v; c.fill = CALC_FILL
            else: c.fill = IN_FILL
            c.font, c.border, c.alignment = font(9), BOX, CTR
            if col == "A": c.number_format = FMT_DATE
            elif col in ("H","I","J","K"): c.number_format = FMT_TIME
            elif col in ("M","Q"): c.number_format = FMT_HR
            elif col in ("L","O","P"): c.number_format = FMT_MIN
            elif col == "S": c.alignment = LEFT
at.auto_filter.ref = f"A2:S{ATT_R1}"
AT_RANGE = f"A{ATT_R0}:S{ATT_R1}"
at.conditional_formatting.add(AT_RANGE, FormulaRule(
    formula=[f'AND($C{ATT_R0}<>"",$N{ATT_R0}<>"",$N{ATT_R0}<>"正常")'],
    fill=GAP_FILL, stopIfTrue=True))
at.conditional_formatting.add(AT_RANGE, FormulaRule(
    formula=[f'AND($C{ATT_R0}<>"",$Q{ATT_R0}>0)'], fill=OT_FILL))
at.conditional_formatting.add(f"N{ATT_R0}:N{ATT_R1}", FormulaRule(
    formula=[f'AND($N{ATT_R0}<>"",$N{ATT_R0}<>"正常")'], fill=ALERT_FILL))

# ================================================================ 8. 月結統計
ms = wb.create_sheet("月結統計")
ms.sheet_view.showGridLines = False
ms.freeze_panes = "E5"
MS_COLS = [("A","員工編號",11), ("B","姓名",11), ("C","職類",9), ("D","院所",10),
           ("E","出勤天數",9), ("F","排班工時",9), ("G","排休",7), ("H","特休",7),
           ("I","病假",7), ("J","事假",7), ("K","公假",7), ("L","國定假日",9),
           ("M","支援他院",9), ("N","加班時數",10), ("O","實際工時",10),
           ("P","遲到次數",9), ("Q","早退次數",9), ("R","未打卡",8),
           ("S","排班完整度檢核",15)]
for col, _, w in MS_COLS: ms.column_dimensions[col].width = w
put(ms, "A1", title_of("月結統計 — 醫護長與管理部(全自動)"), TITLE_F,
    IN_FILL if PENDING else None, LEFT, border=False)
ms.merge_cells("A1:S1")
put(ms, "A2", "期間", font(10, True), SUB_FILL, CTR)
ms["B2"] = PERIOD
ms["B2"].font, ms["B2"].fill, ms["B2"].alignment, ms["B2"].border = (
    font(10, True), CALC_FILL, CTR, BOX)
header_row(ms, 4, [lab for _, lab, _ in MS_COLS], start_col=1, height=30)
MS_R0 = 5
AT_C = f"出勤紀錄!$C${ATT_R0}:$C${ATT_R1}"
AT_N = f"出勤紀錄!$N${ATT_R0}:$N${ATT_R1}"
AT_Q = f"出勤紀錄!$Q${ATT_R0}:$Q${ATT_R1}"
for i in range(N_ASST):
    r = MS_R0 + i; sr = AS_ROW0 + i
    rng = f"醫護長與管理部班表!${A0}{sr}:${A1}{sr}"
    g = f'IF($A{r}="","",'
    vals = {
"A": f'=IF(醫護長與管理部班表!$A{sr}="","",醫護長與管理部班表!$A{sr})',
"B": f'={g}醫護長與管理部班表!$B{sr})',
"C": f'={g}醫護長與管理部班表!$C{sr})',
"D": f'={g}醫護長與管理部班表!$D{sr})',
"E": f'={g}SUMPRODUCT(COUNTIF({rng},{R_WC}),{R_WC_ATT}))',
"F": f'={g}SUMPRODUCT(COUNTIF({rng},{R_WC}),{R_WC_HRS}))',
"G": f'={g}COUNTIF({rng},"OFF"))',
"H": f'={g}COUNTIF({rng},"特"))',
"I": f'={g}COUNTIF({rng},"病"))',
"J": f'={g}COUNTIF({rng},"事"))',
"K": f'={g}COUNTIF({rng},"公"))',
"L": f'={g}COUNTIF({rng},"國"))',
"M": f'={g}COUNTIF({rng},"支"))',
"N": f'={g}SUMIF({AT_C},$A{r},{AT_Q}))',
"O": f'={g}$F{r}+$N{r})',
"P": f'={g}COUNTIFS({AT_C},$A{r},{AT_N},"遲到"))',
"Q": f'={g}COUNTIFS({AT_C},$A{r},{AT_N},"早退"))',
"R": f'={g}COUNTIFS({AT_C},$A{r},{AT_N},"未打卡"))',
"S": (f'={g}IF(COUNTA({rng})={DAYS_FX},"OK",'
      f'"⚠ 缺"&({DAYS_FX}-COUNTA({rng}))&"天"))'),
    }
    for col, _, _w in MS_COLS:
        c = ms[f"{col}{r}"]
        c.value, c.font, c.border, c.alignment, c.fill = (
            vals[col], font(9), BOX, CTR, CALC_FILL)
        if col in ("F","N","O"): c.number_format = "0.0"
MS_TOT = MS_R0 + N_ASST
put(ms, f"A{MS_TOT}", "合計", font(10, True), SUB_FILL, CTR)
ms.merge_cells(f"A{MS_TOT}:D{MS_TOT}")
for col in "EFGHIJKLMNOPQR":
    c = ms[f"{col}{MS_TOT}"]
    c.value = f"=SUM({col}{MS_R0}:{col}{MS_TOT-1})"
    c.font, c.fill, c.border, c.alignment = font(10, True), SUB_FILL, BOX, CTR
    if col in ("F","N","O"): c.number_format = "0.0"
put(ms, f"S{MS_TOT}", "", font(), SUB_FILL, CTR)
ms.conditional_formatting.add(f"S{MS_R0}:S{MS_TOT-1}", FormulaRule(
    formula=[f'AND($A{MS_R0}<>"",S{MS_R0}<>"OK")'], fill=ALERT_FILL))
for col in ("P","Q","R"):
    ms.conditional_formatting.add(f"{col}{MS_R0}:{col}{MS_TOT-1}", FormulaRule(
        formula=[f'AND($A{MS_R0}<>"",{col}{MS_R0}>0)'], fill=GAP_FILL))
for i, t in enumerate([
  "※ 出勤天數與排班工時依「設定」分頁的班別代碼權重自動加總,改權重這裡跟著變。",
  "※ 加班時數由「出勤紀錄」加總,已套用加班門檻。不使用打卡匯入的院所可覆蓋成手動數字。",
  "※ 加班時數為「實際工時 − 排班工時」,未依勞基法第 24 條換算費率,薪資請另行計算。",
  "※ 醫師的月結在「醫師月結」分頁,以診次為單位計算。",
]):
    put(ms, f"A{MS_TOT+2+i}", t, font(9, color="808080"), None, LEFT, border=False)

if not WITH_ASSISTANT:
    for _n in STAFF_SHEETS:
        del wb[_n]

wb.save(OUT)
print("saved:", OUT)
print(f"醫師班表:{N_DOC} 列 × {DAYS_IN_MONTH*3} 診次欄")
print(f"醫護長與管理部班表:{N_ASST} 列 · 出勤紀錄:{N_ASST*DAYS_IN_MONTH} 列 · 打卡匯入:{PUNCH_N} 列")
print("分頁:", wb.sheetnames)
