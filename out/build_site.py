# -*- coding: utf-8 -*-
"""班表看板(階段一:唯讀)。從 build_roster.py 的同一批資料產生一頁式靜態網站。

    python3 build_site.py            # 產生 班表看板.html

為什麼要「模擬 Excel 的公式」而不是直接呼叫 build_roster.expand_month:
Excel 的醫師班表是公式驅動的——格子讀醫師週班表的標記(悅、悅(隔)、悅/睿),
再依單雙週決定當天是哪一間。expand_month 是另一套 Python 實作,而且目前
在 build_roster 裡根本沒被呼叫,等於從未與出貨的 Excel 對過帳。
網站若用它,就成了第二個真相來源——正是路線圖裡列為風險的「兩份資料分叉」。

所以這裡照著公式那條路走:wk_label 產生標記 → 依單雙週解析 → 得到院所。
產生後再用 expand_month 獨立核對一次,兩套實作全數相符才寫檔。
"""
import importlib.util, sys, io, json, contextlib, pathlib, datetime as dt

# 一律以腳本自己所在的資料夾為準,不寫死絕對路徑——這支要在別人的電腦上跑。
HERE    = pathlib.Path(__file__).resolve().parent
# Cloudflare Pages 要的首頁檔名是 index.html。與其叫人手動改名(很容易忘),
# 直接產生一個現成的資料夾,整個拖上去就好。
UPDIR   = HERE / "上傳這個資料夾"
OUT     = str(UPDIR / "index.html")              # 對外託管用(完整文件)
PREVIEW = str(HERE / "班表看板_預覽.html")        # Artifact 預覽用(只有內容)

# ---------------------------------------------------------------- 讀資料
_spec = importlib.util.spec_from_file_location("br", str(HERE / "build_roster.py"))
br = importlib.util.module_from_spec(_spec)
_argv, sys.argv = sys.argv, ["build_roster.py"]
with contextlib.redirect_stdout(io.StringIO()):
    _spec.loader.exec_module(br)
sys.argv = _argv
sys.path.insert(0, str(HERE))
import weekly as wk

YEAR, MONTH, NDAYS = br.YEAR, br.MONTH, br.DAYS_IN_MONTH
SESSIONS = br.SESSIONS
CLINICS = br.CLINICS
DOCTORS = br.DOCTORS
WEEK_CH = "一二三四五六日"

CLINIC_HEX = {"悅": "FBE3EC", "睿": "D9E7F5", "匯": "DCEEDC",
              "曜": "E8DFF2", "寶": "FAE8A0"}
CLINIC_INK = {"悅": "9C4466", "睿": "2C5F8E", "匯": "3B6E3B",
              "曜": "6A4A92", "寶": "8A6310"}

# ---------------------------------------------------------------- 週班表標記
# 與 build_roster 寫進「醫師週班表」的那張表同一份資料、同一個排序規則。
LABEL = {}                                   # (醫師, 星期1-6, 診次0-2) -> 標記字串
for nm, slots in br.DOC_WEEK.items():
    for w in range(1, 7):
        for t in range(3):
            hit = [(cl, fg) for (ww, tt, cl), fg in slots.items()
                   if ww == w and tt == t]
            if not hit:
                continue
            # 排序讓「單數週的院所」永遠排在前面,顯示與解析才對得上
            hit = sorted(hit, key=lambda x: ((nm, w, x[0]) in br.ALT_B, x[0]))
            LABEL[(nm, w, t)] = br.wk_label(hit)

def cell_value(name, date, sidx):
    """醫師班表某一格會顯示什麼。逐字對應 Excel 的那條公式鏈。"""
    wd = date.weekday() + 1                              # 1=一 … 7=日
    parity = ((date - br.ALT_EPOCH).days // 7) % 2       # 第 58 列
    if date in br.HOLIDAY_SET:                           # 第 59 列:國
        return "國"
    if (wd, sidx) in br.ALL_CLOSED:                      # 第 59 列:休
        return "休"
    if wd == 7:                                          # 週日不在週班表的 18 格內
        return "休"
    W = LABEL.get((name, wd, sidx), "")
    if not W:
        return ""
    if len(W) > 1 and W[1] == "/":                       # 悅/睿
        return W[parity * 2]
    if len(W) > 2 and W[2] == "隔":                       # 悅(隔)
        return W[0] if parity == 0 else ""
    if len(W) > 1 and W[1] == "(":                       # 悅(不定) / 悅(註)
        return W[0]
    return W

DAYS = []
for d in range(1, NDAYS + 1):
    date = dt.date(YEAR, MONTH, d)
    DAYS.append({"d": d, "wd": date.weekday() + 1,
                 "w": WEEK_CH[date.weekday()],
                 "hol": br.HOLIDAY_SET.get(date, "")})

# 醫師 × 日 × 診次 → 院所代碼 / 國 / 休 / ""
DOCM = {}
for eid, nm, eng, spec, duty, teams in DOCTORS:
    DOCM[nm] = [[cell_value(nm, dt.date(YEAR, MONTH, d), s) for s in range(3)]
                for d in range(1, NDAYS + 1)]

# ---------------------------------------------------------------- 獨立核對
# expand_month 是另一套實作。兩者全數相符,才敢說網站與 Excel 是同一份班表。
mismatch = []
for eid, nm, *_r in DOCTORS:
    flat = br.expand_month(nm, YEAR, MONTH, NDAYS)
    for d in range(NDAYS):
        for s in range(3):
            a, b = DOCM[nm][d][s], flat[d * 3 + s]
            if a != b:
                mismatch.append((nm, d + 1, SESSIONS[s], a, b))
print(f"與 expand_month 逐格核對:{len(DOCTORS)*NDAYS*3} 格,不符 {len(mismatch)}")
for m in mismatch[:8]:
    print("   ✗", m)
assert not mismatch, "兩套展開邏輯不一致,先查清楚再產網站"

# 院所 × 日 × 診次 → 在診醫師
GRID = {c[0]: [[[] for _ in range(3)] for _ in range(NDAYS)] for c in CLINICS}
for eid, nm, *_r in DOCTORS:
    for d in range(NDAYS):
        for s in range(3):
            v = DOCM[nm][d][s]
            if v in GRID:
                GRID[v][d][s].append(nm)

OPENFLAG = {c[0]: [[1 if (w, t) in br.OPEN[c[0]] else 0 for t in range(3)]
                   for w in range(1, 8)] for c in CLINICS}

DATA = {
    "year": YEAR, "month": MONTH, "ndays": NDAYS,
    "built": dt.date.today().isoformat(),
    "sessions": SESSIONS,
    "week": list(WEEK_CH),
    "clinics": [{"code": c[0], "short": c[1], "full": c[2], "head": c[3],
                 "tel": c[4], "addr": c[5],
                 "bg": CLINIC_HEX[c[0]], "ink": CLINIC_INK[c[0]],
                 "time": wk.SESSION_TIME[c[0]],
                 "open": OPENFLAG[c[0]]} for c in CLINICS],
    "doctors": [{"eid": e, "name": n, "spec": sp, "duty": du, "teams": tm}
                for e, n, _g, sp, du, tm in DOCTORS],
    "days": DAYS,
    "grid": GRID,
    "docm": DOCM,
}

# ---------------------------------------------------------------- 頁面
PAGE = r"""<title>日三班表看板</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500;600&family=Noto+Sans+TC:wght@400;500;700;900&display=swap">
<style>
:root{
  --paper:#FCFCFA; --surface:#FFFFFF; --sunk:#F3F4F2;
  --ink:#152A2C; --body:#33474A; --muted:#728486;
  --line:#E2E4DF; --line-strong:#C6CBC5;
  --petrol:#1F5A57; --petrol-soft:#E4EEEC; --on-petrol:#FFFFFF;
  --alert:#A8433C; --alert-soft:#FBE0DE;
  --sans:"Noto Sans TC",-apple-system,"Segoe UI","PingFang TC","Microsoft JhengHei",sans-serif;
  --mono:"IBM Plex Mono",ui-monospace,Menlo,monospace;
  --c-悅:#FBE3EC; --c-睿:#D9E7F5; --c-匯:#DCEEDC; --c-曜:#E8DFF2; --c-寶:#FAE8A0;
  --i-悅:#8E3A5B; --i-睿:#25547E; --i-匯:#33622F; --i-曜:#5C3E82; --i-寶:#7A560B;
}
@media (prefers-color-scheme:dark){ :root:not([data-theme="light"]){
  --paper:#0F1618; --surface:#161F21; --sunk:#1C2628;
  --ink:#E7EDEB; --body:#BFCCCB; --muted:#8CA0A1;
  --line:#253335; --line-strong:#3A4B4D;
  --petrol:#68AFA6; --petrol-soft:#1B3230; --on-petrol:#0F1618;
  --alert:#E08B84; --alert-soft:#3E2320;
  --c-悅:#43242F; --c-睿:#1E3448; --c-匯:#233A23; --c-曜:#33263F; --c-寶:#3E3315;
  --i-悅:#F0B7CB; --i-睿:#A9CDEA; --i-匯:#A9D6A5; --i-曜:#C9B2E4; --i-寶:#E7C85F;
}}
:root[data-theme="dark"]{
  --paper:#0F1618; --surface:#161F21; --sunk:#1C2628;
  --ink:#E7EDEB; --body:#BFCCCB; --muted:#8CA0A1;
  --line:#253335; --line-strong:#3A4B4D;
  --petrol:#68AFA6; --petrol-soft:#1B3230; --on-petrol:#0F1618;
  --alert:#E08B84; --alert-soft:#3E2320;
  --c-悅:#43242F; --c-睿:#1E3448; --c-匯:#233A23; --c-曜:#33263F; --c-寶:#3E3315;
  --i-悅:#F0B7CB; --i-睿:#A9CDEA; --i-匯:#A9D6A5; --i-曜:#C9B2E4; --i-寶:#E7C85F;
}
*{box-sizing:border-box}
/* 這一頁也會被單獨託管,不能依賴 Artifact 外框提供的重設 */
[hidden]{display:none!important}
body{background:var(--paper); color:var(--body); font-family:var(--sans);
  font-size:15px; line-height:1.6; -webkit-font-smoothing:antialiased;}
h1,h2,h3{color:var(--ink); margin:0; font-weight:900; text-wrap:balance;}
p{margin:0}
button{font-family:inherit; font-size:inherit; cursor:pointer}
:focus-visible{outline:2px solid var(--petrol); outline-offset:2px; border-radius:3px}

.wrap{max-width:900px; margin:0 auto; padding-inline:16px; padding-block:0 60px;}

/* header */
header{padding-block:18px 12px;}
.brand{display:flex; align-items:baseline; gap:9px; flex-wrap:wrap;}
h1{font-size:19px; letter-spacing:.01em}
.period{font-family:var(--mono); font-size:12.5px; color:var(--petrol);
  background:var(--petrol-soft); padding:2px 9px; border-radius:99px;}
.built{font-family:var(--mono); font-size:11px; color:var(--muted); margin-left:auto}

.banner{margin-top:10px; background:var(--petrol-soft); border:1px solid var(--line);
  border-left:3px solid var(--petrol); border-radius:0 4px 4px 0; padding:9px 13px;
  font-size:13px; color:var(--ink);}

/* tabs */
nav.tabs{position:sticky; top:env(safe-area-inset-top,0px); z-index:20;
  background:var(--paper); padding-block:10px 8px; margin-top:6px;
  border-bottom:1px solid var(--line); display:flex; gap:6px; overflow-x:auto;}
nav.tabs button{background:transparent; border:1px solid var(--line-strong);
  color:var(--body); padding:6px 15px; border-radius:99px; white-space:nowrap;
  font-weight:700; font-size:14px;}
nav.tabs button[aria-selected="true"]{background:var(--petrol); border-color:var(--petrol);
  color:var(--on-petrol);}

/* filter */
.filter{display:flex; gap:6px; overflow-x:auto; padding-block:12px 2px;}
.filter button{background:var(--surface); border:1px solid var(--line-strong);
  color:var(--body); padding:5px 13px; border-radius:99px; white-space:nowrap;
  font-size:13.5px; font-weight:500;}
.filter button[aria-pressed="true"]{border-color:var(--ink); color:var(--ink);
  font-weight:700;}
.filter button[data-c][aria-pressed="true"]{border-width:1.5px}

/* day nav */
.daynav{display:flex; align-items:center; gap:10px; margin-top:14px;}
.daynav button{background:var(--surface); border:1px solid var(--line-strong);
  color:var(--ink); width:38px; height:38px; border-radius:9px; font-size:17px;
  line-height:1; display:flex; align-items:center; justify-content:center;}
.daynav button:disabled{opacity:.35; cursor:default}
.dlabel{flex:1; text-align:center}
.dlabel b{display:block; color:var(--ink); font-size:19px; font-weight:900;
  font-variant-numeric:tabular-nums;}
.dlabel span{font-size:12.5px; color:var(--muted); font-family:var(--mono)}
.dlabel .hol{color:var(--alert); font-weight:700; font-family:var(--sans)}

main{padding-top:6px}
section[hidden]{display:none!important}

/* clinic cards */
.cards{display:grid; gap:12px; margin-top:16px;}
@media(min-width:680px){.cards{grid-template-columns:1fr 1fr;}}
.card{background:var(--surface); border:1px solid var(--line); border-radius:7px;
  overflow:hidden;}
.card-h{display:flex; align-items:center; gap:9px; padding:9px 13px;
  border-bottom:1px solid var(--line);}
.dot{width:11px; height:11px; border-radius:3px; flex:none; border:1px solid rgba(0,0,0,.12)}
.card-h b{color:var(--ink); font-size:15px}
.card-h .tel{margin-left:auto; font-family:var(--mono); font-size:11.5px; color:var(--muted)}
.sess{display:grid; grid-template-columns:62px minmax(0,1fr); gap:2px 10px;
  padding:9px 13px; align-items:start;}
.sess + .sess{border-top:1px dashed var(--line)}
.sess .lab{font-family:var(--mono); font-size:11.5px; color:var(--muted); padding-top:3px}
.sess .lab b{display:block; font-family:var(--sans); font-size:13.5px; color:var(--ink);
  font-weight:700}
.names{display:flex; flex-wrap:wrap; gap:5px}
.nm{background:var(--sunk); border:1px solid var(--line); border-radius:5px;
  padding:2px 9px; font-size:13.5px; color:var(--ink);}
.nm.head{font-weight:700}
.nm.head::after{content:"　院長"; font-size:10px; color:var(--muted); font-weight:400}
.empty{font-size:13px; color:var(--muted)}
.empty.bad{color:var(--alert); font-weight:700}
.closed{font-size:13px; color:var(--muted); opacity:.75}

/* week cards */
.wk{background:var(--surface); border:1px solid var(--line); border-radius:7px;
  overflow:hidden; margin-top:12px;}
.wk > table{width:100%; border-collapse:collapse; table-layout:fixed}
.wk th,.wk td{border:1px solid var(--line); padding:4px 5px; vertical-align:top;
  font-size:12px; line-height:1.45;}
.wk thead th{background:var(--sunk); font-family:var(--mono); font-size:10.5px;
  color:var(--muted); font-weight:500; text-align:center}
.wk thead th:first-child{width:46px}
.wk tbody th{background:var(--sunk); color:var(--ink); font-weight:700;
  font-family:var(--mono); font-size:11.5px; text-align:left; white-space:nowrap;
  font-variant-numeric:tabular-nums}
.wk tbody th.today{background:var(--petrol); color:var(--on-petrol)}
.wk td span{display:block; color:var(--ink)}
.wk td.off{background:var(--sunk); color:var(--muted); text-align:center; font-size:11px}
.wk td.hol{background:var(--alert-soft); color:var(--alert); text-align:center;
  font-size:11px; font-weight:700}
.wk td.gap{background:var(--alert-soft); text-align:center; color:var(--alert);
  font-weight:700}

/* month grid */
.scroll{overflow-x:auto; -webkit-overflow-scrolling:touch; margin-top:14px;
  border:1px solid var(--line); border-radius:7px; background:var(--surface);}
table{border-collapse:collapse; width:100%; font-size:12.5px}
th,td{border:1px solid var(--line); padding:5px 7px; vertical-align:top;
  text-align:left;}
thead th{background:var(--sunk); font-family:var(--mono); font-size:11px;
  color:var(--muted); font-weight:500; white-space:nowrap; text-align:center}
thead th.today{background:var(--petrol); color:var(--on-petrol)}
tbody th{background:var(--sunk); color:var(--ink); white-space:nowrap; font-weight:700;
  font-size:12.5px;}
td{min-width:96px}
td.off{background:var(--sunk); color:var(--muted); text-align:center; font-size:11.5px}
td.hol{background:var(--alert-soft); color:var(--alert); text-align:center;
  font-size:11.5px; font-weight:700}
td.gap{background:var(--alert-soft)}
td .who{display:flex; flex-direction:column; gap:1px}
td .who span{color:var(--ink); line-height:1.45}
.ssrow th{font-weight:500; color:var(--muted); font-family:var(--mono); font-size:11px}

/* doctor */
.pick{margin-top:14px}
.pick label{display:block; font-size:12.5px; color:var(--muted); margin-bottom:5px;
  font-family:var(--mono); letter-spacing:.06em; text-transform:uppercase}
select{width:100%; max-width:340px; font-family:inherit; font-size:15px;
  padding:9px 11px; border-radius:7px; border:1px solid var(--line-strong);
  background:var(--surface); color:var(--ink);}
.docmeta{display:flex; flex-wrap:wrap; gap:6px 10px; margin-top:12px; align-items:center}
.docmeta .t{font-size:13px; color:var(--muted)}
.tally{display:flex; flex-wrap:wrap; gap:7px; margin-top:12px}
.tl{border:1px solid var(--line); border-radius:6px; padding:6px 11px;
  background:var(--surface); display:flex; align-items:baseline; gap:7px;}
.tl span{font-size:12px; color:var(--muted)}
.tl b{font-size:17px; color:var(--ink); font-variant-numeric:tabular-nums}
.docdays{display:grid; gap:7px; margin-top:14px;}
@media(min-width:560px){.docdays{grid-template-columns:1fr 1fr;}}
@media(min-width:860px){.docdays{grid-template-columns:1fr 1fr 1fr;}}
.dd{background:var(--surface); border:1px solid var(--line); border-radius:6px;
  padding:8px 11px; display:flex; align-items:center; gap:10px;}
.dd .dnum{font-family:var(--mono); font-size:13px; color:var(--muted); width:54px;
  flex:none; font-variant-numeric:tabular-nums}
.dd .ss{display:flex; gap:4px; flex-wrap:wrap}
.pillc{font-size:12px; padding:2px 8px; border-radius:4px; font-weight:700;
  border:1px solid rgba(0,0,0,.1)}
.note{font-size:12.5px; color:var(--muted); margin-top:14px; max-width:62ch}
footer{margin-top:34px; padding-top:14px; border-top:1px solid var(--line);
  font-size:12px; color:var(--muted); max-width:66ch;}
</style>

<div class="wrap">
<header>
  <div class="brand">
    <h1>日三牙醫體系 班表</h1>
    <span class="period" id="period"></span>
    <span class="built" id="built"></span>
  </div>
  <div class="banner" id="banner" hidden></div>
</header>

<nav class="tabs" role="tablist">
  <button role="tab" id="tab-today" aria-selected="true">今天</button>
  <button role="tab" id="tab-week" aria-selected="false">本週</button>
  <button role="tab" id="tab-month" aria-selected="false">整月</button>
  <button role="tab" id="tab-doc" aria-selected="false">找醫師</button>
</nav>

<div class="filter" id="filter"></div>

<main>
  <section id="v-today">
    <div class="daynav">
      <button id="prev" aria-label="前一天">‹</button>
      <div class="dlabel" id="dlabel"></div>
      <button id="next" aria-label="後一天">›</button>
    </div>
    <div class="cards" id="todayCards"></div>
  </section>

  <section id="v-week" hidden>
    <div id="weekWrap"></div>
    <p class="note">灰底是該院所該時段不開診。粉紅底代表有開診卻沒有醫師,要調班。</p>
  </section>

  <section id="v-month" hidden>
    <div id="monthWrap"></div>
    <p class="note">整月一次看完,橫向可滑動。粉紅底 = 有開診但沒醫師。</p>
  </section>

  <section id="v-doc" hidden>
    <div class="pick">
      <label for="docsel">選一位醫師</label>
      <select id="docsel"></select>
    </div>
    <div class="docmeta" id="docmeta"></div>
    <div class="tally" id="doctally"></div>
    <div class="docdays" id="docdays"></div>
  </section>
</main>

<footer id="foot"></footer>
</div>

<script id="data" type="application/json">__DATA__</script>
<script>
(function(){
  var D = JSON.parse(document.getElementById("data").textContent);
  var CL = {}; D.clinics.forEach(function(c){ CL[c.code] = c; });
  var HEADS = {};
  D.doctors.forEach(function(d){ if(d.duty){ HEADS[d.name] = d.duty; } });

  var state = { view:"today", clinic:"all", day:0, doc:D.doctors[0].name };

  function pad(n){ return (n<10?"0":"") + n; }
  function dstr(i){ return D.month + "/" + D.days[i].d; }
  function isOpen(c, i, s){ return CL[c].open[D.days[i].wd - 1][s] === 1; }

  /* ---- 期間與今天 ---- */
  document.getElementById("period").textContent = D.year + " 年 " + D.month + " 月";
  document.getElementById("built").textContent = "更新 " + D.built;
  var now = new Date();
  var todayIdx = -1;
  if (now.getFullYear() === D.year && (now.getMonth()+1) === D.month) {
    todayIdx = now.getDate() - 1;
    if (todayIdx >= D.ndays) todayIdx = -1;
  }
  state.day = todayIdx >= 0 ? todayIdx : 0;
  if (todayIdx < 0) {
    var b = document.getElementById("banner");
    b.hidden = false;
    b.textContent = "這份看板是 " + D.year + " 年 " + D.month + " 月的班表。今天是 "
      + (now.getMonth()+1) + "/" + now.getDate() + "，不在這個月份裡，所以從 "
      + D.month + "/1 開始顯示。";
  }

  /* ---- 院所篩選 ---- */
  var fil = document.getElementById("filter");
  function mkFilter(){
    fil.innerHTML = "";
    var opts = [{code:"all", short:"全部"}].concat(D.clinics);
    opts.forEach(function(c){
      var b = document.createElement("button");
      b.textContent = c.short;
      b.setAttribute("aria-pressed", state.clinic === c.code ? "true" : "false");
      if (c.code !== "all") {
        b.dataset.c = c.code;
        if (state.clinic === c.code) {
          b.style.background = "var(--c-" + c.code + ")";
          b.style.color = "var(--i-" + c.code + ")";
          b.style.borderColor = "var(--i-" + c.code + ")";
        }
      }
      b.addEventListener("click", function(){ state.clinic = c.code; mkFilter(); render(); });
      fil.appendChild(b);
    });
  }
  function shownClinics(){
    return state.clinic === "all" ? D.clinics : [CL[state.clinic]];
  }

  /* ---- 今天 ---- */
  function renderToday(){
    var i = state.day, dy = D.days[i];
    var lab = document.getElementById("dlabel");
    lab.innerHTML = "";
    var bb = document.createElement("b");
    bb.textContent = D.month + " 月 " + dy.d + " 日";
    var sp = document.createElement("span");
    sp.textContent = "星期" + dy.w + (i === todayIdx ? "　· 今天" : "");
    lab.appendChild(bb); lab.appendChild(sp);
    if (dy.hol) {
      var h = document.createElement("span");
      h.className = "hol"; h.textContent = dy.hol + "　全院休診";
      lab.appendChild(document.createElement("br")); lab.appendChild(h);
    }
    document.getElementById("prev").disabled = (i === 0);
    document.getElementById("next").disabled = (i === D.ndays - 1);

    var box = document.getElementById("todayCards");
    box.innerHTML = "";
    shownClinics().forEach(function(c){
      var card = document.createElement("div"); card.className = "card";
      var h = document.createElement("div"); h.className = "card-h";
      var dot = document.createElement("i"); dot.className = "dot";
      dot.style.background = "var(--c-" + c.code + ")";
      var nm = document.createElement("b"); nm.textContent = c.short;
      var tel = document.createElement("span"); tel.className = "tel"; tel.textContent = c.tel;
      h.appendChild(dot); h.appendChild(nm); h.appendChild(tel);
      card.appendChild(h);
      for (var s = 0; s < 3; s++) {
        var row = document.createElement("div"); row.className = "sess";
        var l = document.createElement("div"); l.className = "lab";
        var lb = document.createElement("b"); lb.textContent = D.sessions[s] + "診";
        l.appendChild(lb);
        l.appendChild(document.createTextNode(c.time[s][0] + "–" + c.time[s][1]));
        row.appendChild(l);
        var who = document.createElement("div"); who.className = "names";
        if (dy.hol) {
          who.innerHTML = '<span class="closed">國定假日休診</span>';
        } else if (!isOpen(c.code, i, s)) {
          who.innerHTML = '<span class="closed">不開診</span>';
        } else {
          var list = D.grid[c.code][i][s];
          if (!list.length) {
            who.innerHTML = '<span class="empty bad">沒有醫師　要調班</span>';
          } else {
            list.forEach(function(n){
              var t = document.createElement("span");
              t.className = "nm" + (HEADS[n] ? " head" : "");
              t.textContent = n;
              who.appendChild(t);
            });
          }
        }
        row.appendChild(who);
        card.appendChild(row);
      }
      box.appendChild(card);
    });
  }

  /* ---- 表格(本週 / 整月共用) ---- */
  function gridTable(from, to){
    var wrap = document.createElement("div"); wrap.className = "scroll";
    var t = document.createElement("table");
    var thead = document.createElement("thead");
    var r1 = document.createElement("tr");
    r1.appendChild(th("院所 / 診次", true));
    for (var i = from; i <= to; i++) {
      var h = th(D.days[i].d + " (" + D.days[i].w + ")");
      h.colSpan = 3;
      if (i === todayIdx) h.className = "today";
      r1.appendChild(h);
    }
    thead.appendChild(r1);
    var r2 = document.createElement("tr"); r2.className = "ssrow";
    r2.appendChild(th(""));
    for (var i2 = from; i2 <= to; i2++) {
      for (var s = 0; s < 3; s++) r2.appendChild(th(D.sessions[s]));
    }
    thead.appendChild(r2);
    t.appendChild(thead);

    var tb = document.createElement("tbody");
    shownClinics().forEach(function(c){
      var tr = document.createElement("tr");
      var rh = document.createElement("th"); rh.textContent = c.short;
      rh.style.background = "var(--c-" + c.code + ")";
      rh.style.color = "var(--i-" + c.code + ")";
      tr.appendChild(rh);
      for (var i3 = from; i3 <= to; i3++) {
        for (var s2 = 0; s2 < 3; s2++) {
          var td = document.createElement("td");
          if (D.days[i3].hol) { td.className = "hol"; td.textContent = "假"; }
          else if (!isOpen(c.code, i3, s2)) { td.className = "off"; td.textContent = "休"; }
          else {
            var list = D.grid[c.code][i3][s2];
            if (!list.length) { td.className = "gap"; td.textContent = "—"; }
            else {
              var w = document.createElement("div"); w.className = "who";
              list.forEach(function(n){
                var sp = document.createElement("span"); sp.textContent = n; w.appendChild(sp);
              });
              td.appendChild(w);
            }
          }
          tr.appendChild(td);
        }
      }
      tb.appendChild(tr);
    });
    t.appendChild(tb);
    wrap.appendChild(t);
    return wrap;
  }
  function th(txt, left){
    var e = document.createElement("th");
    e.textContent = txt;
    if (left) e.style.textAlign = "left";
    return e;
  }

  function weekCard(c, from, to){
    var wrap = document.createElement("div"); wrap.className = "wk";
    var h = document.createElement("div"); h.className = "card-h";
    var dot = document.createElement("i"); dot.className = "dot";
    dot.style.background = "var(--c-" + c.code + ")";
    var nm = document.createElement("b"); nm.textContent = c.short;
    h.appendChild(dot); h.appendChild(nm);
    wrap.appendChild(h);

    var t = document.createElement("table");
    var thead = document.createElement("thead"); var hr = document.createElement("tr");
    hr.appendChild(th(""));
    for (var s = 0; s < 3; s++) hr.appendChild(th(D.sessions[s] + "診"));
    thead.appendChild(hr); t.appendChild(thead);

    var tb = document.createElement("tbody");
    for (var i = from; i <= to; i++) {
      var tr = document.createElement("tr");
      var rh = document.createElement("th");
      rh.textContent = D.days[i].d + " " + D.days[i].w;
      if (i === todayIdx) rh.className = "today";
      tr.appendChild(rh);
      for (var s2 = 0; s2 < 3; s2++) {
        var td = document.createElement("td");
        if (D.days[i].hol) { td.className = "hol"; td.textContent = "假"; }
        else if (!isOpen(c.code, i, s2)) { td.className = "off"; td.textContent = "休"; }
        else {
          var list = D.grid[c.code][i][s2];
          if (!list.length) { td.className = "gap"; td.textContent = "—"; }
          else list.forEach(function(n){
            var sp = document.createElement("span"); sp.textContent = n; td.appendChild(sp);
          });
        }
        tr.appendChild(td);
      }
      tb.appendChild(tr);
    }
    t.appendChild(tb); wrap.appendChild(t);
    return wrap;
  }

  function renderWeek(){
    // 日期直排、診次橫排:七天一間院所剛好放得進手機,不必橫向滑
    var i = state.day;
    var from = Math.max(0, i - (D.days[i].wd - 1));
    var to = Math.min(D.ndays - 1, from + 6);
    var box = document.getElementById("weekWrap");
    box.innerHTML = "";
    var cap = document.createElement("p");
    cap.className = "note"; cap.style.marginTop = "12px";
    cap.textContent = D.month + "/" + D.days[from].d + " – " + D.month + "/" + D.days[to].d;
    box.appendChild(cap);
    shownClinics().forEach(function(c){ box.appendChild(weekCard(c, from, to)); });
  }
  function renderMonth(){
    var box = document.getElementById("monthWrap");
    box.innerHTML = "";
    box.appendChild(gridTable(0, D.ndays - 1));
  }

  /* ---- 找醫師 ---- */
  var sel = document.getElementById("docsel");
  D.doctors.forEach(function(d){
    var o = document.createElement("option");
    o.value = d.name;
    o.textContent = d.name + (d.spec ? "　" + d.spec : "");
    sel.appendChild(o);
  });
  sel.addEventListener("change", function(){ state.doc = sel.value; renderDoc(); });

  function renderDoc(){
    var d = D.doctors.filter(function(x){ return x.name === state.doc; })[0];
    var meta = document.getElementById("docmeta");
    meta.innerHTML = "";
    if (d.duty) meta.appendChild(tag(d.duty, true));
    d.teams.forEach(function(tn){
      var code = D.clinics.filter(function(c){ return c.short === tn; })[0].code;
      meta.appendChild(tag(tn, false, code));
    });
    if (d.spec) {
      var s = document.createElement("span"); s.className = "t"; s.textContent = d.spec;
      meta.appendChild(s);
    }

    var m = D.docm[d.name], per = {}, total = 0, days = 0;
    for (var i = 0; i < D.ndays; i++) {
      var any = false;
      for (var s2 = 0; s2 < 3; s2++) {
        var v = m[i][s2];
        if (CL[v]) { per[v] = (per[v] || 0) + 1; total++; any = true; }
      }
      if (any) days++;
    }
    var tal = document.getElementById("doctally");
    tal.innerHTML = "";
    tal.appendChild(tile("總診次", total));
    tal.appendChild(tile("上班天數", days));
    D.clinics.forEach(function(c){
      if (per[c.code]) tal.appendChild(tile(c.short, per[c.code], c.code));
    });

    var box = document.getElementById("docdays");
    box.innerHTML = "";
    for (var i2 = 0; i2 < D.ndays; i2++) {
      var row = [];
      for (var s3 = 0; s3 < 3; s3++) if (CL[m[i2][s3]]) row.push([s3, m[i2][s3]]);
      if (!row.length) continue;
      var el = document.createElement("div"); el.className = "dd";
      var dn = document.createElement("span"); dn.className = "dnum";
      dn.textContent = D.days[i2].d + " (" + D.days[i2].w + ")";
      var ss = document.createElement("div"); ss.className = "ss";
      row.forEach(function(p){
        var b = document.createElement("span");
        b.className = "pillc";
        b.style.background = "var(--c-" + p[1] + ")";
        b.style.color = "var(--i-" + p[1] + ")";
        b.textContent = D.sessions[p[0]] + "　" + CL[p[1]].short;
        ss.appendChild(b);
      });
      el.appendChild(dn); el.appendChild(ss);
      box.appendChild(el);
    }
    if (!box.children.length) {
      box.innerHTML = '<p class="note">這個月沒有排診。</p>';
    }
  }
  function tag(txt, strong, code){
    var e = document.createElement("span");
    e.className = "pillc";
    if (code) { e.style.background = "var(--c-" + code + ")"; e.style.color = "var(--i-" + code + ")"; }
    else { e.style.background = "var(--petrol-soft)"; e.style.color = "var(--petrol)"; }
    e.textContent = txt;
    return e;
  }
  function tile(lab, val, code){
    var e = document.createElement("div"); e.className = "tl";
    var s = document.createElement("span"); s.textContent = lab;
    var b = document.createElement("b"); b.textContent = val;
    if (code) { e.style.background = "var(--c-" + code + ")"; b.style.color = "var(--i-" + code + ")"; }
    e.appendChild(s); e.appendChild(b);
    return e;
  }

  /* ---- 切換 ---- */
  var TABS = ["today","week","month","doc"];
  TABS.forEach(function(v){
    document.getElementById("tab-" + v).addEventListener("click", function(){
      state.view = v; render();
    });
  });
  document.getElementById("prev").addEventListener("click", function(){
    if (state.day > 0) { state.day--; render(); }
  });
  document.getElementById("next").addEventListener("click", function(){
    if (state.day < D.ndays - 1) { state.day++; render(); }
  });

  function render(){
    TABS.forEach(function(v){
      document.getElementById("tab-" + v)
        .setAttribute("aria-selected", v === state.view ? "true" : "false");
      document.getElementById("v-" + v).hidden = (v !== state.view);
    });
    fil.hidden = (state.view === "doc");
    if (state.view === "today") renderToday();
    else if (state.view === "week") renderWeek();
    else if (state.view === "month") renderMonth();
    else renderDoc();
    try { localStorage.setItem("gemray-board", JSON.stringify(
      {clinic: state.clinic, view: state.view})); } catch(e) {}
  }

  try {
    var saved = JSON.parse(localStorage.getItem("gemray-board") || "{}");
    if (saved.clinic && (saved.clinic === "all" || CL[saved.clinic])) state.clinic = saved.clinic;
    if (TABS.indexOf(saved.view) >= 0) state.view = saved.view;
  } catch(e) {}

  document.getElementById("foot").textContent =
    "資料來源:gemray.tw 五間院所門診表,經體系班表系統展開為本月。本看板唯讀,"
    + "要調整班表請改 Excel 主檔後重新產生。國定假日為草稿,以人事行政總處公告為準。";

  mkFilter();
  renderDoc();
  render();
})();
</script>
"""

body = PAGE.replace("__DATA__",
                    json.dumps(DATA, ensure_ascii=False, separators=(",", ":")))

# 兩個輸出,因為有兩個去處:
#   班表看板.html        —— 完整 HTML 文件,給 Cloudflare Pages / 官網直接託管
#   班表看板_預覽.html   —— 只有內容,給 Artifact 預覽用(發布時外框會自動補上)
SHELL = ("""<!doctype html>
<html lang="zh-Hant">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover">
<meta name="robots" content="noindex,nofollow">
<style>:root{color-scheme:light dark;padding-top:env(safe-area-inset-top,0px);
padding-bottom:env(safe-area-inset-bottom,0px)}body{margin:0}img{max-width:100%}</style>
</head>
<body>
""" + body + """
</body>
</html>
""")
UPDIR.mkdir(exist_ok=True)
with io.open(OUT, "w", encoding="utf-8") as f:
    f.write(SHELL)
with io.open(PREVIEW, "w", encoding="utf-8") as f:
    f.write(body)
page = SHELL

_docs_with_shift = sum(1 for d in DOCM.values()
                       if any(v in GRID for row in d for v in row))
_gaps = sum(1 for c in GRID for i in range(NDAYS) for s in range(3)
            if not DAYS[i]["hol"] and OPENFLAG[c][DAYS[i]["wd"] - 1][s]
            and not GRID[c][i][s])
print(f"saved: {OUT}  ({len(page)/1024:.0f} KB)")
print(f"saved: {PREVIEW}")
print(f"{YEAR}/{MONTH} · {NDAYS} 天 · 醫師 {len(DOCTORS)} 位(本月有排診 {_docs_with_shift} 位)")
print(f"有開診但沒醫師的診次:{_gaps} 個")
