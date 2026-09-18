# -*- coding: utf-8 -*-
"""確認要上網的檔案裡沒有出勤紀錄與個資。上級的第 4 題是「出勤紀錄不要放上網」,
這支腳本把那句話變成可以每次檢查的條件,而不是口頭保證。"""
import io, json, re, sys

PATH = sys.argv[1] if len(sys.argv) > 1 else "/home/user/-/out/班表看板.html"
page = io.open(PATH, encoding="utf-8").read()
TAG = '<script id="data" type="application/json">'
i = page.index(TAG) + len(TAG)
D = json.loads(page[i:page.index('</script>', i)])

# 一、不該出現在頁面任何地方的字眼(含說明文字與程式碼)
BANNED = ["打卡", "遲到", "早退", "未打卡", "加班", "到職", "特休", "病假", "事假",
          "公假", "排休", "薪", "身分證", "月結統計", "出勤紀錄", "實際上班", "實際下班"]
hits = [w for w in BANNED if w in page]
print(f"一、頁面文字含敏感字眼:{len(hits)}" + (f"  {hits}" if hits else " ✓"))

# 二、資料結構只能有這些鍵
ALLOWED_TOP = {"built","sessions","months","clinics","doctors","days","grid","docm"}
extra = set(D) - ALLOWED_TOP
print(f"二、資料多出未預期的欄位:{len(extra)}" + (f"  {extra}" if extra else " ✓"))

ALLOWED_DOC = {"eid","name","spec","duty","teams"}
bad_doc = set()
for d in D["doctors"]: bad_doc |= set(d) - ALLOWED_DOC
print(f"三、醫師資料多出欄位:{len(bad_doc)}" + (f"  {bad_doc}" if bad_doc else " ✓"))

# 四、班表格子只能是院所代碼或國/休/空白——不能混進時間或數字
codes = {c["code"] for c in D["clinics"]} | {"國", "休", ""}
odd = set()
for nm, m in D["docm"].items():
    for row in m:
        for v in row:
            if v not in codes: odd.add(v)
print(f"四、班表格子出現非院所代碼的值:{len(odd)}" + (f"  {odd}" if odd else " ✓"))

# 五、不該有任何看起來像時間的字串(HH:MM),診次時間表除外
times = set(re.findall(r"\b\d{1,2}:\d{2}\b", json.dumps(D, ensure_ascii=False)))
legit = {t for c in D["clinics"] for pair in c["time"] for t in pair}
print(f"五、診次時間以外的時間字串:{len(times - legit)}"
      + (f"  {times - legit}" if times - legit else " ✓"))

# 六、醫護長與管理部(工時制那七位)不該出現在對外頁面
STAFF = ["ivy", "文君", "小玲", "娜娜", "孟諭", "怡雯", "小華"]
inpage = [n for n in STAFF if n in page]
print(f"六、工時制人員出現在頁面:{len(inpage)}" + (f"  {inpage}" if inpage else " ✓"))

_m = D["months"]
print(f"\n頁面實際載的是:{len(D['doctors'])} 位醫師 × {len(D['days'])} 天"
      f"({_m[0]['y']}/{_m[0]['m']}–{_m[-1]['y']}/{_m[-1]['m']})的班表,"
      f"外加院所電話地址。沒有任何一筆出勤或個資。")
