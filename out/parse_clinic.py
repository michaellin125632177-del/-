# -*- coding: utf-8 -*-
"""從官網門診表 HTML 解析每週固定門診(官網 span 巢狀不良,先切 <br> 再抽名字)"""
import re, io, html
FILES = {"悅":"7449aded-____________.html","睿":"3ec7fb99-____________.html",
         "匯":"980ddbb6-____________.html","曜":"596ed7c7-____________.html",
         "寶":"205f6a24-_____________.html"}
BASE = "/root/.claude/uploads/4c86cb05-c0c8-54f2-ad67-3f17161ad8f4/"
ICON = {"icon1-triangle":"植牙","icon2-clover":"牙周","icon1-circle":"矯正",
        "icon1-heart":"兒童","icon1-square":"口外","icon1-star":"根管",
        "icon2-round":"美學","icon1-cross":"兒童矯正"}
def parse(code):
    raw = io.open(BASE + FILES[code], encoding="utf-8").read()
    tbl  = re.search(r"(?is)<table.*?</table>", raw).group(0)
    body = re.search(r"(?is)<tbody>(.*?)</tbody>", tbl).group(1)
    out, times = {}, []
    for ss, row in enumerate(re.findall(r"(?is)<tr>(.*?)</tr>", body)):
        tds = re.findall(r"(?is)<td[^>]*>(.*?)</td>", row)
        times.append(re.sub(r"\s+","", html.unescape(re.sub(r"(?s)<[^>]+>","",tds[0]))))
        for wd, cell in enumerate(tds[1:], start=1):
            frags = re.split(r"(?i)<br\s*/?>", cell)
            entries = []
            for fr in frags:
                cls  = " ".join(re.findall(r'class="([^"]*)"', fr))
                text = re.sub(r"\s+","", html.unescape(re.sub(r"(?s)<[^>]+>","",fr)))
                if not text: continue
                entries.append({
                    "name":  re.sub(r"\(.*?\)|\*", "", text).strip(),
                    "alt":   "隔週" in text,
                    "star":  "*" in text,
                    "note":  next((n for n in re.findall(r"\((.*?)\)", text)
                                   if "隔週" not in n), ""),
                    "spec":  sorted(v for k, v in ICON.items() if k in cls)})
            out[(wd, ss)] = entries
    return out, times
ALL, TIMES = {}, {}
for c in FILES: ALL[c], TIMES[c] = parse(c)

