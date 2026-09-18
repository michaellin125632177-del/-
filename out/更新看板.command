#!/bin/bash
# 日三牙醫體系 · 班表看板月更新（macOS：在 Finder 裡按兩下即可執行）
# 第一次使用如果打不開，請先在這個檔案上按右鍵 →「打開」→ 再按一次「打開」。
cd "$(dirname "$0")" || exit 1

echo "==============================================="
echo "  日三牙醫體系 · 班表看板 月更新"
echo "==============================================="
echo

if ! command -v python3 >/dev/null 2>&1; then
  echo "找不到 python3。請先安裝 Python 3（https://www.python.org/downloads/），"
  echo "裝好後重新按兩下這個檔案。"
  echo; read -r -p "按 Enter 關閉…" _; exit 1
fi

python3 -c "import openpyxl" 2>/dev/null || {
  echo "正在安裝需要的套件 openpyxl…"
  python3 -m pip install --quiet openpyxl || {
    echo "安裝失敗。請把這個畫面截圖給負責人。"
    echo; read -r -p "按 Enter 關閉…" _; exit 1
  }
}

DEFAULT="$(date -v+1m +%Y-%m 2>/dev/null || date -d '+1 month' +%Y-%m)"
echo "要產生哪一個月的班表？"
read -r -p "格式 YYYY-MM，直接按 Enter 就用 $DEFAULT ： " PERIOD
PERIOD="${PERIOD:-$DEFAULT}"

case "$PERIOD" in
  [0-9][0-9][0-9][0-9]-[0-9][0-9]) ;;
  *) echo; echo "「$PERIOD」格式不對，應該像 2026-11 這樣。"
     echo; read -r -p "按 Enter 關閉…" _; exit 1 ;;
esac

echo
echo "產生 $PERIOD 的 Excel 與看板…"
echo
export GEMRAY_PERIOD="$PERIOD"
python3 build_roster.py               || { echo; echo "Excel 產生失敗，請把畫面截圖給負責人。"; read -r -p "按 Enter 關閉…" _; exit 1; }
python3 build_roster.py --no-assistant || { echo; echo "醫師版產生失敗。"; read -r -p "按 Enter 關閉…" _; exit 1; }
python3 build_site.py                 || { echo; echo "看板產生失敗，請把畫面截圖給負責人。"; read -r -p "按 Enter 關閉…" _; exit 1; }

echo
echo "==============================================="
echo "  完成。要上傳的是這個資料夾："
echo "  $(pwd)/上傳這個資料夾"
echo
echo "  下一步：到 Cloudflare Pages 的專案頁面，"
echo "  用「Create deployment」把整個資料夾拖進去。"
echo "  不用改檔名，裡面已經是 index.html 了。"
echo "==============================================="
echo
read -r -p "按 Enter 關閉…" _
