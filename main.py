import csv
from pathlib import Path
from reportlab.lib.pagesizes import mm
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.cidfonts import UnicodeCIDFont

# =========================
# 設定
# =========================
INPUT_DIR = Path("input")
OUTPUT_DIR = Path("pdf_output")
OUTPUT_DIR.mkdir(exist_ok=True)

# ★PDFの向き：短辺が上（幅120 × 高さ235）
PAGE_W = 120 * mm
PAGE_H = 235 * mm

FONT_NAME = "HeiseiKakuGo-W5"
pdfmetrics.registerFont(UnicodeCIDFont(FONT_NAME))

# =========================
# ★オフセット（mm）★
# 基準：PDFの左上（短辺の左上）
# X：右が＋ / 左が-
# Y：下が＋ / 上が-
# =========================
GLOBAL_OFFSET_X_MM = 0
GLOBAL_OFFSET_Y_MM = 0

POSTAL_OFFSET_X_MM = 60
POSTAL_OFFSET_Y_MM = 0

ADDR_OFFSET_X_MM = 80
ADDR_OFFSET_Y_MM = 0

NAME_OFFSET_X_MM = -5
NAME_OFFSET_Y_MM = 0

# =========================
# ★文字間（縦送り）調整（mm）
# =========================
POSTAL_LEADING_MM = 12  # 郵便番号を縦書きにしたい場合に使用（現状は横書き）
ADDR_LEADING_MM = 6    # 住所
NAME_LEADING_MM = 12    # 氏名

# =========================
# ベース配置（mm）※PDF左上 기준
# （必要に応じてここを調整）
# =========================
BASE_POSTAL_X_MM = 12
BASE_POSTAL_Y_MM = 18

BASE_ADDR_X_MM = 18
BASE_ADDR_Y_TOP_MM = 40

BASE_NAME_X_MM = 55
BASE_NAME_Y_TOP_MM = 40


def format_postal(code: str) -> str:
    s = "".join(filter(str.isdigit, str(code)))
    return f"{s[:3]}-{s[3:]}" if len(s) == 7 else str(code).strip()


def sanitize_filename(name: str) -> str:
    return "".join(c if c not in '\\/:*?"<>|' else "_" for c in str(name))


def split_address_for_sample_style(address: str) -> str:
    # ハイフン類を「ー」に寄せる（見た目安定）
    return str(address).replace("-", "ー").replace("−", "ー").replace("―", "ー")


def y_from_top_mm(y_mm_from_top: float) -> float:
    """左上 기준のY(mm) → reportlab座標Y(pt)へ変換（上からの距離を指定できる）"""
    return PAGE_H - (y_mm_from_top * mm)


def draw_vertical_text_from_top(c: canvas.Canvas, x_mm: float, y_top_mm: float, text: str, leading_mm: float):
    """
    1文字ずつ下に描く簡易縦書き
    入力はすべて「左上 기준（上から何mm）」で指定
    """
    x = x_mm * mm
    y = y_from_top_mm(y_top_mm)
    leading = leading_mm * mm

    for ch in str(text):
        if ch == "\n":
            y -= leading
            continue
        if ch == " ":
            y -= leading * 0.6
            continue
        c.drawString(x, y, ch)
        y -= leading


# =========================
# メイン
# =========================
for csv_path in INPUT_DIR.glob("*.csv"):
    print(f"処理中: {csv_path.name}")

    with csv_path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            postal = format_postal(row["郵便番号"])
            address = split_address_for_sample_style(row["住所"])
            name = str(row["氏名"]).strip()

            if not (postal or address or name):
                continue

            out = OUTPUT_DIR / f"宛名_{sanitize_filename(name or 'no_name')}.pdf"
            c = canvas.Canvas(str(out), pagesize=(PAGE_W, PAGE_H))

            # 全体オフセット（左上 기준）
            gx = GLOBAL_OFFSET_X_MM
            gy = GLOBAL_OFFSET_Y_MM

            # 郵便番号（横書き）
            c.setFont(FONT_NAME, 16)
            px = BASE_POSTAL_X_MM + gx + POSTAL_OFFSET_X_MM
            py = BASE_POSTAL_Y_MM + gy + POSTAL_OFFSET_Y_MM
            c.drawString(px * mm, y_from_top_mm(py), f"〒{postal}")

            # 住所（縦書き）
            c.setFont(FONT_NAME, 14)
            ax = BASE_ADDR_X_MM + gx + ADDR_OFFSET_X_MM
            ay = BASE_ADDR_Y_TOP_MM + gy + ADDR_OFFSET_Y_MM
            draw_vertical_text_from_top(c, ax, ay, address, ADDR_LEADING_MM)

            # 氏名（縦書き）
            c.setFont(FONT_NAME, 30)
            nx = BASE_NAME_X_MM + gx + NAME_OFFSET_X_MM
            ny = BASE_NAME_Y_TOP_MM + gy + NAME_OFFSET_Y_MM
            draw_vertical_text_from_top(c, nx, ny, f"{name}様", NAME_LEADING_MM)

            # （任意）郵便番号を縦書きにしたい場合は、上の横書きdrawStringを消して下を使う
            # c.setFont(FONT_NAME, 16)
            # draw_vertical_text_from_top(c, px, py, f"〒{postal}", POSTAL_LEADING_MM)

            c.showPage()
            c.save()

print("短辺が上のPDF配置（左上オフセット基準・文字間個別調整）でPDF生成が完了しました。")
