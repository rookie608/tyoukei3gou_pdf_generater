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

NAME_OFFSET_X_MM = 0
NAME_OFFSET_Y_MM = 0

# =========================
# ★文字間（縦送り）調整（mm）
# =========================
POSTAL_LEADING_MM = 12  # 郵便番号を縦書きにしたい場合に使用（現状は横書き）
ADDR_LEADING_MM = 6     # 住所
NAME_LEADING_MM = 12    # 氏名

# =========================
# ベース配置（mm）※PDF左上 기준
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


def split_address_for_sample_style(address: str) -> str:
    # ハイフン類を「ー」に寄せる（見た目安定）
    return str(address).replace("-", "ー").replace("−", "ー").replace("―", "ー")


def y_from_top_mm(y_mm_from_top: float) -> float:
    """左上 기준のY(mm) → reportlab座標Y(pt)へ変換（上からの距離を指定できる）"""
    return PAGE_H - (y_mm_from_top * mm)


def draw_vertical_text_from_top(
    c: canvas.Canvas,
    x_mm: float,
    y_top_mm: float,
    text: str,
    leading_mm: float,
    font_name: str,
    font_size: float,
    center_each_char: bool = True
):
    """
    1文字ずつ下に描く簡易縦書き（各文字を縦列の中心に中央揃え可能）
    入力はすべて「左上 기준（上から何mm）」で指定
    """
    x_center = x_mm * mm
    y = y_from_top_mm(y_top_mm)
    leading = leading_mm * mm

    for ch in str(text):
        if ch == "\n":
            y -= leading
            continue
        if ch == " ":
            y -= leading * 0.6
            continue

        if center_each_char:
            w = pdfmetrics.stringWidth(ch, font_name, font_size)
            x_draw = x_center - (w / 2.0)
        else:
            x_draw = x_center

        c.drawString(x_draw, y, ch)
        y -= leading


# =========================
# 1つのPDFにまとめて出力
# =========================
merged_pdf_path = OUTPUT_DIR / "宛名まとめ.pdf"
c = canvas.Canvas(str(merged_pdf_path), pagesize=(PAGE_W, PAGE_H))

page_count = 0

for csv_path in INPUT_DIR.glob("*.csv"):
    print(f"処理中: {csv_path.name}")

    with csv_path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            postal = format_postal(row.get("郵便番号", ""))
            address = split_address_for_sample_style(row.get("住所", ""))
            name = str(row.get("氏名", "")).strip()

            if not (postal or address or name):
                continue

            # 全体オフセット（左上 기준）
            gx = GLOBAL_OFFSET_X_MM
            gy = GLOBAL_OFFSET_Y_MM

            # 郵便番号（横書き）
            c.setFont(FONT_NAME, 16)
            px = BASE_POSTAL_X_MM + gx + POSTAL_OFFSET_X_MM
            py = BASE_POSTAL_Y_MM + gy + POSTAL_OFFSET_Y_MM
            c.drawString(px * mm, y_from_top_mm(py), f"〒{postal}")

            # 住所（縦書き）
            addr_font_size = 14
            c.setFont(FONT_NAME, addr_font_size)
            ax = BASE_ADDR_X_MM + gx + ADDR_OFFSET_X_MM
            ay = BASE_ADDR_Y_TOP_MM + gy + ADDR_OFFSET_Y_MM
            draw_vertical_text_from_top(
                c, ax, ay, address, ADDR_LEADING_MM,
                font_name=FONT_NAME, font_size=addr_font_size,
                center_each_char=True
            )

            # 氏名（縦書き）
            name_font_size = 20
            c.setFont(FONT_NAME, name_font_size)
            nx = BASE_NAME_X_MM + gx + NAME_OFFSET_X_MM
            ny = BASE_NAME_Y_TOP_MM + gy + NAME_OFFSET_Y_MM
            draw_vertical_text_from_top(
                c, nx, ny, f"{name}様", NAME_LEADING_MM,
                font_name=FONT_NAME, font_size=name_font_size,
                center_each_char=True
            )

            # 次ページへ（1宛名=1ページ）
            c.showPage()
            page_count += 1

# 1回だけ保存
c.save()

print(f"完了: {merged_pdf_path}（{page_count}ページ）")
