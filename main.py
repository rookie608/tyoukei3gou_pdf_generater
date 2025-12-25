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

PDF_PAGE_LIMIT = 20  # ★1ファイルあたりのページ数

# =========================
# ★オフセット（mm）
# =========================
GLOBAL_OFFSET_X_MM = 0
GLOBAL_OFFSET_Y_MM = 0

POSTAL_OFFSET_X_MM = 50
POSTAL_OFFSET_Y_MM = 0

# ※住所は「右上基準」に切り替える（ADDR_OFFSET_X_MM は右端からの距離として扱う）
ADDR_OFFSET_X_MM = 0
ADDR_OFFSET_Y_MM = 0

NAME_OFFSET_X_MM = 0
NAME_OFFSET_Y_MM = 0

# =========================
# ★文字間（縦送り）調整（mm）
# =========================
ADDR_LEADING_MM = 6
NAME_LEADING_MM = 10

# =========================
# ベース配置（mm）
# =========================
BASE_POSTAL_X_MM = 12
BASE_POSTAL_Y_MM = 18

# 住所（Yは左上基準のまま、Xだけ右上基準にする）
BASE_ADDR_X_MM = 18
BASE_ADDR_Y_TOP_MM = 40
ADDR_COLUMN_GAP_MM = 12

# 氏名
BASE_NAME_X_MM = 55
BASE_NAME_Y_TOP_MM = 40
NAME_COLUMN_GAP_MM = 12


# =========================
# ユーティリティ
# =========================
def format_postal(code: str) -> str:
    s = "".join(filter(str.isdigit, str(code)))
    return f"{s[:3]}-{s[3:]}" if len(s) == 7 else str(code).strip()


def split_address_for_sample_style(address: str) -> str:
    return str(address).replace("-", "ー").replace("−", "ー").replace("―", "ー")


def split_two_blocks_by_space(s: str):
    s = str(s).strip()
    if " " in s:
        a, b = s.split(" ", 1)
        return a.strip(), b.strip()
    if "　" in s:
        a, b = s.split("　", 1)
        return a.strip(), b.strip()
    return s, ""


def y_from_top_mm(y_mm_from_top: float) -> float:
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
    x_center = x_mm * mm
    y = y_from_top_mm(y_top_mm)
    leading = leading_mm * mm

    for ch in str(text):
        if ch == "\n":
            y -= leading
            continue

        if center_each_char:
            w = pdfmetrics.stringWidth(ch, font_name, font_size)
            x_draw = x_center - (w / 2)
        else:
            x_draw = x_center

        c.drawString(x_draw, y, ch)
        y -= leading


# =========================
# メイン（20枚ごとにPDF分割）
# =========================
pdf_index = 1
page_in_current_pdf = 0
c = None


def start_new_pdf(index: int):
    path = OUTPUT_DIR / f"宛名_{index:03d}.pdf"
    return canvas.Canvas(str(path), pagesize=(PAGE_W, PAGE_H)), path


c, current_pdf_path = start_new_pdf(pdf_index)
total_pages = 0

PAGE_W_MM = PAGE_W / mm  # 幅をmm単位で使うため

for csv_path in INPUT_DIR.glob("*.csv"):
    print(f"処理中: {csv_path.name}")

    with csv_path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            postal = format_postal(row.get("郵便番号", ""))
            address_raw = split_address_for_sample_style(row.get("住所", ""))
            sei = str(row.get("氏", "")).strip()
            mei = str(row.get("名", "")).strip()

            if not (postal or address_raw or sei or mei):
                continue

            # --- 新しいPDFが必要か？
            if page_in_current_pdf >= PDF_PAGE_LIMIT:
                c.save()
                pdf_index += 1
                page_in_current_pdf = 0
                c, current_pdf_path = start_new_pdf(pdf_index)

            gx = GLOBAL_OFFSET_X_MM
            gy = GLOBAL_OFFSET_Y_MM

            # 郵便番号
            c.setFont(FONT_NAME, 20)
            px = BASE_POSTAL_X_MM + gx + POSTAL_OFFSET_X_MM
            py = BASE_POSTAL_Y_MM + gy + POSTAL_OFFSET_Y_MM
            c.drawString(px * mm, y_from_top_mm(py), f"〒{postal}")

            # ===== 住所（2列対応 / Xは右上基準）=====
            addr_font_size = 14
            c.setFont(FONT_NAME, addr_font_size)

            addr_a, addr_b = split_two_blocks_by_space(address_raw)

            # ★右上基準：右端からの距離で x を決める
            # addr_x は「右列（前半）」の列中心 X（mm）
            addr_x = PAGE_W_MM - (BASE_ADDR_X_MM + ADDR_OFFSET_X_MM + gx)
            addr_y = BASE_ADDR_Y_TOP_MM + gy + ADDR_OFFSET_Y_MM

            if addr_b:
                # 右列：前半（都道府県など）
                draw_vertical_text_from_top(
                    c, addr_x, addr_y,
                    addr_a, ADDR_LEADING_MM, FONT_NAME, addr_font_size, True
                )
                # 左列：後半（番地など） ※右上基準なので「- GAP」で左へ
                draw_vertical_text_from_top(
                    c, addr_x - ADDR_COLUMN_GAP_MM, addr_y,
                    addr_b, ADDR_LEADING_MM, FONT_NAME, addr_font_size, True
                )
            else:
                draw_vertical_text_from_top(
                    c, addr_x, addr_y,
                    addr_a, ADDR_LEADING_MM, FONT_NAME, addr_font_size, True
                )

            # ===== 氏名（右=氏 / 左=名+様）=====
            name_font_size = 20
            c.setFont(FONT_NAME, name_font_size)

            name_x = BASE_NAME_X_MM + gx + NAME_OFFSET_X_MM
            name_y = BASE_NAME_Y_TOP_MM + gy + NAME_OFFSET_Y_MM

            if sei:
                draw_vertical_text_from_top(
                    c, name_x + NAME_COLUMN_GAP_MM, name_y,
                    sei, NAME_LEADING_MM, FONT_NAME, name_font_size, True
                )
            if mei:
                draw_vertical_text_from_top(
                    c, name_x, name_y,
                    f"{mei}様", NAME_LEADING_MM, FONT_NAME, name_font_size, True
                )

            c.showPage()
            page_in_current_pdf += 1
            total_pages += 1

# 最後のPDFを保存
if c:
    c.save()

print(f"完了：{total_pages}ページを {pdf_index} ファイルに分割して出力しました。")
