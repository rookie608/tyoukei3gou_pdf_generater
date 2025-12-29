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

ADDR_OFFSET_X_MM = 0
ADDR_OFFSET_Y_MM = 0

NAME_OFFSET_X_MM = 0
NAME_OFFSET_Y_MM = 0

# =========================
# ★文字間（縦送り）調整（mm）
# =========================
ADDR_LEADING_MM = 6
NAME_LEADING_MM = 8

# =========================
# ★名の段落下げ設定
# =========================
MEI_INDENT_STEPS = 6
MEI_INDENT_STEPS_NO_MEI = 13

# =========================
# ベース配置（mm）
# =========================
BASE_POSTAL_X_MM = 12
BASE_POSTAL_Y_MM = 18

BASE_ADDR_X_MM = 18
BASE_ADDR_Y_TOP_MM = 40
ADDR_COLUMN_GAP_MM = 12

BASE_NAME_X_MM = 55
BASE_NAME_Y_TOP_MM = 40
NAME_COLUMN_GAP_MM = 12

# =========================
# ★縦書き回転制御
# =========================
ROTATE_90_CHARS = set([
    "、", "。", "，", "．",
    "ー", "-", "−", "―", "—",
    "〜", "~",          # ← ★追加
    "（", "）", "(", ")",
    "「", "」", "『", "』",
])


# ★向きが逆なので180度追加回転
FLIP_180_CHARS = set([
    "（", "）", "(", ")",
    "「", "」", "『", "』",
])

# ★回転文字の微調整（font_size倍率）
# rotate(90)後は y を増やすと「左」に寄る
ROTATE_ADJUST = {
    "ー": {"dx": 0.0, "dy": 0.18},
    "―": {"dx": 0.0, "dy": 0.18},
    "—": {"dx": 0.0, "dy": 0.18},
    "-": {"dx": 0.0, "dy": 0.18},
    "−": {"dx": 0.0, "dy": 0.18},
    "〜": {"dx": 0.0, "dy": 0.18},
    "~": {"dx": 0.0, "dy": 0.18},
}

# =========================
# ★長音「ー」専用スタイル（サイズで区別）
# =========================
CHOONPU_SIZE_SCALE = 0.75   # 0.70〜0.85 推奨
CHOONPU_EXTRA_SHIFT = 0.10  # 左寄せ微調整（font_size倍率）

# =========================
# ユーティリティ
# =========================
def format_postal(code: str) -> str:
    s = "".join(filter(str.isdigit, str(code)))
    return f"{s[:3]}-{s[3:]}" if len(s) == 7 else str(code).strip()


def split_address_for_sample_style(address: str) -> str:
    return str(address)


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

    # ★中央揃え用パラメータ
    PIVOT_Y_FACTOR = 0.35
    DRAW_Y_FACTOR = 0.50

    for ch in str(text):
        if ch == "\n":
            y -= leading
            continue

        w = pdfmetrics.stringWidth(ch, font_name, font_size)

        if center_each_char:
            x_draw = x_center - (w / 2)
        else:
            x_draw = x_center

        if ch in ROTATE_90_CHARS:
            c.saveState()

            pivot_x = x_center
            pivot_y = y + (font_size * PIVOT_Y_FACTOR)
            c.translate(pivot_x, pivot_y)

            angle = 90 + (180 if ch in FLIP_180_CHARS else 0)
            c.rotate(angle)

            # ---- 文字ごとの描画サイズ調整 ----
            draw_font_size = font_size
            extra_dx = 0.0
            extra_dy = 0.0

            if ch == "ー":
                draw_font_size = font_size * CHOONPU_SIZE_SCALE
                extra_dy = font_size * CHOONPU_EXTRA_SHIFT

            c.setFont(font_name, draw_font_size)

            adj = ROTATE_ADJUST.get(ch, {"dx": 0.0, "dy": 0.0})
            dx = font_size * adj["dx"]
            dy = font_size * adj["dy"]

            c.drawString(
                (-w / 2) + dx + extra_dx,
                (-font_size * DRAW_Y_FACTOR) + dy + extra_dy,
                ch
            )

            c.restoreState()
        else:
            c.drawString(x_draw, y, ch)

        y -= leading


# =========================
# メイン処理
# =========================
pdf_index = 1
page_in_current_pdf = 0


def start_new_pdf(index: int):
    path = OUTPUT_DIR / f"宛名_{index:03d}.pdf"
    return canvas.Canvas(str(path), pagesize=(PAGE_W, PAGE_H))


c = start_new_pdf(pdf_index)
total_pages = 0
PAGE_W_MM = PAGE_W / mm

for csv_path in INPUT_DIR.glob("*.csv"):
    with csv_path.open(newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)

        for row in reader:
            postal = format_postal(row.get("郵便番号", ""))
            address_raw = split_address_for_sample_style(row.get("住所", ""))
            sei = str(row.get("氏", "")).strip()
            mei = str(row.get("名", "")).strip()

            if not (postal or address_raw or sei):
                continue

            if page_in_current_pdf >= PDF_PAGE_LIMIT:
                c.save()
                pdf_index += 1
                page_in_current_pdf = 0
                c = start_new_pdf(pdf_index)

            gx = GLOBAL_OFFSET_X_MM
            gy = GLOBAL_OFFSET_Y_MM

            # 郵便番号
            c.setFont(FONT_NAME, 20)
            c.drawString(
                (BASE_POSTAL_X_MM + gx + POSTAL_OFFSET_X_MM) * mm,
                y_from_top_mm(BASE_POSTAL_Y_MM + gy + POSTAL_OFFSET_Y_MM),
                f"〒{postal}"
            )

            # ===== 住所 =====
            addr_font_size = 14
            c.setFont(FONT_NAME, addr_font_size)

            addr_a, addr_b = split_two_blocks_by_space(address_raw)
            addr_x = PAGE_W_MM - (BASE_ADDR_X_MM + ADDR_OFFSET_X_MM + gx)
            addr_y = BASE_ADDR_Y_TOP_MM + gy + ADDR_OFFSET_Y_MM

            draw_vertical_text_from_top(
                c, addr_x, addr_y,
                addr_a, ADDR_LEADING_MM, FONT_NAME, addr_font_size, True
            )
            if addr_b:
                draw_vertical_text_from_top(
                    c, addr_x - ADDR_COLUMN_GAP_MM, addr_y,
                    addr_b, ADDR_LEADING_MM, FONT_NAME, addr_font_size, True
                )

            # ===== 氏名 =====
            name_font_size = 18
            c.setFont(FONT_NAME, name_font_size)

            name_x = BASE_NAME_X_MM + gx + NAME_OFFSET_X_MM
            name_y = BASE_NAME_Y_TOP_MM + gy + NAME_OFFSET_Y_MM

            if sei:
                draw_vertical_text_from_top(
                    c, name_x + NAME_COLUMN_GAP_MM, name_y,
                    sei, NAME_LEADING_MM, FONT_NAME, name_font_size, True
                )

            left_text = f"{mei} ご担当者　様".strip() if mei else "ご担当者　様"
            indent_steps = MEI_INDENT_STEPS if mei else MEI_INDENT_STEPS_NO_MEI
            mei_y = name_y + (NAME_LEADING_MM * indent_steps)

            draw_vertical_text_from_top(
                c, name_x, mei_y,
                left_text, NAME_LEADING_MM, FONT_NAME, name_font_size, True
            )

            c.showPage()
            page_in_current_pdf += 1
            total_pages += 1

c.save()
print(f"完了：{total_pages}ページを {pdf_index} ファイルに分割して出力しました。")
