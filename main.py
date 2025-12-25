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
# ★オフセット（mm）
# =========================
GLOBAL_OFFSET_X_MM = 0
GLOBAL_OFFSET_Y_MM = 0

POSTAL_OFFSET_X_MM = 50
POSTAL_OFFSET_Y_MM = 0

ADDR_OFFSET_X_MM = 70
ADDR_OFFSET_Y_MM = 0

NAME_OFFSET_X_MM = 0
NAME_OFFSET_Y_MM = 0

# =========================
# ★文字間（縦送り）調整（mm）
# =========================
ADDR_LEADING_MM = 6
NAME_LEADING_MM = 10

# =========================
# ベース配置（mm）※PDF左上 기준
# =========================
BASE_POSTAL_X_MM = 12
BASE_POSTAL_Y_MM = 18

# 住所ブロック
BASE_ADDR_X_MM = 18
BASE_ADDR_Y_TOP_MM = 40
ADDR_COLUMN_GAP_MM = 12  # ★住所を2列にする時の列間（調整ポイント）

# 氏名ブロック
BASE_NAME_X_MM = 55
BASE_NAME_Y_TOP_MM = 40
NAME_COLUMN_GAP_MM = 12  # ★氏名の列間（調整ポイント）


# =========================
# ユーティリティ
# =========================
def format_postal(code: str) -> str:
    s = "".join(filter(str.isdigit, str(code)))
    return f"{s[:3]}-{s[3:]}" if len(s) == 7 else str(code).strip()


def split_address_for_sample_style(address: str) -> str:
    # ハイフン類を「ー」に寄せる（見た目安定）
    return str(address).replace("-", "ー").replace("−", "ー").replace("―", "ー")


def split_two_blocks_by_space(s: str):
    """
    半角/全角スペースがあれば最初の1つで2分割して返す。
    なければ (s, "") を返す。
    """
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
    """
    縦書き（1文字ずつ）
    各文字を縦列の中心に中央揃え（数字・記号の左寄り対策）
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
            address_raw = split_address_for_sample_style(row.get("住所", ""))

            sei = str(row.get("氏", "")).strip()
            mei = str(row.get("名", "")).strip()

            if not (postal or address_raw or sei or mei):
                continue

            gx = GLOBAL_OFFSET_X_MM
            gy = GLOBAL_OFFSET_Y_MM

            # 郵便番号（横書き）
            c.setFont(FONT_NAME, 20)
            px = BASE_POSTAL_X_MM + gx + POSTAL_OFFSET_X_MM
            py = BASE_POSTAL_Y_MM + gy + POSTAL_OFFSET_Y_MM
            c.drawString(px * mm, y_from_top_mm(py), f"〒{postal}")

            # ===== 住所（スペースで2分割→2列表示）=====
            addr_font_size = 14
            c.setFont(FONT_NAME, addr_font_size)

            addr_a, addr_b = split_two_blocks_by_space(address_raw)

            # 住所ブロックの基準
            addr_x_base = BASE_ADDR_X_MM + gx + ADDR_OFFSET_X_MM
            addr_y = BASE_ADDR_Y_TOP_MM + gy + ADDR_OFFSET_Y_MM

            if addr_b:
                # 右列：前半（都道府県など）
                draw_vertical_text_from_top(
                    c,
                    addr_x_base + ADDR_COLUMN_GAP_MM,  # 右
                    addr_y,
                    addr_a,
                    ADDR_LEADING_MM,
                    FONT_NAME,
                    addr_font_size,
                    True
                )
                # 左列：後半（番地など）
                draw_vertical_text_from_top(
                    c,
                    addr_x_base,  # 左
                    addr_y,
                    addr_b,
                    ADDR_LEADING_MM,
                    FONT_NAME,
                    addr_font_size,
                    True
                )
            else:
                # スペースが無ければ1列で表示
                draw_vertical_text_from_top(
                    c,
                    addr_x_base,
                    addr_y,
                    addr_a,
                    ADDR_LEADING_MM,
                    FONT_NAME,
                    addr_font_size,
                    True
                )

            # ===== 氏名（2列表示・順序修正：右=氏 / 左=名+様）=====
            name_font_size = 20
            c.setFont(FONT_NAME, name_font_size)

            name_x_base = BASE_NAME_X_MM + gx + NAME_OFFSET_X_MM
            name_y = BASE_NAME_Y_TOP_MM + gy + NAME_OFFSET_Y_MM

            # 右列：氏
            if sei:
                draw_vertical_text_from_top(
                    c,
                    name_x_base + NAME_COLUMN_GAP_MM,  # 右
                    name_y,
                    sei,
                    NAME_LEADING_MM,
                    FONT_NAME,
                    name_font_size,
                    True
                )

            # 左列：名 + 様
            if mei:
                draw_vertical_text_from_top(
                    c,
                    name_x_base,  # 左
                    name_y,
                    f"{mei}様",
                    NAME_LEADING_MM,
                    FONT_NAME,
                    name_font_size,
                    True
                )
            else:
                # 名が無い場合でも「様」だけは出したいならここ（不要なら削除OK）
                draw_vertical_text_from_top(
                    c,
                    name_x_base,
                    name_y,
                    "様",
                    NAME_LEADING_MM,
                    FONT_NAME,
                    name_font_size,
                    True
                )

            c.showPage()
            page_count += 1

c.save()
print(f"完了: {merged_pdf_path}（{page_count}ページ）")
