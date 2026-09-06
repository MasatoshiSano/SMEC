#!/usr/bin/env python3
"""知的で編集的なSMEC標準図解テンプレートを生成する。"""

from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_CONNECTOR, MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml.ns import qn
from pptx.oxml.xmlchemy import OxmlElement
from pptx.util import Inches, Pt


ROOT = Path(__file__).resolve().parents[1]
OUTPUT = ROOT / "docs" / "textbook" / "slides" / "SMEC_visual_diagram_template.pptx"

SLIDE_W = 13.333
SLIDE_H = 7.5
FONT = "Noto Sans"

# Editorial Cobalt
PAPER = "F7F7F4"
WHITE = "FFFFFF"
INK = "17212B"
MUTED = "52606D"
LINE = "D5DBE1"
GRID = "E8EBEE"
COBALT = "245B9E"
COBALT_MID = "7EA2C9"
COBALT_PALE = "E9F0F7"
TEAL = "16776F"
TEAL_PALE = "E7F0ED"
RUST = "B24A3B"
RUST_PALE = "F5EAE7"
GOLD = "8A620F"


def rgb(value: str) -> RGBColor:
    return RGBColor.from_string(value)


def set_east_asian_font(run, typeface: str = FONT) -> None:
    """PowerPointの東アジア文字用フォントも明示する。"""
    r_pr = run._r.get_or_add_rPr()
    east_asian = r_pr.find(qn("a:ea"))
    if east_asian is None:
        east_asian = OxmlElement("a:ea")
        r_pr.append(east_asian)
    east_asian.set("typeface", typeface)


def text(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    value: str,
    *,
    size: float = 15,
    color: str = INK,
    bold: bool = False,
    align: PP_ALIGN = PP_ALIGN.LEFT,
    valign: MSO_ANCHOR = MSO_ANCHOR.TOP,
    margin: float = 0,
):
    shape = slide.shapes.add_textbox(Inches(x), Inches(y), Inches(w), Inches(h))
    frame = shape.text_frame
    frame.clear()
    frame.word_wrap = True
    frame.margin_left = Inches(margin)
    frame.margin_right = Inches(margin)
    frame.margin_top = Inches(margin)
    frame.margin_bottom = Inches(margin)
    frame.vertical_anchor = valign
    paragraph = frame.paragraphs[0]
    paragraph.alignment = align
    paragraph.space_after = Pt(0)
    run = paragraph.add_run()
    run.text = value
    run.font.name = FONT
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = rgb(color)
    set_east_asian_font(run)
    return shape


def rect(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    *,
    fill: str = WHITE,
    line_color: str | None = LINE,
    line_width: float = 0.6,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(x), Inches(y), Inches(w), Inches(h)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(line_width)
    return shape


def circle(
    slide,
    x: float,
    y: float,
    d: float,
    *,
    fill: str = COBALT,
    line_color: str | None = None,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(x), Inches(y), Inches(d), Inches(d)
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(fill)
    if line_color is None:
        shape.line.fill.background()
    else:
        shape.line.color.rgb = rgb(line_color)
        shape.line.width = Pt(0.7)
    return shape


def line(
    slide,
    x1: float,
    y1: float,
    x2: float,
    y2: float,
    *,
    color: str = LINE,
    width: float = 1,
):
    shape = slide.shapes.add_connector(
        MSO_CONNECTOR.STRAIGHT,
        Inches(x1),
        Inches(y1),
        Inches(x2),
        Inches(y2),
    )
    shape.line.color.rgb = rgb(color)
    shape.line.width = Pt(width)
    return shape


def triangle(
    slide,
    x: float,
    y: float,
    w: float = 0.13,
    h: float = 0.16,
    *,
    color: str = MUTED,
    rotation: float = 90,
):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x),
        Inches(y),
        Inches(w),
        Inches(h),
    )
    shape.rotation = rotation
    shape.fill.solid()
    shape.fill.fore_color.rgb = rgb(color)
    shape.line.fill.background()
    return shape


def base_slide(prs: Presentation):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = rgb(PAPER)
    return slide


def pattern_slide(
    prs: Presentation,
    code: str,
    category: str,
    title_value: str,
    use_case: str,
    rule: str,
    *,
    accent: str = COBALT,
):
    slide = base_slide(prs)
    text(
        slide,
        0.62,
        0.34,
        2.50,
        0.22,
        f"{code}  /  {category}",
        size=9,
        color=accent,
        bold=True,
    )
    text(
        slide,
        0.62,
        0.71,
        12.05,
        0.63,
        title_value,
        size=24,
        color=INK,
        bold=True,
    )
    line(slide, 0.62, 1.42, 12.71, 1.42, color=LINE, width=0.8)
    line(slide, 0.62, 6.70, 12.71, 6.70, color=LINE, width=0.6)
    text(
        slide,
        0.62,
        6.82,
        10.90,
        0.20,
        f"用途｜{use_case}　　作図ルール｜{rule}",
        size=8.5,
        color=MUTED,
    )
    text(
        slide,
        12.05,
        6.80,
        0.66,
        0.22,
        f"{len(prs.slides):02d}",
        size=9,
        color=accent,
        bold=True,
        align=PP_ALIGN.RIGHT,
    )
    return slide


def note(slide, x: float, y: float, w: float, title_value: str, body: str) -> None:
    text(slide, x, y, w, 0.24, title_value, size=9.5, color=COBALT, bold=True)
    text(slide, x, y + 0.35, w, 0.72, body, size=12.5, color=MUTED)


def node(
    slide,
    x: float,
    y: float,
    w: float,
    title_value: str,
    body: str = "",
    *,
    accent: str = COBALT,
    fill: str = WHITE,
    h: float = 1.04,
) -> None:
    rect(slide, x, y, w, h, fill=fill, line_color=LINE)
    rect(slide, x, y, 0.06, h, fill=accent, line_color=None)
    text(
        slide,
        x + 0.18,
        y + 0.15,
        w - 0.33,
        0.26,
        title_value,
        size=12.5,
        color=INK,
        bold=True,
    )
    if body:
        text(
            slide,
            x + 0.18,
            y + 0.52,
            w - 0.33,
            h - 0.62,
            body,
            size=10.5,
            color=MUTED,
        )


def add_axes(
    slide,
    x: float,
    y: float,
    w: float,
    h: float,
    x_label: str,
    y_label: str,
) -> None:
    line(slide, x, y + h, x + w, y + h, color=MUTED, width=0.9)
    line(slide, x, y + h, x, y, color=MUTED, width=0.9)
    text(
        slide,
        x + w - 0.85,
        y + h + 0.10,
        0.85,
        0.20,
        x_label,
        size=9,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )
    text(slide, x - 0.10, y - 0.27, 1.15, 0.20, y_label, size=9, color=MUTED)


def add_cover(prs: Presentation) -> None:
    slide = base_slide(prs)
    rect(slide, 0, 0, 0.16, SLIDE_H, fill=COBALT, line_color=None)
    text(slide, 0.74, 0.60, 3.20, 0.25, "SMEC  /  VISUAL LANGUAGE", size=10, color=COBALT, bold=True)
    text(
        slide,
        0.74,
        1.38,
        8.00,
        1.62,
        "複雑な論点を、\n一枚の構造に変える。",
        size=36,
        color=INK,
        bold=True,
    )
    text(
        slide,
        0.78,
        3.46,
        6.95,
        0.70,
        "中小企業診断士 教科書のための\n標準図解スライドシステム",
        size=18,
        color=MUTED,
    )
    line(slide, 0.78, 5.08, 8.35, 5.08, color=LINE, width=0.8)
    text(slide, 0.78, 5.35, 1.50, 0.40, "12", size=28, color=COBALT, bold=True)
    text(slide, 1.58, 5.52, 1.80, 0.20, "汎用図解", size=10, color=MUTED)
    text(slide, 3.40, 5.35, 1.50, 0.40, "06", size=28, color=TEAL, bold=True)
    text(slide, 4.20, 5.52, 2.30, 0.20, "経済学専用", size=10, color=MUTED)
    text(slide, 8.88, 1.12, 3.65, 0.36, "18 PATTERNS", size=11, color=MUTED, bold=True, align=PP_ALIGN.RIGHT)
    text(slide, 9.02, 1.65, 3.55, 1.45, "18", size=82, color=COBALT, bold=True, align=PP_ALIGN.RIGHT)
    text(slide, 9.10, 3.35, 3.45, 0.54, "図解の共通文法", size=20, color=INK, bold=True, align=PP_ALIGN.RIGHT)
    text(
        slide,
        9.10,
        4.18,
        3.45,
        1.20,
        "グリッド\nタイポグラフィ\n意味のある色",
        size=15,
        color=MUTED,
        align=PP_ALIGN.RIGHT,
    )
    text(slide, 0.78, 6.82, 4.50, 0.20, "Editorial Cobalt  /  v2.0", size=8.5, color=MUTED)


def add_principles(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "00",
        "DESIGN PRINCIPLES",
        "統一感は装飾ではなく、反復される判断基準から生まれる",
        "全スライド",
        "色・文字・整列・余白を固定する",
    )
    principles = [
        ("01", "結論を先に書く", "タイトルだけを追っても、論理の流れが分かる"),
        ("02", "図に仕事をさせる", "文章の要約ではなく、関係・変化・比較を見せる"),
        ("03", "色に意味を持たせる", "青は主張、ティールは比較、赤は例外だけに使う"),
        ("04", "余白を情報にする", "囲みを減らし、距離と整列でグループを示す"),
    ]
    for index, (number, title_value, body) in enumerate(principles):
        x = 0.74 + (index % 2) * 6.05
        y = 1.85 + (index // 2) * 2.15
        text(slide, x, y, 0.58, 0.38, number, size=20, color=COBALT, bold=True)
        line(slide, x + 0.72, y + 0.17, x + 1.20, y + 0.17, color=COBALT, width=1.5)
        text(slide, x + 1.42, y - 0.02, 4.25, 0.40, title_value, size=16, color=INK, bold=True)
        text(slide, x + 1.42, y + 0.55, 4.25, 0.70, body, size=12.5, color=MUTED)


def add_selector(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "00",
        "PATTERN SELECTOR",
        "伝えたい関係を定義すれば、選ぶべき図解は一つに絞れる",
        "図解パターンの選択",
        "内容ではなく関係性から型を選ぶ",
    )
    rows = [
        ("WHAT", "何を伝えるか", "PATTERN", "使う型"),
        ("STRUCTURE", "構成・分類", "G01 / G06 / G09", "定義・階層・ツリー"),
        ("DIFFERENCE", "違い・位置づけ", "G02 / G08", "比較・2×2"),
        ("CHANGE", "順序・時間", "G03 / G05 / G07", "プロセス・循環・時系列"),
        ("INFLUENCE", "原因・相互作用", "G04 / G10", "因果・ネットワーク"),
        ("EVIDENCE", "数式・数値", "G11 / G12", "数式・KPI"),
        ("ECONOMICS", "経済曲線", "E01 — E06", "専用グラフ"),
    ]
    widths = [1.72, 2.52, 2.32, 4.65]
    for row_index, row in enumerate(rows):
        y = 1.76 + row_index * 0.63
        x = 0.78
        for col_index, value in enumerate(row):
            fill = COBALT_PALE if row_index == 0 else (PAPER if row_index % 2 else WHITE)
            rect(slide, x, y, widths[col_index], 0.54, fill=fill, line_color=LINE, line_width=0.4)
            text(
                slide,
                x + 0.10,
                y + 0.15,
                widths[col_index] - 0.20,
                0.22,
                value,
                size=9.5 if row_index == 0 else 11.5,
                color=COBALT if row_index == 0 or col_index == 2 else INK,
                bold=row_index == 0 or col_index in (0, 2),
            )
            x += widths[col_index]


def generic_definition(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G01",
        "DEFINITION",
        "GDPは「国内」「一定期間」「新たな付加価値」の3条件で定義される",
        "概念の導入、用語定義",
        "3〜5要素、囲みではなく列で分割",
    )
    items = [
        ("01", "国内", "生産者の国籍ではなく、活動した場所で判断する"),
        ("02", "一定期間", "ある時点の資産ではなく、期間中の生産を測る"),
        ("03", "付加価値", "中間投入を除き、二重計上を避ける"),
    ]
    for index, (number, title_value, body) in enumerate(items):
        x = 0.78 + index * 4.12
        if index:
            line(slide, x - 0.34, 1.92, x - 0.34, 5.70, color=LINE, width=0.7)
        text(slide, x, 1.88, 0.62, 0.38, number, size=19, color=COBALT, bold=True)
        text(slide, x, 2.58, 3.45, 0.44, title_value, size=18, color=INK, bold=True)
        text(slide, x, 3.35, 3.30, 1.28, body, size=13.5, color=MUTED)
        line(slide, x, 5.20, x + 2.60, 5.20, color=COBALT, width=2)


def generic_comparison(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G02",
        "COMPARISON",
        "名目GDPは金額を、実質GDPは生産量の変化を捉える",
        "似た概念・制度・理論の比較",
        "共通の比較軸を同じ行に置く",
    )
    text(slide, 2.55, 1.74, 4.05, 0.35, "名目GDP", size=17, color=COBALT, bold=True)
    text(slide, 7.35, 1.74, 4.05, 0.35, "実質GDP", size=17, color=TEAL, bold=True)
    rows = [
        ("価格", "その年の価格", "基準年の価格"),
        ("物価変動", "含む", "取り除く"),
        ("主な用途", "経済規模の金額比較", "実質的な成長率"),
        ("問い", "いくら生産したか", "どれだけ増えたか"),
    ]
    for row_index, (axis, left, right) in enumerate(rows):
        y = 2.38 + row_index * 0.80
        rect(slide, 0.78, y, 1.42, 0.62, fill=COBALT_PALE if row_index == 3 else PAPER, line_color=LINE)
        text(slide, 0.92, y + 0.20, 1.12, 0.21, axis, size=10, color=MUTED, bold=True)
        rect(slide, 2.20, y, 4.68, 0.62, fill=WHITE, line_color=LINE)
        rect(slide, 6.88, y, 4.70, 0.62, fill=WHITE, line_color=LINE)
        text(slide, 2.55, y + 0.18, 3.95, 0.24, left, size=12.5, color=INK)
        text(slide, 7.35, y + 0.18, 3.95, 0.24, right, size=12.5, color=INK)


def generic_process(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G03",
        "PROCESS",
        "経済問題は「条件→モデル→方向→結論」の順に解けば迷わない",
        "解法、手続、業務フロー",
        "左から右、動詞でラベルを書く",
    )
    y = 3.35
    line(slide, 1.15, y, 12.05, y, color=LINE, width=1.4)
    steps = [
        ("01", "条件を読む", "外生変数を特定"),
        ("02", "モデルを選ぶ", "需要供給／IS-LM"),
        ("03", "方向を描く", "曲線をシフト"),
        ("04", "結論を出す", "価格・数量を判定"),
    ]
    for index, (number, title_value, body) in enumerate(steps):
        x = 1.00 + index * 3.02
        circle(slide, x, y - 0.13, 0.26, fill=COBALT)
        text(slide, x - 0.08, 2.03, 0.50, 0.28, number, size=12, color=COBALT, bold=True)
        text(slide, x - 0.08, 2.47, 2.35, 0.36, title_value, size=15, color=INK, bold=True)
        text(slide, x - 0.08, 3.82, 2.35, 0.56, body, size=11.5, color=MUTED)
        if index < 3:
            text(slide, x + 2.40, 3.17, 0.28, 0.28, "→", size=17, color=COBALT, bold=True)


def generic_cause(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G04",
        "CAUSE & EFFECT",
        "政府支出の増加は所得と消費の連鎖を通じてGDPを押し上げる",
        "政策効果、原因と結果",
        "矢印上に作用を示す動詞を置く",
    )
    items = [
        ("政府支出", "増える"),
        ("企業・家計所得", "受け取る"),
        ("消費支出", "再び増える"),
        ("GDP", "乗数倍になる"),
    ]
    for index, (title_value, verb) in enumerate(items):
        x = 0.75 + index * 3.05
        node(slide, x, 2.63, 2.44, title_value, accent=COBALT if index < 3 else TEAL, h=1.18)
        if index < len(items) - 1:
            line(slide, x + 2.44, 3.22, x + 2.92, 3.22, color=MUTED, width=1)
            triangle(slide, x + 2.78, 3.13, color=MUTED)
            text(slide, x + 2.28, 2.70, 0.95, 0.20, verb, size=8.5, color=MUTED, align=PP_ALIGN.CENTER)
    rect(slide, 3.54, 4.55, 6.22, 0.78, fill=RUST_PALE, line_color=None)
    text(slide, 3.78, 4.78, 5.75, 0.26, "反作用｜利子率上昇が民間投資を減らす場合もある", size=12, color=RUST, bold=True)


def generic_cycle(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G05",
        "CYCLE",
        "景気は回復・好況・後退・不況を繰り返しながら変動する",
        "循環、フィードバック",
        "時計回り、4〜5段階に限定",
    )
    center_x, center_y = 6.05, 3.46
    circle(slide, center_x - 0.62, center_y - 0.62, 1.24, fill=INK)
    text(slide, center_x - 0.44, center_y - 0.12, 0.88, 0.24, "景気", size=15, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    positions = [
        (5.25, 1.78, "回復", "生産・雇用が改善"),
        (8.87, 3.02, "好況", "需要・物価が上昇"),
        (5.25, 4.87, "後退", "在庫・金利が重荷"),
        (1.73, 3.02, "不況", "投資・雇用が減少"),
    ]
    for index, (x, y, title_value, body) in enumerate(positions):
        node(slide, x, y, 2.30, title_value, body, accent=COBALT if index != 2 else RUST, h=0.92)
        next_x, next_y, _, _ = positions[(index + 1) % 4]
        line(slide, x + 1.15, y + 0.46, next_x + 1.15, next_y + 0.46, color=LINE, width=1)


def generic_hierarchy(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G06",
        "HIERARCHY",
        "経済学はマクロとミクロに分かれ、個別論点へ段階的に具体化される",
        "包含関係、抽象度、組織階層",
        "上位から下位へ、幅ではなく字下げで示す",
    )
    levels = [
        (0.90, 1.85, 11.45, "経済学", "一国全体と個別主体の選択を扱う", COBALT),
        (1.72, 2.86, 10.05, "マクロ経済学 / ミクロ経済学", "分析対象による二分", COBALT),
        (2.54, 3.87, 8.65, "市場・政策・国際・消費者・企業", "主要な分析領域", TEAL),
        (3.36, 4.88, 7.25, "36の論点と定型問題", "試験で問われる具体知識", TEAL),
    ]
    for x, y, w, title_value, body, accent in levels:
        line(slide, x - 0.30, y, x - 0.30, y + 0.72, color=accent, width=2)
        text(slide, x, y, w * 0.42, 0.30, title_value, size=14, color=INK, bold=True)
        text(slide, x + w * 0.46, y + 0.02, w * 0.50, 0.28, body, size=10.5, color=MUTED)
        line(slide, x, y + 0.56, x + w, y + 0.56, color=LINE, width=0.6)


def generic_timeline(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G07",
        "TIMELINE",
        "経済理論は危機と政策課題への応答として発展してきた",
        "歴史、制度変更、ライフサイクル",
        "年代は左から右、節点は最大5個",
    )
    y = 3.62
    line(slide, 1.00, y, 12.15, y, color=MUTED, width=1.2)
    events = [
        ("1930s", "世界恐慌", "ケインズ"),
        ("1970s", "石油危機", "期待の導入"),
        ("1980s", "インフレ抑制", "マネタリズム"),
        ("2000s", "ゼロ金利", "非伝統的政策"),
    ]
    for index, (year, event, theory) in enumerate(events):
        x = 1.05 + index * 3.02
        circle(slide, x, y - 0.10, 0.20, fill=COBALT if index != 1 else RUST)
        top = index % 2 == 0
        ty = 2.02 if top else 4.06
        text(slide, x - 0.02, ty, 1.05, 0.28, year, size=13, color=COBALT, bold=True)
        text(slide, x - 0.02, ty + 0.43, 2.34, 0.30, event, size=13.5, color=INK, bold=True)
        text(slide, x - 0.02, ty + 0.86, 2.34, 0.26, theory, size=10.5, color=MUTED)


def generic_matrix(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G08",
        "2 × 2 MATRIX",
        "低成長・高失業の局面では需要刺激策の優先度が高まる",
        "2軸分類、ポジショニング",
        "注目象限だけを淡色で強調",
    )
    x, y, w, h = 2.12, 1.86, 8.45, 4.25
    rect(slide, x, y, w, h, fill=WHITE, line_color=LINE)
    rect(slide, x + w / 2, y + h / 2, w / 2, h / 2, fill=COBALT_PALE, line_color=None)
    line(slide, x + w / 2, y, x + w / 2, y + h, color=LINE, width=0.8)
    line(slide, x, y + h / 2, x + w, y + h / 2, color=LINE, width=0.8)
    labels = [
        (x + 0.30, y + 0.30, "高成長 × 低失業", "過熱・インフレを監視"),
        (x + w / 2 + 0.30, y + 0.30, "高成長 × 高失業", "構造的失業を点検"),
        (x + 0.30, y + h / 2 + 0.30, "低成長 × 低失業", "供給制約を点検"),
        (x + w / 2 + 0.30, y + h / 2 + 0.30, "低成長 × 高失業", "需要刺激策を検討"),
    ]
    for index, (tx, ty, title_value, body) in enumerate(labels):
        text(slide, tx, ty, 3.55, 0.30, title_value, size=13, color=COBALT if index == 3 else INK, bold=True)
        text(slide, tx, ty + 0.48, 3.55, 0.30, body, size=10.5, color=MUTED)
    text(slide, 0.76, 3.73, 1.05, 0.25, "成長率 ↑", size=10, color=MUTED, bold=True, align=PP_ALIGN.CENTER)
    text(slide, 5.36, 6.18, 2.15, 0.22, "失業率 →", size=10, color=MUTED, bold=True, align=PP_ALIGN.CENTER)


def generic_tree(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G09",
        "TREE",
        "物価指数は対象とウェイトの違いで3種類に分類できる",
        "分類、原因分解、論点体系",
        "同じ階層では分類基準を混ぜない",
    )
    node(slide, 0.78, 3.03, 2.05, "物価指数", "何の価格を測るか", accent=COBALT, h=1.00)
    branches = [
        ("CPI", "消費者の購入価格", 1.83),
        ("CGPI", "企業間取引価格", 3.10),
        ("GDPデフレーター", "GDP全体の価格", 4.37),
    ]
    line(slide, 3.30, 2.30, 3.30, 5.22, color=LINE, width=1)
    line(slide, 2.83, 3.53, 3.30, 3.53, color=LINE, width=1)
    for title_value, body, y in branches:
        line(slide, 3.30, y + 0.45, 4.02, y + 0.45, color=LINE, width=1)
        node(slide, 4.02, y, 2.65, title_value, body, accent=TEAL, h=0.90)
        line(slide, 6.67, y + 0.45, 7.25, y + 0.45, color=LINE, width=1)
        text(
            slide,
            7.48,
            y + 0.17,
            4.45,
            0.48,
            {
                "CPI": "ラスパイレス方式・基準年数量",
                "CGPI": "企業が購入する原材料や中間財",
                "GDPデフレーター": "パーシェ方式に近い・比較年数量",
            }[title_value],
            size=11.5,
            color=MUTED,
        )


def generic_network(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G10",
        "NETWORK",
        "市場では家計・企業・政府・海外が価格と所得を介して結びつく",
        "主体間関係、ステークホルダー",
        "中心1・周辺最大5、線の意味を明記",
    )
    circle(slide, 5.72, 2.81, 1.78, fill=INK)
    text(slide, 6.02, 3.47, 1.18, 0.26, "市場", size=17, color=WHITE, bold=True, align=PP_ALIGN.CENTER)
    actors = [
        (0.94, 1.90, "家計", "消費・労働"),
        (9.90, 1.90, "企業", "生産・投資"),
        (0.94, 4.62, "政府", "税・支出"),
        (9.90, 4.62, "海外", "輸出・輸入"),
    ]
    for index, (x, y, title_value, body) in enumerate(actors):
        node(slide, x, y, 2.35, title_value, body, accent=COBALT if index < 2 else TEAL, h=0.94)
        line(slide, 6.61, 3.70, x + 1.18, y + 0.47, color=LINE, width=1)
    text(slide, 5.10, 5.38, 3.10, 0.24, "財・サービス / お金 / 情報", size=10, color=MUTED, align=PP_ALIGN.CENTER)


def generic_formula(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G11",
        "FORMULA",
        "GDPデフレーターは名目値を実質値で割り、物価変動を取り出す",
        "数式、入力と出力",
        "式を中央、変数定義を周囲に配置",
    )
    text(slide, 0.82, 1.88, 2.00, 0.22, "INPUT", size=9, color=MUTED, bold=True)
    text(slide, 0.82, 2.40, 2.30, 0.38, "名目GDP", size=16, color=INK, bold=True)
    text(slide, 0.82, 3.08, 2.30, 0.38, "実質GDP", size=16, color=INK, bold=True)
    line(slide, 3.25, 1.88, 3.25, 5.64, color=LINE, width=0.8)
    text(slide, 3.74, 1.88, 4.95, 0.22, "RELATION", size=9, color=COBALT, bold=True)
    text(slide, 3.74, 2.66, 5.10, 0.86, "GDPデフレーター\n＝ 名目GDP ÷ 実質GDP × 100", size=22, color=COBALT, bold=True, align=PP_ALIGN.CENTER)
    line(slide, 9.17, 1.88, 9.17, 5.64, color=LINE, width=0.8)
    text(slide, 9.64, 1.88, 2.25, 0.22, "OUTPUT", size=9, color=TEAL, bold=True)
    text(slide, 9.64, 2.66, 2.28, 0.38, "物価水準", size=17, color=INK, bold=True)
    text(slide, 9.64, 3.27, 2.28, 0.70, "100より大きい\n→ 基準年より物価上昇", size=12, color=MUTED)


def generic_kpi(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "G12",
        "KPI & TABLE",
        "成長・雇用・物価を同じ基準で並べると、景気の全体像が読める",
        "統計、KPI、数値比較",
        "主指標のみ強調、表とチャートを競合させない",
    )
    indicators = [
        ("実質GDP", "+1.2%", "前年比"),
        ("完全失業率", "2.6%", "季節調整値"),
        ("CPI", "+2.4%", "前年同月比"),
    ]
    for index, (label_value, value, unit) in enumerate(indicators):
        x = 0.80 + index * 2.55
        text(slide, x, 1.87, 2.10, 0.24, label_value, size=10, color=MUTED, bold=True)
        text(slide, x, 2.27, 2.10, 0.50, value, size=25, color=COBALT if index == 0 else INK, bold=True)
        text(slide, x, 2.88, 2.10, 0.22, unit, size=8.5, color=MUTED)
    line(slide, 8.38, 1.82, 8.38, 5.65, color=LINE, width=0.8)
    bars = [1.15, 1.76, 1.42, 2.32]
    for index, bar_h in enumerate(bars):
        x = 9.10 + index * 0.68
        rect(slide, x, 5.14 - bar_h, 0.34, bar_h, fill=COBALT if index == 3 else COBALT_MID, line_color=None)
        text(slide, x - 0.02, 5.25, 0.40, 0.20, f"Q{index + 1}", size=8, color=MUTED, align=PP_ALIGN.CENTER)
    text(slide, 9.08, 1.86, 2.85, 0.26, "実質GDP成長率の推移", size=11.5, color=INK, bold=True)
    rows = [
        ("指標", "確認点"),
        ("GDP", "名目と実質を区別"),
        ("失業率", "労働力人口も確認"),
        ("物価", "指数の対象範囲"),
    ]
    for row_index, (left, right) in enumerate(rows):
        y = 3.62 + row_index * 0.52
        fill = COBALT_PALE if row_index == 0 else PAPER
        rect(slide, 0.80, y, 1.38, 0.45, fill=fill, line_color=LINE, line_width=0.4)
        rect(slide, 2.18, y, 5.47, 0.45, fill=fill, line_color=LINE, line_width=0.4)
        text(slide, 0.94, y + 0.14, 1.08, 0.18, left, size=9.5, color=COBALT if row_index == 0 else INK, bold=True)
        text(slide, 2.38, y + 0.14, 5.07, 0.18, right, size=9.5, color=COBALT if row_index == 0 else MUTED, bold=row_index == 0)


def econ_supply_demand(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E01",
        "SUPPLY & DEMAND",
        "需要と供給の交点が、均衡価格と均衡数量を同時に決める",
        "市場均衡、価格・数量決定",
        "需要＝青、供給＝赤茶で固定",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 7.20, 4.02
    add_axes(slide, x, y, w, h, "数量 Q", "価格 P")
    line(slide, x + 0.30, y + 0.35, x + w - 0.30, y + h - 0.35, color=RUST, width=2.2)
    line(slide, x + 0.30, y + h - 0.35, x + w - 0.30, y + 0.35, color=COBALT, width=2.2)
    text(slide, x + 6.12, y + 3.22, 0.55, 0.22, "S", size=12, color=RUST, bold=True)
    text(slide, x + 6.12, y + 0.48, 0.55, 0.22, "D", size=12, color=COBALT, bold=True)
    ex, ey = x + w / 2, y + h / 2
    circle(slide, ex - 0.06, ey - 0.06, 0.12, fill=INK)
    line(slide, ex, ey, ex, y + h, color=LINE, width=0.6)
    line(slide, x, ey, ex, ey, color=LINE, width=0.6)
    text(slide, ex + 0.16, ey - 0.30, 0.85, 0.22, "E₀", size=10, color=INK, bold=True)
    note(slide, 8.72, 2.02, 3.30, "DEMAND", "価格が上がるほど、需要量は減る")
    note(slide, 8.72, 3.34, 3.30, "SUPPLY", "価格が上がるほど、供給量は増える")
    note(slide, 8.72, 4.66, 3.30, "EQUILIBRIUM", "超過需要も超過供給もない")


def econ_shift(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E02",
        "CURVE SHIFT",
        "需要の増加は曲線を右へ動かし、価格と数量をともに押し上げる",
        "比較静学、外生変数の変化",
        "同一曲線は同一色相の濃淡で表す",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 7.20, 4.02
    add_axes(slide, x, y, w, h, "数量 Q", "価格 P")
    line(slide, x + 0.30, y + 0.35, x + w - 0.30, y + h - 0.35, color=RUST, width=2.2)
    line(slide, x + 0.30, y + h - 0.64, x + w - 1.00, y + 0.42, color=COBALT_MID, width=1.5)
    line(slide, x + 1.18, y + h - 0.64, x + w - 0.12, y + 0.42, color=COBALT, width=2.2)
    text(slide, x + 5.08, y + 0.53, 0.65, 0.22, "D₀", size=10, color=COBALT_MID, bold=True)
    text(slide, x + 6.02, y + 0.53, 0.65, 0.22, "D₁", size=10, color=COBALT, bold=True)
    text(slide, x + 6.12, y + 3.22, 0.55, 0.22, "S", size=10, color=RUST, bold=True)
    line(slide, 5.65, 2.16, 6.36, 2.16, color=COBALT, width=1.2)
    triangle(slide, 6.25, 2.08, color=COBALT)
    notes = [
        ("CAUSE", "所得増加・嗜好変化"),
        ("SHIFT", "D₀ → D₁（右方）"),
        ("RESULT", "均衡価格 ↑ / 均衡数量 ↑"),
    ]
    for index, (heading, body) in enumerate(notes):
        note(slide, 8.72, 2.04 + index * 1.22, 3.40, heading, body)


def econ_keynesian(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E03",
        "KEYNESIAN CROSS",
        "政府支出の増加は総需要を上方へ動かし、均衡GDPを乗数倍押し上げる",
        "45度線分析、乗数効果",
        "45度線は灰、総需要線は青で固定",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 7.20, 4.02
    add_axes(slide, x, y, w, h, "所得 Y", "総需要 AE")
    line(slide, x, y + h, x + w - 0.36, y + 0.36, color=MUTED, width=1.2)
    line(slide, x + 0.30, y + 3.12, x + w - 0.40, y + 1.30, color=COBALT_MID, width=1.5)
    line(slide, x + 0.30, y + 2.48, x + w - 0.40, y + 0.66, color=COBALT, width=2.2)
    text(slide, x + 5.90, y + 1.32, 0.72, 0.22, "AE₀", size=10, color=COBALT_MID, bold=True)
    text(slide, x + 5.90, y + 0.68, 0.72, 0.22, "AE₁", size=10, color=COBALT, bold=True)
    text(slide, x + 5.65, y + 0.22, 0.90, 0.22, "45°", size=9, color=MUTED)
    note(slide, 8.72, 2.05, 3.35, "POLICY", "政府支出 G を増やす")
    note(slide, 8.72, 3.34, 3.35, "SHIFT", "総需要 AE が上方へ移動")
    note(slide, 8.72, 4.63, 3.35, "MULTIPLIER", "ΔY ＝ 1/(1−c) × ΔG")


def econ_is_lm(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E04",
        "IS–LM",
        "拡張的財政政策はISを右へ動かし、GDPと利子率を上昇させる",
        "財政・金融政策の比較",
        "政策→曲線→交点→Y・rの順で示す",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 7.20, 4.02
    add_axes(slide, x, y, w, h, "GDP Y", "利子率 r")
    line(slide, x + 0.42, y + 0.44, x + w - 0.30, y + h - 0.36, color=RUST, width=2.2)
    line(slide, x + 0.30, y + h - 0.42, x + w - 1.00, y + 0.42, color=COBALT_MID, width=1.5)
    line(slide, x + 1.10, y + h - 0.42, x + w - 0.20, y + 0.42, color=COBALT, width=2.2)
    text(slide, x + 5.98, y + 3.24, 0.62, 0.22, "LM", size=10, color=RUST, bold=True)
    text(slide, x + 5.06, y + 0.52, 0.66, 0.22, "IS₀", size=10, color=COBALT_MID, bold=True)
    text(slide, x + 5.96, y + 0.52, 0.66, 0.22, "IS₁", size=10, color=COBALT, bold=True)
    notes = [
        ("POLICY", "政府支出 G が増加"),
        ("SHIFT", "IS₀ → IS₁（右方）"),
        ("RESULT", "GDP ↑ / 利子率 ↑"),
        ("OFFSET", "民間投資は一部減少"),
    ]
    for index, (heading, body) in enumerate(notes):
        note(slide, 8.72, 1.82 + index * 1.05, 3.35, heading, body)


def econ_surplus(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E05",
        "SURPLUS",
        "課税で失われた余剰のうち、税収にならない部分が死荷重になる",
        "課税、独占、関税、余剰分析",
        "余剰・税収・死荷重の色を固定",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 7.20, 4.02
    add_axes(slide, x, y, w, h, "数量 Q", "価格 P")
    line(slide, x + 0.30, y + 0.35, x + w - 0.30, y + h - 0.35, color=RUST, width=2.2)
    line(slide, x + 0.30, y + h - 0.35, x + w - 0.30, y + 0.35, color=COBALT, width=2.2)
    # 図形で余剰の領域を示す（教材用の概念図）
    top = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x + 2.15),
        Inches(y + 0.65),
        Inches(3.00),
        Inches(1.35),
    )
    top.rotation = 180
    top.fill.solid()
    top.fill.fore_color.rgb = rgb(COBALT_PALE)
    top.line.fill.background()
    bottom = slide.shapes.add_shape(
        MSO_SHAPE.ISOSCELES_TRIANGLE,
        Inches(x + 2.15),
        Inches(y + 2.02),
        Inches(3.00),
        Inches(1.35),
    )
    bottom.fill.solid()
    bottom.fill.fore_color.rgb = rgb(TEAL_PALE)
    bottom.line.fill.background()
    text(slide, x + 2.72, y + 1.30, 1.90, 0.22, "消費者余剰", size=10, color=COBALT, bold=True, align=PP_ALIGN.CENTER)
    text(slide, x + 2.72, y + 2.58, 1.90, 0.22, "生産者余剰", size=10, color=TEAL, bold=True, align=PP_ALIGN.CENTER)
    note(slide, 8.72, 2.03, 3.38, "TRANSFER", "余剰の一部は政府税収へ移る")
    note(slide, 8.72, 3.42, 3.38, "LOSS", "取引減少分は誰の利益にもならない")
    rect(slide, 8.72, 4.72, 3.40, 0.76, fill=RUST_PALE, line_color=None)
    text(slide, 8.96, 4.96, 2.96, 0.24, "死荷重 ＝ 純粋な社会的損失", size=12, color=RUST, bold=True)


def econ_cycle(prs: Presentation) -> None:
    slide = pattern_slide(
        prs,
        "E06",
        "BUSINESS CYCLE",
        "実質GDPは長期トレンドの周囲を循環し、山と谷で局面が転換する",
        "景気循環、時系列",
        "実績線1色＋トレンド灰色＋直接ラベル",
        accent=TEAL,
    )
    x, y, w, h = 0.95, 1.82, 8.30, 4.02
    add_axes(slide, x, y, w, h, "時間 t", "実質GDP")
    # 後退局面
    rect(slide, x + 2.15, y, 1.76, h, fill="EFEEEE", line_color=None)
    rect(slide, x + 6.13, y, 1.34, h, fill="EFEEEE", line_color=None)
    line(slide, x + 0.15, y + h - 0.38, x + w - 0.24, y + 0.58, color=MUTED, width=1)
    points = [
        (x + 0.20, y + 3.45),
        (x + 1.20, y + 2.48),
        (x + 2.25, y + 1.50),
        (x + 3.20, y + 2.35),
        (x + 4.20, y + 2.82),
        (x + 5.25, y + 1.68),
        (x + 6.28, y + 0.82),
        (x + 7.20, y + 1.62),
        (x + 8.02, y + 1.98),
    ]
    for first, second in zip(points, points[1:]):
        line(slide, first[0], first[1], second[0], second[1], color=COBALT, width=2.3)
    text(slide, x + 2.05, y + 1.12, 0.70, 0.22, "山", size=10, color=RUST, bold=True)
    text(slide, x + 3.94, y + 2.98, 0.70, 0.22, "谷", size=10, color=TEAL, bold=True)
    text(slide, x + 7.38, y + 0.76, 0.72, 0.22, "実績", size=9, color=COBALT, bold=True)
    text(slide, x + 7.38, y + 1.24, 0.72, 0.22, "トレンド", size=9, color=MUTED)
    text(slide, 9.72, 2.02, 2.35, 0.24, "SHADED AREA", size=9, color=MUTED, bold=True)
    text(slide, 9.72, 2.42, 2.35, 0.56, "景気後退局面", size=14, color=INK, bold=True)
    text(slide, 9.72, 3.42, 2.35, 0.24, "TURNING POINT", size=9, color=MUTED, bold=True)
    text(slide, 9.72, 3.82, 2.35, 0.82, "山：拡張→後退\n谷：後退→拡張", size=12.5, color=INK)


def build_deck(output: Path) -> None:
    prs = Presentation()
    prs.slide_width = Inches(SLIDE_W)
    prs.slide_height = Inches(SLIDE_H)
    prs.core_properties.title = "SMEC標準図解スライドテンプレート Editorial Cobalt"
    prs.core_properties.subject = "汎用図解12種＋経済学専用グラフ6種"
    prs.core_properties.author = "SMEC Project"

    add_cover(prs)
    add_principles(prs)
    add_selector(prs)

    patterns: list[Callable[[Presentation], None]] = [
        generic_definition,
        generic_comparison,
        generic_process,
        generic_cause,
        generic_cycle,
        generic_hierarchy,
        generic_timeline,
        generic_matrix,
        generic_tree,
        generic_network,
        generic_formula,
        generic_kpi,
        econ_supply_demand,
        econ_shift,
        econ_keynesian,
        econ_is_lm,
        econ_surplus,
        econ_cycle,
    ]
    for pattern in patterns:
        pattern(prs)

    output.parent.mkdir(parents=True, exist_ok=True)
    prs.save(output)
    print(f"Generated {len(prs.slides)} slides → {output}")


if __name__ == "__main__":
    build_deck(OUTPUT)
