import io
import os
import re
import zipfile
from typing import Dict, List, Optional, Tuple
from collections import Counter

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont

from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont


# =========================================================
# 고정 머릿말/꼬릿말
# =========================================================
HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

PAGE_TITLE = "유진 sat class report"

# Tab1 (대단원) 표시용
TOPIC_NAMES = {
    1: "1. Linear",
    2: "2. Percent & Unit Conversion",
    3: "3. Quadratic",
    4: "4. Exponential",
    5: "5. Polynomials, radical and rational functions",
    6: "6. Geometry",
    7: "7. Statistics",
}

# =========================================================
# ✅ Tab2: 소단원 매핑 (코드형 단원 -> 소단원명)
# =========================================================
TOPIC_SUB_NAMES = {
    "1.1": "Linear function",
    "1.2": "Linear equation",
    "1.3": "Linear interpretation",
    "1.4": "Linear word problems",
    "1.5": "Linear inequality",
    "1.6": "Identity equation",
    "1.7": "Absolute function and equation",
    "1.8": "System of equations",

    "2.1": "Ratios and Percent",
    "2.2": "Unit conversion",

    "3.1": "Quadratic function",
    "3.2": "Quadratic equation and inequality",
    "3.3": "Sum and product",
    "3.4": "Discriminant",
    "3.5": "Quadratic Word problems",
    "3.6": "Factoring",

    "4.1": "Exponential equation",
    "4.2": "Exponential function",
    "4.3": "Exponential model",

    "5.1": "Polynomial equation and graph",
    "5.2": "Polynomial – Long division and factor/remainder theorem",
    "5.3": "Radical equation and function",
    "5.4": "Rational expression and rational exponent",
    "5.5": "Rational equation and function",
    "5.6": "Isolation",

    "6.1": "Similar and congruent triangles",
    "6.2": "Similar figure",
    "6.3": "Right triangle and trigonometry",
    "6.4": "Volume and surface area",
    "6.5": "Parallel lines",
    "6.6": "Circle",
    "6.7": "Polygon and ETC",

    "7.1": "Probability",
    "7.2": "Conditional probability",
    "7.3": "Scatter plot",
    "7.4": "Sampling method",
    "7.5": "Generalize",
    "7.6": "Mean, median, mode",
    "7.7": "Standard deviation",
    "7.8": "Margin of error",
    "7.9": "Experiment",
    "7.10": "Box plot",
}

FONT_REG = "fonts/NanumGothic-Regular.ttf"
FONT_BOLD = "fonts/NanumGothic-Bold.ttf"


@st.cache_resource
def load_pil_fonts():
    if not os.path.exists(FONT_REG) or not os.path.exists(FONT_BOLD):
        raise FileNotFoundError(
            "폰트 파일을 찾지 못했습니다.\n\n"
            "필요 파일:\n"
            f"- {FONT_REG}\n"
            f"- {FONT_BOLD}\n\n"
            "GitHub 레포에 fonts 폴더를 만들고 폰트 파일을 올려주세요."
        )

    def f(path, size):
        return ImageFont.truetype(path, size=size)

    return {
        "title": f(FONT_BOLD, 32),
        "h2": f(FONT_BOLD, 19),
        "b": f(FONT_BOLD, 17),
        "small_b": f(FONT_BOLD, 14),
        "small": f(FONT_REG, 14),
        "tiny": f(FONT_REG, 12),
    }


@st.cache_resource
def ensure_reportlab_fonts():
    if not os.path.exists(FONT_REG) or not os.path.exists(FONT_BOLD):
        raise FileNotFoundError(
            "폰트 파일을 찾지 못했습니다.\n\n"
            "필요 파일:\n"
            f"- {FONT_REG}\n"
            f"- {FONT_BOLD}\n"
        )
    try:
        pdfmetrics.getFont("NanumGothic")
    except KeyError:
        pdfmetrics.registerFont(TTFont("NanumGothic", FONT_REG))

    try:
        pdfmetrics.getFont("NanumGothic-Bold")
    except KeyError:
        pdfmetrics.registerFont(TTFont("NanumGothic-Bold", FONT_BOLD))


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else "값없음"


def make_zip(files: Dict[str, bytes]) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in files.items():
            zf.writestr(fname, data)
    return bio.getvalue()


def pil_to_png_bytes(img: Image.Image) -> bytes:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


def fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        fv = float(v)
        if abs(fv - round(fv)) < 1e-9:
            return str(int(round(fv)))
        return f"{fv:g}"
    except Exception:
        return str(v)


def norm_module(v) -> Optional[int]:
    s = str(v).strip().upper()
    if s in ["M1", "MODULE1", "1"]:
        return 1
    if s in ["M2", "MODULE2", "2"]:
        return 2
    return None


def parse_wrong_list(val) -> List[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []
    if isinstance(val, (int, float)) and not pd.isna(val):
        n = int(round(float(val)))
        return [n] if n > 0 else []
    s = str(val).strip()
    if s == "" or s.upper() in ["X", "Х", "-"]:
        return []
    if re.fullmatch(r"\d+(\.0+)?", s):
        return [int(float(s))]
    s = s.replace("，", ",").replace(";", ",").replace("/", ",")
    s = re.sub(r"\s+", ",", s)
    nums = re.findall(r"\d+", s)
    out = [int(x) for x in nums]
    out = [n for n in out if n != 0]
    return out


def guess_col(df: pd.DataFrame, exact: List[str] = None, regexes: List[str] = None) -> Optional[str]:
    cols = [str(c).strip() for c in df.columns]
    if exact:
        for e in exact:
            for c in cols:
                if c == e:
                    return c
    if regexes:
        for rx in regexes:
            r = re.compile(rx, re.IGNORECASE)
            for c in cols:
                if r.search(c):
                    return c
    return None


# =========================================================
# ✅ Tab2: "코드형 단원" -> "소단원 표시명" 변환
# =========================================================
def normalize_topic_code(topic_raw: str) -> str:
    """
    topic_raw 예시:
    - "5.3"
    - " 5.3 Radical equation..."
    - "5.3-2)" 같은 잡문자
    -> "5.3" 코드만 뽑기
    """
    s = str(topic_raw or "").strip()
    if not s:
        return ""
    m = re.search(r"(\d+(?:\.\d+)?)", s)
    return m.group(1) if m else ""


def topic_display_sub(topic_raw: str) -> str:
    """
    소단원 표시: "5.3 Radical equation and function"
    매핑 없으면 코드만이라도 반환
    """
    code = normalize_topic_code(topic_raw)
    if not code:
        return ""
    sub = TOPIC_SUB_NAMES.get(code, "")
    if sub:
        return f"{code} {sub}"
    return code


# =========================================================
# 점수/리포트용(탭1) 로더/유틸
# =========================================================
def load_score_excel(uploaded_file) -> Tuple[pd.DataFrame, str, Optional[str]]:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    name_col = guess_col(df, exact=["Name", "이름"], regexes=[r"(학생|student).*이름", r"^name$"])
    if not name_col:
        raise KeyError("점수 파일에서 'Name' (또는 '이름') 컬럼을 찾지 못했습니다.")

    class_col = guess_col(df, exact=["Class", "반", "클래스"], regexes=[r"class"])

    df[name_col] = df[name_col].astype(str).str.strip()
    if class_col:
        df[class_col] = df[class_col].astype(str).str.strip()

    return df, name_col, class_col


# ✅ Quiz(숫자) 먼저 → ReviewQuiz는 항상 맨 뒤
def get_columns(df: pd.DataFrame):
    def num_key(col):
        m = re.search(r"(\d+)", str(col))
        return int(m.group(1)) if m else 9999

    quiz_numbered = []
    quiz_other = []
    review = []
    mock_cols = []
    hw_cols = []

    for c in df.columns:
        s = str(c).strip()

        if re.match(r"^Mocktest\s*\d+", s, re.IGNORECASE):
            mock_cols.append(c)
            continue
        if re.match(r"^Homework\s*\d+", s, re.IGNORECASE):
            hw_cols.append(c)
            continue

        if re.match(r"^ReviewQuiz", s, re.IGNORECASE):
            review.append(c)
            continue

        m = re.match(r"^Quiz\s*0*(\d+)$", s, re.IGNORECASE)
        if m:
            quiz_numbered.append((int(m.group(1)), c))
            continue

        if re.match(r"^Quiz", s, re.IGNORECASE):
            quiz_other.append(c)
            continue

    quiz_cols = [c for _, c in sorted(quiz_numbered, key=lambda x: x[0])]
    quiz_cols += sorted(quiz_other, key=num_key)
    quiz_cols += sorted(review, key=num_key)

    mock_cols = sorted(mock_cols, key=num_key)
    hw_cols = sorted(hw_cols, key=num_key)

    return quiz_cols, mock_cols, hw_cols


def find_avg_row(df: pd.DataFrame, name_col: str) -> pd.Series:
    mask = df[name_col].astype(str).str.strip() == "평균"
    if mask.sum() == 0:
        raise ValueError("평균행(Name='평균')을 찾지 못했습니다.")
    return df.loc[mask].iloc[0]


def students_list(df: pd.DataFrame, name_col: str) -> List[str]:
    names = df[name_col].dropna().astype(str).str.strip()
    names = [n for n in names.tolist() if n not in ["", "nan", "평균"]]
    seen, out = set(), []
    for n in names:
        if n not in seen:
            seen.add(n)
            out.append(n)
    return out


def get_student_row(df: pd.DataFrame, name_col: str, student_name: str) -> pd.Series:
    mask = df[name_col].astype(str).str.strip() == str(student_name).strip()
    if mask.sum() == 0:
        raise ValueError(f"학생을 찾지 못했습니다: {student_name}")
    return df.loc[mask].iloc[0]


def find_wrong_cols_in_score(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    m1 = guess_col(df, regexes=[r"\bM1\b.*(틀린|오답).*문제", r"(틀린|오답).*문제.*\bM1\b"])
    m2 = guess_col(df, regexes=[r"\bM2\b.*(틀린|오답).*문제", r"(틀린|오답).*문제.*\bM2\b"])
    return m1, m2


def build_wrong_map_from_score(df: pd.DataFrame, name_col: str, m1_wrong_col: str, m2_wrong_col: str) -> Dict[str, Dict[int, set]]:
    out: Dict[str, Dict[int, set]] = {}
    for _, r in df.iterrows():
        nm = str(r.get(name_col, "")).strip()
        if nm == "" or nm.lower() == "nan" or nm == "평균":
            continue

        m1_nums = parse_wrong_list(r.get(m1_wrong_col, "")) if m1_wrong_col else []
        m2_nums = parse_wrong_list(r.get(m2_wrong_col, "")) if m2_wrong_col else []

        out[nm] = {
            1: set([n for n in m1_nums if 1 <= n <= 22]),
            2: set([n for n in m2_nums if 1 <= n <= 22]),
        }
    return out


def read_mock_meta(mock_file) -> pd.DataFrame:
    df = pd.read_excel(mock_file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_module = guess_col(df, exact=["모듈", "Module"], regexes=[r"모듈", r"module"])
    col_q = guess_col(
        df,
        exact=["문항번호", "문항", "문제번호", "No", "Q"],
        regexes=[r"(문항|문제).*(번호)?", r"q\s*no", r"^no$"],
    )
    col_topic = guess_col(df, exact=["단원", "Topic"], regexes=[r"단원", r"topic"])

    if not col_module or not col_q or not col_topic:
        raise ValueError("Mock 메타 파일에는 최소 '모듈', '문항번호', '단원' 컬럼이 필요합니다.")

    out = df.copy()
    out["__module__"] = out[col_module].apply(norm_module)
    out["__q__"] = pd.to_numeric(out[col_q], errors="coerce")
    out = out.dropna(subset=["__module__", "__q__"]).copy()
    out["__module__"] = out["__module__"].astype(int)
    out["__q__"] = out["__q__"].astype(int)

    out = out[(out["__q__"] >= 1) & (out["__q__"] <= 22)].copy()
    out["__topic_raw__"] = out[col_topic].astype(str).str.strip()

    out = out[["__module__", "__q__", "__topic_raw__"]].drop_duplicates(subset=["__module__", "__q__"])
    return out


def compute_topic_accuracy(meta_df: pd.DataFrame, wrong_map: Dict[str, Dict[int, set]], student: str) -> Dict[str, Tuple[int, int, float]]:
    # Tab1 자동추천(대단원) 유지
    if student not in wrong_map:
        wrong_map[student] = {1: set(), 2: set()}

    totals: Dict[int, int] = {k: 0 for k in range(1, 8)}
    wrongs: Dict[int, int] = {k: 0 for k in range(1, 8)}

    for _, r in meta_df.iterrows():
        md = int(r["__module__"])
        q = int(r["__q__"])
        code = normalize_topic_code(r["__topic_raw__"])
        if not code:
            continue
        try:
            major = int(code.split(".")[0])
        except:
            continue
        if major not in totals:
            continue

        totals[major] += 1
        if q in wrong_map[student].get(md, set()):
            wrongs[major] += 1

    out: Dict[str, Tuple[int, int, float]] = {}
    for major in range(1, 8):
        total = totals[major]
        wrong = wrongs[major]
        correct = total - wrong
        acc = (correct / total) if total > 0 else 0.0
        out[TOPIC_NAMES[major]] = (correct, total, acc)

    return out


def auto_recommend_topics(topic_acc: Dict[str, Tuple[int, int, float]], threshold: float) -> List[str]:
    items = []
    for topic, (c, t, acc) in topic_acc.items():
        if t <= 0:
            continue
        if acc <= threshold:
            items.append((acc, topic))
    items.sort(key=lambda x: x[0])
    return [t for _, t in items]


def build_topic_units(selected_topics: List[str]) -> List[str]:
    if not selected_topics:
        return []
    return [str(u) for u in selected_topics]


def compute_hw_progress(student_row: pd.Series, hw_cols: list[str]):
    if not hw_cols:
        return None
    vals = pd.to_numeric(student_row[hw_cols], errors="coerce").dropna()
    if len(vals) == 0:
        return None
    avg = float(vals.mean())
    if avg <= 1.0:
        avg *= 100.0
    return avg


def draw_line(draw, x1, y1, x2, y2, color="#D9D9D9", w=2):
    draw.line((x1, y1, x2, y2), fill=color, width=w)


def draw_text(draw, x, y, text, font, fill="#111111"):
    draw.text((x, y), text, font=font, fill=fill)


def right_text(draw, rx, y, text, font, fill="#111111"):
    tw = draw.textlength(text, font=font)
    draw.text((rx - tw, y), text, font=font, fill=fill)


def title_height(draw, class_name, student_name, fonts, max_w):
    one_line = f"{class_name} {student_name} CLASS REPORT"
    if draw.textlength(one_line, font=fonts["title"]) <= max_w:
        return 48, [one_line]
    else:
        return 88, [f"{class_name} {student_name}", "CLASS REPORT"]


def table_height(n_rows, title_gap=30, header_h=30, row_h=30):
    return title_gap + header_h + n_rows * row_h


def render_table(draw, x, y, w, title, rows, fonts, row_h=30):
    draw_text(draw, x, y, title, fonts["h2"], fill="#111111")
    y += 30

    col1 = int(w * 0.60)
    col2 = int(w * 0.20)

    draw.rectangle([x, y, x + w, y + row_h], fill="#F5F6F8", outline=None)
    right_text(draw, x + col1 + col2 - 10, y + 7, "점수", fonts["small_b"], fill="#333333")
    right_text(draw, x + w - 10, y + 7, "class 평균", fonts["small_b"], fill="#333333")
    draw_line(draw, x, y + row_h, x + w, y + row_h, color="#E1E4E8", w=2)
    y += row_h

    for r in rows:
        label = str(r["label"])
        sv = fmt_num(r["student"])
        av = fmt_num(r["avg"])

        draw_text(draw, x + 10, y + 7, label, fonts["small"], fill="#111111")
        right_text(draw, x + col1 + col2 - 10, y + 7, sv, fonts["small"], fill="#111111")
        right_text(draw, x + w - 10, y + 7, av, fonts["small"], fill="#666666")

        draw_line(draw, x, y + row_h, x + w, y + row_h, color="#EDEFF2", w=2)
        y += row_h

    return y


def build_rows(student_row: pd.Series, avg_row: pd.Series, quiz_cols, mock_cols):
    quiz_rows = [{"label": c, "student": student_row.get(c), "avg": avg_row.get(c)} for c in quiz_cols]
    mock_rows = [{"label": c, "student": student_row.get(c), "avg": avg_row.get(c)} for c in mock_cols]
    return quiz_rows, mock_rows


def render_student_report_image(
    class_name: str,
    student_name: str,
    quiz_rows,
    mock_rows,
    hw_progress,
    topic_units: List[str],
    fonts,
):
    W = 877
    margin = 22
    w = W - 2 * margin

    dummy = Image.new("RGB", (W, 200), "white")
    ddraw = ImageDraw.Draw(dummy)

    header_h = 40
    y_title = 50

    th, title_lines = title_height(ddraw, class_name, student_name, fonts, w)

    ROW_H = 30
    GAP = 14

    h_quiz = table_height(len(quiz_rows), title_gap=30, header_h=ROW_H, row_h=ROW_H)
    h_mock = table_height(len(mock_rows), title_gap=30, header_h=ROW_H, row_h=ROW_H)
    h_hw = 30 + 44 + 14

    line_h = 24
    max_lines = min(12, len(topic_units)) if topic_units else 0
    topic_box_h = max(110, 14 + max_lines * line_h + 14)

    content_h = (
        header_h
        + (y_title - header_h)
        + th + 6
        + h_quiz + GAP
        + h_mock + GAP
        + h_hw
        + 30
        + 30
        + topic_box_h
    )

    footer_h = 42
    bottom_pad = 10
    H = int(content_h + footer_h + bottom_pad)

    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    draw_text(draw, margin, 10, HEADER_TEXT, fonts["small_b"], fill="#111111")
    draw_line(draw, margin, 38, W - margin, 38, color="#D9D9D9", w=2)

    y = y_title
    if len(title_lines) == 1:
        draw_text(draw, margin, y, title_lines[0], fonts["title"], fill="#111111")
        y += 48
    else:
        draw_text(draw, margin, y, title_lines[0], fonts["title"], fill="#111111")
        draw_text(draw, margin, y + 38, title_lines[1], fonts["title"], fill="#111111")
        y += 88
    y += 6

    y = render_table(draw, margin, y, w, "Quiz", quiz_rows, fonts, row_h=ROW_H)
    y += GAP
    y = render_table(draw, margin, y, w, "Mocktest (점수 예상)", mock_rows, fonts, row_h=ROW_H)
    y += GAP

    draw_text(draw, margin, y, "Homework 진행도", fonts["h2"], fill="#111111")
    y += 30
    badge_h = 44
    draw.rounded_rectangle([margin, y, margin + w, y + badge_h], radius=18, fill="#F5F6F8", outline=None)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    draw_text(draw, margin + 14, y + 10, hw_txt, fonts["b"], fill="#111111")
    y += badge_h + 14

    draw_text(draw, margin, y, "보강이 필요한 부분 및 유진쌤 Comment", fonts["h2"], fill="#111111")
    y += 30

    draw.rounded_rectangle([margin, y, W - margin, y + topic_box_h], radius=20, fill="#F9FAFB", outline=None)

    yy = y + 14
    for unit in (topic_units[:12] if topic_units else []):
        draw_text(draw, margin + 12, yy, unit, fonts["small"], fill="#111111")
        yy += 24

    footer_y_line = H - 42
    draw_line(draw, margin, footer_y_line, W - margin, footer_y_line, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 30, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    return img


def rl_str_w(text: str, font_name: str, font_size: float) -> float:
    return pdfmetrics.stringWidth(text, font_name, font_size)


def create_editable_pdf_bytes(
    class_name: str,
    student_name: str,
    quiz_rows: List[dict],
    mock_rows: List[dict],
    hw_progress: Optional[float],
    topic_units: List[str],
) -> bytes:
    ensure_reportlab_fonts()

    buf = io.BytesIO()
    page_w = 210 * mm

    top = 12 * mm
    left = 10 * mm
    right = 10 * mm
    usable_w = page_w - left - right

    row_h = 8 * mm
    header_h = 9 * mm
    title_gap = 10 * mm
    section_gap = 6 * mm

    quiz_h = title_gap + header_h + (len(quiz_rows) * row_h) + 4 * mm
    mock_h = title_gap + header_h + (len(mock_rows) * row_h) + 4 * mm
    hw_h = title_gap + 12 * mm + 6 * mm
    topic_title_h = title_gap
    topic_line_h = 7 * mm
    topic_box_min = 28 * mm
    topic_h = topic_title_h + max(topic_box_min, (len(topic_units[:12]) * topic_line_h + 12 * mm))

    header_block = 12 * mm
    footer_block = 12 * mm
    main_title_h = 22 * mm

    page_h = top + header_block + main_title_h + quiz_h + section_gap + mock_h + section_gap + hw_h + section_gap + topic_h + footer_block + 6 * mm

    c = canvas.Canvas(buf, pagesize=(page_w, page_h))
    W, H = page_w, page_h

    title_col = colors.Color(17/255, 17/255, 17/255)
    muted = colors.Color(90/255, 90/255, 90/255)
    line_col = colors.Color(0.85, 0.85, 0.85)
    header_fill = colors.Color(245/255, 246/255, 248/255)
    topic_box_fill = colors.Color(249/255, 250/255, 251/255)

    y = H - top

    c.setFont("NanumGothic-Bold", 11)
    c.setFillColor(title_col)
    c.drawString(left, y - 8*mm, HEADER_TEXT)
    c.setStrokeColor(line_col)
    c.setLineWidth(1)
    c.line(left, y - 11*mm, W - right, y - 11*mm)
    y -= 12 * mm

    c.setFont("NanumGothic-Bold", 20)
    c.setFillColor(title_col)
    title_text = f"{class_name} {student_name} CLASS REPORT"
    if rl_str_w(title_text, "NanumGothic-Bold", 20) <= usable_w:
        c.drawString(left, y - 18, title_text)
        y -= 18 * mm
    else:
        c.drawString(left, y - 18, f"{class_name} {student_name}")
        c.drawString(left, y - 18 - 9*mm, "CLASS REPORT")
        y -= 26 * mm

    def draw_table(title: str, rows: List[dict]):
        nonlocal y
        c.setFont("NanumGothic-Bold", 13)
        c.setFillColor(title_col)
        c.drawString(left, y - 5*mm, title)
        y -= 10 * mm

        c.setFillColor(header_fill)
        c.setStrokeColor(header_fill)
        c.rect(left, y - 9*mm, usable_w, 9*mm, stroke=0, fill=1)

        c.setFillColor(colors.Color(0.2, 0.2, 0.2))
        c.setFont("NanumGothic-Bold", 10.5)

        col_label_w = usable_w * 0.60
        col_score_w = usable_w * 0.20

        c.drawRightString(left + col_label_w + col_score_w - 3*mm, y - 9*mm + 2.5*mm, "점수")
        c.drawRightString(W - right - 3*mm, y - 9*mm + 2.5*mm, "class 평균")

        c.setStrokeColor(line_col)
        c.setLineWidth(1)
        c.line(left, y - 9*mm, W - right, y - 9*mm)
        y -= 9 * mm

        for r in rows:
            label = str(r["label"])
            sv = fmt_num(r["student"])
            av = fmt_num(r["avg"])

            c.setFont("NanumGothic", 10.5)
            c.setFillColor(title_col)
            c.drawString(left + 2*mm, y - 8*mm + 2.5*mm, label)

            c.drawRightString(left + col_label_w + col_score_w - 3*mm, y - 8*mm + 2.5*mm, sv)
            c.setFillColor(muted)
            c.drawRightString(W - right - 3*mm, y - 8*mm + 2.5*mm, av)

            c.setStrokeColor(colors.Color(0.93, 0.93, 0.93))
            c.line(left, y - 8*mm, W - right, y - 8*mm)
            y -= 8 * mm

        y -= 4*mm

    draw_table("Quiz", quiz_rows)
    y -= 6*mm
    draw_table("Mocktest (점수 예상)", mock_rows)
    y -= 6*mm

    c.setFont("NanumGothic-Bold", 13)
    c.setFillColor(title_col)
    c.drawString(left, y - 5*mm, "Homework 진행도")
    y -= 10*mm

    badge_h = 12 * mm
    c.setFillColor(header_fill)
    c.setStrokeColor(header_fill)
    c.roundRect(left, y - badge_h, usable_w, badge_h, 5*mm, stroke=0, fill=1)

    c.setFont("NanumGothic-Bold", 13)
    c.setFillColor(title_col)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    c.drawString(left + 4*mm, y - badge_h + 3.5*mm, hw_txt)

    y -= (badge_h + 4*mm + 6*mm)

    c.setFont("NanumGothic-Bold", 13)
    c.setFillColor(title_col)
    c.drawString(left, y - 5*mm, "보강이 필요한 부분 및 유진쌤 Comment")
    y -= 10*mm

    box_h = max(28*mm, (len(topic_units[:12]) * 7*mm + 12*mm))
    c.setFillColor(topic_box_fill)
    c.setStrokeColor(topic_box_fill)
    c.roundRect(left, y - box_h, usable_w, box_h, 6*mm, stroke=0, fill=1)

    c.setFont("NanumGothic", 11)
    tx = left + 4*mm
    ty = y - 7*mm
    for unit in topic_units[:12]:
        c.setFillColor(title_col)
        c.drawString(tx, ty, unit)
        ty -= 7*mm

    c.setStrokeColor(line_col)
    c.setLineWidth(1)
    c.line(left, 12*mm + 6*mm, W - right, 12*mm + 6*mm)

    c.setFont("NanumGothic", 9.5)
    c.setFillColor(colors.Color(0.35, 0.35, 0.35))
    c.drawString(left, 12*mm, FOOTER_TEXT)

    c.showPage()
    c.save()
    return buf.getvalue()


# =========================================================
# ✅ Tab2 helpers: Export Mocktest에서 Mock별 (M1/M2 틀린문제) 컬럼 찾기
# =========================================================
def read_excel_any(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def detect_mock_numbers_from_columns(cols: List[str]) -> List[int]:
    nums = set()
    for c in cols:
        m = re.search(r"Mocktest\s*(\d+)", str(c), re.IGNORECASE)
        if m:
            nums.add(int(m.group(1)))
    return sorted(nums)


def find_mock_wrong_cols(df: pd.DataFrame, mock_no: int) -> Tuple[Optional[str], Optional[str]]:
    cols = [str(c).strip() for c in df.columns]

    # 다양한 형태 대응:
    # - "Mocktest1 M1 틀린문제"
    # - "Mocktest 1 - m1 틀린문제"
    # - "Mocktest1_m1 틀린문제"
    m1_rx = re.compile(rf"mocktest\s*{mock_no}.*\bm1\b.*(틀린|오답).*문제", re.IGNORECASE)
    m2_rx = re.compile(rf"mocktest\s*{mock_no}.*\bm2\b.*(틀린|오답).*문제", re.IGNORECASE)

    m1_col = None
    m2_col = None
    for c in cols:
        if m1_rx.search(c):
            m1_col = c
        if m2_rx.search(c):
            m2_col = c

    return m1_col, m2_col


def build_meta_topic_map_sub(meta_df: pd.DataFrame) -> Dict[Tuple[int, int], str]:
    """
    (module, q) -> "5.3 Radical equation and function" 형태로 저장
    """
    mp: Dict[Tuple[int, int], str] = {}
    for _, r in meta_df.iterrows():
        md = int(r["__module__"])
        q = int(r["__q__"])
        topic_raw = r["__topic_raw__"]
        mp[(md, q)] = topic_display_sub(topic_raw)
    return mp


def set_to_text_lines(items: List[str], max_lines: int = 30) -> str:
    if not items:
        return ""
    items = items[:max_lines]
    return "\n".join(items)


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title=PAGE_TITLE, layout="wide")
st.title(PAGE_TITLE)

tab1, tab2 = st.tabs(["CLASS REPORT", "틀린 유형 분석"])


# =========================================================
# Tab1: 기존 완성본 그대로 (동작 분리만)
# =========================================================
with tab1:
    st.caption("점수 엑셀 + Mock 메타 → 학생별 PNG ZIP / (편집 가능한) PDF ZIP")

    col1, col2 = st.columns([1.2, 1])
    with col1:
        uploaded_score = st.file_uploader("1) 점수 엑셀 업로드(.xlsx)", type=["xlsx"], key="score_xlsx")
    with col2:
        uploaded_mock_meta = st.file_uploader("2) Mock 메타 업로드(모듈/문항번호/단원)", type=["xlsx"], key="mock_meta_xlsx")

    threshold = st.slider("단원 정답률 기준(자동 추천용)", 0.0, 1.0, 0.70, 0.05, key="t1_threshold")

    if not uploaded_score:
        st.info("점수 엑셀을 업로드해줘.")
    else:
        df_score, name_col, class_col = load_score_excel(uploaded_score)
        quiz_cols, mock_cols, hw_cols = get_columns(df_score)
        avg_row = find_avg_row(df_score, name_col)
        students = students_list(df_score, name_col)

        default_class = ""
        if class_col:
            for s in students:
                sr = get_student_row(df_score, name_col, s)
                v = str(sr.get(class_col, "")).strip()
                if v and v.lower() != "nan":
                    default_class = v
                    break

        class_name = st.text_input("Class 이름(리포트 제목에 표시)", value=default_class or "S2 반", key="t1_class")

        pil_fonts = load_pil_fonts()
        ensure_reportlab_fonts()

        m1_wrong_col, m2_wrong_col = find_wrong_cols_in_score(df_score)
        if not m1_wrong_col or not m2_wrong_col:
            st.error("점수 엑셀에서 오답 열('M1 틀린문제', 'M2 틀린문제' 등)을 찾지 못했어요.")
        else:
            if not uploaded_mock_meta:
                st.warning("Mock 메타 파일을 올려야 단원별 정답률(자동 추천)을 계산할 수 있어요.")
            else:
                meta_df = read_mock_meta(uploaded_mock_meta)
                wrong_map = build_wrong_map_from_score(df_score, name_col, m1_wrong_col, m2_wrong_col)

                topic_acc_by_student: Dict[str, Dict[str, Tuple[int, int, float]]] = {}
                for s in students:
                    topic_acc_by_student[s] = compute_topic_accuracy(meta_df, wrong_map, s)

                topic_options = list(TOPIC_NAMES.values())

                st.divider()
                st.subheader("학생별 보강 단원 선택 (자동 추천 + 수정 가능)")

                if "units_by_student" not in st.session_state:
                    st.session_state["units_by_student"] = {s: [] for s in students}
                units_by_student = st.session_state["units_by_student"]

                if st.button("자동 추천 다시 적용(전체 학생)", key="t1_apply_all"):
                    for s in students:
                        rec = auto_recommend_topics(topic_acc_by_student.get(s, {}), threshold=threshold)
                        units_by_student[s] = rec
                    st.success("자동 추천을 다시 적용했어요.")

                for s in students:
                    c1, c2 = st.columns([1, 4])
                    with c1:
                        st.markdown(f"**{s}**")
                    with c2:
                        default_sel = units_by_student.get(s, [])
                        if not default_sel:
                            rec = auto_recommend_topics(topic_acc_by_student.get(s, {}), threshold=threshold)
                            units_by_student[s] = rec
                            default_sel = rec

                        units_by_student[s] = st.multiselect(
                            label="",
                            options=topic_options,
                            default=default_sel,
                            key=f"units_{s}",
                        )

                st.divider()

                if st.button("학생별 리포트 생성 (PNG + Editable PDF)", type="primary", key="t1_make"):
                    png_files: Dict[str, bytes] = {}
                    pdf_files: Dict[str, bytes] = {}
                    errors = []
                    preview_img = None
                    preview_student = None

                    safe_class = safe_filename(class_name)

                    for s in students:
                        try:
                            student_row = get_student_row(df_score, name_col, s)
                            use_class = class_name.strip() if class_name.strip() else "CLASS"

                            quiz_rows, mock_rows = build_rows(student_row, avg_row, quiz_cols, mock_cols)
                            hw_progress = compute_hw_progress(student_row, hw_cols)

                            selected = units_by_student.get(s, [])
                            topic_units = build_topic_units(selected)

                            img = render_student_report_image(
                                class_name=use_class,
                                student_name=s,
                                quiz_rows=quiz_rows,
                                mock_rows=mock_rows,
                                hw_progress=hw_progress,
                                topic_units=topic_units,
                                fonts=pil_fonts,
                            )

                            base = f"{safe_class}_{safe_filename(s)}"
                            png_files[f"{base}.png"] = pil_to_png_bytes(img)

                            pdf_files[f"{base}.pdf"] = create_editable_pdf_bytes(
                                class_name=use_class,
                                student_name=s,
                                quiz_rows=quiz_rows,
                                mock_rows=mock_rows,
                                hw_progress=hw_progress,
                                topic_units=topic_units,
                            )

                            if preview_img is None:
                                preview_img = img
                                preview_student = s

                        except Exception as e:
                            errors.append(f"{s}: {e}")

                    if errors:
                        st.error("일부 학생 리포트 생성 실패:\n" + "\n".join(errors))

                    if png_files:
                        png_zip = make_zip(png_files)
                        pdf_zip = make_zip(pdf_files)

                        st.success(f"완료! PNG {len(png_files)}개 / PDF {len(pdf_files)}개 생성했습니다.")

                        d1, d2 = st.columns([1, 1])
                        with d1:
                            st.download_button(
                                "📦 PNG ZIP 다운로드",
                                data=png_zip,
                                file_name=f"{safe_class}_reports_png.zip",
                                mime="application/zip",
                                key="download_png_zip",
                            )
                        with d2:
                            st.download_button(
                                "📦 PDF ZIP 다운로드 (편집 가능 텍스트)",
                                data=pdf_zip,
                                file_name=f"{safe_class}_reports_pdf.zip",
                                mime="application/zip",
                                key="download_pdf_zip",
                            )

                        if preview_img is not None:
                            st.image(preview_img, caption=f"미리보기: {preview_student}", use_container_width=True)


# =========================================================
# Tab2: 틀린 유형 분석 (소단원 기준 중복 집계)
# =========================================================
with tab2:
    st.subheader("틀린 유형 분석")
    st.caption("Export Mocktest 파일 + Mock1/2/3 메타(단원코드) → 학생별로 반복해서 틀린 소단원(2회 이상) 표시")

    left, right = st.columns([1.2, 1])
    with left:
        uploaded_export_mock = st.file_uploader("1) Export Mocktest 업로드(.xlsx)", type=["xlsx"], key="t2_export_mock")
    with right:
        st.write("2) Mock 메타 업로드(각각)")
        mock1_meta = st.file_uploader("Mock1 메타(.xlsx)", type=["xlsx"], key="t2_meta1")
        mock2_meta = st.file_uploader("Mock2 메타(.xlsx)", type=["xlsx"], key="t2_meta2")
        mock3_meta = st.file_uploader("Mock3 메타(.xlsx)", type=["xlsx"], key="t2_meta3")

    if not uploaded_export_mock:
        st.info("Export Mocktest 파일을 업로드하면 분석이 시작돼.")
    else:
        try:
            dfm = read_excel_any(uploaded_export_mock)

            name_col2 = guess_col(dfm, exact=["Name", "이름"], regexes=[r"(학생|student).*이름", r"^name$"])
            if not name_col2:
                st.error("Export Mocktest에서 Name/이름 컬럼을 찾지 못했어.")
                st.stop()

            students2 = students_list(dfm, name_col2)

            # 컬럼에서 mock 번호 추출 (Mocktest1/2/3)
            mock_numbers = detect_mock_numbers_from_columns([str(c) for c in dfm.columns])
            mock_numbers = [n for n in mock_numbers if n in [1, 2, 3]]

            if not mock_numbers:
                st.warning(
                    "Export Mocktest에서 Mocktest 번호(Mocktest1/2/3)를 컬럼명에서 찾지 못했어.\n"
                    "예) 'Mocktest1 M1 틀린문제', 'Mocktest2 M2 틀린문제' 같은 형태가 필요해."
                )
                st.stop()

            # 메타 업로드 매핑
            meta_upload_map = {1: mock1_meta, 2: mock2_meta, 3: mock3_meta}
            meta_topic_map_by_mock: Dict[int, Dict[Tuple[int, int], str]] = {}

            missing_meta = []
            for mn in mock_numbers:
                if meta_upload_map.get(mn) is None:
                    missing_meta.append(mn)
                    continue
                mdf = read_mock_meta(meta_upload_map[mn])
                meta_topic_map_by_mock[mn] = build_meta_topic_map_sub(mdf)

            if missing_meta:
                st.warning(f"⚠️ 메타 파일이 없는 Mock: {missing_meta} (해당 Mock은 집계에서 제외됨)")

            usable_mocks = [mn for mn in mock_numbers if mn in meta_topic_map_by_mock]
            if not usable_mocks:
                st.error("사용 가능한 Mock 메타가 하나도 없어. (Mock1/2/3 중 최소 1개는 메타 업로드 필요)")
                st.stop()

            # 각 mock별 오답 컬럼 찾기
            mock_wrong_cols: Dict[int, Tuple[str, str]] = {}
            missing_wrong_cols = []
            for mn in usable_mocks:
                m1c, m2c = find_mock_wrong_cols(dfm, mn)
                if not m1c or not m2c:
                    missing_wrong_cols.append(mn)
                else:
                    mock_wrong_cols[mn] = (m1c, m2c)

            if missing_wrong_cols:
                st.warning(
                    "⚠️ Export Mocktest에서 오답 컬럼을 찾지 못한 Mock이 있어.\n"
                    f"- 못 찾은 Mock: {missing_wrong_cols}\n\n"
                    "컬럼명 예시:\n"
                    "- Mocktest1 M1 틀린문제 / Mocktest1 M2 틀린문제"
                )

            usable_mocks = [mn for mn in usable_mocks if mn in mock_wrong_cols]
            if not usable_mocks:
                st.error("오답 컬럼을 찾을 수 있는 Mock이 없어. Export Mocktest 컬럼명을 확인해줘.")
                st.stop()

            # 분석 실행
            rows_out = []
            for s in students2:
                sr = get_student_row(dfm, name_col2, s)

                # subtopic -> 몇 개의 mock에서 등장했는지 카운트(중복용)
                appear_counter = Counter()

                # mock별 소단원 리스트(표시용)
                per_mock_list: Dict[int, List[str]] = {}

                for mn in usable_mocks:
                    m1c, m2c = mock_wrong_cols[mn]
                    wrong_m1 = set([n for n in parse_wrong_list(sr.get(m1c, "")) if 1 <= n <= 22])
                    wrong_m2 = set([n for n in parse_wrong_list(sr.get(m2c, "")) if 1 <= n <= 22])

                    mp = meta_topic_map_by_mock[mn]

                    subtopics_this_mock = set()
                    for q in wrong_m1:
                        t = mp.get((1, q), "")
                        if t:
                            subtopics_this_mock.add(t)
                    for q in wrong_m2:
                        t = mp.get((2, q), "")
                        if t:
                            subtopics_this_mock.add(t)

                    # 이 mock에서 나온 소단원(중복 제거) -> appear_counter +1
                    for t in subtopics_this_mock:
                        appear_counter[t] += 1

                    per_mock_list[mn] = sorted(list(subtopics_this_mock))

                repeated = [(t, c) for t, c in appear_counter.items() if c >= 2]
                repeated.sort(key=lambda x: (-x[1], x[0]))

                repeated_text = "\n".join([f"{t}  ({c}회)" for t, c in repeated]) if repeated else ""

                out = {"Name": s, "반복 소단원(2회 이상)": repeated_text}
                for mn in [1, 2, 3]:
                    if mn in usable_mocks:
                        out[f"Mock{mn} 틀린 소단원"] = set_to_text_lines(per_mock_list.get(mn, []), max_lines=30)
                    else:
                        out[f"Mock{mn} 틀린 소단원"] = ""

                rows_out.append(out)

            result_df = pd.DataFrame(rows_out)

            st.success(f"완료! 학생 {len(result_df)}명 / 분석 Mock: {usable_mocks}")
            st.dataframe(result_df, use_container_width=True, height=520)

            # 다운로드(CSV)
            csv_bytes = result_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
            st.download_button(
                "⬇️ 결과 CSV 다운로드",
                data=csv_bytes,
                file_name="틀린유형분석_소단원중복.csv",
                mime="text/csv",
                key="t2_download_csv",
            )

        except Exception as e:
            st.error(f"Tab2 오류: {e}")
            st.exception(e)
