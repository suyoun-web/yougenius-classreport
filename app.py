import io
import os
import re
import zipfile
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from PIL import Image, ImageDraw, ImageFont


# =========================================================
# 고정 머릿말/꼬릿말
# =========================================================
HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

PAGE_TITLE = "유진 sat class report"


# =========================================================
# 단원명(major 1~7)
# - Mock 메타의 "단원" 값이 "5.3"처럼 앞에 1~7 숫자가 붙어있을 때 가장 정확
# =========================================================
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
# 폰트 로드 (PIL용) - fonts/ 폴더
# =========================================================
@st.cache_resource
def load_fonts():
    reg_path = "fonts/NanumGothic-Regular.ttf"
    bold_path = "fonts/NanumGothic-Bold.ttf"

    if not os.path.exists(reg_path) or not os.path.exists(bold_path):
        raise FileNotFoundError(
            "폰트 파일을 찾지 못했습니다.\n\n"
            "필요 파일:\n"
            f"- {reg_path}\n"
            f"- {bold_path}\n\n"
            "GitHub 레포에 fonts 폴더를 만들고 폰트 파일을 올려주세요."
        )

    def f(path, size):
        return ImageFont.truetype(path, size=size)

    return {
        "title": f(bold_path, 32),
        "h2": f(bold_path, 19),
        "b": f(bold_path, 17),
        "small_b": f(bold_path, 14),
        "small": f(reg_path, 14),
        "tiny": f(reg_path, 12),
    }


# =========================================================
# 공통 유틸
# =========================================================
def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else "학생"


def pil_to_png_bytes(img: Image.Image) -> bytes:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


def make_zip_of_pngs(png_dict: dict) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in png_dict.items():
            zf.writestr(fname, data)
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


def major_topic_id(topic_str: str) -> Optional[int]:
    s = str(topic_str).strip()
    if s == "" or s.lower() == "nan":
        return None
    m = re.match(r"^\s*(\d+)", s)
    if not m:
        return None
    v = int(m.group(1))
    return v if 1 <= v <= 7 else None


def parse_wrong_list(val) -> List[int]:
    """
    셀 값이
      - "9,19,22"
      - "22.0" (엑셀이 숫자로 저장한 경우)
      - "19 22" / "19;22"
    이런 경우를 최대한 안정적으로 처리.
    """
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return []

    # 숫자(22.0)로 들어오는 경우: 22만
    if isinstance(val, (int, float)) and not pd.isna(val):
        n = int(round(float(val)))
        return [n] if n > 0 else []

    s = str(val).strip()
    if s == "" or s.upper() in ["X", "Х", "-"]:
        return []

    # "22.0" 같은 단일 숫자 문자열이면 22만
    if re.fullmatch(r"\d+(\.0+)?", s):
        return [int(float(s))]

    s = s.replace("，", ",").replace(";", ",").replace("/", ",")
    s = re.sub(r"\s+", ",", s)
    nums = re.findall(r"\d+", s)
    out = [int(x) for x in nums]

    # 혹시 "22.0"이 "22","0"으로 분해되는 케이스 방어: 뒤에 0만 달린 패턴 제거
    # (이미 단일 숫자 문자열은 위에서 잡음. 그래도 혹시 섞이면 0은 제거)
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
# 점수 엑셀 로드 (Sheet1)
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


def get_columns(df: pd.DataFrame):
    quiz_cols = [c for c in df.columns if re.match(r"^(Quiz\d+|QUIZ\d+|ReviewQuiz)", str(c), re.IGNORECASE)]
    mock_cols = [c for c in df.columns if re.match(r"^Mocktest\d+", str(c), re.IGNORECASE)]
    hw_cols = [c for c in df.columns if re.match(r"^Homework\d+", str(c), re.IGNORECASE)]

    def num_key(col):
        m = re.search(r"(\d+)", str(col))
        return int(m.group(1)) if m else 9999

    quiz_cols = sorted(quiz_cols, key=num_key)
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


def detect_latest_mocktest_number(mock_cols: List[str]) -> Optional[int]:
    nums = []
    for c in mock_cols:
        m = re.match(r"^Mocktest(\d+)$", str(c), re.IGNORECASE)
        if m:
            nums.append(int(m.group(1)))
    return max(nums) if nums else None


def find_wrong_cols_in_score(df: pd.DataFrame) -> Tuple[Optional[str], Optional[str]]:
    # 예시 파일: "M1  틀린문제", "M2 틀린문제" (공백 변형 많아서 regex로)
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

        out[nm] = {1: set([n for n in m1_nums if 1 <= n <= 22]),
                   2: set([n for n in m2_nums if 1 <= n <= 22])}
    return out


# =========================================================
# Mock 메타 로드 (모듈/문항번호/단원/난이도)
# =========================================================
def read_mock_meta(mock_file) -> pd.DataFrame:
    df = pd.read_excel(mock_file, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_module = guess_col(df, exact=["모듈", "Module"], regexes=[r"모듈", r"module"])
    col_q = guess_col(df, exact=["문항번호", "문항", "문제번호", "No", "Q"], regexes=[r"(문항|문제).*(번호)?", r"q\s*no", r"^no$"])
    col_topic = guess_col(df, exact=["단원", "Topic"], regexes=[r"단원", r"topic"])
    col_diff = guess_col(df, exact=["난이도", "Difficulty"], regexes=[r"난이도", r"difficulty"])

    if not col_module or not col_q or not col_topic:
        raise ValueError("Mock 메타 파일에는 최소 '모듈', '문항번호', '단원' 컬럼이 필요합니다.")

    out = df.copy()
    out["__module__"] = out[col_module].apply(norm_module)
    out["__q__"] = pd.to_numeric(out[col_q], errors="coerce")
    out = out.dropna(subset=["__module__", "__q__"]).copy()
    out["__module__"] = out["__module__"].astype(int)
    out["__q__"] = out["__q__"].astype(int)

    # 문항번호는 1~22로 가정 (모듈별)
    out = out[(out["__q__"] >= 1) & (out["__q__"] <= 22)].copy()

    out["__topic_raw__"] = out[col_topic].astype(str).str.strip()
    out["__major__"] = out["__topic_raw__"].apply(major_topic_id)

    if col_diff:
        out["__diff__"] = out[col_diff].astype(str).str.strip().str.upper()
    else:
        out["__diff__"] = ""

    out = out[["__module__", "__q__", "__topic_raw__", "__major__", "__diff__"]].drop_duplicates(
        subset=["__module__", "__q__"]
    )
    return out


def compute_topic_accuracy(
    meta_df: pd.DataFrame,
    wrong_map: Dict[str, Dict[int, set]],
    student: str,
) -> Dict[str, Tuple[int, int, float]]:
    if student not in wrong_map:
        wrong_map[student] = {1: set(), 2: set()}

    totals: Dict[int, int] = {k: 0 for k in range(1, 8)}
    wrongs: Dict[int, int] = {k: 0 for k in range(1, 8)}

    for _, r in meta_df.iterrows():
        md = int(r["__module__"])
        q = int(r["__q__"])
        major = r["__major__"]
        if pd.isna(major) or major is None:
            continue
        major = int(major)
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


def build_topic_display_lines(
    selected_topics: List[str],
    topic_acc: Optional[Dict[str, Tuple[int, int, float]]],
    threshold: float,
) -> List[Tuple[str, bool]]:
    out: List[Tuple[str, bool]] = []
    if not selected_topics:
        return [("선택 없음", False)]

    for tname in selected_topics:
        if topic_acc and tname in topic_acc:
            c, tot, acc = topic_acc[tname]
            pct = int(round(acc * 100)) if tot > 0 else 0
            line = f"{tname}  {pct}% ({c}/{tot})"
            is_low = (tot > 0 and acc <= threshold)
            out.append((line, is_low))
        else:
            out.append((tname, False))

    return out


# =========================================================
# Homework 진행도 계산
# =========================================================
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


# =========================================================
# PNG 렌더링 (PIL)
# =========================================================
def draw_line(draw, x1, y1, x2, y2, color="#D9D9D9", w=2):
    draw.line((x1, y1, x2, y2), fill=color, width=w)


def draw_text(draw, x, y, text, font, fill="#111111"):
    draw.text((x, y), text, font=font, fill=fill)


def right_text(draw, rx, y, text, font, fill="#111111"):
    tw = draw.textlength(text, font=font)
    draw.text((rx - tw, y), text, font=font, fill=fill)


def wrap_text(draw, text, font, max_width):
    words = str(text).split(" ")
    lines, cur = [], ""
    for w in words:
        test = (cur + " " + w).strip()
        if draw.textlength(test, font=font) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


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
    topic_lines: List[Tuple[str, bool]],
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
    max_lines = min(12, len(topic_lines))
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

    draw_text(draw, margin, y, "보강필요한 부분 (마지막 Mocktest 기준)", fonts["h2"], fill="#111111")
    y += 30

    draw.rounded_rectangle([margin, y, W - margin, y + topic_box_h], radius=20, fill="#F9FAFB", outline=None)

    yy = y + 14
    red = "#DC2626"
    black = "#111111"
    for (line, is_low) in topic_lines[:12]:
        draw_text(draw, margin + 12, yy, line, fonts["small"], fill=(red if is_low else black))
        yy += 24

    footer_y_line = H - 42
    draw_line(draw, margin, footer_y_line, W - margin, footer_y_line, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 30, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    return img


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title=PAGE_TITLE, layout="wide")
st.title(PAGE_TITLE)
st.caption("점수 엑셀(오른쪽에 M1/M2 틀린문제 포함) + Mock 메타(단원/난이도) → 마지막 Mocktest 기준 자동 보강단원 → 학생별 PNG ZIP")

col1, col2 = st.columns([1.2, 1])

with col1:
    uploaded_score = st.file_uploader("1) 점수 엑셀 업로드(.xlsx)", type=["xlsx"], key="score_xlsx")

with col2:
    uploaded_mock_meta = st.file_uploader("2) Mock 메타 업로드(모듈/문항번호/단원/난이도)", type=["xlsx"], key="mock_meta_xlsx")

threshold = st.slider("단원 정답률 기준(이하면 빨간색)", 0.0, 1.0, 0.70, 0.05)

if not uploaded_score:
    st.info("점수 엑셀을 업로드해줘.")
    st.stop()

try:
    df_score, name_col, class_col = load_score_excel(uploaded_score)
except Exception as e:
    st.error(f"점수 엑셀 로드 실패: {e}")
    st.stop()

quiz_cols, mock_cols, hw_cols = get_columns(df_score)
latest_mock_num = detect_latest_mocktest_number(mock_cols)

try:
    avg_row = find_avg_row(df_score, name_col)
except Exception as e:
    st.error(f"평균행 찾기 실패: {e}")
    st.stop()

students = students_list(df_score, name_col)
if not students:
    st.error("학생 이름을 찾지 못했습니다.")
    st.stop()

default_class = ""
if class_col:
    for s in students:
        sr = get_student_row(df_score, name_col, s)
        v = str(sr.get(class_col, "")).strip()
        if v and v.lower() != "nan":
            default_class = v
            break

class_name = st.text_input("Class 이름(리포트 제목에 표시)", value=default_class or "S2 반")

try:
    fonts = load_fonts()
except Exception as e:
    st.error(f"폰트 로드 실패: {e}")
    st.stop()

m1_wrong_col, m2_wrong_col = find_wrong_cols_in_score(df_score)

st.write("---")
st.subheader("마지막 Mocktest 기준 설정")

if latest_mock_num is not None:
    st.success(f"점수 열 기준으로 마지막 Mocktest는 **Mocktest{latest_mock_num}** 로 인식했어요.")
else:
    st.warning("Mocktest1/2/3 같은 점수 열을 못 찾았어요. (그래도 오답 열이 있으면 보강단원은 계산됩니다.)")

if not m1_wrong_col or not m2_wrong_col:
    st.error(
        "점수 엑셀에서 오답 열을 찾지 못했어요.\n\n"
        "필요한 열 이름 예시:\n"
        "- 'M1 틀린문제'\n"
        "- 'M2 틀린문제'\n\n"
        "지금 엑셀의 해당 열 이름을 위 예시처럼 맞춰줘."
    )
    st.stop()

st.info(f"오답 열 인식: **{m1_wrong_col}**, **{m2_wrong_col}**")

if not uploaded_mock_meta:
    st.warning("Mock 메타 파일을 올려야 단원별 정답률(보강단원 자동추천)을 계산할 수 있어요.")
    st.stop()

try:
    meta_df = read_mock_meta(uploaded_mock_meta)
except Exception as e:
    st.error(f"Mock 메타 로드 실패: {e}")
    st.stop()

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

# 학생 목록 변화 대응
for s in students:
    units_by_student.setdefault(s, [])
for s in list(units_by_student.keys()):
    if s not in students:
        units_by_student.pop(s, None)

colA, colB = st.columns([1, 1])
with colA:
    if st.button("자동 추천 다시 적용(전체 학생)"):
        for s in students:
            rec = auto_recommend_topics(topic_acc_by_student.get(s, {}), threshold=threshold)
            units_by_student[s] = rec
        st.success("자동 추천을 다시 적용했어요.")

with colB:
    st.caption("빨간색 기준: 정답률 ≤ 기준값")

for s in students:
    c1, c2 = st.columns([1, 4])
    with c1:
        st.markdown(f"**{s}**")
    with c2:
        default_sel = units_by_student.get(s, [])

        # 기본값이 비어있으면 자동 추천을 1회 채워줌(학생별)
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

if st.button("학생별 PNG 생성 → ZIP 만들기", type="primary"):
    png_files = {}
    errors = []
    preview_img = None
    preview_student = None

    for s in students:
        try:
            student_row = get_student_row(df_score, name_col, s)
            use_class = class_name.strip() if class_name.strip() else "CLASS"

            quiz_rows, mock_rows = build_rows(student_row, avg_row, quiz_cols, mock_cols)
            hw_progress = compute_hw_progress(student_row, hw_cols)

            topic_acc = topic_acc_by_student.get(s, {})
            selected = units_by_student.get(s, [])
            topic_lines = build_topic_display_lines(selected, topic_acc, threshold=threshold)

            img = render_student_report_image(
                class_name=use_class,
                student_name=s,
                quiz_rows=quiz_rows,
                mock_rows=mock_rows,
                hw_progress=hw_progress,
                topic_lines=topic_lines,
                fonts=fonts,
            )

            png_files[f"{safe_filename(s)}.png"] = pil_to_png_bytes(img)

            if preview_img is None:
                preview_img = img
                preview_student = s

        except Exception as e:
            errors.append(f"{s}: {e}")

    if errors:
        st.error("일부 학생 리포트 생성 실패:\n" + "\n".join(errors))

    if png_files:
        zip_bytes = make_zip_of_pngs(png_files)
        zip_name = f"{safe_filename(class_name)}_reports.zip"

        st.download_button(
            "ZIP 다운로드 (학생별 PNG)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

        st.success(f"완료! 총 {len(png_files)}명의 PNG를 ZIP으로 만들었습니다.")

        if preview_img is not None:
            st.image(preview_img, caption=f"미리보기: {preview_student}", use_container_width=True)
