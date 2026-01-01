import io
import os
import re
import zipfile
import pandas as pd
import streamlit as st

from PIL import Image, ImageDraw, ImageFont


# =========================================================
# 고정 머릿말/꼬릿말
# =========================================================
HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

UNIT_OPTIONS = [
    "I. Linear",
    "IV. Quadratic",
    "V. Exponential",
    "VI. Polynomials, radical and rational functions",
    "VII. Geometry",
    "VIII. Statistics",
]


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
# 데이터 로드 (EXPORT 템플릿 전용)
# =========================================================
def load_export_excel(uploaded_file) -> pd.DataFrame:
    df = pd.read_excel(uploaded_file, engine="openpyxl")
    # 컬럼명 공백 정리
    df.columns = [str(c).strip() for c in df.columns]

    # Name/이름 지원
    name_col = None
    for cand in ["Name", "이름"]:
        if cand in df.columns:
            name_col = cand
            break
    if not name_col:
        raise KeyError("업로드 파일에서 'Name' (또는 '이름') 컬럼을 찾지 못했습니다. EXPORT 시트를 업로드했는지 확인해주세요.")

    # Class/반 지원 (없어도 됨)
    class_col = None
    for cand in ["Class", "반", "클래스"]:
        if cand in df.columns:
            class_col = cand
            break

    # 문자열 정리
    df[name_col] = df[name_col].astype(str).str.strip()
    if class_col:
        df[class_col] = df[class_col].astype(str).str.strip()

    return df, name_col, class_col


def get_columns(df: pd.DataFrame):
    # Quiz: Quiz1.., Quiz2.., ReviewQuiz...
    quiz_cols = [c for c in df.columns if re.match(r"^(Quiz\d+|QUIZ\d+|ReviewQuiz)", str(c), re.IGNORECASE)]
    # Mock: Mocktest1.. Mocktest2..
    mock_cols = [c for c in df.columns if re.match(r"^Mocktest\d+", str(c), re.IGNORECASE)]
    # Homework: Homework1.. (전부)
    hw_cols = [c for c in df.columns if re.match(r"^Homework\d+", str(c), re.IGNORECASE)]

    # 정렬
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
        raise ValueError("평균행(Name='평균')을 찾지 못했습니다. (EXPORT 2행이 평균이도록 만든 파일을 업로드했는지 확인)")
    # 첫 번째 평균행 사용
    return df.loc[mask].iloc[0]


def students_list(df: pd.DataFrame, name_col: str):
    names = df[name_col].dropna().astype(str).str.strip()
    names = [n for n in names.tolist() if n not in ["", "nan", "평균"]]
    # 중복 제거(순서 유지)
    seen = set()
    out = []
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


# =========================================================
# PNG 렌더링 (세로 자동 크롭)
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


def compute_hw_progress(student_row: pd.Series, hw_cols: list[str]):
    if not hw_cols:
        return None
    vals = pd.to_numeric(student_row[hw_cols], errors="coerce").dropna()
    if len(vals) == 0:
        return None
    avg = float(vals.mean())
    # 0~1이면 %로
    if avg <= 1.0:
        avg *= 100.0
    return avg


def build_rows(student_row: pd.Series, avg_row: pd.Series, quiz_cols, mock_cols):
    quiz_rows = [{"label": c, "student": student_row.get(c), "avg": avg_row.get(c)} for c in quiz_cols]
    mock_rows = [{"label": c, "student": student_row.get(c), "avg": avg_row.get(c)} for c in mock_cols]
    return quiz_rows, mock_rows


def render_student_report_image(class_name, student_name, quiz_rows, mock_rows, hw_progress, units, fonts):
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

    unit_txt = ", ".join(units) if units else "선택 없음"
    unit_lines = wrap_text(ddraw, unit_txt, fonts["small"], max_width=w - 24)
    lines_h = len(unit_lines) * 24
    unit_box_h = max(110, 14 + lines_h + 14)

    content_h = (
        header_h
        + (y_title - header_h)
        + th + 6
        + h_quiz + GAP
        + h_mock + GAP
        + h_hw
        + 30
        + unit_box_h
    )

    footer_h = 42
    bottom_pad = 10
    H = int(content_h + footer_h + bottom_pad)

    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # Header
    draw_text(draw, margin, 10, HEADER_TEXT, fonts["small_b"], fill="#111111")
    draw_line(draw, margin, 38, W - margin, 38, color="#D9D9D9", w=2)

    # Title
    y = y_title
    if len(title_lines) == 1:
        draw_text(draw, margin, y, title_lines[0], fonts["title"], fill="#111111")
        y += 48
    else:
        draw_text(draw, margin, y, title_lines[0], fonts["title"], fill="#111111")
        draw_text(draw, margin, y + 38, title_lines[1], fonts["title"], fill="#111111")
        y += 88
    y += 6

    # Quiz -> Mocktest
    y = render_table(draw, margin, y, w, "Quiz", quiz_rows, fonts, row_h=ROW_H)
    y += GAP
    y = render_table(draw, margin, y, w, "Mocktest (점수 예상)", mock_rows, fonts, row_h=ROW_H)
    y += GAP

    # Homework 진행도 (웹앱에서만 계산)
    draw_text(draw, margin, y, "Homework 진행도", fonts["h2"], fill="#111111")
    y += 30
    badge_h = 44
    draw.rounded_rectangle([margin, y, margin + w, y + badge_h],
                           radius=18, fill="#F5F6F8", outline=None)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    draw_text(draw, margin + 14, y + 10, hw_txt, fonts["b"], fill="#111111")
    y += badge_h + 14

    # Units
    draw_text(draw, margin, y, "보강필요한 부분", fonts["h2"], fill="#111111")
    y += 30

    draw.rounded_rectangle([margin, y, W - margin, y + unit_box_h],
                           radius=20, fill="#F9FAFB", outline=None)

    yy = y + 14
    for line in unit_lines[:12]:
        draw_text(draw, margin + 12, yy, line, fonts["small"], fill="#111111")
        yy += 24

    # Footer
    footer_y_line = H - 42
    draw_line(draw, margin, footer_y_line, W - margin, footer_y_line, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 30, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    return img


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="유진 SAT CLASS REPORT", layout="wide")
st.title("유진 SAT CLASS REPORT")
st.caption("EXPORT 템플릿( Class / Name / Quiz… / Mocktest… / Homework… + 평균행 ) 업로드 → 학생별 PNG 생성 → ZIP 다운로드")

uploaded = st.file_uploader("엑셀 업로드(.xlsx) - EXPORT 파일", type=["xlsx"])
if not uploaded:
    st.info("EXPORT 시트를 업로드하면 학생 목록이 나타납니다.")
    st.stop()

try:
    df, name_col, class_col = load_export_excel(uploaded)
except Exception as e:
    st.error(f"엑셀 로드 실패: {e}")
    st.stop()

quiz_cols, mock_cols, hw_cols = get_columns(df)

if not quiz_cols and not mock_cols and not hw_cols:
    st.error("Quiz/Mocktest/Homework 컬럼을 찾지 못했습니다. (예: Quiz1, Mocktest1, Homework1...)")
    st.stop()

try:
    avg_row = find_avg_row(df, name_col)
except Exception as e:
    st.error(f"평균행 찾기 실패: {e}")
    st.stop()

students = students_list(df, name_col)
if not students:
    st.error("학생 이름을 찾지 못했습니다. Name(또는 이름) 열을 확인해주세요.")
    st.stop()

# Class 이름: 입력값 우선, 없으면 파일에서 첫 학생 행의 Class 사용
default_class = ""
if class_col:
    for s in students:
        sr = get_student_row(df, name_col, s)
        v = str(sr.get(class_col, "")).strip()
        if v and v.lower() != "nan":
            default_class = v
            break

class_name = st.text_input("Class 이름(리포트에 표시)", value=default_class or "S2 반")

try:
    fonts = load_fonts()
except Exception as e:
    st.error(f"폰트 로드 실패: {e}")
    st.stop()

st.subheader("학생별 보강 단원 선택 (한 페이지에서 전원 설정)")

if "units_by_student" not in st.session_state:
    st.session_state["units_by_student"] = {s: [] for s in students}

units_by_student = st.session_state["units_by_student"]

# 동기화
for s in students:
    units_by_student.setdefault(s, [])
for s in list(units_by_student.keys()):
    if s not in students:
        units_by_student.pop(s, None)

for s in students:
    c1, c2 = st.columns([1, 4])
    with c1:
        st.markdown(f"**{s}**")
    with c2:
        units_by_student[s] = st.multiselect(
            label="",
            options=UNIT_OPTIONS,
            default=units_by_student[s],
            key=f"units_{s}",
        )

st.divider()

if st.button("학생별 PNG 생성 → ZIP 만들기"):
    png_files = {}
    errors = []
    preview_img = None
    preview_student = None

    for s in students:
        try:
            student_row = get_student_row(df, name_col, s)
            # 파일 Class 열이 있으면 개별 반 이름 사용 가능(하지만 지금은 입력값을 기본으로 씀)
            use_class = class_name.strip() if class_name.strip() else (str(student_row.get(class_col, "")).strip() if class_col else "")

            quiz_rows, mock_rows = build_rows(student_row, avg_row, quiz_cols, mock_cols)
            hw_progress = compute_hw_progress(student_row, hw_cols)

            img = render_student_report_image(
                class_name=use_class or "CLASS",
                student_name=s,
                quiz_rows=quiz_rows,
                mock_rows=mock_rows,
                hw_progress=hw_progress,
                units=units_by_student.get(s, []),
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

        # ✅ ZIP 다운로드 버튼을 미리보기보다 먼저
        st.download_button(
            "ZIP 다운로드 (학생별 PNG)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

        st.success(f"완료! 총 {len(png_files)}명의 PNG를 ZIP으로 만들었습니다.")

        if preview_img is not None:
            st.image(preview_img, caption=f"미리보기: {preview_student}", use_container_width=True)
