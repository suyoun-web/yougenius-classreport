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

    # 화면에 잘 보이도록 좀 크게
    return {
        "title": f(bold_path, 32),
        "h2": f(bold_path, 19),
        "b": f(bold_path, 17),
        "r": f(reg_path, 16),
        "small_b": f(bold_path, 14),
        "small": f(reg_path, 14),
        "tiny": f(reg_path, 12),
    }


# =========================================================
# 엑셀 파싱 (서브헤더 결합 + '이름' 중복 문제 방지)
# =========================================================
def make_unique(colnames):
    seen = {}
    out = []
    for c in colnames:
        if c in seen:
            seen[c] += 1
            out.append(f"{c}__dup{seen[c]}")
        else:
            seen[c] = 1
            out.append(c)
    return out


def load_and_clean(uploaded_file) -> pd.DataFrame:
    raw = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    sub = raw.iloc[0]
    df = raw.iloc[1:].copy()
    cols = list(raw.columns)

    idx_name = None
    for i, c in enumerate(cols):
        if str(c).strip() == "이름":
            idx_name = i
            break

    meta_map = {}
    if idx_name is not None:
        meta_positions = [idx_name - 2, idx_name - 1, idx_name, idx_name + 1, idx_name + 2]
        meta_names = ["레벨", "학교", "이름", "연락처", "연락(이메일/카톡)"]
        for pos, std in zip(meta_positions, meta_names):
            if 0 <= pos < len(cols):
                meta_map[pos] = std

    new_cols = []
    last_main = None

    for i, c in enumerate(cols):
        if i in meta_map:
            new_cols.append(meta_map[i])
            last_main = None
            continue

        main = c
        if isinstance(c, str) and c.startswith("Unnamed"):
            main = last_main if last_main is not None else c
        else:
            last_main = c

        sh = sub[c]
        if pd.isna(sh) or str(sh).strip() == "":
            new_cols.append(str(main).strip())
        else:
            new_cols.append(f"{str(main).strip()}__{str(sh).strip()}")

    df.columns = make_unique(new_cols)
    df = df.reset_index(drop=True)

    for c in df.columns:
        if any(k in str(c) for k in ["__점수", "__Total", "__점수 예상", "Homework"]):
            df[c] = pd.to_numeric(df[c], errors="coerce")

    if "이름" not in df.columns:
        raise KeyError("엑셀에서 '이름' 컬럼을 찾지 못했습니다. C열 헤더가 정확히 '이름'인지 확인해주세요.")

    return df


# =========================================================
# 점수 컬럼 탐지
# =========================================================
def quiz_score_cols(df):
    return [c for c in df.columns if re.match(r"^(QUIZ\d+.*|ReviewQuiz.*)__점수$", str(c))]


def mock_pred_cols(df):
    return [c for c in df.columns if re.match(r"^MOCK TEST.*__점수 예상$", str(c))]


def homework_cols(df):
    return [c for c in df.columns if str(c).startswith("Homework")]


def pretty(label: str) -> str:
    return re.sub(r"\s*\(.*?\)\s*", "", label).strip()


# =========================================================
# 평균행(1개) 찾기
# =========================================================
def find_class_avg_row(df: pd.DataFrame, score_cols: list[str]) -> int:
    best_idx = None
    best_count = -1

    for i in range(len(df)):
        name = df.loc[i, "이름"]
        if isinstance(name, str) and name.strip() != "":
            continue

        cnt = sum(pd.notna(df.loc[i, c]) for c in score_cols)
        if cnt > best_count:
            best_count = cnt
            best_idx = i

    if best_idx is None or best_count <= 0:
        raise ValueError("평균행을 찾지 못했습니다. (이름이 비어있고 점수 칼럼에 숫자가 있는 평균행이 필요)")
    return best_idx


# =========================================================
# 학생 1명 데이터 추출
# =========================================================
def build_onepage_rows(df: pd.DataFrame, student_name: str):
    qcols = quiz_score_cols(df)
    mcols = mock_pred_cols(df)
    hcols = homework_cols(df)
    score_cols = list(dict.fromkeys(qcols + mcols))

    avg_i = find_class_avg_row(df, score_cols)
    avg_row = df.loc[avg_i]

    s_idx = df.index[df["이름"] == student_name].tolist()
    if not s_idx:
        raise ValueError(f"학생을 찾을 수 없음: {student_name}")
    s_row = df.loc[s_idx[0]]

    quiz_rows = []
    for c in qcols:
        main = c.split("__")[0]
        quiz_rows.append({"label": pretty(main).replace("QUIZ", "Quiz"), "student": s_row[c], "avg": avg_row[c]})

    mock_rows = []
    for c in mcols:
        main = c.split("__")[0]
        label = pretty(main).replace("MOCK TEST", "Mocktest")
        mock_rows.append({"label": label, "student": s_row[c], "avg": avg_row[c]})

    hw_progress = None
    if hcols:
        vals = s_row[hcols].dropna()
        if len(vals) > 0:
            hw_progress = float(vals.mean())
            if hw_progress <= 1.0:
                hw_progress *= 100.0

    return quiz_rows, mock_rows, hw_progress


# =========================================================
# PNG 렌더링 (PIL) - "세로 자동 크롭" 버전
# =========================================================
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
    words = text.split(" ")
    lines = []
    cur = ""
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
    draw.rounded_rectangle([margin, y, margin + w, y + badge_h],
                           radius=18, fill="#F5F6F8", outline=None)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    draw_text(draw, margin + 14, y + 10, hw_txt, fonts["b"], fill="#111111")
    y += badge_h + 14

    draw_text(draw, margin, y, "보강필요한 부분", fonts["h2"], fill="#111111")
    y += 30

    draw.rounded_rectangle([margin, y, W - margin, y + unit_box_h],
                           radius=20, fill="#F9FAFB", outline=None)

    yy = y + 14
    for line in unit_lines[:12]:
        draw_text(draw, margin + 12, yy, line, fonts["small"], fill="#111111")
        yy += 24

    footer_y_line = H - 42
    draw_line(draw, margin, footer_y_line, W - margin, footer_y_line, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 30, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    return img


def pil_to_png_bytes(img: Image.Image) -> bytes:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


def safe_filename(name: str) -> str:
    name = str(name).strip()
    name = re.sub(r'[\\/:*?"<>|]+', "_", name)
    name = re.sub(r"\s+", " ", name).strip()
    return name if name else "학생"


def make_zip_of_pngs(png_dict: dict) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        for fname, data in png_dict.items():
            zf.writestr(fname, data)
    return bio.getvalue()


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="유진 SAT CLASS REPORT", layout="wide")
st.title("유진 SAT CLASS REPORT")
st.caption("✅ 고정 높이 없이, 보강필요한 부분까지 출력한 뒤 거기서 딱 잘라 PNG로 저장합니다.")

class_name = st.text_input("Class 이름(리포트에 표시)", value="S2 개념반")

uploaded = st.file_uploader("엑셀 업로드(.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("엑셀 파일을 업로드하면 학생 목록이 나타납니다.")
    st.stop()

try:
    df = load_and_clean(uploaded)
except Exception as e:
    st.error(f"엑셀 인식 실패: {e}")
    st.stop()

students = sorted([s for s in df["이름"].dropna().unique().tolist() if str(s).strip() != ""])
if not students:
    st.error("학생 이름을 찾지 못했습니다. 엑셀에서 C열(이름)에 학생 이름이 들어있는지 확인해주세요.")
    st.stop()

try:
    fonts = load_fonts()
except Exception as e:
    st.error(f"폰트 로드 실패: {e}")
    st.stop()

st.subheader("학생별 보강 단원 선택 (한 페이지에서 전원 설정)")

if "units_by_student" not in st.session_state:
    st.session_state["units_by_student"] = {s: [] for s in students}

units_by_student = st.session_state["units_by_student"]

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
    preview_student = None
    preview_img = None

    for s in students:
        try:
            quiz_rows, mock_rows, hw_progress = build_onepage_rows(df, s)
            img = render_student_report_image(
                class_name=class_name,
                student_name=s,
                quiz_rows=quiz_rows,
                mock_rows=mock_rows,
                hw_progress=hw_progress,
                units=units_by_student.get(s, []),
                fonts=fonts,
            )
            png_files[f"{safe_filename(s)}.png"] = pil_to_png_bytes(img)

            if preview_img is None:
                preview_student = s
                preview_img = img

        except Exception as e:
            errors.append(f"{s}: {e}")

    if errors:
        st.error("일부 학생 리포트 생성 실패:\n" + "\n".join(errors))

    if png_files:
        zip_bytes = make_zip_of_pngs(png_files)

        zip_name = f"{safe_filename(class_name)}_reports_autoCrop.zip"
        st.download_button(
            "ZIP 다운로드 (학생별 PNG)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

        st.success(f"완료! 총 {len(png_files)}명의 PNG를 ZIP으로 만들었습니다.")

        if preview_img is not None:
            st.image(preview_img, caption=f"미리보기: {preview_student}", use_container_width=True)
