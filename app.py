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

    # ✅ 폰트 키움 (A5지만 화면 꽉차게 쓰기)
    return {
        "title": f(bold_path, 30),
        "h2": f(bold_path, 18),
        "b": f(bold_path, 16),
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
# PNG 렌더링 (PIL) - A5 / 여백 최소 / 1컬럼 (Quiz→Mock→HW→Units)
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


def draw_wrapped_title(draw, x, y, max_w, class_name, student_name, fonts):
    one_line = f"{class_name} {student_name} CLASS REPORT"
    if draw.textlength(one_line, font=fonts["title"]) <= max_w:
        draw_text(draw, x, y, one_line, fonts["title"], fill="#111111")
        return y + 46
    else:
        line1 = f"{class_name} {student_name}"
        line2 = "CLASS REPORT"
        draw_text(draw, x, y, line1, fonts["title"], fill="#111111")
        draw_text(draw, x, y + 36, line2, fonts["title"], fill="#111111")
        return y + 82


def render_table_full(draw, x, y, w, title, rows, fonts):
    # ✅ 행높이 키움
    draw_text(draw, x, y, title, fonts["h2"], fill="#111111")
    y += 28

    row_h = 28
    col1 = int(w * 0.58)
    col2 = int(w * 0.20)

    draw.rectangle([x, y, x + w, y + row_h], fill="#F5F6F8", outline=None)
    right_text(draw, x + col1 + col2 - 10, y + 6, "점수", fonts["small_b"], fill="#333333")
    right_text(draw, x + w - 10, y + 6, "class 평균", fonts["small_b"], fill="#333333")
    draw_line(draw, x, y + row_h, x + w, y + row_h, color="#E1E4E8", w=2)
    y += row_h

    for r in rows:
        label = str(r["label"])
        sv = fmt_num(r["student"])
        av = fmt_num(r["avg"])

        draw_text(draw, x + 8, y + 6, label, fonts["small"], fill="#111111")
        right_text(draw, x + col1 + col2 - 10, y + 6, sv, fonts["small"], fill="#111111")
        right_text(draw, x + w - 10, y + 6, av, fonts["small"], fill="#666666")

        draw_line(draw, x, y + row_h, x + w, y + row_h, color="#EDEFF2", w=2)
        y += row_h

    return y


def render_student_report_image(class_name, student_name, quiz_rows, mock_rows, hw_progress, units, fonts):
    # A5 비율
    W, H = 877, 1240

    # ✅ 여백 줄이기
    margin = 26

    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # Header (더 위로)
    draw_text(draw, margin, 12, HEADER_TEXT, fonts["small_b"], fill="#111111")
    draw_line(draw, margin, 40, W - margin, 40, color="#D9D9D9", w=2)

    # Footer (더 아래로)
    draw_line(draw, margin, H - 44, W - margin, H - 44, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 32, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    # Title
    y = 54
    y = draw_wrapped_title(draw, margin, y, W - 2 * margin, class_name, student_name, fonts)

    # Class / Student (원하면 제거 가능)
    draw_text(draw, margin, y, f"Class: {class_name}", fonts["r"], fill="#333333")
    y += 22
    draw_text(draw, margin, y, f"Student: {student_name}", fonts["r"], fill="#333333")
    y += 18

    # ✅ 1컬럼 폭
    w = W - 2 * margin

    # Quiz → Mocktest 순서로 세로로 붙이기
    y += 10
    y = render_table_full(draw, margin, y, w, "Quiz", quiz_rows, fonts)
    y += 16
    y = render_table_full(draw, margin, y, w, "Mocktest (점수 예상)", mock_rows, fonts)

    # Homework 진행도
    y += 14
    draw_text(draw, margin, y, "Homework 진행도", fonts["h2"], fill="#111111")
    y += 28

    badge_h = 40
    draw.rounded_rectangle([margin, y, margin + w, y + badge_h],
                           radius=16, fill="#F5F6F8", outline=None)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    draw_text(draw, margin + 14, y + 9, hw_txt, fonts["b"], fill="#111111")
    y += badge_h + 16

    # Units box (남는 공간 전부)
    draw_text(draw, margin, y, "보강필요한 부분", fonts["h2"], fill="#111111")
    y += 28

    unit_txt = ", ".join(units) if units else "선택 없음"

    bottom_limit = H - 44 - 10  # footer 라인 위까지
    box_h = max(120, bottom_limit - y)
    draw.rounded_rectangle([margin, y, W - margin, y + box_h],
                           radius=18, fill="#F9FAFB", outline=None)

    # wrap
    max_width = w - 24
    words = unit_txt.split(" ")
    lines = []
    cur = ""
    for wword in words:
        test = (cur + " " + wword).strip()
        if draw.textlength(test, font=fonts["small"]) <= max_width:
            cur = test
        else:
            if cur:
                lines.append(cur)
            cur = wword
    if cur:
        lines.append(cur)

    yy = y + 12
    for line in lines[:8]:
        draw_text(draw, margin + 12, yy, line, fonts["small"], fill="#111111")
        yy += 22

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
st.set_page_config(page_title="성적표 PNG ZIP 생성 (A5 / 1컬럼)", layout="wide")
st.title("엑셀 업로드 → 학생별 보강 선택 → PNG ZIP 다운로드 (A5 / 꽉차게)")
st.caption("Quiz 아래에 Mocktest를 배치해서 빈 공간을 줄이고, 여백/폰트를 키워 더 꽉차게 출력합니다.")

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

        zip_name = f"{safe_filename(class_name)}_reports_A5.zip"
        st.download_button(
            "ZIP 다운로드 (학생별 PNG A5)",
            data=zip_bytes,
            file_name=zip_name,
            mime="application/zip",
        )

        st.success(f"완료! 총 {len(png_files)}명의 PNG(A5)를 ZIP으로 만들었습니다.")

        if preview_img is not None:
            st.image(preview_img, caption=f"미리보기(A5): {preview_student}", use_container_width=True)
