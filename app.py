import io
import os
import re
import pandas as pd
import streamlit as st

from PIL import Image, ImageDraw, ImageFont


# =========================================================
# 고정 머릿말/꼬릿말
# =========================================================
HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

REPORT_TITLE = "트리플 유진쌤 MATH CLASS REPORT"

# 보강 단원(요청 6개)
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

    fonts = {
        "title": f(bold_path, 36),
        "h1": f(bold_path, 24),
        "h2": f(bold_path, 18),
        "b": f(bold_path, 16),
        "r": f(reg_path, 16),
        "small_b": f(bold_path, 14),
        "small": f(reg_path, 14),
        "tiny": f(reg_path, 12),
    }
    return fonts


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

    # 0행: 서브헤더(점수/틀린문제/점수 예상 등)
    sub = raw.iloc[0]
    df = raw.iloc[1:].copy()
    cols = list(raw.columns)

    # "이름" 열의 위치 찾기
    idx_name = None
    for i, c in enumerate(cols):
        if str(c).strip() == "이름":
            idx_name = i
            break

    # 이름 기준 앞/뒤 2칸씩을 메타 5칸으로 고정(너 파일 구조 기준)
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
        # 메타 영역은 Unnamed 전파를 막고 강제 이름
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

    # 숫자 변환
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
            continue  # 학생행 제외

        cnt = sum(pd.notna(df.loc[i, c]) for c in score_cols)
        if cnt > best_count:
            best_count = cnt
            best_idx = i

    if best_idx is None or best_count <= 0:
        raise ValueError("평균행을 찾지 못했습니다. (이름이 비어있고 점수 칼럼에 숫자가 있는 평균행이 필요)")
    return best_idx


# =========================================================
# 학생 1명 데이터 추출 (level/school 제거)
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
        quiz_rows.append({
            "label": pretty(main).replace("QUIZ", "Quiz"),
            "student": s_row[c],
            "avg": avg_row[c],
        })

    mock_rows = []
    for c in mcols:
        main = c.split("__")[0]
        label = pretty(main).replace("MOCK TEST", "Mocktest")
        mock_rows.append({
            "label": label,
            "student": s_row[c],
            "avg": avg_row[c],
        })

    # Homework 진행도(평균)
    hw_progress = None
    if hcols:
        vals = s_row[hcols].dropna()
        if len(vals) > 0:
            hw_progress = float(vals.mean())

            # 0~1 비율로 들어온 경우 대비
            if hw_progress <= 1.0:
                hw_progress *= 100.0

    meta = {"이름": s_row.get("이름", "")}
    return quiz_rows, mock_rows, hw_progress, meta


# =========================================================
# PNG 렌더링 (PIL)
# =========================================================
def draw_line(draw, x1, y1, x2, y2, color="#D9D9D9", w=2):
    draw.line((x1, y1, x2, y2), fill=color, width=w)

def draw_text(draw, x, y, text, font, fill="#111111"):
    draw.text((x, y), text, font=font, fill=fill)

def fmt_num(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    try:
        # 정수면 정수로, 아니면 불필요 소수 제거
        fv = float(v)
        if abs(fv - round(fv)) < 1e-9:
            return str(int(round(fv)))
        return f"{fv:g}"
    except Exception:
        return str(v)

def render_table(draw, x, y, w, title, rows, fonts):
    """
    간단한 표: 제목 + 헤더 + 행들
    columns: label / score / avg
    """
    # title
    draw_text(draw, x, y, title, fonts["h2"])
    y += 34

    row_h = 34
    col1 = int(w * 0.52)
    col2 = int(w * 0.24)
    col3 = w - col1 - col2

    # header bg
    draw.rectangle([x, y, x + w, y + row_h], fill="#F5F6F8", outline=None)
    draw_text(draw, x + col1 + col2 - 10, y + 7, "점수", fonts["small_b"], fill="#333333")
    draw_text(draw, x + w - 10, y + 7, "class 평균", fonts["small_b"], fill="#333333")
    draw_text(draw, x + col1 + col2 - 10, y + 7, "점수", fonts["small_b"], fill="#333333")

    # right align helper by measuring text width
    def right(draw, rx, ty, text, font, fill="#111111"):
        tw = draw.textlength(text, font=font)
        draw.text((rx - tw, ty), text, font=font, fill=fill)

    # header underline
    draw_line(draw, x, y + row_h, x + w, y + row_h, color="#E1E4E8", w=2)
    y += row_h

    # rows
    for r in rows:
        label = str(r["label"])
        sv = fmt_num(r["student"])
        av = fmt_num(r["avg"])

        draw_text(draw, x + 8, y + 7, label, fonts["small"])
        right(draw, x + col1 + col2 - 10, y + 7, sv, fonts["small"], fill="#111111")
        right(draw, x + w - 10, y + 7, av, fonts["small"], fill="#666666")

        draw_line(draw, x, y + row_h, x + w, y + row_h, color="#EDEFF2", w=2)
        y += row_h

    return y

def render_student_report_image(class_name, student_name, quiz_rows, mock_rows, hw_progress, units, fonts):
    # A4 느낌(150 dpi 정도) 1240 x 1754
    W, H = 1240, 1754
    margin = 60

    img = Image.new("RGB", (W, H), "white")
    draw = ImageDraw.Draw(img)

    # Header
    draw_text(draw, margin, 30, HEADER_TEXT, fonts["small_b"], fill="#111111")
    draw_line(draw, margin, 70, W - margin, 70, color="#D9D9D9", w=2)

    # Footer
    draw_line(draw, margin, H - 110, W - margin, H - 110, color="#D9D9D9", w=2)
    draw_text(draw, margin, H - 90, FOOTER_TEXT, fonts["tiny"], fill="#444444")

    # Title
    y = 100
    draw_text(draw, margin, y, REPORT_TITLE, fonts["title"], fill="#111111")
    y += 70

    # Class + Student (level/school 제거)
    draw_text(draw, margin, y, f"Class: {class_name}", fonts["r"], fill="#333333")
    y += 32
    draw_text(draw, margin, y, f"Student: {student_name}", fonts["r"], fill="#333333")
    y += 45

    # 2 columns tables
    gap = 50
    col_w = (W - 2 * margin - gap)
    left_w = col_w // 2
    right_w = col_w - left_w

    left_x = margin
    right_x = margin + left_w + gap
    top_y = y

    # Quiz table
    y_left_end = render_table(draw, left_x, top_y, left_w, "Quiz", quiz_rows, fonts)

    # Homework 진행도 (퀴즈 밑으로)
    y_hw = y_left_end + 18
    draw_text(draw, left_x, y_hw, "Homework 진행도", fonts["h2"], fill="#111111")
    y_hw += 34

    # badge
    badge_w = min(520, left_w)
    badge_h = 44
    draw.rounded_rectangle([left_x, y_hw, left_x + badge_w, y_hw + badge_h], radius=16, fill="#F5F6F8", outline=None)
    hw_txt = "데이터 없음" if hw_progress is None else f"{hw_progress:.0f}%"
    draw_text(draw, left_x + 16, y_hw + 9, hw_txt, fonts["b"], fill="#111111")

    # Mock table
    y_right_end = render_table(draw, right_x, top_y, right_w, "Mocktest (점수 예상)", mock_rows, fonts)

    y_next = max(y_hw + badge_h, y_right_end) + 40

    # Units box
    draw_text(draw, margin, y_next, "보강필요한 부분", fonts["h2"], fill="#111111")
    y_next += 40

    unit_txt = ", ".join(units) if units else "선택 없음"
    box_h = 150
    draw.rounded_rectangle([margin, y_next, W - margin, y_next + box_h], radius=18, fill="#F9FAFB", outline=None)

    # wrap text
    max_width = (W - 2 * margin) - 30
    lines = []
    words = unit_txt.split(" ")
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

    yy = y_next + 16
    for line in lines[:4]:
        draw_text(draw, margin + 16, yy, line, fonts["small"], fill="#111111")
        yy += 28

    return img


def combine_images_vertical(images, gap=40, bg="white"):
    if not images:
        return None
    w = max(im.size[0] for im in images)
    total_h = sum(im.size[1] for im in images) + gap * (len(images) - 1)
    out = Image.new("RGB", (w, total_h), bg)
    y = 0
    for im in images:
        out.paste(im, (0, y))
        y += im.size[1] + gap
    return out


def pil_to_png_bytes(img: Image.Image) -> bytes:
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    return bio.getvalue()


# =========================================================
# Streamlit UI
# =========================================================
st.set_page_config(page_title="성적표 PNG 생성", layout="wide")
st.title("엑셀 업로드 → 학생별 보강 선택 → 전체 PNG 다운로드")
st.caption("학생 전원의 요약 리포트를 PNG로 만들어 한 파일로 내려받습니다. (세로로 길게 이어붙인 이미지)")

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

# 폰트 로드
try:
    fonts = load_fonts()
except Exception as e:
    st.error(f"폰트 로드 실패: {e}")
    st.stop()

st.subheader("학생별 보강 단원 선택 (한 페이지에서 전원 설정)")

# 학생별 units 저장
if "units_by_student" not in st.session_state:
    st.session_state["units_by_student"] = {s: [] for s in students}

units_by_student = st.session_state["units_by_student"]

# 학생 목록이 바뀌면 dict 보정
for s in students:
    units_by_student.setdefault(s, [])
for s in list(units_by_student.keys()):
    if s not in students:
        units_by_student.pop(s, None)

# 한 페이지에서 전원 드롭다운
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

# 생성 버튼
if st.button("전체 PNG 생성하기"):
    images = []
    errors = []

    for s in students:
        try:
            quiz_rows, mock_rows, hw_progress, meta = build_onepage_rows(df, s)
            img = render_student_report_image(
                class_name=class_name,
                student_name=s,
                quiz_rows=quiz_rows,
                mock_rows=mock_rows,
                hw_progress=hw_progress,
                units=units_by_student.get(s, []),
                fonts=fonts,
            )
            images.append(img)
        except Exception as e:
            errors.append(f"{s}: {e}")

    if errors:
        st.error("일부 학생 리포트 생성 실패:\n" + "\n".join(errors))

    if images:
        combined = combine_images_vertical(images, gap=40)
        png_bytes = pil_to_png_bytes(combined)

        st.success("PNG 생성 완료!")
        st.image(combined, caption="미리보기(세로로 길게 이어붙인 전체 이미지)", use_container_width=True)

        filename = f"{class_name}_ALL_STUDENTS_REPORT.png"
        st.download_button(
            "전체 PNG 다운로드",
            data=png_bytes,
            file_name=filename,
            mime="image/png",
        )
