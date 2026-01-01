import io
import os
import re
import pandas as pd
import streamlit as st

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# =========================
# 고정 머릿말/꼬릿말
# =========================
HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

# =========================
# 폰트 등록 (나눔고딕) - fonts/ 폴더
# =========================
@st.cache_resource
def register_nanum():
    reg_path = "fonts/NanumGothic-Regular.ttf"
    bold_path = "fonts/NanumGothicBold.ttf"

    if not os.path.exists(reg_path) or not os.path.exists(bold_path):
        raise FileNotFoundError(
            "폰트 파일을 찾지 못했습니다.\n"
            "fonts/NanumGothic-Regular.ttf\n"
            "fonts/NanumGothicBold.ttf\n"
            "경로/파일명을 확인하세요."
        )

    try:
        pdfmetrics.getFont("NG")
        pdfmetrics.getFont("NGB")
    except KeyError:
        pdfmetrics.registerFont(TTFont("NG", reg_path))
        pdfmetrics.registerFont(TTFont("NGB", bold_path))

    return "NG", "NGB"


# =========================
# 엑셀 로드 + 컬럼 정리 (이름 컬럼 자동 탐색)
# =========================
def norm(s: str) -> str:
    return re.sub(r"\s+", "", str(s)).lower()

def pick_col(cols, candidates):
    """
    cols: 실제 df.columns 리스트
    candidates: ['이름','name','학생명'...] 같은 후보
    => 매칭되는 실제 컬럼명을 반환
    """
    col_norm = {c: norm(c) for c in cols}
    cand_norm = [norm(x) for x in candidates]

    for c in cols:
        for cn in cand_norm:
            if cn != "" and cn in col_norm[c]:
                return c
    return None

def load_and_clean(uploaded_file) -> pd.DataFrame:
    # ※ IMPORTANT: 수식 결과값 읽기 위해 openpyxl 사용
    raw = pd.read_excel(uploaded_file, sheet_name=0, engine="openpyxl")
    sub = raw.iloc[0]          # 서브헤더 행(점수/틀린문제/점수 예상)
    df  = raw.iloc[1:].copy()  # 실제 데이터

    # ---- main header + subheader 합치기 ----
    cols = raw.columns.tolist()
    new_cols = []
    last_main = None

    for c in cols:
        main = c
        if isinstance(c, str) and c.startswith("Unnamed"):
            main = last_main
        else:
            last_main = c

        sh = sub[c]
        if pd.isna(sh) or str(sh).strip() == "":
            new_cols.append(str(main).strip())
        else:
            new_cols.append(f"{str(main).strip()}__{str(sh).strip()}")

    df.columns = new_cols
    df = df.reset_index(drop=True)

    # ---- 점수형 컬럼 숫자 변환 ----
    for c in df.columns:
        if any(k in str(c) for k in ["__점수", "__Total", "__점수 예상", "Homework"]):
            df[c] = pd.to_numeric(df[c], errors="coerce")

    # ---- 메타 컬럼 자동 표준화(이름/레벨/학교 등) ----
    cols2 = df.columns.tolist()

    name_col = pick_col(cols2, ["이름", "name", "학생명", "성명"])
    level_col = pick_col(cols2, ["레벨", "level", "class", "반"])
    school_col = pick_col(cols2, ["학교", "school"])
    phone_col = pick_col(cols2, ["연락처", "phone", "전화"])
    contact_col = pick_col(cols2, ["연락(이메일/카톡)", "카톡", "kakao", "email", "이메일"])

    rename_map = {}
    if name_col: rename_map[name_col] = "이름"
    if level_col: rename_map[level_col] = "레벨"
    if school_col: rename_map[school_col] = "학교"
    if phone_col: rename_map[phone_col] = "연락처"
    if contact_col: rename_map[contact_col] = "연락(이메일/카톡)"

    df = df.rename(columns=rename_map)

    # 필수: 이름 컬럼 없으면 여기서 친절하게 종료
    if "이름" not in df.columns:
        raise KeyError(
            "엑셀에서 '이름' 컬럼을 찾지 못했습니다.\n"
            "엑셀의 학생 이름 열 제목이 '이름'/'Name'/'학생명' 중 하나인지 확인하세요."
        )

    return df


# =========================
# 컬럼 판별
# =========================
def quiz_score_cols(df):
    return [c for c in df.columns if re.match(r"^(QUIZ\d+.*|ReviewQuiz.*)__점수$", str(c))]

def mock_pred_cols(df):
    return [c for c in df.columns if re.match(r"^MOCK TEST.*__점수 예상$", str(c))]

def homework_cols(df):
    return [c for c in df.columns if str(c).startswith("Homework")]

def pretty(label: str) -> str:
    return re.sub(r"\s*\(.*?\)\s*", "", label).strip()

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
        raise ValueError("평균행을 찾지 못했습니다. (이름이 비어있고 점수 칼럼에 숫자가 있는 행이 필요)")
    return best_idx

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

    hw_avg = None
    if hcols:
        vals = s_row[hcols].dropna()
        if len(vals) > 0:
            hw_avg = float(vals.mean())

    meta = {
        "이름": s_row.get("이름", ""),
        "레벨": s_row.get("레벨", ""),
        "학교": s_row.get("학교", ""),
    }

    return quiz_rows, mock_rows, hw_avg, meta


# =========================
# PDF
# =========================
def draw_header_footer(c: canvas.Canvas, W, H, margin, fontR, fontB, page_num: int):
    c.setFont(fontB, 11)
    c.setFillColor(colors.HexColor("#111111"))
    c.drawString(margin, H - margin + 2*mm, HEADER_TEXT)

    c.setStrokeColor(colors.HexColor("#D9D9D9"))
    c.setLineWidth(0.6)
    c.line(margin, H - margin - 2*mm, W - margin, H - margin - 2*mm)

    c.line(margin, margin + 10*mm, W - margin, margin + 10*mm)
    c.setFont(fontR, 9.5)
    c.setFillColor(colors.HexColor("#444444"))
    c.drawString(margin, margin + 5*mm, FOOTER_TEXT)
    c.drawRightString(W - margin, margin + 5*mm, f"{page_num}")
    c.setFillColor(colors.black)

def draw_table_clean(c, x, y_top, w, title, rows, fontR, fontB):
    row_h = 7.2 * mm
    col_w = [w * 0.52, w * 0.24, w * 0.24]

    c.setFont(fontB, 11)
    c.setFillColor(colors.HexColor("#111111"))
    c.drawString(x, y_top, title)
    y = y_top - 6*mm

    c.setFillColor(colors.HexColor("#F5F6F8"))
    c.rect(x, y - row_h, w, row_h, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#333333"))
    c.setFont(fontB, 9.8)
    c.drawRightString(x + col_w[0] + col_w[1] - 2*mm, y - row_h + 2.2*mm, "점수")
    c.drawRightString(x + w - 2*mm, y - row_h + 2.2*mm, "class 평균")

    c.setStrokeColor(colors.HexColor("#E1E4E8"))
    c.setLineWidth(0.7)
    c.line(x, y - row_h, x + w, y - row_h)
    y -= row_h

    c.setFont(fontR, 9.8)
    for r in rows:
        c.setFillColor(colors.HexColor("#111111"))
        c.drawString(x + 2*mm, y - row_h + 2.2*mm, str(r["label"]))

        sv = "" if pd.isna(r["student"]) else f"{float(r['student']):g}"
        av = "" if pd.isna(r["avg"]) else f"{float(r['avg']):g}"

        c.drawRightString(x + col_w[0] + col_w[1] - 2*mm, y - row_h + 2.2*mm, sv)
        c.setFillColor(colors.HexColor("#666666"))
        c.drawRightString(x + w - 2*mm, y - row_h + 2.2*mm, av)

        c.setStrokeColor(colors.HexColor("#EDEFF2"))
        c.setLineWidth(0.6)
        c.line(x, y - row_h, x + w, y - row_h)
        y -= row_h

    c.setFillColor(colors.black)
    return y

def make_report_pdf(class_name, meta, quiz_rows, mock_rows, hw_avg, units) -> bytes:
    fontR, fontB = register_nanum()

    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    margin = 16 * mm

    draw_header_footer(c, W, H, margin, fontR, fontB, 1)

    y = H - margin - 12*mm

    c.setFont(fontB, 18)
    c.setFillColor(colors.HexColor("#111111"))
    c.drawString(margin, y, "성적 요약 리포트")
    c.setFillColor(colors.black)

    y -= 10*mm
    c.setFont(fontR, 11)
    c.setFillColor(colors.HexColor("#333333"))
    c.drawString(margin, y, f"Class: {class_name}")
    y -= 6*mm
    c.drawString(margin, y, f"Student: {meta.get('이름','')}")
    y -= 6*mm
    c.drawString(margin, y, f"Level: {meta.get('레벨','')}")
    y -= 6*mm
    c.drawString(margin, y, f"School: {meta.get('학교','')}")
    c.setFillColor(colors.black)

    y -= 10*mm
    badge_w = 70*mm
    badge_h = 10*mm
    c.setFillColor(colors.HexColor("#F5F6F8"))
    c.roundRect(margin, y - badge_h + 2*mm, badge_w, badge_h, 3*mm, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#111111"))
    c.setFont(fontB, 10.5)
    hw_txt = "데이터 없음" if hw_avg is None else f"{hw_avg:.0f}%"
    c.drawString(margin + 3*mm, y - badge_h + 5*mm, f"Homework 평균  {hw_txt}")
    c.setFillColor(colors.black)

    y_tables_top = y - 12*mm
    gap = 10*mm
    col_w = (W - 2*margin - gap) / 2

    left_x = margin
    right_x = margin + col_w + gap

    y_left_end = draw_table_clean(c, left_x, y_tables_top, col_w, "Quiz", quiz_rows, fontR, fontB)
    y_right_end = draw_table_clean(c, right_x, y_tables_top, col_w, "Mocktest (점수 예상)", mock_rows, fontR, fontB)

    y_next = min(y_left_end, y_right_end) - 10*mm

    c.setFont(fontB, 12)
    c.setFillColor(colors.HexColor("#111111"))
    c.drawString(margin, y_next, "보강필요한 부분")
    c.setFillColor(colors.black)

    y_next -= 7*mm
    unit_txt = ", ".join(units) if units else "선택 없음"
    box_h = 22*mm
    c.setFillColor(colors.HexColor("#F9FAFB"))
    c.roundRect(margin, y_next - box_h + 2*mm, W - 2*margin, box_h, 3*mm, stroke=0, fill=1)
    c.setFillColor(colors.HexColor("#111111"))
    c.setFont(fontR, 10.5)

    max_chars = 110
    lines = [unit_txt[i:i+max_chars] for i in range(0, len(unit_txt), max_chars)]
    yy = y_next - 4*mm
    for line in lines[:3]:
        c.drawString(margin + 3*mm, yy, line)
        yy -= 6*mm

    c.showPage()
    c.save()
    return buf.getvalue()


# =========================
# Streamlit UI
# =========================
st.set_page_config(page_title="성적표 PDF 생성", layout="wide")
st.title("엑셀 업로드 → 학생 선택 → PDF 성적표")

class_name = st.text_input("Class 이름(성적표에 표시)", value="S2 개념반")

uploaded = st.file_uploader("엑셀 업로드(.xlsx)", type=["xlsx"])
if not uploaded:
    st.info("엑셀 파일을 업로드하면 학생 목록이 자동으로 뜹니다.")
    st.stop()

try:
    df = load_and_clean(uploaded)
except Exception as e:
    st.error(f"엑셀 파싱/인식 실패: {e}")
    st.write("디버그용: 인식된 컬럼 목록")
    # 가능하면 컬럼을 보여줘서 바로 원인 파악 가능
    try:
        raw_dbg = pd.read_excel(uploaded, sheet_name=0, engine="openpyxl")
        st.write(list(raw_dbg.columns))
    except Exception:
        pass
    st.stop()

students = sorted([s for s in df["이름"].dropna().unique().tolist() if str(s).strip() != ""])
if not students:
    st.error("학생 이름을 찾지 못했습니다. '이름' 열 제목을 확인해주세요.")
    st.write("현재 컬럼:", list(df.columns))
    st.stop()

student = st.selectbox("학생 선택", students)

DEFAULT_UNITS = [
    "Linear equations", "Inequalities", "Functions", "Quadratics",
    "Polynomials", "Factoring", "Exponents", "Radicals",
    "Geometry", "Trigonometry", "Word problems"
]
units = st.multiselect("보강필요한 부분(드롭다운)", DEFAULT_UNITS)

quiz_rows, mock_rows, hw_avg, meta = build_onepage_rows(df, student)

c1, c2 = st.columns(2)
with c1:
    st.subheader("Quiz 미리보기")
    st.dataframe(
        pd.DataFrame([{"Quiz": r["label"], "점수": r["student"], "class 평균": r["avg"]} for r in quiz_rows]),
        use_container_width=True,
        hide_index=True
    )
with c2:
    st.subheader("Mocktest(점수 예상) 미리보기")
    st.dataframe(
        pd.DataFrame([{"Mocktest": r["label"], "점수": r["student"], "class 평균": r["avg"]} for r in mock_rows]),
        use_container_width=True,
        hide_index=True
    )

st.subheader("Homework")
st.write("평균:", ("데이터 없음" if hw_avg is None else f"{hw_avg:.0f}%"))

pdf_bytes = make_report_pdf(class_name, meta, quiz_rows, mock_rows, hw_avg, units)
filename = f"{class_name}_{meta.get('이름','학생')}_report.pdf"

st.download_button("PDF 다운로드", data=pdf_bytes, file_name=filename, mime="application/pdf")
