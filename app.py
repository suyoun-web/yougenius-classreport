import io
import os
import re
import zipfile
import tempfile
from datetime import datetime
from collections import defaultdict, Counter

import pandas as pd
import streamlit as st

# PDF 생성
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# 이미지 트림
from PIL import Image, ImageChops

# ✅ PyMuPDF는 환경에 없을 수 있으니 안전하게 import
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except Exception:
    HAS_PYMUPDF = False


# =========================
# 0. 앱 기본 설정
# =========================
st.set_page_config(page_title="유진 SAT class report", layout="wide")
st.title("유진 SAT class report")


# =========================
# 1. 공통 상수/유틸
# =========================
FONT_REGULAR = "fonts/NanumGothic-Regular.ttf"
FONT_BOLD = "fonts/NanumGothic-Bold.ttf"

HEADER_TEXT = "YOU, GENIUS 유지니어스 MATH with 유진쌤"
FOOTER_TEXT = "Kakaotalk : yujinj524 / Phone : 010-6395-8733"

REINFORCE_TOPICS = [
    "I. Linear",
    "IV. Quadratic",
    "V. Exponential",
    "VI. Polynomials, radical and rational functions",
    "VII. Geometry",
    "VIII. Statistics",
]

TOPIC_NAMES_MAJOR = {
    "1": "I. Linear",
    "3": "IV. Quadratic",
    "4": "V. Exponential",
    "5": "VI. Polynomials, radical and rational functions",
    "6": "VII. Geometry",
    "7": "VIII. Statistics",
}


def _clean(x) -> str:
    if x is None:
        return ""
    if isinstance(x, float) and pd.isna(x):
        return ""
    return str(x).replace("\r", "").strip()


def _ensure_fonts():
    if not (os.path.exists(FONT_REGULAR) and os.path.exists(FONT_BOLD)):
        return False
    try:
        pdfmetrics.registerFont(TTFont("NanumGothic", FONT_REGULAR))
    except:
        pass
    try:
        pdfmetrics.registerFont(TTFont("NanumGothic-Bold", FONT_BOLD))
    except:
        pass
    return True


def _parse_wrong_list(val) -> set[int]:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return set()
    s = _clean(val)
    if s == "" or s.upper() in ["X", "-", "—", "–"]:
        return set()
    s = s.replace("，", ",").replace(";", ",")
    parts = [p.strip() for p in s.split(",") if p.strip()]
    out = set()
    for p in parts:
        try:
            out.add(int(float(p)))
        except:
            pass
    return out


def _is_number(v) -> bool:
    try:
        if v is None or v == "":
            return False
        float(v)
        return True
    except:
        return False


def _extract_num(s: str) -> int:
    m = re.search(r"(\d+)", _clean(s))
    return int(m.group(1)) if m else 9999


def _quiz_sort_key(colname: str):
    s = _clean(colname).lower().replace(" ", "")
    if "review" in s and "quiz" in s:
        return (1, _extract_num(s), s)  # reviewquiz는 뒤로
    if "quiz" in s:
        return (0, _extract_num(s), s)
    return (2, 9999, s)


def _mock_sort_key(colname: str):
    return (_extract_num(colname), _clean(colname).lower())


def _trim_white(img: Image.Image, bg=(255, 255, 255)) -> Image.Image:
    if img.mode != "RGB":
        img = img.convert("RGB")
    bg_img = Image.new("RGB", img.size, bg)
    diff = ImageChops.difference(img, bg_img)
    bbox = diff.getbbox()
    if bbox:
        return img.crop(bbox)
    return img


def _pdf_first_page_to_png(pdf_path: str, zoom: float = 2.2) -> bytes:
    """
    ✅ PyMuPDF 있을 때만 사용
    """
    doc = fitz.open(pdf_path)
    page = doc[0]
    pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
    png = pix.tobytes("png")
    # 트림
    try:
        img = Image.open(io.BytesIO(png))
        img2 = _trim_white(img)
        out = io.BytesIO()
        img2.save(out, format="PNG")
        return out.getvalue()
    except:
        return png


def _df_read_xlsx(uploaded) -> pd.DataFrame:
    df = pd.read_excel(uploaded)
    df.columns = [_clean(c) for c in df.columns]
    return df


def _find_avg_row(df: pd.DataFrame):
    if "Name" in df.columns:
        cand = df[df["Name"].astype(str).map(_clean) == "평균"]
        if len(cand) >= 1:
            return cand.iloc[0]
    return None


def _detect_columns_export(df: pd.DataFrame):
    cols = list(df.columns)

    name_col = "Name" if "Name" in cols else None
    class_col = "Class" if "Class" in cols else None

    quiz_cols = [c for c in cols if "quiz" in _clean(c).lower() and c not in ["Name", "Class"]]
    mock_cols = [c for c in cols if "mock" in _clean(c).lower() and c not in ["Name", "Class"] and "틀린" not in _clean(c)]
    hw_cols = [c for c in cols if _clean(c).lower().replace(" ", "").startswith("homework")]

    quiz_cols = sorted(quiz_cols, key=_quiz_sort_key)
    mock_cols = sorted(mock_cols, key=_mock_sort_key)
    hw_cols = sorted(hw_cols, key=_mock_sort_key)

    return name_col, class_col, quiz_cols, mock_cols, hw_cols


def _safe_filename(s: str) -> str:
    s = _clean(s)
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s.strip() if s.strip() else "file"


# =========================
# Tab1 PDF 생성
# =========================
def create_student_report_pdf(
    out_pdf_path: str,
    class_name: str,
    student_name: str,
    quizzes: list[tuple[str, str, str]],
    mocks: list[tuple[str, str, str]],
    homework: list[tuple[str, str]],
    reinforce_text: str,
    generated_date: str,
):
    _ensure_fonts()
    c = canvas.Canvas(out_pdf_path, pagesize=A5)
    W, H = A5

    title_col = colors.Color(15/255, 23/255, 42/255)
    muted = colors.Color(100/255, 116/255, 139/255)
    stroke = colors.Color(203/255, 213/255, 225/255)
    pill = colors.Color(241/255, 245/255, 249/255)
    row_bg = colors.Color(248/255, 250/255, 252/255)

    L = 6 * mm
    R = 6 * mm
    T = H - 7 * mm
    B = 6 * mm
    usable_w = W - L - R

    c.setFillColor(title_col)
    c.setFont("NanumGothic-Bold", 9.5)
    c.drawCentredString(W/2, T, HEADER_TEXT)

    title = f"{class_name} {student_name} CLASS REPORT"
    c.setFont("NanumGothic-Bold", 15.5)
    c.drawString(L, T - 10*mm, title)

    c.setFillColor(muted)
    c.setFont("NanumGothic", 8.5)
    c.drawRightString(W - R, T - 9.5*mm, f"Generated: {generated_date}")

    c.setStrokeColor(title_col)
    c.setLineWidth(1.6)
    c.line(L, T - 12.8*mm, W - R, T - 12.8*mm)

    y = T - 18*mm

    def draw_section_box(y_top, header, rows, col2_name="Student", col3_name="Class Avg"):
        box_pad = 4.2*mm
        header_h = 7.0*mm
        row_h = 6.3*mm
        h = box_pad + header_h + len(rows)*row_h + box_pad

        x = L
        w = usable_w

        c.setFillColor(colors.white)
        c.setStrokeColor(stroke)
        c.setLineWidth(1)
        c.roundRect(x, y_top - h, w, h, 5*mm, stroke=1, fill=1)

        c.setFillColor(pill)
        c.setStrokeColor(pill)
        c.rect(x + 3*mm, y_top - box_pad - header_h, w - 6*mm, header_h, stroke=0, fill=1)

        c.setFillColor(title_col)
        c.setFont("NanumGothic-Bold", 11.5)
        c.drawString(x + 4.3*mm, y_top - box_pad - 5.0*mm, header)

        c.setFillColor(muted)
        c.setFont("NanumGothic-Bold", 8.3)
        cx_label = x + 4.3*mm
        cx_s = x + w*0.66
        cx_a = x + w - 6.0*mm
        c.drawRightString(cx_s, y_top - box_pad - 5.0*mm, col2_name)
        c.drawRightString(cx_a, y_top - box_pad - 5.0*mm, col3_name)

        yy = y_top - box_pad - header_h - 1.2*mm
        for i, (lab, sval, aval) in enumerate(rows):
            ry = yy - (i+1)*row_h
            if i % 2 == 0:
                c.setFillColor(row_bg)
                c.setStrokeColor(row_bg)
                c.rect(x + 3*mm, ry, w - 6*mm, row_h, stroke=0, fill=1)

            c.setFillColor(title_col)
            c.setFont("NanumGothic", 9.8)
            c.drawString(cx_label, ry + 1.9*mm, str(lab))

            c.setFillColor(title_col)
            c.setFont("NanumGothic-Bold", 10.3)
            c.drawRightString(cx_s, ry + 1.9*mm, str(sval))

            c.setFillColor(muted)
            c.setFont("NanumGothic", 9.5)
            c.drawRightString(cx_a, ry + 1.9*mm, str(aval))

        return y_top - h - 4.0*mm

    if quizzes:
        y = draw_section_box(y, "Quiz Scores", quizzes)

    if mocks:
        y = draw_section_box(y, "Mocktest Scores", mocks, col2_name="Score", col3_name="Class Avg")

    if homework:
        done = 0
        total = 0
        for _, v in homework:
            total += 1
            if _clean(v) != "" and _clean(v) != "0":
                done += 1
        pct = int(round((done/total)*100)) if total else 0

        box_h = 22*mm
        x = L
        w = usable_w

        c.setFillColor(colors.white)
        c.setStrokeColor(stroke)
        c.setLineWidth(1)
        c.roundRect(x, y - box_h, w, box_h, 5*mm, stroke=1, fill=1)

        c.setFillColor(title_col)
        c.setFont("NanumGothic-Bold", 11.5)
        c.drawString(x + 4.3*mm, y - 6.8*mm, "Homework 진행도")

        c.setFillColor(title_col)
        c.setFont("NanumGothic-Bold", 16)
        c.drawRightString(x + w - 6.0*mm, y - 8.0*mm, f"{pct}%")

        c.setFillColor(muted)
        c.setFont("NanumGothic", 9.5)
        c.drawString(x + 4.3*mm, y - 13.8*mm, f"{done}/{total} completed")

        y = y - box_h - 4.0*mm

    box_h = 26*mm
    x = L
    w = usable_w
    c.setFillColor(colors.white)
    c.setStrokeColor(stroke)
    c.setLineWidth(1)
    c.roundRect(x, y - box_h, w, box_h, 5*mm, stroke=1, fill=1)

    c.setFillColor(title_col)
    c.setFont("NanumGothic-Bold", 10.8)
    c.drawString(x + 4.3*mm, y - 6.8*mm, "보강이 필요한 부분 및 유진쌤 Comment")

    c.setFillColor(title_col)
    c.setFont("NanumGothic", 9.3)
    reinforce_text = _clean(reinforce_text)
    if reinforce_text:
        lines = []
        for part in reinforce_text.split(","):
            p = part.strip()
            if p:
                lines.append("• " + p)
        text_y = y - 12.2*mm
        for ln in lines[:4]:
            c.drawString(x + 5.5*mm, text_y, ln)
            text_y -= 4.4*mm

    c.setFillColor(title_col)
    c.setFont("NanumGothic", 8.3)
    c.drawCentredString(W/2, B, FOOTER_TEXT)

    c.showPage()
    c.save()


def tab1_class_report():
    st.subheader("Tab1) Class Report 생성 (학생별 PDF/PNG ZIP)")

    if not _ensure_fonts():
        st.warning("⚠️ fonts 폴더에 NanumGothic-Regular.ttf / NanumGothic-Bold.ttf 를 넣어줘야 해요.")

    if not HAS_PYMUPDF:
        st.warning("⚠️ PNG 변환은 PyMuPDF(fitz)가 필요해요. 지금은 PDF ZIP만 생성 가능합니다. (requirements/runtime.txt 확인)")

    colA, colB = st.columns([1.2, 1])
    with colA:
        export_file = st.file_uploader("EXPORT 파일 업로드 (.xlsx)", type=["xlsx"], key="t1_export")
    with colB:
        class_name_input = st.text_input("Class 이름(파일에 없어도 OK)", value="", key="t1_class")
        gen_date = st.date_input("Generated 날짜", value=datetime.now().date(), key="t1_date")

    if export_file is None:
        st.info("EXPORT 파일을 업로드하면 학생 목록이 뜹니다.")
        return

    try:
        df = _df_read_xlsx(export_file)
    except Exception as e:
        st.error(f"엑셀 읽기 실패: {e}")
        return

    name_col, class_col, quiz_cols, mock_cols, hw_cols = _detect_columns_export(df)
    if name_col is None:
        st.error("EXPORT에 Name 컬럼이 없어요. (AppScript EXPORT 형식인지 확인)")
        return

    avg_row = _find_avg_row(df)

    df2 = df.copy()
    df2[name_col] = df2[name_col].astype(str).map(_clean)
    students_df = df2[(df2[name_col] != "") & (df2[name_col] != "평균")].copy()
    students = students_df[name_col].tolist()
    if not students:
        st.error("학생 이름을 찾지 못했어요.")
        return

    inferred_class = ""
    if class_col and class_col in students_df.columns:
        inferred_class = _clean(students_df.iloc[0][class_col])
    class_name = _clean(class_name_input) or inferred_class or "CLASS"

    st.markdown("#### 보강 단원 선택(전체 학생)")
    if "reinforce_map" not in st.session_state:
        st.session_state.reinforce_map = {s: "" for s in students}

    bulk_col1, bulk_col2, bulk_col3 = st.columns([1.4, 1.4, 0.8])
    with bulk_col1:
        target_students = st.multiselect("적용할 학생 선택", options=students, default=[], key="t1_bulk_students")
    with bulk_col2:
        target_topics = st.multiselect("보강 단원 선택(복수 가능)", options=REINFORCE_TOPICS, default=[], key="t1_bulk_topics")
    with bulk_col3:
        if st.button("적용", use_container_width=True, key="t1_apply"):
            txt = ", ".join(target_topics).strip()
            for s in target_students:
                st.session_state.reinforce_map[s] = txt

    st.divider()

    if st.button("🚀 Report 생성", type="primary", use_container_width=True, key="t1_build"):
        if not _ensure_fonts():
            st.error("폰트가 없어서 PDF 생성이 안 돼요. fonts 폴더 확인!")
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_paths = []
            png_paths = []
            prog = st.progress(0)

            for i, stu in enumerate(students):
                row = students_df[students_df[name_col] == stu].iloc[0]

                def get_val(col):
                    v = row.get(col, "")
                    return _clean(v) if _clean(v) != "" else "-"

                def get_avg(col):
                    if avg_row is None:
                        return "-"
                    v = avg_row.get(col, "")
                    if _is_number(v):
                        return f"{float(v):.2f}"
                    vv = _clean(v)
                    return vv if vv != "" else "-"

                quizzes = [(c, get_val(c), get_avg(c)) for c in quiz_cols]
                mocks = [(c, get_val(c), get_avg(c)) for c in mock_cols]
                homework = [(c, get_val(c)) for c in hw_cols]

                reinforce_text = st.session_state.reinforce_map.get(stu, "")

                pdf_name = f"{_safe_filename(class_name)}_{_safe_filename(stu)}.pdf"
                pdf_path = os.path.join(tmpdir, pdf_name)

                create_student_report_pdf(
                    out_pdf_path=pdf_path,
                    class_name=class_name,
                    student_name=stu,
                    quizzes=quizzes,
                    mocks=mocks,
                    homework=homework,
                    reinforce_text=reinforce_text,
                    generated_date=gen_date.strftime("%Y-%m-%d"),
                )

                pdf_paths.append(pdf_path)

                if HAS_PYMUPDF:
                    png_name = f"{_safe_filename(class_name)}_{_safe_filename(stu)}.png"
                    png_path = os.path.join(tmpdir, png_name)
                    png_bytes = _pdf_first_page_to_png(pdf_path, zoom=2.2)
                    with open(png_path, "wb") as f:
                        f.write(png_bytes)
                    png_paths.append(png_path)

                prog.progress((i + 1) / len(students))

            # PDF ZIP
            pdf_zip = io.BytesIO()
            with zipfile.ZipFile(pdf_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for p in pdf_paths:
                    z.write(p, arcname=os.path.basename(p))
            pdf_zip.seek(0)

            st.success(f"✅ 생성 완료: PDF {len(pdf_paths)}개" + (f" / PNG {len(png_paths)}개" if HAS_PYMUPDF else ""))

            # 다운로드 먼저
            st.download_button(
                "📦 PDF ZIP 다운로드",
                data=pdf_zip,
                file_name=f"{_safe_filename(class_name)}_CLASS_REPORT_PDF.zip",
                mime="application/zip",
                use_container_width=True,
                key="t1_dl_pdf",
            )

            # PNG ZIP은 PyMuPDF 있을 때만
            if HAS_PYMUPDF:
                png_zip = io.BytesIO()
                with zipfile.ZipFile(png_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                    for p in png_paths:
                        z.write(p, arcname=os.path.basename(p))
                png_zip.seek(0)

                st.download_button(
                    "🖼️ PNG ZIP 다운로드",
                    data=png_zip,
                    file_name=f"{_safe_filename(class_name)}_CLASS_REPORT_PNG.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="t1_dl_png",
                )

            st.markdown("#### 미리보기(1명)")
            preview_student = st.selectbox("미리볼 학생", options=students, key="t1_preview")
            if HAS_PYMUPDF:
                prev_png = os.path.join(tmpdir, f"{_safe_filename(class_name)}_{_safe_filename(preview_student)}.png")
                if os.path.exists(prev_png):
                    st.image(prev_png, use_container_width=True)
            else:
                st.info("PNG 미리보기는 PyMuPDF 설치 후 가능해요.")


# =========================
# Tab2 (자리만 유지 - 너가 이미 요청한 “두번째 탭”)
# =========================
def tab2_placeholder():
    st.subheader("Tab2) Mock 중복 단원 분석")
    st.info("Tab2는 다음 단계에서(네가 export_mock/메타 파일 형식 확정되면) 그대로 붙이면 됩니다. 지금 오류는 Tab1 PNG 변환 라이브러리(PyMuPDF) 문제였어요.")


# =========================
# 탭 구성
# =========================
tab1, tab2 = st.tabs(["📄 Class Report", "📌 Mock 중복 단원 분석"])

with tab1:
    tab1_class_report()

with tab2:
    tab2_placeholder()
