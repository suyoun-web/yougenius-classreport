import io
import os
import re
import zipfile
import tempfile
from datetime import datetime
from collections import defaultdict, Counter

import pandas as pd
import streamlit as st

# PDF/PNG
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A5
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import fitz  # PyMuPDF
from PIL import Image, ImageChops

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
    # reviewquiz는 항상 뒤로
    if "review" in s and "quiz" in s:
        return (1, _extract_num(s), s)
    # quiz 숫자
    if "quiz" in s:
        return (0, _extract_num(s), s)
    return (2, 9999, s)


def _mock_sort_key(colname: str):
    return (_extract_num(colname), _clean(colname).lower())


def _trim_white(img: Image.Image, bg=(255, 255, 255)) -> Image.Image:
    """
    PNG 렌더 후 여백 자동 트림(흰 배경 기준)
    """
    if img.mode != "RGB":
        img = img.convert("RGB")
    bg_img = Image.new("RGB", img.size, bg)
    diff = ImageChops.difference(img, bg_img)
    bbox = diff.getbbox()
    if bbox:
        return img.crop(bbox)
    return img


def _pdf_first_page_to_png(pdf_path: str, zoom: float = 2.2) -> bytes:
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

    # 정렬
    quiz_cols = sorted(quiz_cols, key=_quiz_sort_key)
    mock_cols = sorted(mock_cols, key=_mock_sort_key)
    hw_cols = sorted(hw_cols, key=_mock_sort_key)

    # 마지막 mock틀린문제(전체 export에서 쓰던)
    m1_wrong_col = None
    m2_wrong_col = None
    for c in cols:
        s = _clean(c).replace(" ", "")
        if s == "m1틀린문제":
            m1_wrong_col = c
        if s == "m2틀린문제":
            m2_wrong_col = c

    return name_col, class_col, quiz_cols, mock_cols, hw_cols, m1_wrong_col, m2_wrong_col


def _safe_filename(s: str) -> str:
    s = _clean(s)
    s = re.sub(r"[\\/:*?\"<>|]+", "_", s)
    return s.strip() if s.strip() else "file"


# =========================
# 2. Tab1: Class Report (PDF/PNG ZIP)
# =========================
def create_student_report_pdf(
    out_pdf_path: str,
    class_name: str,
    student_name: str,
    quizzes: list[tuple[str, str, str]],   # (label, student, avg)
    mocks: list[tuple[str, str, str]],
    homework: list[tuple[str, str]],       # (label, student)
    reinforce_text: str,
    generated_date: str,
):
    _ensure_fonts()
    c = canvas.Canvas(out_pdf_path, pagesize=A5)
    W, H = A5

    # palette
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

    # Header
    c.setFillColor(title_col)
    c.setFont("NanumGothic-Bold", 9.5)
    c.drawCentredString(W/2, T, HEADER_TEXT)

    # Title
    title = f"{class_name} {student_name} CLASS REPORT"
    c.setFont("NanumGothic-Bold", 15.5)
    c.drawString(L, T - 10*mm, title)

    # Date
    c.setFillColor(muted)
    c.setFont("NanumGothic", 8.5)
    c.drawRightString(W - R, T - 9.5*mm, f"Generated: {generated_date}")

    # Divider
    c.setStrokeColor(title_col)
    c.setLineWidth(1.6)
    c.line(L, T - 12.8*mm, W - R, T - 12.8*mm)

    y = T - 18*mm

    def draw_section_box(y_top, header, rows, col2_name="Student", col3_name="Class Avg"):
        """
        rows: list[(label, student_val, avg_val)]
        """
        box_pad = 4.2*mm
        header_h = 7.0*mm
        row_h = 6.3*mm
        h = box_pad + header_h + len(rows)*row_h + box_pad

        x = L
        w = usable_w

        # card
        c.setFillColor(colors.white)
        c.setStrokeColor(stroke)
        c.setLineWidth(1)
        c.roundRect(x, y_top - h, w, h, 5*mm, stroke=1, fill=1)

        # header strip
        c.setFillColor(pill)
        c.setStrokeColor(pill)
        c.rect(x + 3*mm, y_top - box_pad - header_h, w - 6*mm, header_h, stroke=0, fill=1)

        c.setFillColor(title_col)
        c.setFont("NanumGothic-Bold", 11.5)
        c.drawString(x + 4.3*mm, y_top - box_pad - 5.0*mm, header)

        # column labels
        c.setFillColor(muted)
        c.setFont("NanumGothic-Bold", 8.3)
        cx_label = x + 4.3*mm
        cx_s = x + w*0.66
        cx_a = x + w - 6.0*mm
        c.drawRightString(cx_s, y_top - box_pad - 5.0*mm, col2_name)
        c.drawRightString(cx_a, y_top - box_pad - 5.0*mm, col3_name)

        # rows
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

    # Quiz section
    if quizzes:
        y = draw_section_box(y, "Quiz Scores", quizzes)
    # Mock section (아래로 배치)
    if mocks:
        y = draw_section_box(y, "Mocktest Scores", mocks, col2_name="Score", col3_name="Class Avg")

    # Homework progress (간단 카드)
    if homework:
        # progress = 완료(0이 아닌 값) / 전체
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

        # 작은 리스트
        c.setFillColor(muted)
        c.setFont("NanumGothic", 8.5)
        txt = " / ".join([f"{lab}:{_clean(v) if _clean(v)!='' else '-'}" for lab, v in homework[:7]])
        c.drawString(x + 4.3*mm, y - 18.5*mm, txt[:120])

        y = y - box_h - 4.0*mm

    # Reinforce / Comment
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
    # 선택없으면 그냥 비워둠
    reinforce_text = _clean(reinforce_text)
    if reinforce_text:
        # 줄바꿈
        lines = []
        for part in reinforce_text.split(","):
            p = part.strip()
            if p:
                lines.append("• " + p)
        text_y = y - 12.2*mm
        for ln in lines[:4]:
            c.drawString(x + 5.5*mm, text_y, ln)
            text_y -= 4.4*mm

    y = y - box_h - 2.0*mm

    # Footer
    c.setFillColor(title_col)
    c.setFont("NanumGothic", 8.3)
    c.drawCentredString(W/2, B, FOOTER_TEXT)

    c.showPage()
    c.save()


def tab1_class_report():
    st.subheader("Tab1) Class Report 생성 (학생별 PDF/PNG ZIP)")

    if not _ensure_fonts():
        st.warning("⚠️ fonts 폴더에 NanumGothic-Regular.ttf / NanumGothic-Bold.ttf 를 넣어줘야 해요.")

    colA, colB = st.columns([1.2, 1])
    with colA:
        export_file = st.file_uploader("EXPORT 파일 업로드 (.xlsx)", type=["xlsx"], key="t1_export")
    with colB:
        class_name_input = st.text_input("Class 이름(파일에 없어도 OK)", value="", key="t1_class")
        gen_date = st.date_input("Generated 날짜", value=datetime.now().date(), key="t1_date")

    st.caption("Tip: ZIP 다운로드 버튼이 위에 나오도록, 생성 완료 후 다운로드 먼저 보여줍니다. 미리보기는 1명만 보여줘요.")

    if export_file is None:
        st.info("EXPORT 파일을 업로드하면 학생 목록이 뜹니다.")
        return

    try:
        df = _df_read_xlsx(export_file)
    except Exception as e:
        st.error(f"엑셀 읽기 실패: {e}")
        return

    name_col, class_col, quiz_cols, mock_cols, hw_cols, _, _ = _detect_columns_export(df)
    if name_col is None:
        st.error("EXPORT에 Name 컬럼이 없어요. (AppScript EXPORT 형식인지 확인)")
        return

    # 평균행/학생행
    avg_row = _find_avg_row(df)

    df2 = df.copy()
    df2[name_col] = df2[name_col].astype(str).map(_clean)
    students_df = df2[(df2[name_col] != "") & (df2[name_col] != "평균")].copy()
    students = students_df[name_col].tolist()
    if not students:
        st.error("학생 이름을 찾지 못했어요.")
        return

    # class_name
    inferred_class = ""
    if class_col and class_col in students_df.columns:
        inferred_class = _clean(students_df.iloc[0][class_col])
    class_name = _clean(class_name_input) or inferred_class or "CLASS"

    # 보강 단원 선택(전체 학생 한번에)
    st.markdown("#### 보강 단원 선택(전체 학생)")
    st.write("아래에서 학생들을 선택하고, 보강 단원을 선택한 뒤 적용하면 됩니다. (선택 안 하면 빈칸)")

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

    with st.expander("학생별 보강 단원 개별 수정(선택사항)", expanded=False):
        pick = st.selectbox("학생 선택", options=students, key="t1_one_pick")
        one_topics = st.multiselect(
            "이 학생 보강 단원",
            options=REINFORCE_TOPICS,
            default=[t.strip() for t in st.session_state.reinforce_map.get(pick, "").split(",") if t.strip()],
            key="t1_one_topics",
        )
        if st.button("이 학생만 저장", key="t1_one_save"):
            st.session_state.reinforce_map[pick] = ", ".join(one_topics).strip()
            st.success("저장 완료")

    st.divider()

    # 생성 버튼
    if st.button("🚀 Report 생성 (PDF + PNG)", type="primary", use_container_width=True, key="t1_build"):
        if not _ensure_fonts():
            st.error("폰트가 없어서 PDF 생성이 안 돼요. fonts 폴더 확인!")
            st.stop()

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_paths = []
            png_paths = []
            prog = st.progress(0)

            for i, stu in enumerate(students):
                row = students_df[students_df[name_col] == stu].iloc[0]

                # 학생/평균 값 가져오기
                def get_val(col):
                    v = row.get(col, "")
                    return _clean(v) if _clean(v) != "" else "-"

                def get_avg(col):
                    if avg_row is None:
                        return "-"
                    v = avg_row.get(col, "")
                    # 평균은 소수점 2자리 표시
                    if _is_number(v):
                        return f"{float(v):.2f}"
                    vv = _clean(v)
                    return vv if vv != "" else "-"

                quizzes = [(c, get_val(c), get_avg(c)) for c in quiz_cols]
                mocks = [(c, get_val(c), get_avg(c)) for c in mock_cols]
                homework = [(c, get_val(c)) for c in hw_cols]

                reinforce_text = st.session_state.reinforce_map.get(stu, "")

                pdf_name = f"{_safe_filename(class_name)}_{_safe_filename(stu)}.pdf"
                png_name = f"{_safe_filename(class_name)}_{_safe_filename(stu)}.png"
                pdf_path = os.path.join(tmpdir, pdf_name)
                png_path = os.path.join(tmpdir, png_name)

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

                # PNG 렌더
                png_bytes = _pdf_first_page_to_png(pdf_path, zoom=2.2)
                with open(png_path, "wb") as f:
                    f.write(png_bytes)

                pdf_paths.append(pdf_path)
                png_paths.append(png_path)

                prog.progress((i + 1) / len(students))

            # ZIP 만들기
            pdf_zip = io.BytesIO()
            with zipfile.ZipFile(pdf_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for p in pdf_paths:
                    z.write(p, arcname=os.path.basename(p))
            pdf_zip.seek(0)

            png_zip = io.BytesIO()
            with zipfile.ZipFile(png_zip, "w", compression=zipfile.ZIP_DEFLATED) as z:
                for p in png_paths:
                    z.write(p, arcname=os.path.basename(p))
            png_zip.seek(0)

            st.success(f"✅ 생성 완료: PDF {len(pdf_paths)}개 / PNG {len(png_paths)}개")

            # 다운로드 먼저
            d1, d2 = st.columns([1, 1])
            with d1:
                st.download_button(
                    "📦 PDF ZIP 다운로드",
                    data=pdf_zip,
                    file_name=f"{_safe_filename(class_name)}_CLASS_REPORT_PDF.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="t1_dl_pdf",
                )
            with d2:
                st.download_button(
                    "🖼️ PNG ZIP 다운로드",
                    data=png_zip,
                    file_name=f"{_safe_filename(class_name)}_CLASS_REPORT_PNG.zip",
                    mime="application/zip",
                    use_container_width=True,
                    key="t1_dl_png",
                )

            # 미리보기는 1명만
            st.markdown("#### 미리보기(1명)")
            preview_student = st.selectbox("미리볼 학생", options=students, key="t1_preview")
            prev_png = os.path.join(tmpdir, f"{_safe_filename(class_name)}_{_safe_filename(preview_student)}.png")
            # 임시폴더가 버튼 누른 컨텍스트라서 여기에서 바로 보여주기 위해 bytes를 다시 만들자
            # (zip 만들 때 이미 파일이 존재하지만, tmpdir scope 끝나면 사라짐 → 지금은 scope 안이라 OK)
            if os.path.exists(prev_png):
                st.image(prev_png, use_container_width=True)
            else:
                st.info("미리보기 이미지를 찾지 못했어요.")


# =========================
# 3. Tab2: Mock1~3 중복 단원 분석
# =========================
def _norm_module(v):
    s = _clean(v).upper().replace(" ", "")
    if s in ["M1", "MODULE1", "1"]:
        return 1
    if s in ["M2", "MODULE2", "2"]:
        return 2
    return None


def _major_topic(topic_str: str) -> str:
    s = _clean(topic_str)
    if s == "" or s.lower() == "nan":
        return ""
    m = re.match(r"^\s*(\d+)", s)
    if m:
        return m.group(1)
    rm = re.match(r"^\s*([IVX]+)\.?", s.upper())
    if rm:
        return rm.group(1)
    return s


def _display_major(k: str) -> str:
    k = _clean(k)
    return TOPIC_NAMES_MAJOR.get(k, k)


def _detect_mock_wrong_cols_export_mock(df: pd.DataFrame) -> dict[int, dict[int, str]]:
    """
    EXPORT_MOCK의 컬럼명:
      Mocktest1 m1 틀린문제
      Mocktest1 m2 틀린문제
      ...
    """
    out = defaultdict(dict)
    pat = re.compile(r"mocktest\s*(\d+).*(m1|m2).*(틀린\s*문제|틀린문제)", re.IGNORECASE)
    for c in df.columns:
        s = _clean(c).replace("\n", " ")
        m = pat.search(s)
        if not m:
            continue
        mock_no = int(m.group(1))
        mod = 1 if m.group(2).lower() == "m1" else 2
        out[mock_no][mod] = c
    return dict(out)


def _read_mock_meta(file) -> dict[int, dict[int, str]]:
    """
    메타: 모듈 / 문항번호 / 단원
    return {1:{q:topic}, 2:{q:topic}}
    """
    df = pd.read_excel(file)
    df.columns = [_clean(c) for c in df.columns]
    need = {"모듈", "문항번호", "단원"}
    if not need.issubset(set(df.columns)):
        raise ValueError(f"메타 파일에 필요한 컬럼 누락: {sorted(list(need))} / 현재: {list(df.columns)}")

    mapping = {1: {}, 2: {}}
    for _, r in df.iterrows():
        mod = _norm_module(r.get("모듈"))
        if mod not in (1, 2):
            continue
        try:
            q = int(float(str(r.get("문항번호")).strip()))
        except:
            continue
        t = _clean(r.get("단원"))
        if t:
            mapping[mod][q] = t
    return mapping


def _df_to_excel_bytes(df: pd.DataFrame, sheet="RESULT") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name=sheet)
    buf.seek(0)
    return buf.getvalue()


def tab2_repeated_topics():
    st.subheader("Tab2) Mock 1~3 중복 약점 단원 분석")
    st.caption("EXPORT_MOCK + Mock1/Mock2/Mock3 메타 업로드 → 학생별로 중복해서 틀린 단원을 추천합니다.")

    c1, c2, c3 = st.columns([1.1, 1.1, 1])
    with c1:
        rep_threshold = st.selectbox("몇 개 Mock에서 겹치면 추천할까?", [2, 3], index=0, key="t2_thr")
    with c2:
        use_major = st.checkbox("단원은 큰 단원(major)로 묶기", value=True, key="t2_major")
    with c3:
        st.write("")

    st.divider()

    left, right = st.columns([1, 1])
    with left:
        export_mock = st.file_uploader("1) EXPORT_MOCK 업로드 (.xlsx)", type=["xlsx"], key="t2_export_mock")
    with right:
        st.markdown("2) Mock 메타 업로드 (Mock1/2/3 각각)")
        meta1 = st.file_uploader("Mock1 메타 (.xlsx)", type=["xlsx"], key="t2_meta1")
        meta2 = st.file_uploader("Mock2 메타 (.xlsx)", type=["xlsx"], key="t2_meta2")
        meta3 = st.file_uploader("Mock3 메타 (.xlsx)", type=["xlsx"], key="t2_meta3")

    st.divider()

    if st.button("🚀 분석하기", type="primary", use_container_width=True, key="t2_run"):
        if export_mock is None:
            st.error("EXPORT_MOCK를 먼저 업로드해줘.")
            st.stop()

        try:
            df = pd.read_excel(export_mock)
            df.columns = [_clean(c) for c in df.columns]
        except Exception as e:
            st.error(f"EXPORT_MOCK 읽기 실패: {e}")
            st.stop()

        if "Name" not in df.columns:
            st.error("EXPORT_MOCK에 Name 컬럼이 필요해요.")
            st.stop()
        if "Class" not in df.columns:
            df["Class"] = ""

        df["Name"] = df["Name"].astype(str).map(_clean)
        df = df[(df["Name"] != "") & (df["Name"] != "평균")].copy()
        if df.empty:
            st.error("학생 데이터가 비어있어요.")
            st.stop()

        mock_wrong_cols = _detect_mock_wrong_cols_export_mock(df)
        if not mock_wrong_cols:
            st.error("EXPORT_MOCK에서 'MocktestN m1 틀린문제 / m2 틀린문제' 컬럼을 찾지 못했어요.")
            st.stop()

        # 메타 업로드된 것만 사용
        uploaded = {1: meta1, 2: meta2, 3: meta3}
        meta_files = {k: v for k, v in uploaded.items() if v is not None}
        if len(meta_files) < 2 and rep_threshold >= 2:
            st.warning("중복(>=2) 분석은 메타 파일이 최소 2개 이상 업로드되어야 의미가 좋아요. (권장)")

        mock_metas = {}
        for mock_no, f in meta_files.items():
            try:
                mock_metas[mock_no] = _read_mock_meta(f)
            except Exception as e:
                st.error(f"Mock{mock_no} 메타 읽기 실패: {e}")
                st.stop()

        # 실제 계산은 업로드된 메타 번호와 export 컬럼 번호의 교집합만
        available_mocks = sorted(set(mock_wrong_cols.keys()) & set(mock_metas.keys()))
        if not available_mocks:
            st.error("EXPORT_MOCK에 있는 Mock 번호와 업로드한 메타(Mock1/2/3)가 매칭되지 않아요.")
            st.stop()

        rows = []
        for _, r in df.iterrows():
            cls = _clean(r.get("Class", ""))
            name = _clean(r.get("Name", ""))

            per_mock_topics = {}
            for mk in available_mocks:
                cols = mock_wrong_cols.get(mk, {})
                col_m1 = cols.get(1)
                col_m2 = cols.get(2)

                wrong_m1 = _parse_wrong_list(r.get(col_m1)) if col_m1 else set()
                wrong_m2 = _parse_wrong_list(r.get(col_m2)) if col_m2 else set()

                meta = mock_metas[mk]
                topics = set()

                for q in wrong_m1:
                    t = meta.get(1, {}).get(q, "")
                    if t:
                        topics.add(_major_topic(t) if use_major else _clean(t))
                for q in wrong_m2:
                    t = meta.get(2, {}).get(q, "")
                    if t:
                        topics.add(_major_topic(t) if use_major else _clean(t))

                topics = {t for t in topics if _clean(t) != ""}
                per_mock_topics[mk] = topics

            # mock 몇 개에서 등장했는지 카운트
            cnt = Counter()
            for mk, tset in per_mock_topics.items():
                for t in tset:
                    cnt[t] += 1

            repeated = [t for t, ccc in cnt.items() if ccc >= int(rep_threshold)]
            repeated.sort(key=lambda x: (-cnt[x], str(x)))

            if use_major:
                repeated_disp = [_display_major(t) for t in repeated]
            else:
                repeated_disp = repeated

            rows.append({
                "Class": cls,
                "Name": name,
                f"중복 단원(>= {rep_threshold} mocks)": ", ".join(repeated_disp),
                "중복 단원 카운트": ", ".join([f"{(_display_major(t) if use_major else t)}({cnt[t]})" for t in repeated]),
            })

        out = pd.DataFrame(rows)
        st.subheader("✅ 학생별 중복 약점 단원")
        st.dataframe(out, use_container_width=True, hide_index=True)

        # 다운로드
        xlsx_bytes = _df_to_excel_bytes(out, sheet="REPEATED_TOPICS")
        st.download_button(
            "⬇️ 결과 엑셀 다운로드",
            data=xlsx_bytes,
            file_name="mock_repeated_topics.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="t2_dl_xlsx",
        )

        csv_bytes = out.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "⬇️ 결과 CSV 다운로드",
            data=csv_bytes,
            file_name="mock_repeated_topics.csv",
            mime="text/csv",
            use_container_width=True,
            key="t2_dl_csv",
        )

    st.info(
        "메타 파일 형식(최소): **모듈 / 문항번호 / 단원**\n"
        "- 모듈: M1, M2 또는 1, 2\n"
        "- 문항번호: 1~22\n"
        "- 단원: 예) 5.3 / 7.1 / VII. Geometry 등\n\n"
        "EXPORT_MOCK 컬럼 예시:\n"
        "- Mocktest1 m1 틀린문제 / Mocktest1 m2 틀린문제\n"
        "- Mocktest2 m1 틀린문제 / Mocktest2 m2 틀린문제\n"
        "- Mocktest3 m1 틀린문제 / Mocktest3 m2 틀린문제"
    )


# =========================
# 4. 탭 구성
# =========================
tab1, tab2 = st.tabs(["📄 Class Report", "📌 Mock 중복 단원 분석"])

with tab1:
    tab1_class_report()

with tab2:
    tab2_repeated_topics()
