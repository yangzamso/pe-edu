from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st


# ─── 유틸 함수 ──────────────────────────────────────────────────────────────

def parse_zip_range(filename: str) -> tuple[int, int] | None:
    """파일명 어디서든 숫자3-숫자3 패턴 추출. 예: 체육_001-050.zip → (1, 50)"""
    stem = Path(filename).stem
    match = re.search(r"(\d{3})-(\d{3})", stem)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None


def sanitize(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r'[/\\:*?"<>|]+', "_", text)
    return text.rstrip(". ") or "unnamed"


def make_stem(row: pd.Series, cols: list[str], sep: str) -> str:
    return sep.join(sanitize(row[c]) for c in cols)


def numeric_sort_key(name: str) -> tuple[int, str]:
    stem = Path(name).stem
    try:
        return (int(stem), "")
    except ValueError:
        return (10**9, stem)


def jpg_files_in_zip(zf: zipfile.ZipFile) -> list[str]:
    return [
        n for n in zf.namelist()
        if not n.endswith("/") and Path(n).suffix.lower() in (".jpg", ".jpeg", ".png", ".pdf")
    ]


# ─── 페이지 설정 ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="파일명 일괄 변경기", page_icon="✏️", layout="centered")
st.title("✏️ 파일명 일괄 변경기")
st.caption("엑셀 명단 기준으로 ZIP 안의 파일명을 일괄 변경합니다.")

# ─── STEP 1: 엑셀 업로드 ────────────────────────────────────────────────────

st.header("1단계 · 엑셀 명단 업로드")
excel_upload = st.file_uploader("엑셀 파일 (.xlsx)", type=["xlsx"])

if excel_upload is None:
    st.info("엑셀 파일을 먼저 업로드해주세요.")
    st.stop()

try:
    df = pd.read_excel(excel_upload, header=0)
except Exception as e:
    st.error(f"엑셀 읽기 오류: {e}")
    st.stop()

st.caption(f"총 **{len(df)}행** 로드 (헤더 제외)")
st.dataframe(df.head(5), use_container_width=True)

# ─── STEP 2: 파일명 조합 설정 ───────────────────────────────────────────────

st.header("2단계 · 파일명 조합 설정")

all_cols = list(df.columns)
c1, c2 = st.columns([3, 1])
with c1:
    selected = st.multiselect(
        "사용할 열 선택 (선택한 순서 = 파일명 순서)",
        options=all_cols,
        default=all_cols[:1] if all_cols else [],
        help="선택한 순서대로 구분자를 사이에 두고 파일명이 조합됩니다.",
    )
with c2:
    sep_label = st.selectbox("구분자", ["_ (언더바)", "- (하이픈)", "없음"], index=0)
    sep = {"_ (언더바)": "_", "- (하이픈)": "-", "없음": ""}[sep_label]

custom_suffix = st.text_input(
    "다운로드 ZIP 파일명 접미사 (선택)",
    placeholder="예: _변환완료  →  001-050_변환완료.zip",
)

if not selected:
    st.warning("열을 1개 이상 선택해주세요.")
    st.stop()

st.subheader("파일명 미리보기")
preview = df.head(5).copy()
preview["생성될 파일명"] = preview.apply(
    lambda r: make_stem(r, selected, sep) + ".jpg", axis=1
)
st.dataframe(preview[selected + ["생성될 파일명"]], use_container_width=True)

# ─── STEP 3: ZIP 업로드 ─────────────────────────────────────────────────────

st.header("3단계 · ZIP 파일 업로드")
zip_uploads = st.file_uploader(
    "ZIP 파일 (파일명 형식: 001-050.zip, 여러 개 동시 업로드 가능)",
    type=["zip"],
    accept_multiple_files=True,
)

if not zip_uploads:
    st.info("ZIP 파일을 업로드해주세요.")
    st.stop()

# ZIP 유효성 검사 & 매핑 테이블 구성
zip_rows: list[dict] = []
has_error = False

for uf in zip_uploads:
    rng = parse_zip_range(uf.name)
    if rng is None:
        st.error(f"❌ **{uf.name}**: 파일명 형식 오류 (올바른 형식 예시: 001-050.zip)")
        has_error = True
        continue

    start, end = rng
    uf.seek(0)

    try:
        with zipfile.ZipFile(uf) as zf:
            files_in_zip = sorted(jpg_files_in_zip(zf), key=numeric_sort_key)
            actual_count = len(files_in_zip)
    except zipfile.BadZipFile:
        st.error(f"❌ **{uf.name}**: ZIP 파일이 손상되었습니다.")
        has_error = True
        continue

    excel_max = len(df)

    if start - 1 >= excel_max:
        st.error(f"❌ **{uf.name}**: 범위 시작({start})이 엑셀 데이터 행 수({excel_max})를 초과합니다.")
        has_error = True
        continue

    expected_count = end - start + 1
    real_end = min(end, excel_max)
    process_count = min(actual_count, real_end - start + 1)

    if actual_count != expected_count:
        note = f"파일수 불일치 (예상 {expected_count}개, 실제 {actual_count}개) → {process_count}개 처리"
        status = "⚠️"
    else:
        note = ""
        status = "✅"

    zip_rows.append({
        "ZIP 파일": uf.name,
        "범위": f"{start} ~ {end}",
        "파일 수": actual_count,
        "처리 수": process_count,
        "엑셀 행": f"{start} ~ {real_end}행",
        "상태": status,
        "비고": note,
        "_start": start,
        "_end": end,
        "_process": process_count,
        "_uf": uf,
    })

if zip_rows:
    display_cols = ["ZIP 파일", "범위", "파일 수", "처리 수", "엑셀 행", "상태", "비고"]
    st.dataframe(pd.DataFrame(zip_rows)[display_cols], use_container_width=True)

if has_error:
    st.error("오류가 있는 ZIP 파일을 수정한 뒤 다시 업로드해주세요.")
    st.stop()

# ─── STEP 4: 변환 ───────────────────────────────────────────────────────────

st.header("4단계 · 변환 시작")

if st.button("🚀 변환 시작", type="primary"):
    all_results: list[dict] = []
    progress = st.progress(0.0)
    total_zips = len(zip_rows)

    merged = io.BytesIO()
    with zipfile.ZipFile(merged, "w", zipfile.ZIP_DEFLATED) as z_merged:
        for z_idx, info in enumerate(zip_rows):
            uf: io.BytesIO = info["_uf"]
            start: int = info["_start"]
            process_count: int = info["_process"]

            uf.seek(0)
            with zipfile.ZipFile(uf) as z_in:
                files_sorted = sorted(jpg_files_in_zip(z_in), key=numeric_sort_key)
                for i in range(process_count):
                    orig = files_sorted[i]
                    excel_idx = start - 1 + i
                    row = df.iloc[excel_idx]
                    ext = Path(orig).suffix
                    new_name = make_stem(row, selected, sep) + ext
                    z_merged.writestr(new_name, z_in.read(orig))
                    all_results.append({
                        "원본 ZIP": uf.name,
                        "원본 파일명": Path(orig).name,
                        "변경 후 파일명": new_name,
                    })
            progress.progress((z_idx + 1) / total_zips)

    merged.seek(0)
    overall_start = min(info["_start"] for info in zip_rows)
    overall_end = max(info["_end"] for info in zip_rows)
    default_dl_name = f"{overall_start:03d}-{overall_end:03d}.zip"
    dl_name = f"{Path(default_dl_name).stem}{custom_suffix}.zip" if custom_suffix else default_dl_name

    st.session_state["merged_data"] = merged.getvalue()
    st.session_state["dl_name"] = dl_name
    st.session_state["all_results"] = all_results

# 변환 완료 후 다운로드 버튼 표시 (session_state로 유지)
if "merged_data" in st.session_state:
    st.success(f"✅ 변환 완료 — 총 **{len(st.session_state['all_results'])}개** 파일 처리됨")

    st.subheader("ZIP 다운로드")
    st.download_button(
        label=f"⬇️ {st.session_state['dl_name']} 다운로드",
        data=st.session_state["merged_data"],
        file_name=st.session_state["dl_name"],
        mime="application/zip",
        key="dl_merged",
    )

    st.subheader("변환 결과 상세")
    st.dataframe(pd.DataFrame(st.session_state["all_results"]), use_container_width=True)
