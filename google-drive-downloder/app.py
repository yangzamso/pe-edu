from __future__ import annotations

import re
import tempfile
from io import BytesIO
from pathlib import Path
from urllib.parse import parse_qs, urlparse
from zipfile import ZIP_DEFLATED, ZipFile

import pandas as pd
import requests
import streamlit as st


EXPECTED_COLUMNS = ("고유번호", "파일")
CHUNK_SIZE = 32768


def sanitize_filename(value: object) -> str:
    text = str(value).strip()
    text = re.sub(r'[<>:"/\\|?*]+', "_", text)
    return text.rstrip(". ") or "unnamed"


def extract_drive_file_id(url: str) -> str | None:
    parsed = urlparse(url)
    query_id = parse_qs(parsed.query).get("id")
    if query_id:
        return query_id[0]

    patterns = [
        r"/file/d/([a-zA-Z0-9_-]+)",
        r"/d/([a-zA-Z0-9_-]+)",
    ]
    for pattern in patterns:
        match = re.search(pattern, url)
        if match:
            return match.group(1)
    return None


def load_dataframe(source) -> pd.DataFrame:
    df = pd.read_excel(source)
    missing = [column for column in EXPECTED_COLUMNS if column not in df.columns]
    if missing:
        raise ValueError(f"필수 컬럼이 없습니다: {', '.join(missing)}")
    return df[list(EXPECTED_COLUMNS)].dropna(how="all")


def extract_filename_from_headers(response: requests.Response) -> str | None:
    content_disposition = response.headers.get("content-disposition", "")
    filename_matches = [
        re.search(r"filename\*=UTF-8''([^;]+)", content_disposition, re.IGNORECASE),
        re.search(r'filename="([^"]+)"', content_disposition, re.IGNORECASE),
        re.search(r"filename=([^;]+)", content_disposition, re.IGNORECASE),
    ]
    for match in filename_matches:
        if match:
            return requests.utils.unquote(match.group(1).strip().strip('"'))
    return None


def get_confirm_token(response: requests.Response) -> str | None:
    for cookie_name, cookie_value in response.cookies.items():
        if cookie_name.startswith("download_warning"):
            return cookie_value

    text = response.text
    patterns = [
        r'confirm=([0-9A-Za-z_]+)',
        r'"downloadUrl":"[^"]*confirm=([0-9A-Za-z_]+)',
        r'name="confirm" value="([0-9A-Za-z_]+)"',
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1)
    return None


def detect_google_drive_access_issue(response: requests.Response) -> str | None:
    final_url = str(response.url)
    content_type = response.headers.get("content-type", "").lower()

    if "accounts.google.com" in final_url:
        return "구글 로그인이 필요한 링크입니다. 파일 공유를 '링크가 있는 모든 사용자'로 변경해야 합니다."

    if "text/html" not in content_type:
        return None

    text = response.text[:20000]
    if "You need access" in text or "액세스 권한이 필요" in text:
        return "구글드라이브 접근 권한이 없습니다. 파일 공유 권한을 확인하세요."
    if "Google Drive - Virus scan warning" in text or "too large for Google to scan" in text:
        return None
    if "Google Drive" in text and "download" not in text.lower():
        return "구글드라이브가 파일 대신 안내 페이지를 반환했습니다. 공유 링크 형식과 권한을 확인하세요."
    return None


def fetch_google_drive_file(file_id: str) -> tuple[requests.Response, str]:
    session = requests.Session()
    base_url = "https://drive.google.com/uc"
    params = {"export": "download", "id": file_id}

    response = session.get(base_url, params=params, stream=True, timeout=60)
    response.raise_for_status()

    access_issue = detect_google_drive_access_issue(response)
    if access_issue:
        raise ValueError(access_issue)

    content_type = response.headers.get("content-type", "")
    if "text/html" in content_type.lower():
        confirm_token = get_confirm_token(response)
        if confirm_token:
            response.close()
            response = session.get(
                base_url,
                params={**params, "confirm": confirm_token},
                stream=True,
                timeout=60,
            )
            response.raise_for_status()

            access_issue = detect_google_drive_access_issue(response)
            if access_issue:
                raise ValueError(access_issue)

    original_name = extract_filename_from_headers(response)
    if not original_name:
        raise ValueError("원본 파일명을 구글드라이브 응답에서 찾지 못했습니다.")
    return response, original_name


def save_original_file(url: str, target_stem: str, workspace: Path) -> Path:
    file_id = extract_drive_file_id(url)
    if not file_id:
        raise ValueError("구글드라이브 파일 ID를 링크에서 찾지 못했습니다.")

    response, original_name = fetch_google_drive_file(file_id)
    original_suffix = "".join(Path(original_name).suffixes) or Path(original_name).suffix
    output_path = workspace / f"{target_stem}{original_suffix}"

    with response:
        with output_path.open("wb") as file_obj:
            for chunk in response.iter_content(CHUNK_SIZE):
                if chunk:
                    file_obj.write(chunk)

    return output_path


def build_zip(file_paths: list[Path]) -> BytesIO:
    buffer = BytesIO()
    with ZipFile(buffer, "w", compression=ZIP_DEFLATED) as zip_file:
        for path in file_paths:
            zip_file.write(path, arcname=path.name)
    buffer.seek(0)
    return buffer


st.set_page_config(page_title="구글 자료 다운로더", page_icon="📥", layout="centered")

st.title("구글 자료 다운로더")
st.write(
    "엑셀 파일을 업로드하면 `고유번호`, `파일` 컬럼을 읽어 구글드라이브 원본 파일을 그대로 다운로드하고 "
    "`고유번호 + 원본 확장자` 이름으로 ZIP 파일을 생성합니다."
)
st.markdown(
    """
    - 다운로드할 구글드라이브 파일의 `공유` 버튼을 누른 뒤 `링크가 있는 모든 사용자`로 변경하고 `뷰어` 권한으로 설정해주세요.
    - 공유 설정을 변경하지 않으면 다운로드가 되지 않습니다.
    - 다운로드가 모두 완료된 뒤 약 15초 정도 지나면 `ZIP 다운로드` 버튼이 나타날 수 있습니다. 버튼을 눌러 압축 파일을 받은 뒤 압축을 풀어서 사용해주세요.
    """
)

uploaded_file = st.file_uploader("엑셀 파일 업로드", type=["xlsx"])

if uploaded_file is None:
    st.info("엑셀 파일을 업로드한 뒤 내용을 확인하고 `다운로드 시작` 버튼을 눌러주세요.")
    st.stop()

st.caption(f"업로드한 파일: {uploaded_file.name}")

try:
    preview_df = load_dataframe(uploaded_file)
except Exception as exc:
    st.error(f"엑셀을 읽는 중 오류가 발생했습니다: {exc}")
    st.stop()

st.dataframe(preview_df.head(10), use_container_width=True)
st.caption(f"총 {len(preview_df)}건")

if st.button("다운로드 시작", type="primary"):
    results: list[dict[str, str]] = []
    downloaded_files: list[Path] = []
    failed_items: list[dict[str, str]] = []
    progress = st.progress(0.0)
    status_box = st.empty()
    failed_box = st.empty()

    failed_box.info("실패한 항목이 생기면 여기에 표시됩니다.")

    with tempfile.TemporaryDirectory() as temp_dir:
        workspace = Path(temp_dir)

        for index, row in preview_df.reset_index(drop=True).iterrows():
            unique_id = sanitize_filename(row["고유번호"])
            url = str(row["파일"]).strip()
            status_box.write(f"[{index + 1}/{len(preview_df)}] {unique_id} 처리 중")

            try:
                saved_path = save_original_file(url, unique_id, workspace)
                downloaded_files.append(saved_path)
                results.append(
                    {
                        "고유번호": unique_id,
                        "결과": "성공",
                        "저장파일": saved_path.name,
                    }
                )
            except Exception as exc:
                failed_item = {
                    "고유번호": unique_id,
                    "실패사유": str(exc),
                }
                failed_items.append(failed_item)
                results.append(
                    {
                        "고유번호": unique_id,
                        "결과": f"실패: {exc}",
                        "저장파일": "",
                    }
                )
                failed_box.dataframe(pd.DataFrame(failed_items), use_container_width=True)

            progress.progress((index + 1) / len(preview_df))

        result_df = pd.DataFrame(results)
        success_count = int((result_df["결과"] == "성공").sum())
        fail_count = len(result_df) - success_count

        status_box.write("처리가 완료되었습니다.")
        st.dataframe(result_df, use_container_width=True)

        if downloaded_files:
            zip_buffer = build_zip(downloaded_files)
            st.download_button(
                "ZIP 다운로드",
                data=zip_buffer,
                file_name="downloaded_files.zip",
                mime="application/zip",
            )

        if success_count:
            st.success(f"{success_count}건 다운로드 성공")
        if fail_count:
            st.warning(f"{fail_count}건 다운로드 실패")
            failed_df = pd.DataFrame(failed_items)
            st.subheader("실패 목록")
            st.dataframe(failed_df, use_container_width=True)
            st.download_button(
                "실패 목록 CSV 다운로드",
                data=failed_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="failed_items.csv",
                mime="text/csv",
            )
