# PE ID 이미지 다운로더

`data.xlsx` 또는 업로드한 엑셀 파일의 `고유번호`, `파일` 컬럼을 읽어서 구글드라이브 파일을 다운로드하고, 파일명을 `고유번호.확장자`로 저장한 뒤 ZIP으로 내려받는 스트림릿 앱입니다.

## 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

## 엑셀 형식

- `고유번호`: 저장할 파일명
- `파일`: 구글드라이브 공유 링크

예시:

| 고유번호 | 파일 |
| --- | --- |
| 211 - 01 - 202603 | https://drive.google.com/open?id=... |

