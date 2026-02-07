@echo off
echo ====================================
echo 엑셀 자동채점 앱 실행
echo ====================================

REM 가상환경 활성화 (venv 폴더가 있을 경우)
if exist venv\Scripts\activate.bat (
    call venv\Scripts\activate.bat
)

REM Streamlit 앱 실행
streamlit run app.py
