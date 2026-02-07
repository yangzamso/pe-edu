"""
엑셀 자동채점 시스템 - Streamlit 웹 앱
"""
import streamlit as st
import datetime
import io
import pandas as pd
from grader import ExcelGrader

# 페이지 설정
st.set_page_config(
    page_title="P.E 자동 채점",
    page_icon="📊",
    layout="wide"
)

# 커스텀 CSS
st.markdown("""
    <style>
    h1 {
        font-size: 24px !important;
        line-height: 1.6 !important;
        padding-top: 1rem !important;
        padding-bottom: 0.5rem !important;
    }
    h2 {
        font-size: 20px !important;
        padding-top: 1rem !important;
        padding-bottom: 0.5rem !important;
    }
    h3 {
        font-size: 18px !important;
        padding-top: 1rem !important;
        padding-bottom: 0.5rem !important;
    }
    /* 메인 컨테이너 너비 제한 (1280px) */
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
        max-width: 1280px !important;
    }
    /* 버튼 스타일 조정 */
    .stButton button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

def main():
    st.title("📊 P.E 자동 채점")
    
    # 세션 상태 초기화
    if 'results_df' not in st.session_state:
        st.session_state.results_df = None
    if 'excel_data' not in st.session_state:
        st.session_state.excel_data = None
        
    sheet_info_text = None
    
    # 레이아웃 분할 (좌 1 : 우 2)
    left_col, right_col = st.columns([1, 2], gap="large")
    
    # --- 좌측 컬럼: 입력 및 액션 ---
    with left_col:
        st.subheader("1. 파일 데이터 입력")
        
        # 1. 파일 업로드
        uploaded_file = st.file_uploader(
            "채점할 엑셀 파일 (.xlsx)",
            type=['xlsx'],
            help="답안 시트가 포함된 엑셀 파일을 선택하세요."
        )
        
        if uploaded_file is not None:
            # 임시 파일 저장 (매번 새로 저장)
            temp_file_path = f"temp_{uploaded_file.name}"
            with open(temp_file_path, "wb") as f:
                f.write(uploaded_file.getvalue())
            
            # Grader 초기화
            grader = ExcelGrader(temp_file_path)
            
            if grader.load_workbook():
                sheet_name = grader.workbook.sheetnames[0]
                row_count = grader.answer_sheet.max_row
                # 2. 채점 실행 버튼
                st.subheader("2. 채점 실행")
                if st.button("🚀 채점 시작"):
                    with st.spinner("채점 중입니다..."):
                        try:
                            # 분석 및 결과 생성
                            results_df = grader.analyze_answer_sheet()
                            excel_data = grader.generate_scored_excel()
                            
                            # 세션에 저장
                            st.session_state.results_df = results_df
                            st.session_state.excel_data = excel_data
                            
                        except Exception as e:
                            st.error(f"오류 발생: {str(e)}")
            else:
                st.error("파일 로드 실패")

        # 3. 다운로드 버튼 (채점 결과가 있을 때만 표시)
        if st.session_state.excel_data is not None:
            st.subheader("3. 결과 다운로드")
            st.caption("원본 엑셀 양식을 유지하며, 채점 결과와 점수가 자동 계산되어 저장됩니다.")
            
            today_str = datetime.datetime.now().strftime("%Y.%m.%d")
            filename = f"PE-Training-Test-{today_str}.xlsx"
            
            st.download_button(
                label="📥 채점 결과 다운로드",
                data=st.session_state.excel_data.getvalue(),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

    # --- 우측 컬럼: 결과 대시보드 ---
    with right_col:
        st.subheader("📋 채점 결과")
        
        if st.session_state.results_df is not None:
            df = st.session_state.results_df.copy()
            
            # 통계 계산
            total_students = len(df)
            avg_score = df['총점(100점)'].mean()
            
            # 상단에 통계 정보 표시
            st.info(f"👥 총 **{total_students}명** 응시  |  📈 평균 점수: **{avg_score:.1f}점**")
            
            # 순번 컬럼 추가 (1부터 시작)
            df.insert(0, '순번', range(1, len(df) + 1))
            
            # 컬럼 순서 및 이름 정리
            display_cols = ['순번', '학생명', '객관식(25점)', '주관식(75점)', '총점(100점)']
            
            # 컬럼 설정 (공통 사용)
            column_configuration = {
                "순번": st.column_config.NumberColumn(
                    "순번",
                    width=20,
                    format="%d"
                ),
                "학생명": st.column_config.TextColumn(
                    "학생명",
                    width=180
                ),
                "객관식(25점)": st.column_config.NumberColumn(
                    "객관식(25점)",
                    format="%.1f"
                ),
                "주관식(75점)": st.column_config.NumberColumn(
                    "주관식(75점)",
                    format="%.1f"
                ),
                "총점(100점)": st.column_config.NumberColumn(
                    "총점(100점)",
                    format="%.1f"
                )
            }
            
            # 데이터프레임 표시 (컬럼 설정 추가)
            st.dataframe(
                df[display_cols],
                use_container_width=True,
                hide_index=True,
                height=600,
                column_config=column_configuration
            )
            
        else:
            # 데이터가 없을 때 안내 문구
            st.info("👈 왼쪽에서 파일을 업로드하고 '채점 시작' 버튼을 눌러주세요.")
            
            # 빈 테이블 프레임 보여주기
            empty_data = pd.DataFrame(columns=['순번', '학생명', '객관식(25점)', '주관식(75점)', '총점(100점)'])
            
            # 동일한 컬럼 설정 적용
            st.dataframe(
                empty_data, 
                use_container_width=True, 
                hide_index=True,
                column_config={
                    "순번": st.column_config.NumberColumn("순번", width=20),
                    "학생명": st.column_config.TextColumn("학생명", width=180),
                    "객관식(25점)": st.column_config.NumberColumn("객관식(25점)"),
                    "주관식(75점)": st.column_config.NumberColumn("주관식(75점)"),
                    "총점(100점)": st.column_config.NumberColumn("총점(100점)")
                }
            )

if __name__ == '__main__':
    main()
