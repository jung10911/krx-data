import streamlit as st
import pandas as pd
import requests
from io import BytesIO

# 페이지 기본 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 조회기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기")
st.write("**종목명(ISU_NM)**을 여러 개 입력하면, 해당 기업들의 **종가(TDD_CLSPRC)**와 **시가총액(MKTCAP)**을 조회하고 **엑셀(Excel)**로 다운로드할 수 있습니다.")

# API 인증키 설정
# (주의: 깃허브 업로드 시 아래 줄을 지우고 API_KEY = st.secrets["KRX_API_KEY"] 로 변경하세요!)
API_KEY = "E76EEC8AF3D142F2BCA4A0EDB7510FEC9DA32064"

# 복수 기업명 입력 섹션
input_names = st.text_area(
    "조회할 기업명을 입력하세요 (쉼표 또는 줄바꿈으로 구분)", 
    "삼성전자, SK하이닉스, NAVER"
)

if st.button("데이터 조회 및 엑셀 생성"):
    # 1. 입력된 문자열을 리스트로 변환 및 공백 제거
    target_companies = [name.strip() for name in input_names.replace('\n', ',').split(',') if name.strip()]

    if not target_companies:
        st.warning("최소 1개 이상의 기업명을 입력해주세요.")
    else:
        with st.spinner("KRX Open API에서 데이터를 불러오는 중..."):
            try:
                # 2. KRX API 호출 설정
                # ※ 주의: 신청하신 정확한 API 엔드포인트 URL로 반드시 변경해야 합니다. 
                # (아래는 전종목 시세를 가져오는 가상의 기본 URL 예시입니다)
                url = "https://data-dbg.krx.co.kr/svc/apis/sto/stk_isur_QtiqIsuList" 
                headers = {
                    "AUTH_KEY": API_KEY,
                    "Content-Type": "application/json"
                }
                
                # API 요청 (필요시 날짜 등 params 추가)
                response = requests.get(url, headers=headers)
                response.raise_for_status()
                data = response.json()

                # 3. 사진의 구조에 맞춘 데이터 추출 (OutBlock_1)
                stock_list = data.get("OutBlock_1", [])
                
                if not stock_list:
                    st.error("API에서 데이터를 불러오지 못했습니다. URL이나 호출 방식을 확인해주세요.")
                else:
                    df = pd.DataFrame(stock_list)
                    
                    # 4. 사용자가 입력한 종목명(ISU_NM)으로 복수 필터링
                    filtered_df = df[df['ISU_NM'].isin(target_companies)]
                    
                    if filtered_df.empty:
                        st.warning("입력하신 기업명과 일치하는 데이터가 시장에 없습니다.")
                    else:
                        # 5. 핵심단어 데이터만 선택 및 컬럼명 한글화
                        result_df = filtered_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].copy()
                        result_df.columns = ['종목명', '종가', '시가총액']
                        
                        st.success("조회 완료!")
                        st.dataframe(result_df, use_container_width=True)
                        
                        # 6. 엑셀(Excel) 다운로드 기능 구현
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='KRX_Data')
                        
                        st.download_button(
                            label="📥 엑셀(Excel) 파일 다운로드",
                            data=output.getvalue(),
                            file_name="krx_stock_data.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
            except Exception as e:
                st.error(f"오류가 발생했습니다: {e}")
