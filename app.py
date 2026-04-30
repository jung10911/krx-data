import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 추출기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기")
st.write("**종목명(ISU_NM)**을 입력하여 **종가(TDD_CLSPRC)**와 **시가총액(MKTCAP)** 데이터를 추출합니다.")

# 1. API 인증키 및 설정
# 깃허브 배포 시 st.secrets["KRX_API_KEY"] 사용 권장
API_KEY = "E76EEC8AF3D142F2BCA4A0EDB7510FEC9DA32064"

# 2. 사용자 입력 섹션
input_names = st.text_area(
    "조회할 기업명을 입력하세요 (쉼표 또는 줄바꿈으로 구분)", 
    "삼성전자, SK하이닉스"
)

# 3. 데이터 조회 로직
if st.button("데이터 조회 및 엑셀 생성"):
    # 입력값 정리
    target_companies = [name.strip() for name in input_names.replace('\n', ',').split(',') if name.strip()]

    if not target_companies:
        st.warning("조회할 기업명을 입력해주세요.")
    else:
        with st.spinner("KRX 데이터를 불러오는 중..."):
            try:
                # KRX API는 '기준일자(basDd)'가 필수입니다. 
                # 장 종료 전이라면 어제 날짜로, 종료 후라면 오늘 날짜로 시도합니다.
                now = datetime.datetime.now()
                search_date = now.strftime("%Y%m%d")
                
                # [중요] 404 에러 방지를 위한 정확한 URL (전종목 시세 기준)
                # API 서비스에 따라 주소 끝부분이 다를 수 있으니 마이페이지에서 확인 필수입니다.
                url = "https://data-dbg.krx.co.kr/svc/apis/sto/stk_isur_qtiq_list" 
                
                headers = {
                    "AUTH_KEY": API_KEY
                }
                
                # 필수 파라미터 추가
                params = {
                    "basDd": search_date
                }
                
                response = requests.get(url, headers=headers, params=params)
                
                # 404 등 오류 발생 시 예외 처리
                if response.status_code == 404:
                    st.error("API 주소를 찾을 수 없습니다(404). KRX 마이페이지에서 'Request URL'을 다시 확인해 주세요.")
                else:
                    response.raise_for_status()
                    data = response.json()

                    # 4. 데이터 추출 및 필터링 (OutBlock_1)
                    stock_list = data.get("OutBlock_1", [])
                    
                    if not stock_list:
                        st.error("해당 날짜에 데이터가 없습니다. (주말/공휴일 확인)")
                    else:
                        df = pd.DataFrame(stock_list)
                        
                        # 종목명(ISU_NM) 일치 여부 확인
                        filtered_df = df[df['ISU_NM'].isin(target_companies)]
                        
                        if filtered_df.empty:
                            st.warning("일치하는 종목명이 없습니다. 정확한 이름을 입력했는지 확인하세요.")
                        else:
                            # 필요한 컬럼만 추출 (종목명, 종가, 시가총액)
                            result_df = filtered_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].copy()
                            result_df.columns = ['종목명', '종가', '시가총액']
                            
                            # 결과 출력
                            st.success(f"{len(result_df)}건의 데이터를 찾았습니다.")
                            st.dataframe(result_df, use_container_width=True)
                            
                            # 5. 엑셀 다운로드 파일 생성
                            output = BytesIO()
                            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                                result_df.to_excel(writer, index=False, sheet_name='KRX_Data')
                            
                            st.download_button(
                                label="📥 결과 엑셀 다운로드",
                                data=output.getvalue(),
                                file_name=f"KRX_Data_{search_date}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                            
            except Exception as e:
                st.error(f"오류 발생: {e}")
pandas
requests
xlsxwriter
