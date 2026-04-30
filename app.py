import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 추출기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기")
st.write("**종목명(ISU_NM)**을 입력하여 **종가(TDD_CLSPRC)**와 **시가총액(MKTCAP)** 데이터를 추출합니다.")

# API 인증키
API_KEY = "E76EEC8AF3D142F2BCA4A0EDB7510FEC9DA32064"

# 기업명 입력
input_names = st.text_area(
    "조회할 기업명을 입력하세요 (쉼표 또는 줄바꿈으로 구분)", 
    "삼성전자, SK하이닉스"
)

if st.button("데이터 조회 및 엑셀 생성"):
    target_companies = [name.strip() for name in input_names.replace('\n', ',').split(',') if name.strip()]

    if not target_companies:
        st.warning("조회할 기업명을 입력해주세요.")
    else:
        with st.spinner("KRX 데이터를 불러오는 중..."):
            try:
                # 오늘 날짜로 조회
                now = datetime.datetime.now()
                search_date = now.strftime("%Y%m%d")
                
                # 전종목 시세 API URL
                url = "https://data-dbg.krx.co.kr/svc/apis/sto/stk_isur_qtiq_list" 
                
                headers = {
                    "AUTH_KEY": API_KEY
                }
                params = {
                    "basDd": search_date
                }
                
                response = requests.get(url, headers=headers, params=params)
                
                if response.status_code == 404:
                    st.error("API 주소를 찾을 수 없습니다(404).")
                else:
                    response.raise_for_status()
                    data = response.json()

                    stock_list = data.get("OutBlock_1", [])
                    
                    if not stock_list:
                        st.error("해당 날짜에 데이터가 없습니다. (주말/공휴일이 아닌지 확인하세요)")
                    else:
                        df = pd.DataFrame(stock_list)
                        
                        # 종목명 필터링
                        filtered_df = df[df['ISU_NM'].isin(target_companies)]
                        
                        if filtered_df.empty:
                            st.warning("일치하는 종목명이 없습니다.")
                        else:
                            # 핵심단어 추출
                            result_df = filtered_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].copy()
                            result_df.columns = ['종목명', '종가', '시가총액']
                            
                            st.success(f"{len(result_df)}건의 데이터를 찾았습니다.")
                            st.dataframe(result_df, use_container_width=True)
                            
                            # 엑셀 다운로드
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
