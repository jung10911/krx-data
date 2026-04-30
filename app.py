import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 추출기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기 (코스피/코스닥 통합)")
st.write("**종목명(ISU_NM)**을 입력하여 **종가(TDD_CLSPRC)**와 **시가총액(MKTCAP)** 데이터를 추출합니다.")

# API 인증키
API_KEY = "E76EEC8AF3D142F2BCA4A0EDB7510FEC9DA32064"

# 1. 날짜 선택 및 기업명 입력
col1, col2 = st.columns([1, 2])
with col1:
    selected_date = st.date_input("조회 기준일자를 선택하세요", datetime.date.today())
with col2:
    input_names = st.text_area(
        "조회할 기업명을 입력하세요 (쉼표 또는 줄바꿈으로 구분)", 
        "나우로보틱스\n쎄크\n에이유브랜즈\n더즌"
    )

if st.button("데이터 조회 및 엑셀 생성"):
    # 입력한 순서대로 리스트 생성 (target_companies)
    target_companies = [name.strip() for name in input_names.replace('\n', ',').split(',') if name.strip()]

    if not target_companies:
        st.warning("조회할 기업명을 최소 1개 이상 입력해주세요.")
    else:
        with st.spinner("KRX 코스피 및 코스닥 데이터를 모두 불러오는 중..."):
            try:
                search_date_str = selected_date.strftime("%Y%m%d")
                
                url_kospi = "https://data-dbg.krx.co.kr/svc/apis/sto/stk_bydd_trd"
                url_kosdaq = "https://data-dbg.krx.co.kr/svc/apis/sto/ksq_bydd_trd"
                
                headers = {"AUTH_KEY": API_KEY}
                params = {"basDd": search_date_str}
                
                def fetch_market_data(url):
                    response = requests.get(url, headers=headers, params=params)
                    if response.status_code == 200:
                        data = response.json()
                        return data.get("OutBlock_1", [])
                    else:
                        return []

                kospi_data = fetch_market_data(url_kospi)
                kosdaq_data = fetch_market_data(url_kosdaq)
                
                combined_data = kospi_data + kosdaq_data
                
                if not combined_data:
                    st.error(f"{selected_date} 일자의 데이터가 없습니다. (주말, 공휴일이거나 데이터 업데이트 전일 수 있습니다.)")
                else:
                    df = pd.DataFrame(combined_data)
                    
                    # 1차 필터링: 입력한 기업명만 남기기
                    filtered_df = df[df['ISU_NM'].isin(target_companies)].copy()
                    
                    if filtered_df.empty:
                        st.warning("입력하신 기업과 일치하는 종목이 코스피/코스닥 양쪽 시장 모두에 없습니다.")
                    else:
                        # 🔥 [핵심 추가] 사용자가 입력한 순서대로 정렬하기
                        # ISU_NM 컬럼을 target_companies 순서를 가진 '카테고리'로 변환 후 정렬
                        filtered_df['ISU_NM'] = pd.Categorical(filtered_df['ISU_NM'], categories=target_companies, ordered=True)
                        filtered_df = filtered_df.sort_values('ISU_NM').reset_index(drop=True)
                        
                        # 핵심단어(종목명, 종가, 시가총액) 추출
                        result_df = filtered_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].copy()
                        result_df.columns = ['종목명', '종가', '시가총액']
                        
                        st.success(f"총 {len(result_df)}건의 데이터를 성공적으로 찾았습니다!")
                        st.dataframe(result_df, use_container_width=True)
                        
                        # 엑셀 다운로드
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            result_df.to_excel(writer, index=False, sheet_name='KRX_Data')
                        
                        st.download_button(
                            label="📥 결과 엑셀 다운로드",
                            data=output.getvalue(),
                            file_name=f"KRX_Stock_Data_{search_date_str}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        
            except Exception as e:
                st.error(f"데이터 조회 중 오류가 발생했습니다: {e}")
