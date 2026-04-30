import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 추출기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기 (빈칸 유지형)")
st.write("**종목명(ISU_NM)**을 입력하면, 상장폐지/거래정지 등으로 데이터가 없는 기업도 제외하지 않고 **빈칸으로 남겨두어 엑셀 복사-붙여넣기에 최적화**된 결과를 제공합니다.")

# API 인증키
API_KEY = "E76EEC8AF3D142F2BCA4A0EDB7510FEC9DA32064"

# 1. 날짜 선택 및 기업명 입력
col1, col2 = st.columns([1, 2])
with col1:
    selected_date = st.date_input("조회 기준일자를 선택하세요", datetime.date.today())
with col2:
    input_names = st.text_area(
        "조회할 기업명을 입력하세요 (쉼표 또는 줄바꿈으로 구분)", 
        "삼성전자\n상장폐지기업예시\nSK하이닉스\n거래정지기업예시"
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
                    api_df = pd.DataFrame(combined_data)
                    
                    # 🔥 [핵심 수정] 사용자가 입력한 순서 그대로 기본 뼈대 만들기
                    input_df = pd.DataFrame({'ISU_NM': target_companies})
                    
                    # API 데이터에서 필요한 컬럼만 추출 (혹시 모를 중복 제거)
                    if not api_df.empty and 'ISU_NM' in api_df.columns:
                        api_filtered = api_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].drop_duplicates(subset=['ISU_NM'])
                        
                        # 입력된 기업명 뼈대에 API 데이터를 병합 (Left Join)
                        # API에 없는 기업은 TDD_CLSPRC, MKTCAP 값이 NaN(결측치)으로 들어감
                        result_df = pd.merge(input_df, api_filtered, on='ISU_NM', how='left')
                    else:
                        # API 데이터가 정상적이지 않을 경우 모두 빈칸 처리
                        result_df = input_df
                        result_df['TDD_CLSPRC'] = ""
                        result_df['MKTCAP'] = ""
                    
                    # NaN(결측치)를 빈칸("")으로 변경하여 화면 및 엑셀에서 깔끔하게 보이도록 처리
                    result_df.fillna("", inplace=True)
                    
                    # 컬럼명 한글로 변경
                    result_df.columns = ['종목명', '종가', '시가총액']
                    
                    st.success(f"총 {len(result_df)}개의 입력값을 모두 처리했습니다. (API에서 찾은 데이터: {sum(result_df['종가'] != '')}건)")
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
