import streamlit as st
import pandas as pd
import requests
import datetime
from io import BytesIO

# 페이지 설정
st.set_page_config(page_title="KRX 주가 및 시가총액 추출기", layout="wide")

st.title("📊 KRX 기업 주가 및 시가총액 추출기 (숫자 자동 포맷형)")
st.write("**종목명(ISU_NM)**을 입력하면, 빈칸은 유지하면서 엑셀 다운로드 시 **숫자(천 단위 콤마)** 형식으로 바로 출력되도록 자동 변환합니다.")

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
                    api_df = pd.DataFrame(combined_data)
                    
                    # 사용자가 입력한 순서 그대로 기본 뼈대 만들기
                    input_df = pd.DataFrame({'ISU_NM': target_companies})
                    
                    if not api_df.empty and 'ISU_NM' in api_df.columns:
                        api_filtered = api_df[['ISU_NM', 'TDD_CLSPRC', 'MKTCAP']].drop_duplicates(subset=['ISU_NM']).copy()
                        
                        # 🔥 [핵심 추가] 문자를 실제 숫자형 데이터로 완벽히 변환
                        api_filtered['TDD_CLSPRC'] = pd.to_numeric(api_filtered['TDD_CLSPRC'], errors='coerce')
                        api_filtered['MKTCAP'] = pd.to_numeric(api_filtered['MKTCAP'], errors='coerce')
                        
                        result_df = pd.merge(input_df, api_filtered, on='ISU_NM', how='left')
                    else:
                        result_df = input_df
                        result_df['TDD_CLSPRC'] = None
                        result_df['MKTCAP'] = None
                    
                    # 컬럼명 한글로 변경
                    result_df.columns = ['종목명', '종가', '시가총액']
                    
                    st.success(f"총 {len(result_df)}개의 기업 데이터를 준비했습니다!")
                    
                    # 스트림릿 화면에서도 콤마가 보이게 포맷팅 출력 (NaN 값은 빈칸 처리)
                    st.dataframe(result_df.style.format({'종가': '{:,.0f}', '시가총액': '{:,.0f}'}, na_rep=""), use_container_width=True)
                    
                    # 엑셀 다운로드 파일 생성
                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                        result_df.to_excel(writer, index=False, sheet_name='KRX_Data')
                        
                        # 🔥 [핵심 추가] xlsxwriter를 활용해 엑셀 파일 자체에 콤마 서식 강제 적용
                        workbook = writer.book
                        worksheet = writer.sheets['KRX_Data']
                        
                        # 숫자 서식 지정 (천 단위 콤마)
                        num_format = workbook.add_format({'num_format': '#,##0'})
                        
                        # B열(종가), C열(시가총액)의 열 너비를 15로 늘리고 숫자 서식 일괄 적용
                        worksheet.set_column('B:C', 15, num_format)
                        
                        # A열(종목명)은 글자가 잘리지 않게 열 너비를 20으로 여유 있게 세팅
                        worksheet.set_column('A:A', 20)
                    
                    st.download_button(
                        label="📥 결과 엑셀 다운로드",
                        data=output.getvalue(),
                        file_name=f"KRX_Stock_Data_{search_date_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                        
            except Exception as e:
                st.error(f"데이터 조회 중 오류가 발생했습니다: {e}")
