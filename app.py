import streamlit as st
import pandas as pd
import io
import traceback
import xlsxwriter

st.set_page_config(page_title="출장 여비지급명세서 생성기", layout="wide")

st.title("📄 여비지급명세서 자동 생성 웹 서비스")
st.markdown("공무원분들이 안전하고 편리하게 인터넷 창에서 바로 사용할 수 있는 버전입니다. \n파일을 열고 직원을 선택하면 미리보기 후 깔끔한 엑셀(수기입력 포함)을 다운로드할 수 있습니다.")

with st.expander("📖 공무원 여비 규정 지급 기준 요약 (2025.12.30. 개정, 2026.1.2. 적용 기준)", expanded=False):
    st.markdown("""
    신규 임용자나 타 부서에서 오신 분들의 빠른 업무 파악을 위해, 프로그램이 어떻게 금액을 판단하는지 **[법정 여비 지급 기준]** 생략본을 안내해 드립니다!
    
    ### 🚙 1. 관내 출장 (근무지내 또는 12km 미만)
    - **4시간 이상**: 기본 **20,000원** (공용차량 사용 시 10,000원 차감 ➜ `10,000원` 지급)
    - **4시간 미만**: 기본 **10,000원** (공용차량 사용 시 10,000원 차감 ➜ `0원` 지급)
    
    ### 🚅 2. 관외 출장 (근무지외)
    *관련 근거: 「공무원 여비 규정」 제16조제1항 및 [별표 2]*
    - **일비**: 1일당 **25,000원** (공용차량 사용 시 50% 감액 ➜ 1일당 `12,500원` 지급)
    - **식비**: 1일당 **25,000원**
    - **숙박비 및 기타 교통비**: 실비 상한 한도 내 정산 내역이 다르므로, 엑셀 파일 다운로드 후 영수증에 맞게 **직접 수기 입력**
    
    > 💡 **참고**: 엑셀 추출 시 프로그램이 알아서 시간과 일수('O일'), 차량이용 여부를 스캔하여 위의 금액을 맞춰줍니다. (만약 예외 사항이 있다면 다운로드한 엑셀에서 숫자만 수정하시면 합계에 그대로 자동 반영됩니다.)
    """)

uploaded_file = st.file_uploader("1. 인사랑 출장내역서 엑셀 파일 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])

@st.cache_data
def process_raw_data(file_bytes):
    df = pd.read_excel(file_bytes, header=None)
    
    header_r_idx = -1
    col_map = {}
    for r_idx in range(min(15, len(df))):
        row_vals = df.iloc[r_idx].fillna('').astype(str).tolist()
        row_vals_clean = [str(x).replace(" ", "") for x in row_vals]
        
        if "성명" in row_vals_clean and "부서" in row_vals_clean and "구분" in row_vals_clean:
            header_r_idx = r_idx
            for c_idx, val in enumerate(row_vals_clean):
                if val: col_map[val] = c_idx
            break
            
    if header_r_idx == -1:
        return None, "엑셀에서 '성명', '부서', '구분' 등의 기본 항목 제목을 찾을 수 없습니다."
        
    name_c_idx = col_map.get("성명")
    data_start_idx = header_r_idx + 1
    
    for r_idx in range(header_r_idx + 1, len(df)):
        val = str(df.iloc[r_idx, name_c_idx]).strip()
        if val and val != "nan" and val not in ["일자", "일시", "사유", "지정시간"]:
            data_start_idx = r_idx
            break
            
    data_df = df.iloc[data_start_idx:].copy()
    
    dept_idx = col_map.get("부서", -1)
    rank_idx = col_map.get("직급", -1)
    gubun_idx = col_map.get("구분", -1)
    vehicle_idx = col_map.get("공무용차량", -1)
    purpose_idx = col_map.get("출장목적", -1)
    dest_idx = col_map.get("출장지", -1)
    time_idx = col_map.get("총출장시간", -1)
    grade_idx = col_map.get("여비등급", -1)
    start_dt_idx = col_map.get("출장시작", -1) 
    end_dt_idx = col_map.get("출장종료", -1)
    
    parsed_data = []
    for _, row in data_df.iterrows():
        name = str(row.iloc[name_c_idx]).strip()
        if not name or name == "nan": continue
        
        def get_val(idx):
            if idx == -1: return ""
            v = str(row.iloc[idx]).strip()
            return "" if v == "nan" else v
            
        department = get_val(dept_idx)
        rank = get_val(rank_idx)
        gubun = get_val(gubun_idx)
        vehicle = get_val(vehicle_idx)
        purpose = get_val(purpose_idx)
        dest = get_val(dest_idx)
        total_time = get_val(time_idx)
        grade = get_val(grade_idx)
        
        start_date = get_val(start_dt_idx)
        start_time = get_val(start_dt_idx + 1) if start_dt_idx != -1 else ""
        end_date = get_val(end_dt_idx)
        end_time = get_val(end_dt_idx + 1) if end_dt_idx != -1 else ""
        period_str = f"{start_date} {start_time}\n~ {end_date} {end_time}".strip()
        
        parsed_data.append({
            "성명": name,
            "부서": department,
            "직급": rank,
            "구분": gubun,
            "공무용차량": vehicle,
            "출장목적": purpose,
            "출장지": dest,
            "총출장시간": total_time,
            "여비등급": grade,
            "출장기간": period_str,
            "sort_key": f"{start_date} {start_time}"
        })
        
    final_df = pd.DataFrame(parsed_data)
    if not final_df.empty:
        final_df = final_df.sort_values(by="sort_key", ascending=True).drop(columns=["sort_key"]).reset_index(drop=True)
    return final_df, None

def write_sheet(workbook, df, title, is_empty=False):
    worksheet = workbook.add_worksheet(title[:31])
    
    # 인쇄 옵션 (가로방향, 여백, 너비맞춤)
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)
    
    font_name = '맑은 고딕'
    num_fmt = '#,##0;-#,##0;"-"'  # 음수/0 표기용 확실한 회계서식
    
    format_title = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'font_name': font_name})
    format_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7', 'font_name': font_name})
    format_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': font_name, 'text_wrap': True})
    format_money = workbook.add_format({'num_format': num_fmt, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    format_money_bold = workbook.add_format({'bold': True, 'num_format': num_fmt, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    format_total_label = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    format_claim = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': font_name})
    
    # 열 너비 세팅 ('현지교통비' 빠지면서 17열로 감소, 0~16)
    # 요청하신 수치 반영 (출장기간: 18 / 여비등급~합계: 9 / 청구액: 13)
    widths = [5, 25, 8, 18, 10, 10, 10, 20, 30, 9, 9, 9, 9, 9, 9, 9, 13]
    for i, w in enumerate(widths):
        worksheet.set_column(i, i, w)
        
    worksheet.merge_range(0, 0, 1, 16, title, format_title)
    
    headers = [
        "순번", "소속/직급", "성명", "출장기간", "총출장시간", "구분", "공용차량", "출장지",
        "출장목적", "여비등급", "일비", "식비", "숙박비", "교통비", "기타", "합계", "청구액(수령액)"
    ]
    for col_num, header in enumerate(headers):
        worksheet.write(3, col_num, header, format_header)
        
    start_row = 4 
    current_row = start_row
    
    if is_empty or len(df) == 0:
        worksheet.merge_range(4, 0, 4, 16, "해당 조건의 출장 내역이 없습니다.", format_center)
        return
        
    for i, (_, row) in enumerate(df.iterrows()):
        gubun = str(row['구분']).strip()
        time_str = str(row['총출장시간']).strip()
        vehicle = str(row['공무용차량']).strip()
        
        ilbi = 0
        sikbi = 0
        
        if gubun == '근무지내' or gubun == '관내':
            hours = 0
            if '시간' in time_str:
                hours_part = time_str.split('시간')[0]
                if hours_part.isdigit():
                    hours = int(hours_part)
            
            base = 20000 if hours >= 4 else 10000
            if vehicle == '사용':
                base -= 10000
            ilbi = max(base, 0)
            
        elif gubun == '근무지외' or gubun == '관외':
            # 관외출장의 경우 총출장시간이 'O일' 인지 파악하여 일수 계산 (기본 1일)
            days = 1
            if '일' in time_str:
                days_part = time_str.split('일')[0].strip()
                if days_part.isdigit():
                    days = int(days_part)
                    
            # 개정된 공무원 여비 규정(2025. 12. 30. 개정, 26년 적용) 제16조제1항 단서 반영
            # 일비: 기본 25,000원 / 공무용차량 사용 시 50% 감액 (12,500원)
            daily_ilbi = 25000
            if vehicle == '사용':
                daily_ilbi = 12500 
                
            ilbi = max(daily_ilbi, 0) * days
            
            # 식비: 기본 25,000원 단가 배정
            sikbi = 25000 * days
            
        worksheet.write(current_row, 0, i + 1, format_center)
        worksheet.write(current_row, 1, f"{row.get('부서', '')}\n{row.get('직급', '')}", format_center)
        worksheet.write(current_row, 2, str(row.get('성명', '')), format_center)
        worksheet.write(current_row, 3, str(row.get('출장기간', '')), format_center)
        worksheet.write(current_row, 4, time_str, format_center)
        worksheet.write(current_row, 5, gubun, format_center)
        worksheet.write(current_row, 6, vehicle, format_center)
        worksheet.write(current_row, 7, str(row.get('출장지', '')), format_center)
        worksheet.write(current_row, 8, str(row.get('출장목적', '')), format_center)
        worksheet.write(current_row, 9, str(row.get('여비등급', '')), format_center)
        
        worksheet.write_number(current_row, 10, ilbi, format_money) # 일비
        worksheet.write_number(current_row, 11, sikbi, format_money) # 식비 (표 기준 기본단가 자동기입)
        worksheet.write_number(current_row, 12, 0, format_money)    # 숙박비
        worksheet.write_number(current_row, 13, 0, format_money)    # 교통비 (노란 바탕 없앰)
        worksheet.write_number(current_row, 14, 0, format_money)    # 기타 ("현지교통비" 아예 날아감)
        
        # 합계: K열(10)부터 O열(14)까지 더해서 P열(15)에 넣음
        formula_str = f"=SUM(K{current_row+1}:O{current_row+1})" 
        worksheet.write_formula(current_row, 15, formula_str, format_money)
        
        current_row += 1

    total_row = current_row
    excel_total_row_num = total_row + 1 
    
    # --- 소계 행 ---
    formula_count = f'="소        계     (총 "&COUNTA(A{start_row+1}:A{excel_total_row_num-1})&" 건)"'
    worksheet.merge_range(total_row, 0, total_row, 8, "", format_total_label) 
    worksheet.write_formula(total_row, 0, formula_count, format_total_label)
    
    worksheet.write(total_row, 9, "", format_total_label) # 여비등급
    
    # 10~15열 (일비 ~ 합계) - 총합 수식
    for col_idx in range(10, 16):
        col_letter = chr(ord('A') + col_idx) 
        sum_formula = f"=SUM({col_letter}{start_row+1}:{col_letter}{excel_total_row_num-1})"
        worksheet.write_formula(total_row, col_idx, sum_formula, format_money_bold)

    # 16열 (청구액) 세로 병합 (합계: P열/15번의 합)
    target_name = str(df.iloc[0]['성명']).strip() if not df.empty else ""
    claim_formula = f'="{target_name} (인)" & CHAR(10) & CHAR(10) & TEXT(SUM(P{start_row+1}:P{excel_total_row_num-1}), "#,##0") & "원"'
    
    worksheet.merge_range(start_row, 16, total_row, 16, "", format_claim)
    worksheet.write_formula(start_row, 16, claim_formula, format_claim)

if uploaded_file is not None:
    try:
        full_df, err = process_raw_data(uploaded_file.getvalue())
        
        if err:
            st.error(err)
        else:
            st.success(f"✅ 파일 인식 성공! 분석된 데이터 총 {len(full_df)}건")
            
            names_list = full_df['성명'].unique().tolist()
            names_list.sort() 
            
            st.markdown("---")
            st.subheader("2. 작업 대상자 선택")
            
            selected_name = st.selectbox("어떤 직원분의 출장 내역을 작업하시겠습니까?", options=["선택해주세요..."] + names_list)
            
            if selected_name and selected_name != "선택해주세요...":
                user_df = full_df[full_df['성명'] == selected_name].copy()
                
                in_count = len(user_df[user_df['구분'].str.contains('근무지내|관내', na=False)])
                out_count = len(user_df[user_df['구분'].str.contains('근무지외|관외', na=False)])
                
                st.info(f"**{selected_name}**님의 이번 달 출장 내역: 총 **{len(user_df)}건** 준비됨 (관내 {in_count}건 / 관외 {out_count}건)")
                
                st.subheader("👀 출장 내역 추출 미리보기")
                st.dataframe(user_df, use_container_width=True, hide_index=True)
                
                st.markdown("---")
                st.subheader("3. 결과 엑셀 저장 옵션")
                
                col1, col2 = st.columns(2)
                with col1:
                    report_style = st.radio("출력 방식 선택", ["관내/관외 통합", "관내/관외 분리 (시트 분리)"])
                
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                
                if report_style == "관내/관외 통합":
                    write_sheet(workbook, user_df, "통합 여비지급명세서")
                else:
                    in_df = user_df[user_df['구분'].str.contains('근무지내|관내', na=False)]
                    out_df = user_df[user_df['구분'].str.contains('근무지외|관외', na=False)]
                    
                    if not in_df.empty:
                        write_sheet(workbook, in_df, "관내 여비지급명세서")
                    elif "분리" in report_style:
                        write_sheet(workbook, pd.DataFrame(), "관내 여비지급명세서", is_empty=True)
                        
                    if not out_df.empty:
                        write_sheet(workbook, out_df, "관외 여비지급명세서")
                    elif "분리" in report_style:
                        write_sheet(workbook, pd.DataFrame(), "관외 여비지급명세서", is_empty=True)
                
                workbook.close()
                processed_data = output.getvalue()
                
                with col2:
                    st.write("")
                    st.write("")
                    st.download_button(
                        label=f"📥 {selected_name}_여비지급명세서 다운로드",
                        data=processed_data,
                        file_name=f"{selected_name}_여비지급명세서_{'통합' if report_style == '관내/관외 통합' else '분리'}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                st.caption("* 다운로드한 엑셀 파일에서 '식비', '교통비' 등 임의의 열의 수치를 수기 변경하면 합계가 꼬임 없이 자동으로 재계산됩니다.")
                    
    except Exception as e:
        st.error(f"오류가 발생했습니다: {str(e)}")
        st.code(traceback.format_exc())
