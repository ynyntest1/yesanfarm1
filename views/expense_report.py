import streamlit as st
import pandas as pd
import io
import traceback
import xlsxwriter


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

    dept_idx     = col_map.get("부서", -1)
    rank_idx     = col_map.get("직급", -1)
    gubun_idx    = col_map.get("구분", -1)
    vehicle_idx  = col_map.get("공무용차량", -1)
    purpose_idx  = col_map.get("출장목적", -1)
    dest_idx     = col_map.get("출장지", -1)
    time_idx     = col_map.get("총출장시간", -1)
    grade_idx    = col_map.get("여비등급", -1)
    start_dt_idx = col_map.get("출장시작", -1)
    end_dt_idx   = col_map.get("출장종료", -1)

    parsed_data = []
    for _, row in data_df.iterrows():
        name = str(row.iloc[name_c_idx]).strip()
        if not name or name == "nan": continue

        def get_val(idx):
            if idx == -1: return ""
            v = str(row.iloc[idx]).strip()
            return "" if v == "nan" else v

        department = get_val(dept_idx)
        rank       = get_val(rank_idx)
        gubun      = get_val(gubun_idx)
        vehicle    = get_val(vehicle_idx)
        purpose    = get_val(purpose_idx)
        dest       = get_val(dest_idx)
        total_time = get_val(time_idx)
        grade      = get_val(grade_idx)

        start_date = get_val(start_dt_idx)
        start_time = get_val(start_dt_idx + 1) if start_dt_idx != -1 else ""
        end_date   = get_val(end_dt_idx)
        end_time   = get_val(end_dt_idx + 1) if end_dt_idx != -1 else ""
        period_str = f"{start_date} {start_time}\n~ {end_date} {end_time}".strip()

        parsed_data.append({
            "성명": name, "부서": department, "직급": rank,
            "구분": gubun, "공무용차량": vehicle, "출장목적": purpose,
            "출장지": dest, "총출장시간": total_time, "여비등급": grade,
            "출장기간": period_str
        })

    return pd.DataFrame(parsed_data), None


def write_sheet(workbook, df, title, is_empty=False):
    worksheet = workbook.add_worksheet(title[:31])
    worksheet.set_landscape()
    worksheet.fit_to_pages(1, 0)
    worksheet.set_margins(left=0.25, right=0.25, top=0.75, bottom=0.75)

    font_name = '맑은 고딕'
    num_fmt   = '#,##0;-#,##0;"-"'

    fmt_title  = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'valign': 'vcenter', 'font_name': font_name})
    fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#DDEBF7', 'font_name': font_name})
    fmt_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': font_name, 'text_wrap': True})
    fmt_money  = workbook.add_format({'num_format': num_fmt, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    fmt_money_bold = workbook.add_format({'bold': True, 'num_format': num_fmt, 'align': 'right', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    fmt_total  = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_name': font_name})
    fmt_claim  = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_name': font_name})

    widths = [5, 25, 8, 18, 10, 10, 10, 20, 30, 9, 9, 9, 9, 9, 9, 9, 13]
    for i, w in enumerate(widths):
        worksheet.set_column(i, i, w)

    worksheet.merge_range(0, 0, 1, 16, title, fmt_title)

    headers = ["순번", "소속/직급", "성명", "출장기간", "총출장시간", "구분", "공용차량",
               "출장지", "출장목적", "여비등급", "일비", "식비", "숙박비", "교통비", "기타", "합계", "청구액(수령액)"]
    for col_num, header in enumerate(headers):
        worksheet.write(3, col_num, header, fmt_header)

    start_row   = 4
    current_row = start_row

    if is_empty or len(df) == 0:
        worksheet.merge_range(4, 0, 4, 16, "해당 조건의 출장 내역이 없습니다.", fmt_center)
        return

    for i, (_, row) in enumerate(df.iterrows()):
        gubun     = str(row['구분']).strip()
        time_str  = str(row['총출장시간']).strip()
        vehicle   = str(row['공무용차량']).strip()
        ilbi = sikbi = 0

        if gubun in ('근무지내', '관내'):
            hours = 0
            if '시간' in time_str:
                p = time_str.split('시간')[0]
                if p.isdigit(): hours = int(p)
            base = 20000 if hours >= 4 else 10000
            if vehicle == '사용': base -= 10000
            ilbi = max(base, 0)

        elif gubun in ('근무지외', '관외'):
            days = 1
            if '일' in time_str:
                p = time_str.split('일')[0].strip()
                if p.isdigit(): days = int(p)
            daily = 12500 if vehicle == '사용' else 25000
            ilbi  = daily * days
            sikbi = 25000 * days

        worksheet.write(current_row, 0, i + 1, fmt_center)
        worksheet.write(current_row, 1, f"{row.get('부서','')}\n{row.get('직급','')}", fmt_center)
        worksheet.write(current_row, 2, str(row.get('성명','')), fmt_center)
        worksheet.write(current_row, 3, str(row.get('출장기간','')), fmt_center)
        worksheet.write(current_row, 4, time_str, fmt_center)
        worksheet.write(current_row, 5, gubun, fmt_center)
        worksheet.write(current_row, 6, vehicle, fmt_center)
        worksheet.write(current_row, 7, str(row.get('출장지','')), fmt_center)
        worksheet.write(current_row, 8, str(row.get('출장목적','')), fmt_center)
        worksheet.write(current_row, 9, str(row.get('여비등급','')), fmt_center)
        worksheet.write_number(current_row, 10, ilbi,  fmt_money)
        worksheet.write_number(current_row, 11, sikbi, fmt_money)
        worksheet.write_number(current_row, 12, 0, fmt_money)
        worksheet.write_number(current_row, 13, 0, fmt_money)
        worksheet.write_number(current_row, 14, 0, fmt_money)
        worksheet.write_formula(current_row, 15, f"=SUM(K{current_row+1}:O{current_row+1})", fmt_money)
        current_row += 1

    total_row = current_row
    etr = total_row + 1
    worksheet.merge_range(total_row, 0, total_row, 8, "", fmt_total)
    worksheet.write_formula(total_row, 0,
        f'="소        계     (총 "&COUNTA(A{start_row+1}:A{etr-1})&" 건)"', fmt_total)
    worksheet.write(total_row, 9, "", fmt_total)
    for col_idx in range(10, 16):
        col_letter = chr(ord('A') + col_idx)
        worksheet.write_formula(total_row, col_idx,
            f"=SUM({col_letter}{start_row+1}:{col_letter}{etr-1})", fmt_money_bold)

    target_name = str(df.iloc[0]['성명']).strip() if not df.empty else ""
    worksheet.merge_range(start_row, 16, total_row, 16, "", fmt_claim)
    worksheet.write_formula(start_row, 16,
        f'="{target_name} (인)"&CHAR(10)&CHAR(10)&TEXT(SUM(P{start_row+1}:P{etr-1}),"#,##0")&"원"',
        fmt_claim)


def show():
    st.title("📄 여비지급명세서 자동 생성")
    st.markdown("인사랑 출장내역서 엑셀 파일을 업로드하면 여비지급명세서를 자동으로 생성합니다.")

    with st.expander("📖 공무원 여비 규정 지급 기준 요약 (2025.12.30. 개정, 2026.1.2. 적용)", expanded=False):
        st.markdown("""
        ### 🚙 1. 관내 출장 (근무지내 또는 12km 미만)
        - **4시간 이상**: 기본 **20,000원** (공용차량 사용 시 10,000원 차감 → `10,000원`)
        - **4시간 미만**: 기본 **10,000원** (공용차량 사용 시 10,000원 차감 → `0원`)

        ### 🚅 2. 관외 출장 (근무지외)
        - **일비**: 1일당 **25,000원** (공용차량 사용 시 50% 감액 → `12,500원`)
        - **식비**: 1일당 **25,000원**
        - **숙박비 및 교통비**: 다운로드 후 수기 입력
        """)

    uploaded_file = st.file_uploader("1. 인사랑 출장내역서 엑셀 파일 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])

    if uploaded_file is not None:
        try:
            full_df, err = process_raw_data(uploaded_file.getvalue())
            if err:
                st.error(err)
            else:
                st.success(f"✅ 파일 인식 성공! 총 {len(full_df)}건")
                names_list = sorted(full_df['성명'].unique().tolist())

                st.markdown("---")
                st.subheader("2. 작업 대상자 선택")
                selected_name = st.selectbox("직원 선택", options=["선택해주세요..."] + names_list)

                if selected_name and selected_name != "선택해주세요...":
                    user_df   = full_df[full_df['성명'] == selected_name].copy()
                    in_count  = len(user_df[user_df['구분'].str.contains('근무지내|관내', na=False)])
                    out_count = len(user_df[user_df['구분'].str.contains('근무지외|관외', na=False)])

                    st.info(f"**{selected_name}**님: 총 **{len(user_df)}건** (관내 {in_count}건 / 관외 {out_count}건)")
                    st.subheader("👀 출장 내역 미리보기")
                    st.dataframe(user_df, use_container_width=True, hide_index=True)

                    st.markdown("---")
                    st.subheader("3. 결과 엑셀 저장 옵션")
                    col1, col2 = st.columns(2)
                    with col1:
                        report_style = st.radio("출력 방식", ["관내/관외 통합", "관내/관외 분리 (시트 분리)"])

                    output   = io.BytesIO()
                    workbook = xlsxwriter.Workbook(output, {'in_memory': True})

                    if report_style == "관내/관외 통합":
                        write_sheet(workbook, user_df, "통합 여비지급명세서")
                    else:
                        in_df  = user_df[user_df['구분'].str.contains('근무지내|관내', na=False)]
                        out_df = user_df[user_df['구분'].str.contains('근무지외|관외', na=False)]
                        if not in_df.empty:
                            write_sheet(workbook, in_df, "관내 여비지급명세서")
                        else:
                            write_sheet(workbook, pd.DataFrame(), "관내 여비지급명세서", is_empty=True)
                        if not out_df.empty:
                            write_sheet(workbook, out_df, "관외 여비지급명세서")
                        else:
                            write_sheet(workbook, pd.DataFrame(), "관외 여비지급명세서", is_empty=True)

                    workbook.close()

                    with col2:
                        st.write("")
                        st.write("")
                        st.download_button(
                            label=f"📥 {selected_name}_여비지급명세서 다운로드",
                            data=output.getvalue(),
                            file_name=f"{selected_name}_여비지급명세서_{'통합' if report_style == '관내/관외 통합' else '분리'}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            type="primary"
                        )
                    st.caption("* 다운로드한 엑셀에서 숫자를 직접 수정하면 합계가 자동으로 재계산됩니다.")

        except Exception as e:
            st.error(f"오류가 발생했습니다: {str(e)}")
            st.code(traceback.format_exc())
