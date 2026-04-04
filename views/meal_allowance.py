import streamlit as st
import pandas as pd
import io
import xlsxwriter
import calendar
from datetime import datetime
from views.team_settings import load_settings, save_settings

def amount_to_korean(num):
    units = ["", "십", "백", "천"]
    nums = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
    
    def to_korean_chunk(n):
        s = str(n)
        l = len(s)
        res = ""
        for i, c in enumerate(s):
            if c != '0':
                res += nums[int(c)] + units[l - 1 - i]
        return res
    if num == 0: return "영"
    res = ""
    part = num % 10000
    if part > 0: res = to_korean_chunk(part)
    num //= 10000
    if num > 0:
        part = num % 10000
        if part > 0: res = to_korean_chunk(part) + "만" + res
    num //= 10000
    if num > 0:
        part = num % 10000
        if part > 0: res = to_korean_chunk(part) + "억" + res
    return res

def show():
    st.title("💰 급식비 자동 문서 작성")
    st.markdown("새올에서 다운로드한 **[초과근무목록 액셀 파일]**을 올리면, 내부결재용 공문 텍스트와 지출증빙 엑셀(2p, 3p)을 자동으로 생성합니다.")
    
    with st.expander("📖 급식비 지급 규정 요약 보기 (클릭하여 열기) *지방자치단체회계관리에관한훈령 제13조[별표2의2] (행정안전부)"):
        st.markdown("""
        **[급식비 지급 기준 및 필수 확인사항]**
        - 1. 집행 단가 및 방법
            - 💰 단가: 1인당 1식 9,000원 이내 집행
            - 💳 결제: 지방자치단체구매카드 사용 원칙
            - 🚫 주의: 카드 사용 불가 시 계좌이체(채권자 직접 송금)만 가능하며, 공무원 개인에게 현금 지급 절대 금지

        - 2. 지급 대상 및 요건 (초과근무 시)
            - ⏰ 시간 요건:(평일) 정규 근무 시작 1시간 전 출근 또는 종료 후 1시간 이상 근무자 / (휴일) 1시간 이상 근무자
            - 📑 증빙 필수: 초과근무실적, 출퇴근 인증, PC 접속기록 등 객관적 사실 확인 후 집행
            - 🕒 유연근무: 근무시간 전/후 1시간 이상 근무 시 지급 (단, 09:00~18:00 중에는 지급 제외)

        - 3. 지급 제외 대상 (중복 지급 방지)
            - ❌ 시간외근무수당 수령자 중 교대근무자
            - ❌ 야간근무수당·휴일근무수당 지급 대상자
            - ❌ 「공무원 여비 규정」에 따라 식비를 이미 지원받은 자(출장 등)
        """)
        
    uploaded_file = st.file_uploader("1. 새올 초과근무목록 엑셀 업로드 (.xls, .xlsx)", type=['xls', 'xlsx'])
    
    if uploaded_file:
        try:
            df = pd.read_excel(io.BytesIO(uploaded_file.getvalue()))
            
            # 실제 데이터 헤더 찾기
            if "성명" in df.columns and "출근(실제)" in df.columns:
                # pandas가 알아서 첫 줄을 헤더로 잘 읽어온 경우
                data_df = df.copy()
            else:
                # 헤더가 위에 다른 설명글로 밀려있는 경우 루프 돌면서 찾기
                header_idx = -1
                for r in range(15):
                    row_vals = [str(x).strip().replace(' ', '') for x in df.iloc[r].fillna('')]
                    if "성명" in row_vals and "출근(실제)" in row_vals:
                        header_idx = r
                        break
                
                if header_idx == -1:
                    st.error("엑셀에서 '성명', '출근(실제)' 열을 찾을 수 없습니다. 원본 파일을 확인해주세요.")
                    return
                    
                columns = df.iloc[header_idx].tolist()
                data_df = df.iloc[header_idx+1:].copy()
                data_df.columns = columns
            
            # 유효한 근무자만 필터링
            data_df = data_df[data_df["성명"].notna() & (data_df["성명"] != "") & (data_df["성명"].astype(str) != "nan")]
            
            
            st.success(f"데이터 총 {len(data_df)}건(식) 인식 완료!")
            
            # 신규 인원 대기열(unassigned) 및 직급 정보 자동 추가 로직
            t_set = load_settings()
            
            rank_col = next((c for c in data_df.columns if "직급" in str(c) or "직위" in str(c)), None)
            rank_dict = t_set.get("rank_dict", {})
            updated_settings = False
            
            if rank_col:
                for _, r in data_df.iterrows():
                    n = str(r.get("성명", "")).strip()
                    rk = str(r.get(rank_col, "")).strip()
                    if n and rk and rk not in ["nan", "None", ""]:
                        rank_dict[n] = rk
                t_set["rank_dict"] = rank_dict
                updated_settings = True
            
            current_org = t_set.get("org_data", {})
            current_u = t_set.get("unassigned", [])
            assigned_names = set()
            for d_val in current_org.values():
                for t_val in d_val.values():
                    for name in t_val:
                        assigned_names.add(name)
                        
            extracted_names = [str(x) for x in data_df["성명"].unique() if str(x) not in ["nan", "None", ""]]
            new_people = [n for n in extracted_names if n not in assigned_names and n not in current_u]
            if new_people:
                current_u.extend(new_people)
                t_set["unassigned"] = current_u
                updated_settings = True
                st.toast(f"엑셀에서 {len(new_people)}명의 새로운 인원을 발견하여 [팀 설정]의 대기열에 추가했습니다!")
                
            if updated_settings:
                save_settings(t_set)
            
            # ────────────────────────────────────────────────────────────────
            st.markdown("---")
            st.subheader("2. 대상자 필터링 (부서 및 인원 선택)")
            st.markdown("과를 먼저 선택하고 세부 팀을 고르면 소속 직원이 자동으로 모두 선택됩니다.(팀설정에서 변경가능)")
            
            # 단계 1: 과 단위 선택 (버튼형)
            if "selected_div" not in st.session_state:
                st.session_state["selected_div"] = None

            div_list = ["기술지원과", "스마트농업과", "미래농업과"]
            st.markdown("##### 🏢 과 단위 선택")
            cols_div = st.columns(len(div_list))
            for i, div_name in enumerate(div_list):
                with cols_div[i]:
                    btn_type = "primary" if st.session_state.get("selected_div") == div_name else "secondary"
                    # 버튼을 누르면 상태를 업데이트하고 페이지 새로고침
                    if st.button(div_name, use_container_width=True, type=btn_type, key=f"div_btn_{i}"):
                        st.session_state["selected_div"] = div_name
                        st.session_state["selected_team"] = None  # 과가 바뀌면 팀 초기화
                        st.rerun()

            selected_div = st.session_state.get("selected_div")
            if not selected_div:
                st.info("👆 위에서 소속 과 상자를 눌러주세요.")
                return
                
            # 팀 환경설정 데이터 먼저 불러오기 (조직도 기반 동적 렌더링을 위해)
            t_set = load_settings()
            
            # 부서 매핑 (조직도 정보에서 추출)
            loaded_org = t_set.get("org_data", {})
            org_chart = {div: list(teams.keys()) for div, teams in loaded_org.items()}
            
            # 만약 조직도가 비어있다면 기본값 부여
            if not org_chart:
                org_chart = {
                    "기술지원과": ["기획운영팀", "인력육성팀", "농기계팀", "농업지원팀"],
                    "스마트농업과": ["식량작물팀", "과수기술팀", "스마트원예팀", "스마트팜지원팀"],
                    "미래농업과": ["축산개발팀", "귀농지원팀", "농촌자원팀", "먹거리지원팀", "학교급식팀"]
                }
            
            filtered_depts = org_chart.get(selected_div, [])
                
            # 단계 2: 팀 선택 (버튼형)
            st.markdown(f"##### 🏷️ [{selected_div}] 소속 팀 선택")
            if "selected_team" not in st.session_state:
                st.session_state["selected_team"] = None
                
            cols_team = st.columns(len(filtered_depts) if filtered_depts else 1)
            for i, team_name in enumerate(filtered_depts):
                with cols_team[i % len(cols_team)]:
                    btn_type = "primary" if st.session_state.get("selected_team") == team_name else "secondary"
                    if st.button(team_name, use_container_width=True, type=btn_type, key=f"team_btn_{i}"):
                        st.session_state["selected_team"] = team_name
                        st.rerun()
                        
            selected_team = st.session_state.get("selected_team")
            if not selected_team:
                st.info("👆 위에서 소속 팀 상자를 눌러주세요.")
                return
            
            # 단계 3: 팀원 자동 선택 (팀 설정의 조직도 기반)
            team_members = loaded_org.get(selected_div, {}).get(selected_team, [])
            
            if not team_members:
                # 만약 조직도에 등록된 인원이 전혀 없다면 (기존 방식 유지)
                dept_filtered_df = data_df[data_df["부서"].astype(str).str.contains(selected_team, na=False)]
                all_names = [str(x) for x in dept_filtered_df["성명"].unique() if str(x) not in ["nan", "None", ""]]
            else:
                # 조직도에 등록된 이름 그대로 사용!
                all_names = team_members
            
            # 파란색 칩으로 매력적인 다중 선택
            selected_names = st.multiselect(
                "👤 이번 청구 대상 인원 확인 및 제외 (팀 전체가 기본으로 꽉 채워집니다)", 
                options=all_names, 
                default=[n for n in all_names if n in data_df["성명"].values]
            )
            
            # 엑셀 데이터 필터링: 선택된 명단에 해당하는 사람의 모든 초과근무 추출
            filtered_data_df = data_df[data_df["성명"].astype(str).isin(selected_names)]
            
            st.markdown("---")
            # 최소 수당시간 맞춤형 필터 (팀마다 1분, 60분 등 조건이 다름)
            st.markdown("##### ⏱️ 초과근무 수당시간 필터 조건 선택")
            if "min_hours_filter" not in st.session_state:
                st.session_state["min_hours_filter"] = 60
                
            f_col1, f_col2 = st.columns(2)
            with f_col1:
                btn_type_60 = "primary" if st.session_state["min_hours_filter"] == 60 else "secondary"
                if st.button("60분 이상 근무 내역만 추출 (기본)", use_container_width=True, type=btn_type_60, key="hour_filter_60"):
                    st.session_state["min_hours_filter"] = 60
                    st.rerun()
            with f_col2:
                btn_type_1 = "primary" if st.session_state["min_hours_filter"] == 1 else "secondary"
                if st.button("1분 이상 근무 내역 모두 추출", use_container_width=True, type=btn_type_1, key="hour_filter_1"):
                    st.session_state["min_hours_filter"] = 1
                    st.rerun()
                    
            min_hours = st.session_state["min_hours_filter"]
            filtered_data_df = filtered_data_df[pd.to_numeric(filtered_data_df["수당시간(분)"], errors='coerce').fillna(0) >= min_hours]
            
            if len(filtered_data_df) == 0:
                st.warning("현재 선택하신 부서, 인원, 그리고 시간 조건에 해당하는 초과근무 내역이 하나도 없습니다.")
                return
                
            # ────────────────────────────────────────────────────────────────
            # 팀 환경설정 데이터 불러오기
            t_set = load_settings()
            
            st.markdown("---")
            st.subheader("3. 서류 기본 정보 입력")
            
            # 조직도 정보를 바탕으로 확인자/작성자 인공지능 자동완성! (첫번째가 팀장, 마지막이 서기/담당)
            rank_map = t_set.get("rank_dict", {})
            
            default_confirmer = ""
            default_writer = ""
            if team_members:
                t_leader = team_members[0]
                l_rank = rank_map.get(t_leader, "지방농업주사")
                default_confirmer = f"{l_rank} {t_leader}"
                
                t_writer = team_members[-1]
                w_rank = rank_map.get(t_writer, "지방농업서기")
                default_writer = f"{w_rank} {t_writer}"
            
            col1, col2, col3, col4 = st.columns(4)
            with col1: team_name = st.text_input("팀명", value=selected_team)
            with col2: target_month = st.text_input("기준 월", value=str(datetime.now().month))
            with col3: confirm_person = st.text_input("확인자 (직급 이름)", value=default_confirmer, help="예: 지방농업주사 홍길동")
            with col4: write_person = st.text_input("작성자 (직급 이름)", value=default_writer, help="예: 지방농업서기 신사임당")
            # ────────────────────────────────────────────────────────────────
            st.markdown("---")
            st.subheader("4. 급식장소 상세 배정")
            total_cases = len(filtered_data_df)
            total_amount = total_cases * 9000
            st.warning(f"아래 표에 선택하신 **{total_cases}건**의 내역이 나타났습니다. /  💰 **이번 달 총 급식비 예상액: {total_amount:,}원** (단가 9,000원 기준) /   **'급식장소'** 열을 더블클릭하여 직접 입력해주세요.")
            
            disp_df = filtered_data_df[['근무일자', '부서', '성명', '휴일구분', '출근(실제)', '퇴근(실제)', '수당시간(분)', '근무내역']].copy()
            
            # 성명 순서를 팀원 등록 순서(team_members) 에 우선 맞추기 (1순위 정렬)
            if team_members:
                name_order = {name: i for i, name in enumerate(team_members)}
                disp_df['_name_order'] = disp_df['성명'].map(lambda x: name_order.get(x, 999))
            else:
                disp_df['_name_order'] = 0
                
            # 근무일자 오름차순 (2순위 정렬)
            disp_df = disp_df.sort_values(by=['_name_order', '근무일자']).drop(columns=['_name_order']).reset_index(drop=True)
            disp_df.index = disp_df.index + 1
            disp_df.insert(0, '순번', disp_df.index)
            
            # 다른 팀이나 인원을 지웠다 켰을때, 혹은 분 조건을 바꿀때 테이블 리셋
            current_view_key = f"{selected_div}_{selected_team}_{','.join(selected_names)}_{min_hours}"
            if st.session_state.get("current_view_id") != current_view_key:
                if "meal_data_editor" in st.session_state:
                    del st.session_state["meal_data_editor"]
                disp_df["급식장소"] = ""
                st.session_state["meal_df"] = disp_df
                st.session_state["current_view_id"] = current_view_key
                
            # 식당 이름 직접 입력 및 인덱스 숨기기 (dynamic 삭제하여 엔터 입력 가능하게 함)
            edited_df = st.data_editor(
                st.session_state["meal_df"], 
                key="meal_data_editor",
                use_container_width=True, 
                height=400,
                hide_index=True,
                disabled=['순번', '근무일자', '부서', '성명', '휴일구분', '출근(실제)', '퇴근(실제)', '수당시간(분)', '근무내역'],
                column_config={
                    "급식장소": st.column_config.TextColumn(
                        "급식장소 (직접 입력)",
                        help="엔터를 치면 바로 입력이 완료됩니다. 복사(Ctrl+C) 후 여러 칸을 잡고 붙여넣기(Ctrl+V) 가능합니다.",
                        required=True
                    )
                }
            )
            
            # 주의: Streamlit 내부 data_editor state 관리를 무너뜨리므로 수동 저장 제거
            # 빈칸 빼고 고유 식당 목록 추출
            unique_rests = [r for r in edited_df["급식장소"].unique() if str(r).strip() not in ["", "nan", "None"]]
            
            if len(unique_rests) > 0:
                st.markdown("---")
                st.subheader("5. 최종 출력물 확인 및 다운로드")
                
                # 공문 생성
                total_cases = len(edited_df)
                total_amount = total_cases * 9000
                kor_amount = amount_to_korean(total_amount)
                rests_str = ", ".join(unique_rests)
                
                cur_year = datetime.now().year
                try:
                    int_month = int(target_month)
                    last_day = calendar.monthrange(cur_year, int_month)[1]
                except (ValueError, TypeError):
                    last_day = 31
                
                text_gongmun = f"""

{target_month}월 {team_name} 급식비

{target_month}월 {team_name} 급식비를 아래와 같이 지출하고자 합니다.

1. 건    명: {target_month}월 {team_name} 급식비
2. 기    간: {cur_year}. {target_month}. 1. ~ {target_month}. {last_day}.
3. 지 출 처: {rests_str}
4. 지출금액: 금{total_amount:,}원(금{kor_amount}원)

붙임 급식비 집행내역 1부. 끝."""

                with st.expander("📝 1) 기안문(공문) 텍스트 복사하기 [열기]", expanded=True):
                    st.code(text_gongmun, language="text")
                    st.caption("복사 버튼을 눌러 새올 기안 양식에 바로 붙여넣으세요.")
                
                # 엑셀 생성
                output = io.BytesIO()
                workbook = xlsxwriter.Workbook(output, {'in_memory': True})
                
                # --- Sheet 1: 급식비 집행내역 ---
                ws1 = workbook.add_worksheet("급식비 집행내역")
                ws1.set_paper(9) # A4 용지
                ws1.set_landscape()
                ws1.fit_to_pages(1, 0) # 너비 맞춤
                ws1.set_margins(left=0.5, right=0.5, top=0.75, bottom=0.75)
                
                fmt_title = workbook.add_format({'bold': True, 'font_size': 20, 'align': 'center', 'valign': 'vcenter'})
                fmt_header = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#f2f2f2'})
                fmt_center = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                fmt_money = workbook.add_format({'align': 'right', 'valign': 'vcenter', 'border': 1, 'num_format': '#,##0'})
                fmt_normal = workbook.add_format({'align': 'left'})
                fmt_normal_border = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'border': 1})
                fmt_right = workbook.add_format({'align': 'right'})
                fmt_sign = workbook.add_format({'align': 'center', 'valign': 'vcenter'}) # 테두리 없는 서명용
                
                ws1.set_column(0, 0, 25)
                ws1.set_column(1, 1, 12)
                ws1.set_column(2, 2, 12)
                ws1.set_column(3, 3, 15)
                ws1.set_column(4, 4, 18)
                ws1.set_column(5, 5, 12)
                
                ws1.merge_range("A1:F2", f"{team_name} 급식비 집행내역", fmt_title)
                # 3행 공백
                ws1.merge_range("A4:C4", f"급식기간 : {cur_year}. {target_month}. 1. ~ {target_month}. {last_day}.", fmt_normal)
                ws1.merge_range("D4:F4", "(단위 : 원, 명)", fmt_right)
                
                ws1.merge_range("A5:A6", "급식처", fmt_header)
                ws1.merge_range("B5:B6", "급식인원", fmt_header)
                ws1.merge_range("C5:C6", "급식단가", fmt_header)
                ws1.merge_range("D5:D6", "금액", fmt_header)
                ws1.merge_range("E5:F5", "지급방법(카드결제)", fmt_header)
                ws1.write("E6", "사업자번호", fmt_header)
                ws1.write("F6", "대표자", fmt_header)
                
                ws1.write("A7", "계", fmt_center)
                ws1.write("B7", total_cases, fmt_center)
                ws1.write("C7", 9000, fmt_money)
                ws1.write("D7", val_total := total_cases * 9000, fmt_money)
                ws1.write("E7", "-", fmt_center)
                ws1.write("F7", "-", fmt_center)
                
                curr_row = 7
                for r_name in unique_rests:
                    cnt = len(edited_df[edited_df['급식장소'] == r_name])
                    
                    ws1.write(curr_row, 0, r_name, fmt_center)
                    ws1.write(curr_row, 1, cnt, fmt_center)
                    ws1.write(curr_row, 2, 9000, fmt_money)
                    ws1.write(curr_row, 3, cnt * 9000, fmt_money)
                    # 더 이상 사업자번호/대표자를 관리하지 않으므로 빈칸으로 처리
                    ws1.write(curr_row, 4, "", fmt_center)
                    ws1.write(curr_row, 5, "", fmt_center)
                    curr_row += 1
                
                curr_row += 2
                
                def split_rank_name(text):
                    parts = text.split(maxsplit=1)
                    if len(parts) == 2:
                        return parts[0], parts[1]
                    return text, ""
                
                def space_name(n):
                    n_str = str(n)
                    return f"{n_str[0]} {n_str[1]} {n_str[2]}" if len(n_str) == 3 else n_str
                    
                c_rank, c_name = split_rank_name(confirm_person)
                w_rank, w_name = split_rank_name(write_person)
                
                ws1.write(f"D{curr_row+1}", "확인자 :", fmt_sign)
                ws1.write(f"E{curr_row+1}", c_rank, fmt_sign)
                ws1.write(f"F{curr_row+1}", f"{space_name(c_name)} (인)", fmt_sign)
                
                ws1.write(f"D{curr_row+3}", "작성자 :", fmt_sign)
                ws1.write(f"E{curr_row+3}", w_rank, fmt_sign)
                ws1.write(f"F{curr_row+3}", f"{space_name(w_name)} (인)", fmt_sign)
                
                # --- Sheet 2: 세부집행내역 ---
                ws2 = workbook.add_worksheet("세부집행내역")
                ws2.set_paper(9) # A4 용지
                ws2.fit_to_pages(1, 0) # 너비 맞춤
                ws2.set_landscape()
                ws2.set_margins(left=0.25, right=0.25, top=0.5, bottom=0.5)
                
                ws2.set_column(0, 0, 5) 
                ws2.set_column(1, 1, 12) 
                ws2.set_column(2, 2, 10) 
                ws2.set_column(3, 3, 6) 
                ws2.set_column(4, 4, 8) 
                ws2.set_column(5, 5, 8) 
                ws2.set_column(6, 6, 8) 
                ws2.set_column(7, 7, 30)
                ws2.set_column(8, 8, 15)
                ws2.set_column(9, 9, 10)
                
                ws2.merge_range("A1:J2", f"{team_name} 업무추진 급식비 세부집행내역 [{target_month}. 1. ~ {target_month}. {last_day}.]", fmt_title)
                
                name_only = write_person.split()[-1] if ' ' in write_person else write_person
                # 3행은 공백
                ws2.merge_range("H4:J4", f"새올초과근무내역 확인필 : {space_name(name_only)} (인)", fmt_right)
                
                ws2.set_row(4, 6) # 5행 얇은 공백
                
                fmt_box = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1})
                fmt_box_gray = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#f2f2f2'})
                
                unique_workers = len(edited_df['성명'].unique())
                ws2.merge_range("A6:B6", "근무자:", fmt_box_gray)
                ws2.write("C6", f"{unique_workers}명", fmt_box)
                ws2.merge_range("D6:E6", "근무내역:", fmt_box_gray)
                ws2.merge_range("F6:H6", f"{total_cases}식 × 1일 급식비 9,000원", fmt_box)
                ws2.write("I6", "급식비:", fmt_box_gray)
                ws2.write("J6", f"{val_total:,} 원", fmt_box)
                
                ws2.set_row(6, 6) # 7행 얇은 공백
                
                headers = ["번호", "근무일자", "근무자", "구분", "출근", "퇴근", "수당시간", "근무내역", "급식장소", "급식비(원)"]
                for i, h in enumerate(headers):
                    ws2.write(7, i, h, fmt_header)
                    
                curr_row = 8
                for idx, r in edited_df.reset_index(drop=True).iterrows():
                    ws2.write(curr_row, 0, idx + 1, fmt_center)
                    date_str = str(r.get('근무일자','')).split()[0]
                    ws2.write(curr_row, 1, date_str, fmt_center)
                    ws2.write(curr_row, 2, str(r.get('성명','')), fmt_center)
                    
                    is_weekend = str(r.get('휴일구분','')).strip()
                    gubun = "휴일" if is_weekend and is_weekend not in ['nan', '0', '평일'] else "평일"
                    ws2.write(curr_row, 3, gubun, fmt_center)
                    
                    t_in = str(r.get('출근(실제)','')).replace(' ', '')
                    t_out = str(r.get('퇴근(실제)','')).replace(' ', '')
                    ws2.write(curr_row, 4, t_in[-5:] if len(t_in)>=5 else t_in, fmt_center)
                    ws2.write(curr_row, 5, t_out[-5:] if len(t_out)>=5 else t_out, fmt_center)
                    
                    ws2.write(curr_row, 6, str(r.get('수당시간(분)','')), fmt_center)
                    ws2.write(curr_row, 7, str(r.get('근무내역','')), fmt_normal_border)
                    ws2.write(curr_row, 8, str(r.get('급식장소','')), fmt_center)
                    ws2.write(curr_row, 9, 9000, fmt_money)
                    curr_row += 1
                    
                workbook.close()
                st.write("")
                st.download_button("📥 2) 급식비 지출증빙 엑셀 다운로드 (2~3p 양식)",
                                   data=output.getvalue(),
                                   file_name=f"{target_month}월_{team_name}급식비.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   type="primary",
                                   use_container_width=True)
            else:
                st.info("⬆️ 위 표에서 '급식장소' 항목을 하나 이상 입력하시면 다음 단계가 열립니다.")
                
        except Exception as e:
            st.error(f"오류가 발생했습니다: {str(e)}")
