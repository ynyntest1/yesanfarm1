import streamlit as st
import pandas as pd
import io
import os
import xlsxwriter
import calendar
from datetime import datetime
import streamlit.components.v1 as components
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

@st.cache_data(show_spinner=False)
def _parse_excel_cached(file_bytes: bytes, label: str):
    """새올 초과근무목록 엑셀을 DataFrame으로 변환. 같은 파일은 캐시 재사용."""
    df = pd.read_excel(io.BytesIO(file_bytes))
    if "성명" in df.columns and "출근(실제)" in df.columns:
        data = df.copy()
    else:
        header_idx = -1
        for r in range(15):
            row_vals = [str(x).strip().replace(' ', '') for x in df.iloc[r].fillna('')]
            if "성명" in row_vals and "출근(실제)" in row_vals:
                header_idx = r
                break
        if header_idx == -1:
            return None, f"[{label}] 엑셀에서 '성명', '출근(실제)' 열을 찾을 수 없습니다."
        data = df.iloc[header_idx+1:].copy()
        data.columns = df.iloc[header_idx].tolist()
    data = data[data["성명"].notna() & (data["성명"].astype(str).str.strip() != "") & (data["성명"].astype(str) != "nan")].copy()
    data["고용형태"] = label
    return data, None


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
        
    # ── 파일 업로드 (공무원 / 공무직 구분) ──────────────────────────────────
    st.markdown("#### 1. 새올 초과근무목록 업로드")
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        file_gongmuwon = st.file_uploader("🧑‍💼 공무원 초과근무목록", type=['xls', 'xlsx'], key="upload_gongmuwon")
    with col_up2:
        file_gongmujik = st.file_uploader("👷 공무직 초과근무목록", type=['xls', 'xlsx'], key="upload_gongmujik")

    if not file_gongmuwon and not file_gongmujik:
        st.info("위에서 파일을 하나 이상 업로드해주세요. 공무원·공무직 동시 업로드도 가능합니다.")
        return

    try:
        frames = []
        if file_gongmuwon:
            df_gm, err_gm = _parse_excel_cached(file_gongmuwon.getvalue(), "공무원")
            if err_gm: st.error(err_gm)
            elif df_gm is not None: frames.append(df_gm)
        if file_gongmujik:
            df_gj, err_gj = _parse_excel_cached(file_gongmujik.getvalue(), "공무직")
            if err_gj: st.error(err_gj)
            elif df_gj is not None: frames.append(df_gj)

        if not frames:
            return

        data_df = pd.concat(frames, ignore_index=True)

        cnt_gm = len(data_df[data_df["고용형태"] == "공무원"]) if file_gongmuwon else 0
        cnt_gj = len(data_df[data_df["고용형태"] == "공무직"]) if file_gongmujik else 0
        parts = []
        if cnt_gm: parts.append(f"공무원 {cnt_gm}건")
        if cnt_gj: parts.append(f"공무직 {cnt_gj}건")
        st.success(f"✅ 데이터 인식 완료! 총 {len(data_df)}건 ({' + '.join(parts)})")
        
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
        st.subheader("2. 대상자 선택 및 순서 지정")

        # ── 팀 설정 로드 ─────────────────────────────────────────────────
        t_set      = load_settings()
        loaded_org = t_set.get("org_data", {})

        # ── 단계 1: 과 자동 감지 (엑셀 부서 컬럼 기반, 하드코딩 제거) ──
        excel_divs = sorted({
            str(x).strip() for x in data_df["부서"].dropna()
            if str(x).strip() not in ["", "nan"]
        })

        if "selected_div" not in st.session_state:
            st.session_state["selected_div"] = None

        st.markdown("##### 🏢 과 선택")
        cols_div = st.columns(max(len(excel_divs), 1))
        for i, div_name in enumerate(excel_divs):
            with cols_div[i]:
                btn_type = "primary" if st.session_state.get("selected_div") == div_name else "secondary"
                if st.button(div_name, use_container_width=True, type=btn_type, key=f"div_btn_{i}"):
                    st.session_state["selected_div"] = div_name
                    st.session_state["selected_team"] = None
                    st.rerun()

        selected_div = st.session_state.get("selected_div")
        if not selected_div:
            st.info("👆 위에서 소속 과를 선택해주세요.")
            return

        # ── 단계 2: 팀 선택 (team_settings org_data 기반) ───────────────
        teams_in_div = list(loaded_org.get(selected_div, {}).keys())

        if not teams_in_div:
            st.warning(f"[팀 설정] 탭에서 {selected_div} 소속 팀을 먼저 등록해주세요.")
            return

        st.markdown(f"##### 🏷️ [{selected_div}] 팀 선택")
        if "selected_team" not in st.session_state:
            st.session_state["selected_team"] = None

        cols_team = st.columns(max(len(teams_in_div), 1))
        for i, tname in enumerate(teams_in_div):
            with cols_team[i % len(cols_team)]:
                btn_type = "primary" if st.session_state.get("selected_team") == tname else "secondary"
                if st.button(tname, use_container_width=True, type=btn_type, key=f"team_btn_{i}"):
                    st.session_state["selected_team"] = tname
                    st.rerun()

        selected_team = st.session_state.get("selected_team")
        if not selected_team:
            st.info("👆 위에서 소속 팀을 선택해주세요.")
            return

        # ── 단계 3: 팀원 목록 구성 ──────────────────────────────────────
        # 기존 팀원 (team_settings 저장 순서)
        existing_members = loaded_org.get(selected_div, {}).get(selected_team, [])

        # 엑셀에 있는 전체 인원 (중복 제거, 순서 유지)
        excel_names_unique = list(dict.fromkeys(
            str(x).strip() for x in data_df["성명"].dropna()
            if str(x).strip() not in ["", "nan"]
        ))

        # 조직도 전체에서 이미 배정된 인원
        all_assigned = {
            name
            for div_data in loaded_org.values()
            for team_data in div_data.values()
            for name in team_data
        }

        # 신규 인원 = 엑셀에 있지만 어디에도 배정 안 된 사람
        new_members = [n for n in excel_names_unique if n not in all_assigned]

        # ── 단계 4: 드래그 순서 지정 컴포넌트 ──────────────────────────
        st.markdown(f"##### 👥 [{selected_team}] 팀원 순서 지정")
        st.caption("드래그로 순서를 바꾸면 자동 저장됩니다. 💼 아이콘이 팀장, 파란 점선이 신규 인원입니다.")

        _sorter_dir = os.path.join(
            os.path.dirname(os.path.dirname(__file__)),
            "components", "member_sorter"
        )
        member_sorter_component = components.declare_component("member_sorter", path=_sorter_dir)

        sort_result = member_sorter_component(
            members     = existing_members,
            new_members = new_members,
            team_key    = f"{selected_div}__{selected_team}",
            key         = f"sorter_{selected_div}_{selected_team}"
        )

        # 드래그 결과가 오면 team_settings 즉시 저장 (GitHub 자동 커밋)
        if sort_result is not None:
            new_order = sort_result.get("ordered_members", [])
            loaded_org.setdefault(selected_div, {})[selected_team] = new_order
            t_set["org_data"] = loaded_org
            t_set["unassigned"] = [u for u in t_set.get("unassigned", []) if u not in new_order]
            save_settings(t_set)
            selected_names = new_order
        else:
            # 첫 렌더 (아직 드래그 없음) → 저장된 순서 + 신규 인원
            selected_names = existing_members + new_members
        
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
        st.markdown("---")
        st.subheader("3. 서류 기본 정보 입력")

        # 순서 기반으로 확인자(팀장=첫번째) / 작성자(막내=마지막) 자동완성
        rank_map = t_set.get("rank_dict", {})

        default_confirmer = ""
        default_writer    = ""
        # existing_members 기준 (드래그 저장 전엔 기존 순서 사용)
        ref_list = existing_members if existing_members else selected_names
        if ref_list:
            leader = ref_list[0]
            default_confirmer = f"{rank_map.get(leader, '지방농업주사')} {leader}"
            writer = ref_list[-1]
            default_writer    = f"{rank_map.get(writer, '지방농업서기')} {writer}"

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
        
        disp_df = filtered_data_df[['근무일자', '부서', '고용형태', '성명', '휴일구분', '출근(실제)', '퇴근(실제)', '수당시간(분)', '근무내역']].copy()
        
        # 성명 순서를 드래그 지정 순서(selected_names) 에 맞추기 (1순위 정렬)
        if selected_names:
            name_order = {name: i for i, name in enumerate(selected_names)}
            disp_df['_name_order'] = disp_df['성명'].map(lambda x: name_order.get(x, 999))
        else:
            disp_df['_name_order'] = 0
            
        # 근무일자 오름차순 (2순위 정렬)
        disp_df = disp_df.sort_values(by=['_name_order', '근무일자']).drop(columns=['_name_order']).reset_index(drop=True)
        disp_df.index = disp_df.index + 1
        disp_df.insert(0, '순번', disp_df.index)
        
        # 다른 팀이나 인원을 지웠다 켰을때, 혹은 분 조건을 바꿀때 테이블 리셋
        current_view_key = f"{selected_div}_{selected_team}_{','.join(str(n) for n in selected_names)}_{min_hours}"
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
            disabled=['순번', '근무일자', '부서', '고용형태', '성명', '휴일구분', '출근(실제)', '퇴근(실제)', '수당시간(분)', '근무내역'],
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

        # ── 급식장소별 실시간 소계 ────────────────────────────────────────────
        if len(unique_rests) > 0:
            st.markdown("---")
            st.markdown("##### 📊 급식장소별 소계 (실시간)")
            rest_metric_cols = st.columns(min(len(unique_rests), 4))
            for _i, _rname in enumerate(unique_rests):
                _cnt = int((edited_df["급식장소"] == _rname).sum())
                _amt = _cnt * 9000
                with rest_metric_cols[_i % min(len(unique_rests), 4)]:
                    st.metric(label=f"🍽️ {_rname}", value=f"{_amt:,}원", delta=f"{_cnt}명")
            # 미입력 건수 경고
            _empty_cnt = int(edited_df["급식장소"].astype(str).str.strip().isin(["", "nan", "None"]).sum())
            if _empty_cnt > 0:
                st.caption(f"⚠️ 급식장소 미입력 내역: **{_empty_cnt}건** — 모두 입력 후 다운로드해주세요.")
            else:
                st.success(f"✅ 모든 {len(edited_df)}건에 급식장소가 입력되었습니다. 아래에서 출력물을 확인하세요.")

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
            
            ws2.set_column(0, 0, 5)   # 번호
            ws2.set_column(1, 1, 12)  # 근무일자
            ws2.set_column(2, 2, 10)  # 근무자
            ws2.set_column(3, 3, 7)   # 고용형태
            ws2.set_column(4, 4, 6)   # 구분
            ws2.set_column(5, 5, 8)   # 출근
            ws2.set_column(6, 6, 8)   # 퇴근
            ws2.set_column(7, 7, 8)   # 수당시간
            ws2.set_column(8, 8, 28)  # 근무내역
            ws2.set_column(9, 9, 15)  # 급식장소
            ws2.set_column(10, 10, 10) # 급식비
            
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
            
            headers = ["번호", "근무일자", "근무자", "고용형태", "구분", "출근", "퇴근", "수당시간", "근무내역", "급식장소", "급식비(원)"]
            for i, h in enumerate(headers):
                ws2.write(7, i, h, fmt_header)

            # 고용형태별 셀 색상
            fmt_gongmuwon = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#EFF6FF'})
            fmt_gongmujik = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'border': 1, 'bg_color': '#F0FDF4'})

            curr_row = 8
            for idx, r in edited_df.reset_index(drop=True).iterrows():
                고용형태 = str(r.get('고용형태', '공무원')).strip()
                fmt_type = fmt_gongmujik if 고용형태 == '공무직' else fmt_gongmuwon

                ws2.write(curr_row, 0, idx + 1, fmt_center)
                date_str = str(r.get('근무일자','')).split()[0]
                ws2.write(curr_row, 1, date_str, fmt_center)
                ws2.write(curr_row, 2, str(r.get('성명','')), fmt_center)
                ws2.write(curr_row, 3, 고용형태, fmt_type)

                is_weekend = str(r.get('휴일구분','')).strip()
                gubun = "휴일" if is_weekend and is_weekend not in ['nan', '0', '평일'] else "평일"
                ws2.write(curr_row, 4, gubun, fmt_center)

                t_in  = str(r.get('출근(실제)','')).replace(' ', '')
                t_out = str(r.get('퇴근(실제)','')).replace(' ', '')
                ws2.write(curr_row, 5, t_in[-5:]  if len(t_in)>=5  else t_in,  fmt_center)
                ws2.write(curr_row, 6, t_out[-5:] if len(t_out)>=5 else t_out, fmt_center)

                ws2.write(curr_row, 7,  str(r.get('수당시간(분)','')), fmt_center)
                ws2.write(curr_row, 8,  str(r.get('근무내역','')),    fmt_normal_border)
                ws2.write(curr_row, 9,  str(r.get('급식장소','')),    fmt_center)
                ws2.write(curr_row, 10, 9000, fmt_money)
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
