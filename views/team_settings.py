import streamlit as st
import json
import os
import pandas as pd
import streamlit.components.v1 as components

DATA_DIR = "data"
SETTING_FILE = os.path.join(DATA_DIR, "team_settings.json")

def load_settings():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    if os.path.exists(SETTING_FILE):
        with open(SETTING_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "team_name": "스마트팜지원팀",
        "confirmer": "지방농업주사 김경민",
        "writer": "지방농업서기 신충호",
        "org_data": {
            "기술지원과": {"기획운영팀": [], "인력육성팀": [], "농기계팀": [], "농업지원팀": []},
            "스마트농업과": {"식량작물팀": [], "과수기술팀": [], "스마트원예팀": [], "스마트팜지원팀": []},
            "미래농업과": {"축산개발팀": [], "귀농지원팀": [], "농촌자원팀": [], "먹거리지원팀": [], "학교급식팀": []}
        },
        "unassigned": [],
        "restaurants": [
            {"식당명": "홍익궁중전통육개장 예산점", "사업자번호": "654-28-00325", "대표자": "전미선"},
            {"식당명": "대흥식당 딸", "사업자번호": "311-09-70205", "대표자": "김준명"},
            {"식당명": "아파트분식", "사업자번호": "311-05-93175", "대표자": "김정희"}
        ]
    }

def save_settings(data):
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

def show():
    st.title("⚙️ 팀 환경설정")
    st.markdown("급식비 탭에서 엑셀 파일을 올릴 때마다 새로운 인원이 있으면 자동으로 아래의 **[대기열]** 바구니에 담깁니다. 이름을 마우스로 끌어서 알맞은 과/팀 상자에 넣어주세요! 팀 이름(글씨)을 더블클릭하면 직접 수정도 가능합니다.")
    
    settings = load_settings()
    
    settings = load_settings()
    
    current_org = settings.get("org_data", {})
    current_u = settings.get("unassigned", [])
    
    # 컴포넌트 렌더링
    st.info("💡 **팁:** 상자 안의 팀 이름을 더블클릭하면 수정할 수 있습니다. 각 과 상자 아래의 '+ 팀 추가' 버튼으로 새 팀을 단일 추가할 수 있습니다.")
    _component_dir = os.path.join(os.path.dirname(os.path.dirname(__file__)), "components", "org_chart")
    org_component = components.declare_component("org_chart", path=_component_dir)
    
    # 컴포넌트의 결과물 받기 (사용자가 저장 버튼을 누르면 이 리턴값이 바뀜)
    ret_val = org_component(org_data=current_org, unassigned=current_u)
    
    if ret_val is not None:
        # 데이터가 프론트에서 넘어왔다면 저장 후 자동 재렌더링
        settings["org_data"] = ret_val.get("org_data", {})
        settings["unassigned"] = ret_val.get("unassigned", [])
        save_settings(settings)
        # st.rerun()을 호출하면 튀는 현상이 있으므로 가급적 무시하거나 한번 찝어줌
        st.success("조직도 구조가 JSON 파일에 영구 저장되었습니다!")
