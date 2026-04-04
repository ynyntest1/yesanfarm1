import streamlit as st
import json
import os
import base64
import requests
import pandas as pd
import streamlit.components.v1 as components

DATA_DIR = "data"
SETTING_FILE = os.path.join(DATA_DIR, "team_settings.json")

# ── GitHub API 설정 ─────────────────────────────────────────────────────────
_GITHUB_REPO      = "ynyntest1/yesanfarm1"
_GITHUB_FILE_PATH = "data/team_settings.json"
_GITHUB_BRANCH    = "main"
_GITHUB_API_URL   = f"https://api.github.com/repos/{_GITHUB_REPO}/contents/{_GITHUB_FILE_PATH}"

def _get_pat() -> str | None:
    """Streamlit Secrets에서 PAT 토큰 반환. 없으면 None."""
    try:
        return st.secrets["GITHUB_PAT"]
    except Exception:
        return None

def _github_headers(token: str) -> dict:
    return {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}

# ── 기본값 ───────────────────────────────────────────────────────────────────
_DEFAULT_SETTINGS = {
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

# ── load_settings ────────────────────────────────────────────────────────────
def load_settings() -> dict:
    """
    우선순위: ① GitHub API → ② 로컬 파일 → ③ 하드코딩 기본값
    """
    token = _get_pat()
    if token:
        try:
            r = requests.get(_GITHUB_API_URL, headers=_github_headers(token),
                             params={"ref": _GITHUB_BRANCH}, timeout=5)
            if r.status_code == 200:
                content = base64.b64decode(r.json()["content"]).decode("utf-8")
                return json.loads(content)
        except Exception:
            pass  # 네트워크 오류 등 → 로컬 폴백

    # 로컬 폴백
    if os.path.exists(SETTING_FILE):
        try:
            with open(SETTING_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except json.JSONDecodeError:
            pass

    return _DEFAULT_SETTINGS.copy()

# ── save_settings ────────────────────────────────────────────────────────────
def save_settings(data: dict) -> bool:
    """
    우선순위: ① GitHub API 커밋 → ② 로컬 파일 저장
    반환값: True = GitHub 저장 성공, False = 로컬 저장으로 폴백
    """
    token = _get_pat()
    if token:
        try:
            headers = _github_headers(token)
            # 현재 파일의 SHA 조회 (PUT 시 필수)
            r = requests.get(_GITHUB_API_URL, headers=headers,
                             params={"ref": _GITHUB_BRANCH}, timeout=5)
            sha = r.json().get("sha") if r.status_code == 200 else None

            content_b64 = base64.b64encode(
                json.dumps(data, ensure_ascii=False, indent=2).encode("utf-8")
            ).decode("utf-8")

            payload = {
                "message": "팀 설정 업데이트 (자동 커밋)",
                "content": content_b64,
                "branch": _GITHUB_BRANCH,
            }
            if sha:
                payload["sha"] = sha

            put_r = requests.put(_GITHUB_API_URL, headers=headers,
                                 json=payload, timeout=10)
            if put_r.status_code in (200, 201):
                return True  # GitHub 저장 성공
        except Exception:
            pass  # 실패 시 로컬 폴백

    # 로컬 폴백
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)
    with open(SETTING_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return False  # 로컬에만 저장됨

def show():
    st.title("⚙️ 팀 환경설정")
    st.markdown("급식비 탭에서 엑셀 파일을 올릴 때마다 새로운 인원이 있으면 자동으로 아래의 **[대기열]** 바구니에 담깁니다. 이름을 마우스로 끌어서 알맞은 과/팀 상자에 넣어주세요! 팀 이름(글씨)을 더블클릭하면 직접 수정도 가능합니다.")
    
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
        saved_to_github = save_settings(settings)
        # st.rerun()을 호출하면 튀는 현상이 있으므로 가급적 무시하거나 한번 찝어줌
        if saved_to_github:
            st.success("✅ 조직도 구조가 GitHub에 영구 저장되었습니다! (자동 커밋)")
        else:
            st.warning("⚠️ GitHub 저장 실패 → 로컬 파일에 임시 저장되었습니다. (GITHUB_PAT 시크릿 확인 필요)")
