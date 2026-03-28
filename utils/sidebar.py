import streamlit as st

# ✏️ 홈 이외의 메뉴 추가/수정은 여기 SUB_PAGES 에만 하면 됩니다
HOME = {"icon": "🏠", "title": "홈", "path": "app.py"}

SUB_PAGES = [
    {"icon": "📄", "title": "여비 명세서", "path": "pages/1_여비지급명세서.py"},
    {"icon": "💰", "title": "급량비",        "path": "pages/2_급량비.py"},
    {"icon": "🍱", "title": "식사명단 취합", "path": "pages/3_식사명단취합.py"},
]

def show_sidebar():
    st.markdown("""
        <style>
        [data-testid="stSidebarNav"] { display: none !important; }

        /* 공통 버튼 스타일 */
        section[data-testid="stSidebar"] .stButton button {
            font-weight: 600;
            white-space: pre-line;
            line-height: 1.6;
            border: 1.5px solid #e0e0e0;
            border-radius: 10px;
            background: #ffffff;
            color: #333;
            transition: all 0.15s ease;
        }
        section[data-testid="stSidebar"] .stButton button:hover {
            border-color: #4e8df5;
            color: #4e8df5;
            background: #f0f5ff;
        }

        /* 홈 버튼: 크고 두드러지게 */
        section[data-testid="stSidebar"] [data-testid="stButton-nav_home"] button {
            height: 80px;
            font-size: 1rem;
            border-width: 2px;
        }

        /* 서브메뉴 버튼: 작은 박스형 */
        section[data-testid="stSidebar"] .stButton:not([data-testid="stButton-nav_home"]) button {
            height: 68px;
            font-size: 0.8rem;
        }
        </style>
    """, unsafe_allow_html=True)

    with st.sidebar:
        st.markdown("#### 📋 메뉴 선택")
        st.markdown("---")

        # 홈 — 풀너비 단독 버튼
        if st.button(
            f"{HOME['icon']}  {HOME['title']}",
            key="nav_home",
            use_container_width=True
        ):
            st.switch_page(HOME["path"])

        st.markdown("<div style='margin-top:8px'></div>", unsafe_allow_html=True)

        # 나머지 — 2열 격자 박스
        cols = st.columns(2)
        for i, page in enumerate(SUB_PAGES):
            with cols[i % 2]:
                if st.button(
                    f"{page['icon']}\n{page['title']}",
                    key=f"nav_{i}",
                    use_container_width=True
                ):
                    st.switch_page(page["path"])

        st.markdown("---")
