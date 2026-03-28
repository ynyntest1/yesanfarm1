import streamlit as st
from views.expense_report import show as show_expense_report
from views.meal_allowance import show as show_meal_allowance
from views.meal_list import show as show_meal_list

st.set_page_config(
    page_title="업무 자동화 도구 모음",
    page_icon="🏠",
    layout="wide"
)

# ─── 앱 상단 고정 헤더 ────────────────────────────────────────────────
st.title("🏠 업무 자동화 도구 모음")
st.markdown("---")
st.markdown("아래 탭을 클릭하면 해당 기능으로 이동할 수 있습니다.")

# ─── 탭 네비게이션 ────────────────────────────────────────────────────
tab_home, tab_expense, tab_meal, tab_meallist = st.tabs([
    "🏠 홈",
    "📄 여비지급명세서",
    "💰 급량비",
    "🍱 식사명단 취합",
])

# ─── 홈 ──────────────────────────────────────────────────────────────
with tab_home:

    st.markdown("""
    <style>
    .menu-card-active {
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 24px;
        background: white;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
        min-height: 210px;
        transition: all 0.18s ease;
    }
    .menu-card-active:hover {
        border-color: #4e8df5;
        box-shadow: 0 6px 20px rgba(78,141,245,0.18);
        transform: translateY(-3px);
    }
    .menu-card-disabled {
        border: 1px solid #e0e0e0;
        border-radius: 12px;
        padding: 24px;
        background: #fafafa;
        box-shadow: 0 2px 8px rgba(0,0,0,0.04);
        min-height: 210px;
        opacity: 0.6;
    }
    
    /* ─── 탭 스타일 커스텀 (모서리가 둥근 사각형) ─── */
    button[data-baseweb="tab"] {
        border: 1px solid #e0e0e0 !important;
        border-radius: 12px !important; /* 둥근 사각형 */
        margin-right: 8px !important;
        padding: 8px 20px !important;
        background-color: #ffffff !important;
        transition: all 0.2s ease !important;
    }
    button[data-baseweb="tab"]:hover {
        background-color: #f7f9fc !important;
        border-color: #c0c8d0 !important;
    }
    /* 선택된 활성 탭 */
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #4e8df5 !important;
        border-color: #4e8df5 !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] p {
        color: white !important;
        font-weight: 700 !important;
    }
    /* 기본 탭 밑줄(하이라이트 선) 제거 */
    div[data-baseweb="tab-highlight"] {
        display: none !important;
    }
    /* 탭 리스트 하단선 제거 및 간격 조정 */
    div[data-baseweb="tab-board"] {
        border-bottom: none !important;
    }
    </style>
    """, unsafe_allow_html=True)

    MENU_ITEMS = [
        {"icon": "📄", "title": "여비지급명세서",
         "desc": "인사랑 출장내역서 엑셀을 업로드하면\n관내/관외 여비지급명세서를 자동으로 생성합니다.",
         "status": "운영중"},
        {"icon": "💰", "title": "급량비 계산기",
         "desc": "급량비 지급 대상자 명단을 입력하면\n자동으로 계산 및 정리해드립니다.",
         "status": "개발중"},
        {"icon": "🍱", "title": "식사명단 취합",
         "desc": "여러 팀의 식사 신청 명단을\n한 번에 취합하고 정리합니다.",
         "status": "개발중"},
    ]

    cols = st.columns(len(MENU_ITEMS))
    for idx, item in enumerate(MENU_ITEMS):
        is_active   = item["status"] == "운영중"
        badge_bg    = "#eff6ff" if is_active else "#fff3e0"
        badge_color = "#2563eb" if is_active else "#e65100"
        badge_text  = "✅ 운영중" if is_active else "🚧 개발중"
        desc_html   = item["desc"].replace("\n", "<br>")
        card_inner  = f"""
            <div style="font-size:2.5rem; margin-bottom:10px;">{item['icon']}</div>
            <div style="font-size:1.1rem; font-weight:700; margin-bottom:6px;
                        color:#1a1a1a;">{item['title']}</div>
            <div style="font-size:0.85rem; color:#666; margin-bottom:14px;">{desc_html}</div>
            <span style="font-size:0.75rem; padding:3px 10px; border-radius:20px;
                         background:{badge_bg}; color:{badge_color}; font-weight:700;">
                {badge_text}
            </span>
        """
        with cols[idx]:
            card_class = "menu-card-active" if is_active else "menu-card-disabled"
            st.markdown(f'<div class="{card_class} card-nav" data-target="{idx + 1}">{card_inner}</div>', unsafe_allow_html=True)

    # 카드 클릭 시 탭 이동을 위한 자바스크립트 주입 (화면엔 안 보임)
    st.components.v1.html("""
    <script>
    setTimeout(() => {
        const doc = window.parent.document;
        const cards = doc.querySelectorAll('.card-nav');
        cards.forEach(card => {
            if (card.classList.contains('menu-card-active')) {
                card.style.cursor = 'pointer';
                card.onclick = function() {
                    const tabs = doc.querySelectorAll('button[data-baseweb="tab"]');
                    const targetIdx = parseInt(card.getAttribute('data-target'), 10);
                    if(tabs && tabs.length > targetIdx) {
                        tabs[targetIdx].click();
                    }
                };
            }
        });
    }, 150);
    </script>
    """, height=0)

    st.markdown("---")
    st.caption("💡 사용에 불편이 있을 수 있습니다. (개발중)")

# ─── 각 기능 탭 ───────────────────────────────────────────────────────
with tab_expense:
    show_expense_report()

with tab_meal:
    show_meal_allowance()

with tab_meallist:
    show_meal_list()
