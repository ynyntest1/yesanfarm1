import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(__file__))
from utils.sidebar import show_sidebar

st.set_page_config(
    page_title="업무 자동화 도구 모음",
    page_icon="🏠",
    layout="wide"
)

show_sidebar()

# ─── 카드 스타일 ─────────────────────────────────────────────────────
st.markdown("""
<style>
.menu-card-link {
    text-decoration: none;
    color: #1a1a1a;
    display: block;
    border: 1px solid #e0e0e0;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 4px;
    background: white;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    min-height: 200px;
    transition: all 0.18s ease;
    cursor: pointer;
}
.menu-card-link:hover,
.menu-card-link:visited,
.menu-card-link:active {
    text-decoration: none !important;
    color: #1a1a1a !important;
}
.menu-card-link * {
    text-decoration: none !important;
}
.menu-card-link:hover {
    border-color: #4e8df5;
    box-shadow: 0 6px 20px rgba(78,141,245,0.18);
    transform: translateY(-3px);
}
.menu-card-disabled {
    display: block;
    border: 1px solid #e0e0e0;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 4px;
    background: #fafafa;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
    min-height: 200px;
    opacity: 0.6;
    cursor: not-allowed;
}
</style>
""", unsafe_allow_html=True)

st.title("🏠 업무 자동화 도구 모음")
st.markdown("아래 카드를 클릭하면 해당 기능으로 이동합니다.")
st.markdown("---")

# ─── 메뉴 카드 정의 — 새 기능 추가 시 여기에만 추가하면 됩니다 ────────
MENU_ITEMS = [
    {
        "icon": "📄",
        "title": "여비지급명세서",
        "desc": "인사랑 출장내역서 엑셀을 업로드하면\n관내/관외 여비지급명세서를 자동으로 생성합니다.",
        "url": "/1_expense_report",
        "status": "운영중",
    },
    {
        "icon": "💰",
        "title": "급량비 계산기",
        "desc": "급량비 지급 대상자 명단을 입력하면\n자동으로 계산 및 정리해드립니다.",
        "url": "/2_meal_allowance",
        "status": "개발중",
    },
    {
        "icon": "🍱",
        "title": "식사명단 취합",
        "desc": "여러 팀의 식사 신청 명단을\n한 번에 취합하고 정리합니다.",
        "url": "/3_meal_list",
        "status": "개발중",
    },
]

# ─── 카드 렌더링 ─────────────────────────────────────────────────────
cols = st.columns(len(MENU_ITEMS))

for idx, item in enumerate(MENU_ITEMS):
    is_active = item["status"] == "운영중"
    status_badge = "운영중" if is_active else "🚧 개발중"
    badge_bg    = "#eff6ff" if is_active else "#fff3e0"
    badge_color = "#2563eb" if is_active else "#e65100"
    badge_prefix = "✅ " if is_active else ""
    desc_html   = item["desc"].replace("\n", "<br>")

    card_inner = f"""
        <div style="font-size: 2.5rem; margin-bottom: 10px;">{item['icon']}</div>
        <div style="font-size: 1.1rem; font-weight: 700; margin-bottom: 6px; color: #1a1a1a;">{item['title']}</div>
        <div style="font-size: 0.85rem; color: #666; margin-bottom: 14px;">{desc_html}</div>
        <span style="
            font-size: 0.75rem;
            padding: 3px 10px;
            border-radius: 20px;
            background: {badge_bg};
            color: {badge_color};
            font-weight: 700;
        ">{badge_prefix}{status_badge}</span>
    """

    if is_active:
        card_html = f'<a href="{item["url"]}" class="menu-card-link" target="_self">{card_inner}</a>'
    else:
        card_html = f'<div class="menu-card-disabled">{card_inner}</div>'

    with cols[idx]:
        st.markdown(card_html, unsafe_allow_html=True)

st.markdown("---")
st.caption("💡 사용에 불편이 있을 수 있습니다. (개발중)")
