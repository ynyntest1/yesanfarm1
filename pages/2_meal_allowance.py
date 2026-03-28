import streamlit as st
import sys, os
sys.path.insert(0, os.path.dirname(os.path.dirname(__file__)))
from utils.sidebar import show_sidebar

st.set_page_config(page_title="급량비 계산기", layout="wide")
show_sidebar()

st.title("💰 급량비 계산기")
st.info("🚧 현재 개발 중인 기능입니다. 곧 업데이트될 예정입니다!")
