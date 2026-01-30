import streamlit as st
import importlib.util
import sys
import os

# ============================================================================
# PAGE CONFIG (MUST BE FIRST)
# ============================================================================
st.set_page_config(layout="wide", page_title="Multi-Market Dashboard 🌏", page_icon="💰")

# ============================================================================
# MARKET SELECTOR
# ============================================================================
st.sidebar.header("🌏 MARKET SELECTION")
selected_market = st.sidebar.radio(
    "Choose Market:",
    ["Taiwan 🇹🇼", "Vietnam 🇻🇳"],
    horizontal=False
)

st.sidebar.divider()
st.sidebar.info(f"**Active Market:** {selected_market}")
st.sidebar.caption("All features from original dashboards are preserved.")

# ============================================================================
# DYNAMIC MODULE LOADING
# ============================================================================
current_folder = os.path.dirname(os.path.abspath(__file__))

if selected_market == "Taiwan 🇹🇼":
    dashboard_file = os.path.join(current_folder, "dashboard_tw.py")

    # Read and execute Taiwan dashboard (skip st.set_page_config line)
    with open(dashboard_file, 'r', encoding='utf-8') as f:
        code = f.read()

    # Remove st.set_page_config to avoid duplicate config error
    lines = code.split('\n')
    filtered_lines = []
    skip_next = False

    for line in lines:
        if 'st.set_page_config' in line:
            skip_next = True
            continue
        if skip_next and line.strip() and not line.strip().startswith('#'):
            skip_next = False
        if not skip_next:
            filtered_lines.append(line)

    modified_code = '\n'.join(filtered_lines)
    exec(modified_code)

else:  # Vietnam
    dashboard_file = os.path.join(current_folder, "dashboard_vn.py")

    # Read and execute Vietnam dashboard (skip st.set_page_config line)
    with open(dashboard_file, 'r', encoding='utf-8') as f:
        code = f.read()

    # Remove st.set_page_config to avoid duplicate config error
    lines = code.split('\n')
    filtered_lines = []
    skip_next = False

    for line in lines:
        if 'st.set_page_config' in line:
            skip_next = True
            continue
        if skip_next and line.strip() and not line.strip().startswith('#'):
            skip_next = False
        if not skip_next:
            filtered_lines.append(line)

    modified_code = '\n'.join(filtered_lines)
    exec(modified_code)
