

import streamlit as st

st.set_page_config(page_title="RFP Proposal AI Generator", layout="wide")

# --- Custom Styling ---
st.markdown("""
<style>
.main-header {
    text-align: center;
    color: #000;
    font-size: 3em;
    font-weight: 800;
    padding-top: 20px;
    padding-bottom: 5px;
}
.highlight-text { color: #1A75E0; }
.sub-tagline {
    text-align: center;
    color: #555;
    font-size: 1.1em;
    padding-bottom: 40px;
}
div[data-testid="stButton"] > button {
    border: 1px solid #d1d1d1;
    border-radius: 10px;
    padding: 12px 0;
    font-size: 15px;
    font-weight: 600;
    background-color: #f8f9fa;
    color: #333;
    transition: all 0.2s ease-in-out;
}
div[data-testid="stButton"] > button:hover {
    border-color: #1A75E0;
    background-color: #EAF3FF;
    color: #1A75E0;
}
</style>
""", unsafe_allow_html=True)

# --- Header ---
st.markdown("<div class='main-header'>Automate Your <span class='highlight-text'>Proposal Response</span></div>", unsafe_allow_html=True)
st.markdown("<p class='sub-tagline'>Respond to RFPs in minutes with AI-driven content generation.</p>", unsafe_allow_html=True)

# --- Service Buttons ---
st.markdown("## 🧩 Select Service Type")

service_types = {
    "Integration": "🔗",
    "Core Assessment": "📋"
}

cols = st.columns(2)
for i, (label, icon) in enumerate(service_types.items()):
    with cols[i % 2]:
        if st.button(f"{icon} {label}", use_container_width=True):
            # Redirect to the correct page
            if label == "Integration":
                st.switch_page("pages/1_Integration.py") 
            elif label == "Core Assessment":
                st.switch_page("pages/2_Core_Assessment.py")
