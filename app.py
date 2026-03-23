from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st


st.set_page_config(
    page_title="Dashboard",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# Current time in Greece
now = datetime.now(ZoneInfo("Europe/Athens"))
current_hour = now.hour
current_minute = now.minute

# Greeting logic
if current_hour < 7:
    greeting = "Τι κάνεις εδώ τέτοια ώρα;"
elif 7 <= current_hour < 12:
    greeting = "Καλημέρα!"
else:
    greeting = "Καλησπέρα!"


# Remove ALL Streamlit spacing issues
st.markdown(
    """
    <style>
    /* Remove Streamlit padding */
    .block-container {
        padding: 0 !important;
        margin: 0 !important;
        max-width: 100% !important;
    }

    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}

    html, body, .stApp {
        height: 100%;
        margin: 0;
        padding: 0;
        background-color: white;
    }

    /* Full screen center */
    .greeting-wrap {
        position: fixed;
        top: 0;
        left: 0;
        width: 100vw;
        height: 100vh;

        display: flex;
        align-items: center;
        justify-content: center;
    }

    .greeting-text {
        font-size: 120px;
        font-weight: 800;
        color: #111111;
        text-align: center;

        opacity: 0;
        animation: fadeInOut 3s ease-in-out forwards;
    }

    @keyframes fadeInOut {
        0%   { opacity: 0; transform: scale(0.96); }
        20%  { opacity: 1; transform: scale(1); }
        80%  { opacity: 1; transform: scale(1); }
        100% { opacity: 0; transform: scale(1.02); }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Render greeting
st.markdown(
    f"""
    <div class="greeting-wrap">
        <div class="greeting-text">{greeting}</div>
    </div>
    """,
    unsafe_allow_html=True,
)