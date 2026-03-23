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

# Greeting logic
if 7 <= current_hour < 12:
    greeting = "Καλημέρα!"
elif current_hour >= 12:
    greeting = "Καλησπέρα!"
else:
    greeting = ""

# Hide Streamlit default UI
st.markdown(
    """
    <style>
    #MainMenu {visibility: hidden;}
    header {visibility: hidden;}
    footer {visibility: hidden;}

    .stApp {
        background-color: white;
    }

    .greeting-wrap {
        height: 100vh;
        display: flex;
        align-items: center;
        justify-content: center;
    }

    .greeting-text {
        font-size: 110px;
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

if greeting:
    st.markdown(
        f"""
        <div class="greeting-wrap">
            <div class="greeting-text">{greeting}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )