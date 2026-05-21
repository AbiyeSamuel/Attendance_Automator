import streamlit as st
import pandas as pd

# 1. Set up the page layout
st.set_page_config(page_title="Attendance Automator", page_icon="🏫", layout="centered")

st.title("Weekly Attendance Automator 🏫")
st.write("Upload your daily `.dat` log files below to generate the populated weekly Excel template.")

# 2. Create the drag-and-drop file uploader
# We set accept_multiple_files=True so you can upload all 3 days at once
uploaded_files = st.file_uploader("Drop your .dat files here", type=["dat"], accept_multiple_files=True)

# 3. Check if files are uploaded
if uploaded_files:
    st.success(f"Awesome! You have uploaded {len(uploaded_files)} file(s).")

    # Just to show you it works, let's list the file names
    for file in uploaded_files:
        st.write(f"- {file.name}")

    st.info("The time-compression and Excel export logic will go here in the next step!")