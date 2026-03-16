import streamlit as st
import sys
import os
import shutil
from io import StringIO
from datetime import datetime
from main import run_full_process


# --------------------------------------------------
# PAGE CONFIG
# --------------------------------------------------

st.set_page_config(
    page_title="RFID Tag Report Generator",
    layout="wide"
)


# --------------------------------------------------
# SESSION STATE
# --------------------------------------------------

if "terminal_output" not in st.session_state:
    st.session_state.terminal_output = ""


# --------------------------------------------------
# TERMINAL REDIRECT
# --------------------------------------------------

class StreamlitRedirect:

    def write(self, string):
        timestamp = datetime.now().strftime("%H:%M:%S")
        st.session_state.terminal_output += f"[{timestamp}] {string}"

    def flush(self):
        pass


# --------------------------------------------------
# TERMINAL FUNCTIONS
# --------------------------------------------------

def append_terminal(text):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.terminal_output += f"[{timestamp}] {text}\n"


def clear_terminal():
    st.session_state.terminal_output = ""


# --------------------------------------------------
# SAVE UPLOADED FILES
# --------------------------------------------------

def save_uploaded_files(files, folder):

    os.makedirs(folder, exist_ok=True)

    for file in files:
        path = os.path.join(folder, file.name)

        with open(path, "wb") as f:
            f.write(file.getbuffer())


# --------------------------------------------------
# RUN MAIN PROCESS
# --------------------------------------------------

def run_process(master_files, loco_files, division):

    master_folder = "master_temp"
    loco_folder = "loco_temp"

    try:

        clear_terminal()

        append_terminal("=" * 60)
        append_terminal("RFID TAG REPORT GENERATOR STARTED")
        append_terminal("=" * 60)

        append_terminal(f"Division : {division}")
        append_terminal("-" * 60)

        # remove old temp folders
        if os.path.exists(master_folder):
            shutil.rmtree(master_folder)

        if os.path.exists(loco_folder):
            shutil.rmtree(loco_folder)

        # save uploaded files
        append_terminal("Saving Master Files...")
        save_uploaded_files(master_files, master_folder)

        append_terminal("Saving Loco Files...")
        save_uploaded_files(loco_files, loco_folder)

        append_terminal("Files Saved Successfully")
        append_terminal("-" * 60)

        # redirect stdout
        old_stdout = sys.stdout
        sys.stdout = StreamlitRedirect()

        run_full_process(master_folder, loco_folder, division)

        sys.stdout = old_stdout

        append_terminal("REPORTS GENERATED SUCCESSFULLY")

        st.success("Reports Generated Successfully!")

    except Exception as e:

        sys.stdout = old_stdout
        append_terminal(f"ERROR : {str(e)}")
        st.error(str(e))


# --------------------------------------------------
# UI
# --------------------------------------------------

st.title("RFID Tag Verification and Validation Report Generator")

st.caption("Automated Railway Tag Validation Tool")

st.divider()


# --------------------------------------------------
# FILE UPLOAD
# --------------------------------------------------

col1, col2 = st.columns(2)

with col1:

    master_files = st.file_uploader(
        "Upload Master DB Files",
        accept_multiple_files=True
    )

with col2:

    loco_files = st.file_uploader(
        "Upload Loco Log Files",
        accept_multiple_files=True
    )


division = st.text_input("Division Name")

st.divider()


# --------------------------------------------------
# BUTTONS
# --------------------------------------------------

col1, col2 = st.columns(2)

generate = col1.button("Generate Reports", use_container_width=True)

clear = col2.button("Clear Terminal", use_container_width=True)

if clear:
    clear_terminal()


# --------------------------------------------------
# PROCESS EXECUTION
# --------------------------------------------------

if generate:

    if not master_files or not loco_files or not division:

        st.warning("Please upload files and enter division")

    else:

        with st.spinner("Processing... Please wait"):

            run_process(master_files, loco_files, division)


# --------------------------------------------------
# TERMINAL OUTPUT
# --------------------------------------------------

st.subheader("Process Output Terminal")

st.text_area(
    "terminal",
    st.session_state.terminal_output,
    height=400,
    label_visibility="collapsed"
)