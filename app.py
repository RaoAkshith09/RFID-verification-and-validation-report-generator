import streamlit as st
import sys
import os
from io import StringIO
from datetime import datetime
from main import run_full_process
import tkinter as tk
from tkinter import filedialog


# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------

st.set_page_config(
    page_title="RFID Tag Report Generator",
    layout="wide"
)

st.title("RFID Tag Verification and Validation Report Generator")
st.caption("Automated RFID Tag Report Generation Tool")


# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------

if "terminal_output" not in st.session_state:
    st.session_state.terminal_output = ""

if "master_folder" not in st.session_state:
    st.session_state.master_folder = ""

if "loco_folder" not in st.session_state:
    st.session_state.loco_folder = ""


# -------------------------------------------------
# TERMINAL REDIRECT
# -------------------------------------------------

class StreamlitRedirect:

    def write(self, text):

        timestamp = datetime.now().strftime("%H:%M:%S")

        st.session_state.terminal_output += f"[{timestamp}] {text}"

    def flush(self):
        pass


def append_terminal(text):

    timestamp = datetime.now().strftime("%H:%M:%S")

    st.session_state.terminal_output += f"[{timestamp}] {text}\n"


def clear_terminal():

    st.session_state.terminal_output = ""


# -------------------------------------------------
# FOLDER BROWSER
# -------------------------------------------------

def browse_master():

    root = tk.Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="Select Master Folder")

    if folder:
        st.session_state.master_folder = folder


def browse_loco():

    root = tk.Tk()
    root.withdraw()

    folder = filedialog.askdirectory(title="Select Loco Log Folder")

    if folder:
        st.session_state.loco_folder = folder


# -------------------------------------------------
# PROCESS FUNCTION
# -------------------------------------------------

def run_process(master_folder, loco_folder, division):

    try:

        clear_terminal()

        append_terminal("="*60)
        append_terminal("RFID TAG REPORT GENERATOR - PROCESS STARTED")
        append_terminal("="*60)

        append_terminal(f"Master Folder : {master_folder}")
        append_terminal(f"Loco Folder   : {loco_folder}")
        append_terminal(f"Division      : {division}")

        append_terminal("-"*60)

        # Redirect stdout
        old_stdout = sys.stdout
        sys.stdout = StreamlitRedirect()

        run_full_process(master_folder, loco_folder, division)

        sys.stdout = old_stdout

        append_terminal("REPORT GENERATION COMPLETED")

        st.success("Reports Generated Successfully!")

    except Exception as e:

        sys.stdout = old_stdout

        append_terminal(f"ERROR : {str(e)}")

        st.error(str(e))


# -------------------------------------------------
# INPUT SECTION
# -------------------------------------------------

st.divider()

col1, col2 = st.columns([4,1])

with col1:
    master_folder = st.text_input(
        "Master Data Folder",
        value=st.session_state.master_folder
    )

with col2:
    st.button("Browse", on_click=browse_master)


col3, col4 = st.columns([4,1])

with col3:
    loco_folder = st.text_input(
        "Loco Log Folder",
        value=st.session_state.loco_folder
    )

with col4:
    st.button("Browse ", on_click=browse_loco)


division = st.text_input("Division Name")

st.divider()


# -------------------------------------------------
# BUTTONS
# -------------------------------------------------

col1, col2 = st.columns(2)

generate = col1.button("Generate Reports", use_container_width=True)

clear = col2.button("Clear Terminal", use_container_width=True)

if clear:
    clear_terminal()


# -------------------------------------------------
# RUN PROCESS
# -------------------------------------------------

if generate:

    if not master_folder or not loco_folder or not division:

        st.warning("Please fill all inputs")

    else:

        if not os.path.exists(master_folder):

            st.error("Master folder does not exist")

        elif not os.path.exists(loco_folder):

            st.error("Loco folder does not exist")

        else:

            with st.spinner("Processing... Please wait"):

                run_process(master_folder, loco_folder, division)


# -------------------------------------------------
# TERMINAL
# -------------------------------------------------

st.divider()

st.subheader("Process Output Terminal")

st.text_area(
    "",
    st.session_state.terminal_output,
    height=400
)