import streamlit as st
import sys
import os
from io import StringIO
from datetime import datetime
from main import run_full_process


# ---------------------------------------------------
# PAGE CONFIG
# ---------------------------------------------------

st.set_page_config(
    page_title="RFID Tag Report Generator",
    layout="wide"
)

st.title("RFID Tag Verification and Validation Report Generator")
st.caption("Automated RFID Tag Report Generation Tool")


# ---------------------------------------------------
# SESSION STATE
# ---------------------------------------------------

if "terminal_output" not in st.session_state:
    st.session_state.terminal_output = ""


# ---------------------------------------------------
# TERMINAL REDIRECT CLASS
# ---------------------------------------------------

class StreamlitRedirect:

    def write(self, text):

        timestamp = datetime.now().strftime("%H:%M:%S")

        st.session_state.terminal_output += f"[{timestamp}] {text}"

    def flush(self):
        pass


# ---------------------------------------------------
# TERMINAL FUNCTIONS
# ---------------------------------------------------

def append_terminal(text):

    timestamp = datetime.now().strftime("%H:%M:%S")

    st.session_state.terminal_output += f"[{timestamp}] {text}\n"


def clear_terminal():

    st.session_state.terminal_output = ""


# ---------------------------------------------------
# PROCESS FUNCTION
# ---------------------------------------------------

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

        # Debug checks
        append_terminal(f"Master folder exists: {os.path.exists(master_folder)}")
        append_terminal(f"Loco folder exists  : {os.path.exists(loco_folder)}")

        # Redirect stdout
        old_stdout = sys.stdout
        sys.stdout = StreamlitRedirect()

        # Run main processing
        run_full_process(master_folder, loco_folder, division)

        # Restore stdout
        sys.stdout = old_stdout

        append_terminal("-"*60)
        append_terminal("REPORT GENERATION COMPLETED SUCCESSFULLY")

        st.success("Reports Generated Successfully!")

    except Exception as e:

        sys.stdout = old_stdout

        append_terminal(f"ERROR : {str(e)}")

        st.error(str(e))


# ---------------------------------------------------
# INPUT SECTION
# ---------------------------------------------------

st.divider()

col1, col2 = st.columns(2)

with col1:

    master_folder = st.text_input(
        "Master Data Folder",
        placeholder="Example: C:/Users/Akshith/Desktop/master_data"
    )

with col2:

    loco_folder = st.text_input(
        "Loco Log Folder",
        placeholder="Example: C:/Users/Akshith/Desktop/loco_logs"
    )

division = st.text_input(
    "Division Name",
    placeholder="Example: Nagpur"
)

st.divider()


# ---------------------------------------------------
# BUTTONS
# ---------------------------------------------------

col1, col2 = st.columns(2)

generate_btn = col1.button(
    "Generate Reports",
    use_container_width=True
)

clear_btn = col2.button(
    "Clear Terminal",
    use_container_width=True
)

if clear_btn:
    clear_terminal()


# ---------------------------------------------------
# PROCESS EXECUTION
# ---------------------------------------------------

if generate_btn:

    if not master_folder or not loco_folder or not division:

        st.warning("Please fill all inputs")

    else:

        # Validate paths
        if not os.path.exists(master_folder):

            st.error("Master folder does not exist")
            st.write(master_folder)
            st.stop()

        if not os.path.exists(loco_folder):

            st.error("Loco folder does not exist")
            st.write(loco_folder)
            st.stop()

        with st.spinner("Processing... Please wait"):

            run_process(master_folder, loco_folder, division)


# ---------------------------------------------------
# TERMINAL OUTPUT
# ---------------------------------------------------

st.divider()

st.subheader("Process Output Terminal")

st.text_area(
    "",
    st.session_state.terminal_output,
    height=400
)