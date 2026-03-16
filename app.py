import streamlit as st
import sys
from io import StringIO
from datetime import datetime
from main import run_full_process

# ------------------ Streamlit Page Config ------------------

st.set_page_config(
    page_title="RFID Tag Report Generator",
    layout="wide"
)

# ------------------ Session State ------------------

if "terminal_output" not in st.session_state:
    st.session_state.terminal_output = ""

# ------------------ Terminal Redirect ------------------

class StreamlitRedirect:
    def __init__(self):
        self.buffer = StringIO()

    def write(self, string):
        timestamp = datetime.now().strftime("%H:%M:%S")
        text = f"[{timestamp}] {string}"
        st.session_state.terminal_output += text

    def flush(self):
        pass

# ------------------ Functions ------------------

def append_terminal(text):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.terminal_output += f"[{timestamp}] {text}\n"

def clear_terminal():
    st.session_state.terminal_output = ""

def run_process(master_folder, loco_folder, division):

    try:

        clear_terminal()

        append_terminal("="*60)
        append_terminal("RFID TAG REPORT GENERATOR - PROCESSING STARTED")
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

        append_terminal("REPORTS GENERATED SUCCESSFULLY")

        st.success("Reports Generated Successfully!")

    except Exception as e:

        sys.stdout = old_stdout
        append_terminal(f"ERROR : {str(e)}")
        st.error(str(e))

# ------------------ UI ------------------

st.title("RFID Tag Verification and Validation Report Generator")
st.caption("Automated Report Generation Tool")

st.divider()

# ------------------ Inputs ------------------

col1, col2 = st.columns(2)

with col1:
    master_folder = st.text_input("Master Data Folder")

with col2:
    loco_folder = st.text_input("Loco Log Folder")

division = st.text_input("Division Name")

st.divider()

# ------------------ Buttons ------------------

col1, col2 = st.columns([1,1])

generate = col1.button("Generate Reports", use_container_width=True)

clear = col2.button("Clear Terminal", use_container_width=True)

if clear:
    clear_terminal()

# ------------------ Run Process ------------------

if generate:

    if not master_folder or not loco_folder or not division:
        st.warning("Please fill all inputs")
    else:
        with st.spinner("Processing... Please wait"):
            run_process(master_folder, loco_folder, division)

# ------------------ Terminal ------------------

st.subheader("Process Output Terminal")

terminal_box = st.empty()

terminal_box.text_area(
    "",
    st.session_state.terminal_output,
    height=400
)