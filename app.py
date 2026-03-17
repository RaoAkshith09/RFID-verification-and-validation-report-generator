import streamlit as st
import sys
import os
import zipfile
import shutil
from datetime import datetime
from main import run_full_process  # your processing function

# -------------------------------------------------
# PAGE CONFIG
# -------------------------------------------------
st.set_page_config(
    page_title="RFID Tag Report Generator",
    layout="wide"
)

# -------------------------------------------------
# SESSION STATE
# -------------------------------------------------
if "terminal_output" not in st.session_state:
    st.session_state.terminal_output = ""

# -------------------------------------------------
# TERMINAL REDIRECT
# -------------------------------------------------
class StreamlitRedirect:
    def write(self, string):
        timestamp = datetime.now().strftime("%H:%M:%S")
        st.session_state.terminal_output += f"[{timestamp}] {string}"

    def flush(self):
        pass

# -------------------------------------------------
# TERMINAL FUNCTIONS
# -------------------------------------------------
def append_terminal(text):
    timestamp = datetime.now().strftime("%H:%M:%S")
    st.session_state.terminal_output += f"[{timestamp}] {text}\n"

def clear_terminal():
    st.session_state.terminal_output = ""

# -------------------------------------------------
# ZIP EXTRACTION FUNCTION
# -------------------------------------------------
def extract_zip(uploaded_zip, extract_to):
    # remove old folder
    if os.path.exists(extract_to):
        shutil.rmtree(extract_to)

    os.makedirs(extract_to, exist_ok=True)

    zip_path = os.path.join(extract_to, uploaded_zip.name)

    # save uploaded zip
    with open(zip_path, "wb") as f:
        f.write(uploaded_zip.getbuffer())

    # extract zip
    with zipfile.ZipFile(zip_path, "r") as zip_ref:
        zip_ref.extractall(extract_to)

    return extract_to

# -------------------------------------------------
# MAIN PROCESS RUNNER
# -------------------------------------------------
def run_process(master_zip, loco_zip, division):
    master_folder = "master_data"
    loco_folder = "loco_logs"

    try:
        clear_terminal()
        append_terminal("=" * 60)
        append_terminal("RFID TAG REPORT GENERATOR STARTED")
        append_terminal("=" * 60)
        append_terminal(f"Division : {division}")
        append_terminal("-" * 60)

        append_terminal("Extracting Master Data ZIP...")
        master_path = extract_zip(master_zip, master_folder)

        append_terminal("Extracting Loco Log ZIP...")
        loco_path = extract_zip(loco_zip, loco_folder)
        append_terminal("Files extracted successfully")
        append_terminal("-" * 60)

        # redirect stdout to Streamlit terminal
        old_stdout = sys.stdout
        sys.stdout = StreamlitRedirect()

        # Run main processing
        output_folder = run_full_process(master_path, loco_path, division, append_terminal)

        # Restore stdout
        sys.stdout = old_stdout

        # Create ZIP of output folder
        zip_name = "Tag_verification_reports"
        zip_path = f"{zip_name}.zip"

        if os.path.exists(zip_path):
            os.remove(zip_path)

        # create zip
        shutil.make_archive(zip_name, 'zip', output_folder)

        # download button
        with open(zip_path, "rb") as f:
            st.download_button(
                label="Download Tag Verification Reports",
                data=f,
                file_name=zip_path,
                mime="application/zip"
            )

        append_terminal("REPORTS GENERATED SUCCESSFULLY")
        st.success("Reports Generated Successfully!")

    except Exception as e:
        sys.stdout = old_stdout
        append_terminal(f"ERROR : {str(e)}")
        st.error(str(e))

# -------------------------------------------------
# UI
# -------------------------------------------------
st.title("RFID Tag Verification and Validation Report Generator")
st.caption("Automated Railway RFID Validation Tool")
st.divider()

# File upload
col1, col2 = st.columns(2)
with col1:
    master_zip = st.file_uploader("Upload Master Data ZIP", type="zip")
with col2:
    loco_zip = st.file_uploader("Upload Loco Log ZIP", type="zip")

division = st.text_input("Division Name")
st.divider()

# Buttons
col1, col2 = st.columns(2)
generate = col1.button("Generate Reports", use_container_width=True)
clear = col2.button("Clear Terminal", use_container_width=True)

if clear:
    clear_terminal()

# Run process
if generate:
    if not master_zip or not loco_zip or not division:
        st.warning("Please upload both ZIP files and enter division")
    else:
        with st.spinner("Processing... Please wait"):
            run_process(master_zip, loco_zip, division)

# Terminal output
st.subheader("Process Output Terminal")
st.text_area(
    "terminal",
    st.session_state.terminal_output,
    height=400,
    label_visibility="collapsed"
)