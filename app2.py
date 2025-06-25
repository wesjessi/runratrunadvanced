# app.py

import streamlit as st
import os
import uuid
import shutil
import tempfile

from runratrun6 import main_process

# Streamlit App Title
st.title("Running Data Analysis App")

# Sidebar for Input Mode
st.sidebar.header("Configuration")

mode = st.sidebar.radio(
    "Select data type:",
    [
        "Daily Files (Original Mode)",
        "Continuous Multi-Day File",
        "Multi-Day with Manual Cycle Start",
        "Custom Range"
    ]
)

# Common input directory (or file uploads)
input_dir = None
uploaded_files = []
use_local = st.sidebar.radio("Data Source:", ["Local Directory", "Upload Files"])

if use_local == "Local Directory":
    input_dir = st.sidebar.text_input("Path to directory containing .xlsx files", value=r"C:\path\to\files")
else:
    # File uploader
    uploaded_files = st.sidebar.file_uploader("Upload .xlsx files", accept_multiple_files=True)

if mode == "Daily Files (Original Mode)":
    st.write("**Original Daily Files** mode selected.")
    user_params = {}
elif mode == "Continuous Multi-Day File":
    st.write("**Continuous Multi-Day** mode selected.")
    start_cycle = st.sidebar.radio("Starting cycle?", ["Active", "Inactive"])
    first_cycle_start_str = st.sidebar.text_input("First cycle start (MM/DD/YYYY HH:MM)", value="01/10/2025 18:00")

    user_params = {
        "start_cycle": start_cycle,
        "first_cycle_start_str": first_cycle_start_str
    }
    
elif mode == "Custom Range":
    st.write("**Custom Range**: only analyze rows X through Y of each sheet")
    custom_start = st.sidebar.number_input("Start row", min_value=1, value=1)
    custom_end   = st.sidebar.number_input("End   row", min_value=1, value=60)
    user_params = {
        "custom_start": custom_start,
        "custom_end":   custom_end
    }
    
else:
    st.write("**Multi-Day with Manual Cycle Start** mode selected.")
    first_file_name = st.sidebar.text_input("Earliest file name", value="Day1.xlsx")
    start_row = st.sidebar.number_input("Row that begins the first cycle", min_value=1, value=20)
    user_params = {
        "first_file": first_file_name,
        "start_row": start_row
    }


# Button to start processing
if st.sidebar.button("Start Processing"):
    if use_local == "Local Directory":
        if not os.path.isdir(input_dir):
            st.error("Invalid directory. Please provide a valid path.")
        else:
            # We have a local dir
            output_dir = tempfile.mkdtemp()
            try:
                if mode == "Daily Files (Original Mode)":
                    main_process(input_dir, output_dir, mode="daily", user_params=user_params)
                elif mode == "Continuous Multi-Day File":
                    main_process(input_dir, output_dir, mode="continuous", user_params=user_params)
                elif mode == "Multi-Day with Manual Cycle Start": 
                    main_process(input_dir, output_dir, mode="daily_manual_start", user_params=user_params)
                elif mode == "Custom Range":
                    main_process(input_dir, output_dir, mode="custom", user_params=user_params)
                else:
                    st.error(f"Unrecognized mode: {mode}")
                    st.session_state["processed"] = False
                    raise ValueError(f"Unknown mode: {mode}")
                    
                st.success("Processing completed!")
                st.session_state["processed"] = True
                st.session_state["output_dir"] = output_dir
                
            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.session_state["processed"] = False 
    else:
        # Upload scenario
        if len(uploaded_files) == 0:
            st.error("Please upload at least one .xlsx file.")
        else:
            # Create a temp input dir
            session_id = str(uuid.uuid4())
            input_dir = os.path.join("uploaded_files", session_id)
            os.makedirs(input_dir, exist_ok=True)

            # Save uploads
            for uf in uploaded_files:
                with open(os.path.join(input_dir, uf.name), "wb") as f:
                    f.write(uf.getbuffer())

            st.sidebar.success(f"Uploaded files saved to {input_dir}")

            output_dir = tempfile.mkdtemp()
            try:
                if mode == "Daily Files (Original Mode)":
                    main_process(input_dir, output_dir, mode="daily", user_params=user_params)
                elif mode == "Continuous Multi-Day File":
                    main_process(input_dir, output_dir, mode="continuous", user_params=user_params)
                elif mode == "Multi-Day with Manual Cycle Start": 
                    main_process(input_dir, output_dir, mode="daily_manual_start", user_params=user_params)
                elif mode == "Custom Range":
                    main_process(input_dir, output_dir, mode="custom", user_params=user_params)
                else: 
                    st.error(f"Unrecognized mode: {mode}")
                    st.session_state["processed"] = False
                    raise ValueError(f"Unknown mode: {mode}")
                st.success("Processing completed!")
                st.session_state["processed"] = True
                st.session_state["output_dir"] = output_dir
            except Exception as e:
                st.error(f"Error: {str(e)}")
                st.session_state["processed"] = False

            finally:
                # Optionally clean up the input_dir if you no longer need it
                pass

# Download section
if st.session_state.get("processed", False):
    st.write("Download your output files below:")
    output_dir = st.session_state.get("output_dir")
    if output_dir and os.path.exists(output_dir):
        # 1. Build a small priority lookup:
        priority = {
            "Active_Data.xlsx": 0,
            "Inactive_Data.xlsx": 1,
            "Active_Hourly_ByHour.xlsx": 2,
            "Inactive_Hourly_ByHour.xlsx": 3,
            "Active_Hourly_ByDay.xlsx": 4, 
            "Inactive_Hourly_ByDay.xlsx": 5,
            # add more if you have other well-known filenames…
        }

        # 2. Gather all files, then sort by priority (default=99) and then alphabetically
        all_files = [
            f for f in os.listdir(output_dir)
            if os.path.isfile(os.path.join(output_dir, f))
        ]
        all_files.sort(key=lambda fn: (priority.get(fn, 99), fn.lower()))

        # 3. Loop in that order
        for filename in all_files:
            path = os.path.join(output_dir, filename)
            with open(path, "rb") as f:
                data = f.read()
            st.download_button(
                label=f"Download {filename}",
                data=data,
                file_name=filename
            )
