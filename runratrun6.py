# runratrun6.py

import os
import glob
import re
import uuid
import shutil
import pandas as pd
import openpyxl
from datetime import datetime, timedelta

# -----------------------------
# CONSTANTS & HELPER FUNCTIONS
# -----------------------------
ROW_SPLIT = 720            # For daily approach: ~12 hours in rows (if 1 row/min)
WHEEL_CIRCUMFERENCE = 1.081

def extract_date_from_filename(filename):
    """
    Extract a date of the form MM-DD-YY from the filename.
    Returns a datetime object or None if not found/parsable.
    """
    pattern = r'(\d{1,2}-\d{1,2}-\d{2})'  # or (\d{1,2}-\d{1,2}-\d{4}) for 4-digit year
    match = re.search(pattern, filename)
    if match:
        date_str = match.group(1)
        try:
            return datetime.strptime(date_str, '%m-%d-%y')
        except ValueError:
            return None
    return None


def calculate_running_bouts(wheel_turns):
    """
    Identifies 'running bouts' if Activity >=3 and the previous minute was <3.
    """
    bouts = []
    for i in range(len(wheel_turns)):
        if i == 0:
            # First row special check
            if wheel_turns[i] >= 3:
                bouts.append(1)
            else:
                bouts.append(0)
        else:
            if wheel_turns[i - 1] < 3 and wheel_turns[i] >= 3:
                bouts.append(1)
            else:
                bouts.append(0)
    return bouts


def calculate_metrics(data, sheet, date, segment, debug_data):
    """
    Calculate various metrics from the data:
      - total bouts
      - minutes running
      - total wheel turns
      - distance, speed, etc.
    """
    metrics = {}
    valid_wheel_turns = data['Activity'][data['Activity'] >= 3]

    metrics['Total_Bouts'] = data['Running_Bout'].sum()
    metrics['Minutes_Running'] = len(valid_wheel_turns)
    metrics['Total_Wheel_Turns'] = valid_wheel_turns.sum()
    metrics['Distance_m'] = metrics['Total_Wheel_Turns'] * WHEEL_CIRCUMFERENCE

    if metrics['Total_Bouts'] > 0:
        metrics['Avg_Distance_per_Bout'] = metrics['Distance_m'] / metrics['Total_Bouts']
        metrics['Avg_Bout_Length'] = metrics['Minutes_Running'] / metrics['Total_Bouts']
    else:
        metrics['Avg_Distance_per_Bout'] = 0
        metrics['Avg_Bout_Length'] = 0

    metrics['Speed'] = (metrics['Distance_m'] / metrics['Minutes_Running']) if metrics['Minutes_Running'] > 0 else 0

    # Append debug info
    debug_data.append({
        'Rat': sheet,
        'Date': date,
        'Segment': segment,
        'Wheel Turns Sum (>=3)': valid_wheel_turns.sum(),
        'Minutes Running (>= 3)': len(valid_wheel_turns),
        'Distance (meters)': metrics['Distance_m'],
        'Total Bouts': data['Running_Bout'].sum()
    })

    return metrics


def save_data_to_excel(output_dir, data_dict, filename):
    """
    Saves metrics to an Excel file with one sheet per metric.
    Each sheet has rows = rats, columns = date.
    """
    with pd.ExcelWriter(os.path.join(output_dir, filename)) as writer:
        for metric, rat_data in data_dict.items():
            if rat_data:
                df = pd.DataFrame(rat_data).T  # pivot so rows=rat, columns=date
                df.to_excel(writer, sheet_name=metric)


def save_hourly_data(output_dir, hourly_data, filename, phase_label):
    """
    Saves hourly data to an Excel file with separate sheets for each hour.
    """
    with pd.ExcelWriter(os.path.join(output_dir, filename)) as writer:
        has_data = False
        for hour, data in hourly_data.items():
            if data:  # Only process if there is data for this hour
                has_data = True
                rows = {'Rat': []}
                for rat, dates in data.items():
                    rows['Rat'].append(rat)
                    for day_index, (date, metrics) in enumerate(dates.items(), start=1):
                        for metric, value in metrics.items():
                            col_name = f"{metric} Day {day_index}"
                            if col_name not in rows:
                                rows[col_name] = []
                            rows[col_name].append(value)
                # Pad columns to the same length
                max_len = max(len(col) for col in rows.values())
                for col in rows:
                    while len(rows[col]) < max_len:
                        rows[col].append(None)
                pd.DataFrame(rows).to_excel(writer, sheet_name=f"{phase_label} Hour {hour + 1}", index=False)

        if not has_data:
            pd.DataFrame({"Message": ["No data available"]}).to_excel(writer, sheet_name="No Data")


# ------------------------------------
# MAIN PROCESS & SUB-FUNCTIONS
# ------------------------------------
def main_process(input_dir, output_dir, mode="daily", user_params=None):
    """
    This single main_process handles:
      1) Creating a temp folder
      2) Moving *.xlsx files there
      3) Calling the correct sub-function based on 'mode'
      4) Cleaning up temp folder
    """
    if user_params is None:
        user_params = {}

    # --- STEP A: Create a unique temp subfolder ---
    session_id = str(uuid.uuid4())
    temp_input_dir = os.path.join(input_dir, session_id)
    os.makedirs(temp_input_dir, exist_ok=True)

    try:
        # --- STEP B: Move all .xlsx files from input_dir into temp_input_dir
        original_xlsx_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
        if not original_xlsx_files:
            raise FileNotFoundError(f"No Excel files found in directory: {input_dir}")

        for filepath in original_xlsx_files:
            filename = os.path.basename(filepath)
            shutil.move(filepath, os.path.join(temp_input_dir, filename))

        # --- STEP C: Call the appropriate sub-function based on 'mode' ---
        if mode == "daily":
            process_daily_files(temp_input_dir, output_dir, user_params)
        elif mode == "continuous":
            process_continuous_file(temp_input_dir, output_dir, user_params)
        elif mode == "daily_manual_start":
            process_daily_with_manual_cycle_start(temp_input_dir, output_dir, user_params)
        else:
            raise ValueError(f"Unrecognized mode: {mode}")

    finally:
        # --- STEP D: Cleanup ---
        shutil.rmtree(temp_input_dir, ignore_errors=True)
        print(f"Cleaned up temporary folder: {temp_input_dir}")


def process_daily_files(input_dir, output_dir, user_params):
    """
    Original daily logic that uses ROW_SPLIT (719) to guess which half is Active or Inactive.
    Also includes offset logic if you want it (not shown below).
    """
    # Gather files
    excel_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in {input_dir} for daily processing.")

    # Sort files by date in filename
    excel_files = sorted(
        excel_files,
        key=lambda f: (extract_date_from_filename(os.path.basename(f)) or datetime.min)
    )

    # Data structures
    active_data = {m: {} for m in ['Total_Bouts','Minutes_Running','Total_Wheel_Turns','Distance_m','Avg_Distance_per_Bout','Avg_Bout_Length','Speed']}
    inactive_data = {m: {} for m in ['Total_Bouts','Minutes_Running','Total_Wheel_Turns','Distance_m','Avg_Distance_per_Bout','Avg_Bout_Length','Speed']}
    hourly_data_active = {hour: {} for hour in range(12)}
    hourly_data_inactive = {hour: {} for hour in range(12)}
    debug_data = []

    # Process each file
    for file in excel_files:
        xl = pd.ExcelFile(file)
        file_date_str = os.path.basename(file).split('.')[0]

        for sheet in xl.sheet_names:
            try:
                df_rat = xl.parse(sheet, usecols="A")  # only read 1 column for 'Activity'
                if 'Activity' in df_rat.columns and not df_rat['Activity'].isnull().all():
                    # Truncate if there are extra rows (more than 1440)
                    if len(df_rat) > 1440:
                        df_rat = df_rat.iloc[:1440]
                    df_rat['Running_Bout'] = calculate_running_bouts(df_rat['Activity'])

                    # Decide if first half or second half is Active
                    sum_first = df_rat['Activity'].iloc[:ROW_SPLIT].sum()
                    sum_second = df_rat['Activity'].iloc[ROW_SPLIT:].sum()
                    first_cycle = 'Active' if sum_first > sum_second else 'Inactive'

                    if first_cycle == 'Active':
                        df_active = df_rat.iloc[:ROW_SPLIT]
                        df_inactive = df_rat.iloc[ROW_SPLIT:]
                    else:
                        df_inactive = df_rat.iloc[:ROW_SPLIT]
                        df_active = df_rat.iloc[ROW_SPLIT:]

                    # Hourly breakdown (24 hours => 24 chunks of 60 rows)
                    for hour in range(24):
                        if hour < 12:
                            phase = first_cycle
                        else: 
                            phase = "Inactive" if first_cycle == "Active" else "Active"                             
                        start_hr = hour * 60
                        end_hr = start_hr + 60
                        hourly_df = df_rat.iloc[start_hr:end_hr].copy()

                        # Pad with zeros if it's short
                        if len(hourly_df) < 60:
                            padding = pd.DataFrame(0, index=range(60 - len(hourly_df)), columns=hourly_df.columns)
                            hourly_df = pd.concat([hourly_df, padding], ignore_index=True)

                        metrics = calculate_metrics(hourly_df, sheet, file_date_str, f"Hour {hour+1}", debug_data)
                        if phase == 'Active':
                            if sheet not in hourly_data_active[hour % 12]:
                                hourly_data_active[hour % 12][sheet] = {}
                            hourly_data_active[hour % 12][sheet][file_date_str] = metrics
                        else:
                            if sheet not in hourly_data_inactive[hour % 12]:
                                hourly_data_inactive[hour % 12][sheet] = {}
                            hourly_data_inactive[hour % 12][sheet][file_date_str] = metrics

                    # Calculate total active/inactive
                    active_metrics = calculate_metrics(df_active, sheet, file_date_str, "Active", debug_data)
                    inactive_metrics = calculate_metrics(df_inactive, sheet, file_date_str, "Inactive", debug_data)

                    # Store
                    for metric, val in active_metrics.items():
                        if sheet not in active_data[metric]:
                            active_data[metric][sheet] = {}
                        active_data[metric][sheet][file_date_str] = val

                    for metric, val in inactive_metrics.items():
                        if sheet not in inactive_data[metric]:
                            inactive_data[metric][sheet] = {}
                        inactive_data[metric][sheet][file_date_str] = val

                else:
                    debug_data.append({'File': file, 'Sheet': sheet, 'Error': 'Missing or invalid Activity column'})

            except Exception as e:
                debug_data.append({'File': file, 'Sheet': sheet, 'Error': str(e)})

        # (Optional) Save partial results here or after the loop

    # Save final active/inactive
    save_data_to_excel(output_dir, active_data, 'Active_Data.xlsx')
    save_data_to_excel(output_dir, inactive_data, 'Inactive_Data.xlsx')

    # Save hourly
    save_hourly_data(output_dir, hourly_data_active, 'Active_Hourly_Data.xlsx', "Active")
    save_hourly_data(output_dir, hourly_data_inactive, 'Inactive_Hourly_Data.xlsx', "Inactive")

    # Debug
    debug_df = pd.DataFrame(debug_data)
    debug_df.to_excel(os.path.join(output_dir, "Debug_Output.xlsx"), index=False)


def process_daily_with_manual_cycle_start(input_dir, output_dir, user_params):
    """
    Merges the 'original daily logic' (each sheet = one rat)
    with a manual, concatenated approach:
      1) Collect all .xlsx files in input_dir.
      2) For each file, open it with pd.ExcelFile(...).
      3) For each sheet (rat) in the file:
         - read column A, skipping row 1 if it's a text header
         - convert everything to numeric, filling invalid with 0
         - store the resulting DataFrame in a dictionary big_dfs[sheet].
      4) After reading all files, for each rat (sheet) in big_dfs:
         - concat all partial DataFrames -> one big DataFrame
         - slice it by 24-hour blocks (1440 rows)
         - split each block into 12-hour halves
         - use sum-based approach to decide which is active vs. inactive
         - compute metrics -> store in active_data / inactive_data
      5) Save final Active_Data.xlsx, Inactive_Data.xlsx, and Debug_Output.xlsx.
    """

    import pandas as pd
    import os, glob
    from math import ceil

    debug_data = []

    # 1) Gather .xlsx files
    excel_files = glob.glob(os.path.join(input_dir, "*.xlsx"))
    if not excel_files:
        raise FileNotFoundError(f"No Excel files found in {input_dir} for daily_manual_start.")

    # 2) Sort them by date or name if you like:
    excel_files = sorted(
        excel_files,
        key=lambda f: (extract_date_from_filename(os.path.basename(f)) or datetime.min)
    )

    # We'll store DataFrames by sheet name (rat name)
    big_dfs = {}  # e.g. big_dfs["Rat1"] = [df1, df2, df3...]

    # 3) Loop through each file and each sheet to read full data
    for f in excel_files:
        xl = pd.ExcelFile(f)
        for sheet_name in xl.sheet_names:
            try:
                # Skip non-rat sheets
                if not sheet_name.lower().startswith("rat"):
                    continue

                # Read the entire column (don’t skip rows here)
                df_rat = xl.parse(
                    sheet_name,
                    usecols="A",
                    header=None  # Adjust header if needed
                )
                df_rat.columns = ["Activity"]
                df_rat["Activity"] = pd.to_numeric(df_rat["Activity"], errors="coerce").fillna(0)
                
                # Enforce exactly 1440 rows by truncating if there are extra
                if len(df_rat) > 1440:
                    df_rat = df_rat.iloc[:1440]

                if sheet_name not in big_dfs:
                    big_dfs[sheet_name] = []
                big_dfs[sheet_name].append(df_rat)

            except Exception as e:
                debug_data.append({
                    "File": f,
                    "Sheet": sheet_name,
                    "Error": str(e)
                })

    # Now process each rat's data separately
    ROWS_PER_DAY = 1440

    # Prepare final dictionaries for metrics
    metrics_list = [
        'Total_Bouts','Minutes_Running','Total_Wheel_Turns',
        'Distance_m','Avg_Distance_per_Bout','Avg_Bout_Length','Speed'
    ]
    active_data = {m: {} for m in metrics_list}
    inactive_data = {m: {} for m in metrics_list}
    
    # Prepare dictionaries for hourly metrics (grouped into 12 buckets as in your original mode)
    hourly_data_active = {h: {} for h in range(12)}
    hourly_data_inactive = {h: {} for h in range(12)}
    
    # For each rat (sheet) in big_dfs
    for sheet in big_dfs.keys():
        # 1. Concatenate all DataFrames for the rat
        master_df = pd.concat(big_dfs[sheet], ignore_index=True)

        # 2. Apply the user-defined skip only once
        start_row = user_params.get("start_row", 1) - 1  # Convert to 0-based index
        master_df = master_df.iloc[start_row:].reset_index(drop=True)

        total_rows = len(master_df)

        # 3. Calculate Running_Bout, etc.
        master_df["Running_Bout"] = calculate_running_bouts(master_df["Activity"])

        # 4. Determine the number of days
        number_of_days = (total_rows + ROWS_PER_DAY - 1) // ROWS_PER_DAY

        # Continue with day segmentation, splitting into 12-hour halves, etc.
        # (Your existing logic for splitting and calculating metrics would follow here)


        row_pointer = 0
        for day_idx in range(1, number_of_days+1):
            # slice 1440 rows for day i
            day_block = master_df.iloc[row_pointer : row_pointer + ROWS_PER_DAY].copy()
            actual_len = len(day_block)

            # if short, pad with zeros
            if actual_len < ROWS_PER_DAY:
                missing = ROWS_PER_DAY - actual_len
                padding = pd.DataFrame({"Activity": [0]*missing})
                # compute running bouts for padding if needed
                padding["Running_Bout"] = [0]*missing
                day_block = pd.concat([day_block, padding], ignore_index=True)

            # split into 12-hour halves (720 rows each)
            ROWS_HALF = 720
            first_half = day_block.iloc[:ROWS_HALF]
            second_half = day_block.iloc[ROWS_HALF:]

            sum_first = first_half["Activity"].sum()
            sum_second = second_half["Activity"].sum()
            
            if sum_first > sum_second:
                df_active = first_half
                df_inactive = second_half
                first_half_active = True
            else:
                df_inactive = first_half
                df_active = second_half
                first_half_active = False

            # compute active metrics
            active_metrics = calculate_metrics(
                data=df_active,
                sheet=sheet,
                date=f"Day{day_idx}",
                segment="Active",
                debug_data=debug_data
            )
            # store
            for m, val in active_metrics.items():
                if sheet not in active_data[m]:
                    active_data[m][sheet] = {}
                active_data[m][sheet][f"Day{day_idx}"] = val

            # compute inactive metrics
            inactive_metrics = calculate_metrics(
                data=df_inactive,
                sheet=sheet,
                date=f"Day{day_idx}",
                segment="Inactive",
                debug_data=debug_data
            )
            for m, val in inactive_metrics.items():
                if sheet not in inactive_data[m]:
                    inactive_data[m][sheet] = {}
                inactive_data[m][sheet][f"Day{day_idx}"] = val
                
            # Hourly Segmentation Start
            # --------------------------
            # For each day, break the 1440 rows into 24 hourly segments (60 rows each)
            # and assign the phase based on whether the hour falls in the first or second half.
            for hour in range(24):
                start_hr = hour * 60
                end_hr = start_hr + 60
                hourly_df = day_block.iloc[start_hr:end_hr].copy()

                # If the hourly segment is short, pad with zeros
                if len(hourly_df) < 60:
                    padding = pd.DataFrame({"Activity": [0]*(60 - len(hourly_df))})
                    padding["Running_Bout"] = [0]*(60 - len(hourly_df))
                    hourly_df = pd.concat([hourly_df, padding], ignore_index=True)

                # Determine phase for this hourly block
                if hour < 12:
                    phase = "Active" if first_half_active else "Inactive"
                else:
                    phase = "Inactive" if first_half_active else "Active"

                hr_metrics = calculate_metrics(
                    hourly_df,
                    sheet=sheet,
                    date=f"Day{day_idx}",
                    segment=f"Hour {hour+1}",
                    debug_data=debug_data
                )
                # Use hour % 12 as the key (to group similar hours as in original mode)
                hour_key = hour % 12
                if phase == "Active":
                    if sheet not in hourly_data_active[hour_key]:
                        hourly_data_active[hour_key][sheet] = {}
                    hourly_data_active[hour_key][sheet][f"Day{day_idx}"] = hr_metrics
                else:
                    if sheet not in hourly_data_inactive[hour_key]:
                        hourly_data_inactive[hour_key][sheet] = {}
                    hourly_data_inactive[hour_key][sheet][f"Day{day_idx}"] = hr_metrics

            # End of day processing
            row_pointer += ROWS_PER_DAY

    # finally, save the results
    save_data_to_excel(output_dir, active_data, "Active_Data.xlsx")
    save_data_to_excel(output_dir, inactive_data, "Inactive_Data.xlsx")
    
    # Save hourly exports
    save_hourly_data(output_dir, hourly_data_active, "Active_Hourly_Data.xlsx", "Active")
    save_hourly_data(output_dir, hourly_data_inactive, "Inactive_Hourly_Data.xlsx", "Inactive")

    debug_df = pd.DataFrame(debug_data)
    debug_df.to_excel(os.path.join(output_dir, "Debug_Output.xlsx"), index=False)


def process_continuous_file(input_dir, output_dir, user_params):
    import pandas as pd
    
    all_files = glob.glob(os.path.join(input_dir, '*.xlsx'))
    if len(all_files) == 0:
        raise FileNotFoundError("No Excel files found for continuous mode.")
    
    file_path = all_files[0]  # assume single file
    
    # Gather user parameters
    start_cycle = user_params.get("start_cycle", "Active")
    # New parameter:
    first_cycle_start_str = user_params.get("first_cycle_start_str", None)

    # Set the start datetime using the first cycle start input
    start_datetime = datetime.strptime(first_cycle_start_str, "%m/%d/%Y %H:%M")
    
    # 1) Read the entire file with no header
    df_raw = pd.read_excel(file_path, header=None)
    
    # 2) Extract column names from the first row
    colnames = df_raw.iloc[0].tolist()  # e.g. ["Channel Name:", "RAT 1", "RAT 2", ...]
    colnames[0] = "Timestamp"          # rename the first to "Timestamp"
    
    # 3) skip the first 3 lines of metadata:
    df_data = df_raw.iloc[3:].copy()
    df_data.columns = colnames
    
    # Check timestamp
    if "Timestamp" not in df_data.columns:
        raise ValueError("Continuous file must have a 'Timestamp' column.")
    df_data["Timestamp"] = pd.to_datetime(df_data["Timestamp"])
    df_data = df_data.sort_values("Timestamp").reset_index(drop=True)

    # Prepare output structures
    metrics_list = ['Total_Bouts','Minutes_Running','Total_Wheel_Turns',
                    'Distance_m','Avg_Distance_per_Bout','Avg_Bout_Length','Speed']
    active_data = {m: {} for m in metrics_list}
    inactive_data = {m: {} for m in metrics_list}
    hourly_data_active = {}
    hourly_data_inactive = {}
    debug_data = []

    # Identify rat columns
    rat_columns = [c for c in df_data.columns if c.lower().startswith("rat")]

    # For each rat column
    for rat_col in rat_columns:
        sub_df = pd.DataFrame({
            "Timestamp": df_data["Timestamp"],
            "Activity": df_data[rat_col]
        })
        sub_df["Running_Bout"] = calculate_running_bouts(sub_df["Activity"])

        # Segmenting into 12-hour blocks from 'start_datetime'
        current_cycle = start_cycle
        segment_index = 0
        segment_start = start_datetime
        max_time = sub_df["Timestamp"].max()

        while segment_start < max_time:
            segment_end = segment_start + timedelta(hours=12)
            chunk = sub_df[(sub_df["Timestamp"] >= segment_start) & (sub_df["Timestamp"] < segment_end)]

            if len(chunk) > 0:
                seg_metrics = calculate_metrics(
                    chunk,
                    sheet=rat_col,
                    date=segment_start.strftime("%m-%d-%Y"),
                    segment=current_cycle,
                    debug_data=debug_data
                )
                # Store metrics in active or inactive
                if current_cycle == "Active":
                    for m, val in seg_metrics.items():
                        if rat_col not in active_data[m]:
                            active_data[m][rat_col] = {}
                        active_data[m][rat_col][f"Seg_{segment_index}"] = val
                else:
                    for m, val in seg_metrics.items():
                        if rat_col not in inactive_data[m]:
                            inactive_data[m][rat_col] = {}
                        inactive_data[m][rat_col][f"Seg_{segment_index}"] = val

            # Move to next segment
            segment_start = segment_end
            segment_index += 1
            # Possibly flip cycle each 12h
            current_cycle = "Inactive" if current_cycle == "Active" else "Active"
            
        # PART B) Hour-by-hour breakdown
        current_cycle = start_cycle
        hour_index = 0
        cycle_hours_used = 0
        hour_start = start_datetime

        while hour_start < max_time:
            hour_end = hour_start + timedelta(hours=1)
    
            # Ensure that the dictionaries have a key for the current hour index
            if hour_index not in hourly_data_active:
                hourly_data_active[hour_index] = {}
            if hour_index not in hourly_data_inactive:
                hourly_data_inactive[hour_index] = {}
    
            hour_chunk = sub_df[
                (sub_df["Timestamp"] >= hour_start) &
                (sub_df["Timestamp"] < hour_end)
            ]

            if len(hour_chunk) > 0:
                hour_metrics = calculate_metrics(
                    hour_chunk,
                    sheet=rat_col,
                    date=hour_start.strftime("%m-%d-%Y"),
                    segment=current_cycle,
                    debug_data=debug_data
                )
                # Store in hourly_data_* depending on cycle
                if current_cycle == "Active":
                    # Ensure rat_col entry exists within the current hour
                    if rat_col not in hourly_data_active[hour_index]:
                        hourly_data_active[hour_index][rat_col] = {}
                    hourly_data_active[hour_index][rat_col][hour_start.strftime("%m-%d-%Y")] = hour_metrics
                else:
                    if rat_col not in hourly_data_inactive[hour_index]:
                        hourly_data_inactive[hour_index][rat_col] = {}
                    hourly_data_inactive[hour_index][rat_col][hour_start.strftime("%m-%d-%Y")] = hour_metrics

            # Move to next hour
            hour_start = hour_end
            hour_index += 1
            cycle_hours_used += 1

            # Flip cycle every 12 hours
            if cycle_hours_used == 12:
                current_cycle = "Inactive" if current_cycle == "Active" else "Active"
                cycle_hours_used = 0

    # Finally, save data
    save_data_to_excel(output_dir, active_data, "Active_Data.xlsx")
    save_data_to_excel(output_dir, inactive_data, "Inactive_Data.xlsx")
    
    # B) Hourly "Active_Hourly_Data.xlsx" / "Inactive_Hourly_Data.xlsx"
    save_hourly_data(output_dir, hourly_data_active, "Active_Hourly_Data.xlsx", "Active")
    save_hourly_data(output_dir, hourly_data_inactive, "Inactive_Hourly_Data.xlsx", "Inactive")

    debug_df = pd.DataFrame(debug_data)
    debug_df.to_excel(os.path.join(output_dir, "Debug_Output.xlsx"), index=False)
