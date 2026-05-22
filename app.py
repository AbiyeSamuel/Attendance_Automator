import streamlit as st
import pandas as pd
import datetime
import io

# --- 1. PAGE CONFIGURATION & ROBUST CSS ---
st.set_page_config(page_title="Attendance Automator", page_icon="🏫", layout="wide")

st.markdown("""
    <style>
        /* Light and clean app background */
        [data-testid="stAppViewContainer"] {
            background-color: #f8fafc;
        }

        /* Style the File Uploader Dropzones */
        [data-testid="stFileUploadDropzone"] {
            background-color: #ffffff;
            border: 2px dashed #3b82f6;
            border-radius: 12px;
            padding: 20px;
        }

        /* Style standard buttons (Generate Report) */
        [data-testid="baseButton-secondary"] {
            border: 2px solid #3b82f6;
            color: #3b82f6;
            border-radius: 8px;
            font-weight: bold;
            transition: all 0.3s ease;
        }
        [data-testid="baseButton-secondary"]:hover {
            background-color: #3b82f6;
            color: white;
            box-shadow: 0 4px 10px rgba(59, 130, 246, 0.3);
        }

        /* Style Primary buttons (Download) */
        [data-testid="baseButton-primary"] {
            background-color: #10b981;
            color: white;
            border-radius: 8px;
            font-weight: bold;
            border: none;
            transition: all 0.3s ease;
        }
        [data-testid="baseButton-primary"]:hover {
            background-color: #059669;
            box-shadow: 0 4px 10px rgba(16, 185, 129, 0.3);
        }

        /* Clean up metric cards */
        [data-testid="stMetricValue"] {
            color: #1e293b;
            font-weight: 800;
        }
    </style>
""", unsafe_allow_html=True)


# --- 2. THE TIME COMPRESSION ENGINE ---
def scale_time(time_str):
    try:
        t = datetime.datetime.strptime(time_str, '%H:%M:%S')

        m_start = datetime.datetime.strptime('10:00:00', '%H:%M:%S')
        m_end = datetime.datetime.strptime('11:30:00', '%H:%M:%S')
        a_start = datetime.datetime.strptime('14:00:00', '%H:%M:%S')
        a_end = datetime.datetime.strptime('14:50:00', '%H:%M:%S')

        target_seconds = 1800

        if m_start <= t <= m_end:
            offset = t - m_start
            source_duration = (m_end - m_start).total_seconds()
            new_offset = datetime.timedelta(seconds=int(offset.total_seconds() * (target_seconds / source_duration)))
            return (m_start + new_offset).strftime('%H:%M:%S')

        elif a_start <= t <= a_end:
            offset = t - a_start
            source_duration = (a_end - a_start).total_seconds()
            new_offset = datetime.timedelta(seconds=int(offset.total_seconds() * (target_seconds / source_duration)))
            return (a_start + new_offset).strftime('%H:%M:%S')

        return time_str
    except Exception:
        return time_str

    # --- 3. GRADING LOGIC ---


def grade_attendance(time_str, session):
    try:
        t = datetime.datetime.strptime(time_str, '%H:%M:%S').time()
        if session == 'Morning':
            if t <= datetime.time(10, 30, 0):
                return 'PRESENT', 0.5
            elif t <= datetime.time(11, 0, 0):
                return 'LATE', 0.25
            else:
                return 'ABSENT', 0
        else:  # Afternoon
            if t <= datetime.time(14, 30, 0):
                return 'PRESENT', 0.5
            elif t <= datetime.time(14, 50, 0):
                return 'LATE', 0.25
            else:
                return 'ABSENT', 0
    except:
        return 'ABSENT', 0


# --- 4. APP UI HEADER ---
st.title("🏫 Weekly Attendance Automator")
st.markdown(
    "Upload your biometric `.dat` logs and your Names list to generate a **highly professional, graded attendance report**.")
st.divider()

# --- 5. UPLOAD ZONES ---
col1, col2 = st.columns(2)
with col1:
    uploaded_files = st.file_uploader("📂 1. Upload Biometric Logs (.dat)", type=["dat"], accept_multiple_files=True)
with col2:
    names_file = st.file_uploader("📝 2. Upload Student Names (.csv or .xlsx)", type=["csv", "xlsx"])

if uploaded_files and names_file:
    all_data = []

    # Process the DAT files
    for file in uploaded_files:
        lines = file.getvalue().decode("utf-8").splitlines()
        for line in lines:
            parts = line.strip().split()
            if len(parts) >= 3:
                parts[2] = scale_time(parts[2])
                all_data.append(parts)

    if all_data:
        raw_df = pd.DataFrame(all_data).iloc[:, :3]
        raw_df.columns = ["ID", "Date", "Time"]
        raw_df['ID'] = raw_df['ID'].astype(str)

        # Remove exact duplicates
        raw_df = raw_df.drop_duplicates()

        st.markdown("<br>", unsafe_allow_html=True)

        # Process Button
        if st.button("🚀 Generate Professional Report", use_container_width=True):
            with st.spinner("Crunching numbers and generating your report..."):
                try:
                    # 1. Map Names based on CSV or XLSX
                    if names_file.name.endswith('.csv'):
                        names_df = pd.read_csv(names_file)
                    else:
                        names_df = pd.read_excel(names_file)

                    id_col = 'Attendance_ID'
                    name_col = 'Full_Name'

                    id_to_name = {str(k).split('.')[0]: str(v) for k, v in zip(names_df[id_col], names_df[name_col])}
                    raw_df['Full_Name'] = raw_df['ID'].map(id_to_name).fillna("Unknown ID")

                    # 2. Assign Sessions
                    raw_df['Session'] = raw_df['Time'].apply(
                        lambda x: 'Morning' if int(x.split(':')[0]) < 13 else 'Afternoon'
                    )

                    # 3. Grade the attendance
                    raw_df[['Status', 'Points']] = raw_df.apply(
                        lambda row: pd.Series(grade_attendance(row['Time'], row['Session'])), axis=1
                    )

                    # 4. Build Detailed Grid
                    first_scans = raw_df.sort_values('Time').groupby(
                        ['ID', 'Full_Name', 'Date', 'Session']).first().reset_index()
                    grid_df = first_scans.pivot_table(
                        index=['ID', 'Full_Name'],
                        columns=['Date', 'Session'],
                        values='Status',
                        aggfunc='first',
                        fill_value='ABSENT'
                    )

                    # 5. Build Review Metrics
                    points_df = first_scans.groupby(['ID', 'Full_Name'])['Points'].sum().reset_index()
                    total_days = len(first_scans['Date'].unique())
                    total_possible_points = total_days * 1.0

                    points_df['Total Points Earned'] = points_df['Points']
                    points_df['Total Possible'] = total_possible_points
                    points_df['Attendance (%)'] = ((points_df['Points'] / total_possible_points) * 100).round(1).astype(
                        str) + '%'
                    points_df.drop(columns=['Points'], inplace=True)

                    # 6. Build Executive Summary Data
                    total_delegates = len(points_df)
                    average_attendance = (points_df['Total Points Earned'].sum() / (
                                total_delegates * total_possible_points)) * 100

                    summary_data = {
                        "Metric": ["Total Class Days Logged", "Total Active Delegates",
                                   "Average Overall Attendance Rate"],
                        "Value": [total_days, total_delegates, f"{average_attendance:.1f}%"]
                    }
                    summary_df = pd.DataFrame(summary_data)

                    # --- EXCEL WRITER ---
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        summary_df.to_excel(writer, sheet_name='Executive_Summary', index=False)
                        points_df.to_excel(writer, sheet_name='Delegate_Review', index=False)
                        grid_df.to_excel(writer, sheet_name='Detailed_Grid')
                    output.seek(0)

                    st.success("✨ Success! Your Professional Report is ready.")
                    st.divider()

                    # --- BEAUTIFUL METRICS UI PREVIEW ---
                    st.write("### 📊 Executive Summary Preview")
                    met1, met2, met3 = st.columns(3)
                    met1.metric(label="Total Class Days", value=total_days)
                    met2.metric(label="Active Delegates", value=total_delegates)
                    met3.metric(label="Average Attendance", value=f"{average_attendance:.1f}%")

                    st.markdown("<br>", unsafe_allow_html=True)

                    # Target 'primary' button style for the download
                    st.download_button(
                        label="📥 Download Final Attendance Report",
                        data=output,
                        file_name="Weekly_Attendance_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"An error occurred: {e}")