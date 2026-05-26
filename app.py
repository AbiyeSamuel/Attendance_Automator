import streamlit as st
import pandas as pd
import datetime
import io
import subprocess
import sys
from openpyxl import load_workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Font, PatternFill

# Auto-install missing dependencies
try:
    import openpyxl
except ImportError:
    st.warning("📦 Installing missing 'openpyxl'...")
    subprocess.check_call([sys.executable, "-m", "pip", "install", "openpyxl"])
    st.success("✅ openpyxl installed. Please rerun the app.")
    st.stop()

st.set_page_config(page_title="Attendance Automator Pro", page_icon="🏫", layout="wide")

st.markdown("""
    <style>
        h1 { font-weight: 800; }
        [data-testid="stFileUploadDropzone"] { border: 2px dashed #3b82f6; border-radius: 12px; background-color: rgba(59, 130, 246, 0.05); }
        [data-testid="baseButton-secondary"] { border: 2px solid #3b82f6; color: #3b82f6; border-radius: 8px; font-weight: bold; }
        [data-testid="baseButton-secondary"]:hover { background-color: #3b82f6; color: white; }
        [data-testid="baseButton-primary"] { background-color: #10b981; color: white; border-radius: 8px; }
    </style>
""", unsafe_allow_html=True)

# --- Helper Functions ---
def scale_time(time_str: str) -> str:
    try:
        t = datetime.datetime.strptime(time_str.split('.')[0], '%H:%M:%S')
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
        return time_str.split('.')[0]
    except:
        return time_str

def grade_attendance(time_str: str, session: str):
    try:
        t = datetime.datetime.strptime(time_str.split('.')[0], '%H:%M:%S').time()
        if session == 'Morning':
            if t <= datetime.time(10, 30, 0):
                return 'PRESENT', 0.5
            elif t <= datetime.time(11, 0, 0):
                return 'LATE', 0.25
            else:
                return 'ABSENT', 0
        else:
            if t <= datetime.time(14, 30, 0):
                return 'PRESENT', 0.5
            elif t <= datetime.time(14, 50, 0):
                return 'LATE', 0.25
            else:
                return 'ABSENT', 0
    except:
        return 'ABSENT', 0

def parse_dat_file(file) -> pd.DataFrame:
    records = []
    for line in file.getvalue().decode("utf-8", errors='replace').splitlines():
        parts = line.strip().split()
        if len(parts) >= 3:
            records.append([parts[0], parts[1], scale_time(parts[2])])
    df = pd.DataFrame(records, columns=["ID", "Date", "Time"])
    return df.drop_duplicates()

def add_charts_and_insights(workbook, summary_df, dist_df):
    """Add charts, dashboard, and insights to the existing workbook."""
    # --- 1. Create Executive Dashboard sheet ---
    dash_sheet = workbook.create_sheet("Executive_Dashboard")
    dash_sheet['A1'] = "KEY PERFORMANCE INDICATORS"
    dash_sheet['A1'].font = Font(bold=True, size=14)
    kpi_data = [
        ["Metric", "Value"],
        ["Total Class Days", summary_df.loc[0, "Value"]],
        ["Active Delegates", summary_df.loc[1, "Value"]],
        ["Avg Attendance Rate", summary_df.loc[2, "Value"]],
    ]
    for row_idx, row in enumerate(kpi_data, start=3):
        for col_idx, val in enumerate(row, start=1):
            dash_sheet.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 3:
                dash_sheet.cell(row=row_idx, column=col_idx).font = Font(bold=True)

    # Distribution data for chart
    for i, row in dist_df.iterrows():
        dash_sheet.cell(row=13 + i, column=1, value=row["Attendance Band"])
        dash_sheet.cell(row=13 + i, column=2, value=row["Number of Students"])
    bar_chart = BarChart()
    bar_chart.title = "Attendance Distribution"
    bar_chart.x_axis.title = "Attendance Range"
    bar_chart.y_axis.title = "Number of Students"
    data = Reference(dash_sheet, min_col=2, min_row=13, max_row=12 + len(dist_df))
    cats = Reference(dash_sheet, min_col=1, min_row=13, max_row=12 + len(dist_df))
    bar_chart.add_data(data, titles_from_data=True)
    bar_chart.set_categories(cats)
    dash_sheet.add_chart(bar_chart, "D3")

    # --- 2. Add line chart to existing Daily_Trend sheet ---
    if "Daily_Trend" in workbook.sheetnames:
        trend_sheet = workbook["Daily_Trend"]
        line_chart = LineChart()
        line_chart.title = "Daily Attendance Rate Trend"
        line_chart.y_axis.title = "Attendance Rate (%)"
        line_chart.x_axis.title = "Date"
        max_row = trend_sheet.max_row
        data = Reference(trend_sheet, min_col=2, min_row=1, max_row=max_row, max_col=2)
        cats = Reference(trend_sheet, min_col=1, min_row=2, max_row=max_row)
        line_chart.add_data(data, titles_from_data=True)
        line_chart.set_categories(cats)
        trend_sheet.add_chart(line_chart, "D2")

    # --- 3. Add stacked bar chart to Session_Heatmap_Data sheet ---
    if "Session_Heatmap_Data" in workbook.sheetnames:
        heat_sheet = workbook["Session_Heatmap_Data"]
        bar_chart2 = BarChart()
        bar_chart2.title = "Session Performance: Morning vs Afternoon"
        bar_chart2.type = "col"
        bar_chart2.grouping = "stacked"
        max_row = heat_sheet.max_row
        data2 = Reference(heat_sheet, min_col=2, min_row=1, max_row=max_row, max_col=3)
        cats2 = Reference(heat_sheet, min_col=1, min_row=2, max_row=max_row)
        bar_chart2.add_data(data2, titles_from_data=True)
        bar_chart2.set_categories(cats2)
        heat_sheet.add_chart(bar_chart2, "F2")

    # --- 4. Add conditional formatting & pie chart to Risk_Analysis_Table ---
    if "Risk_Analysis_Table" in workbook.sheetnames:
        risk_sheet = workbook["Risk_Analysis_Table"]
        red_fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
        green_fill = PatternFill(start_color="CCFFCC", end_color="CCFFCC", fill_type="solid")
        # Risk Level is column C (index 3)
        for row in range(2, risk_sheet.max_row + 1):
            cell = risk_sheet.cell(row=row, column=3)
            if cell.value == "High Risk":
                cell.fill = red_fill
            elif cell.value == "Moderate Risk":
                cell.fill = yellow_fill
            elif cell.value == "Low Risk":
                cell.fill = green_fill

        # Build risk counts from the sheet data (or use provided dist_df but more accurate to use sheet)
        risk_counts = {}
        for row in range(2, risk_sheet.max_row + 1):
            risk = risk_sheet.cell(row=row, column=3).value
            risk_counts[risk] = risk_counts.get(risk, 0) + 1
        # Write counts to sheet
        start_row = risk_sheet.max_row + 3
        risk_sheet.cell(row=start_row, column=1, value="Risk Level")
        risk_sheet.cell(row=start_row, column=2, value="Count")
        row_idx = start_row + 1
        for risk, count in risk_counts.items():
            risk_sheet.cell(row=row_idx, column=1, value=risk)
            risk_sheet.cell(row=row_idx, column=2, value=count)
            row_idx += 1
        # Pie chart
        pie_chart = PieChart()
        pie_chart.title = "Risk Level Distribution"
        pie_data = Reference(risk_sheet, min_col=2, min_row=start_row, max_row=start_row + len(risk_counts), max_col=2)
        pie_cats = Reference(risk_sheet, min_col=1, min_row=start_row + 1, max_row=start_row + len(risk_counts))
        pie_chart.add_data(pie_data, titles_from_data=True)
        pie_chart.set_categories(pie_cats)
        risk_sheet.add_chart(pie_chart, "E2")

    # --- 5. Create Insights sheet with auto-generated text ---
    insights_sheet = workbook.create_sheet("Insights")
    insights_sheet.column_dimensions['A'].width = 80
    insights_sheet['A1'] = "AUTOMATED INSIGHTS & RECOMMENDATIONS"
    insights_sheet['A1'].font = Font(bold=True, size=12)
    row = 3
    avg_att = float(summary_df.loc[2, "Value"].strip('%'))
    total_students = summary_df.loc[1, "Value"]
    high_risk = risk_counts.get("High Risk", 0)
    moderate_risk = risk_counts.get("Moderate Risk", 0)
    perfect_att = 0
    # Count perfect attendance from risk_sheet (attendance % = 100.0%)
    if "Risk_Analysis_Table" in workbook.sheetnames:
        for r in range(2, workbook["Risk_Analysis_Table"].max_row + 1):
            att_cell = workbook["Risk_Analysis_Table"].cell(row=r, column=2)
            if att_cell.value == "100.0%":
                perfect_att += 1

    # Get best/worst day from Daily_Trend sheet
    if "Daily_Trend" in workbook.sheetnames:
        trend_sheet = workbook["Daily_Trend"]
        dates = []
        rates = []
        for r in range(2, trend_sheet.max_row + 1):
            dates.append(trend_sheet.cell(row=r, column=1).value)
            rates.append(float(trend_sheet.cell(row=r, column=2).value))
        if rates:
            best_idx = rates.index(max(rates))
            worst_idx = rates.index(min(rates))
            best_day = dates[best_idx]
            worst_day = dates[worst_idx]
            best_rate = max(rates)
            worst_rate = min(rates)
        else:
            best_day = worst_day = "N/A"
            best_rate = worst_rate = 0
    else:
        best_day = worst_day = "N/A"
        best_rate = worst_rate = 0

    # Morning/Afternoon comparison from Session_Heatmap_Data
    if "Session_Heatmap_Data" in workbook.sheetnames:
        heat_sheet = workbook["Session_Heatmap_Data"]
        morning_vals = []
        afternoon_vals = []
        for r in range(2, heat_sheet.max_row + 1):
            morning_vals.append(float(heat_sheet.cell(row=r, column=2).value))
            afternoon_vals.append(float(heat_sheet.cell(row=r, column=3).value))
        avg_morning = sum(morning_vals) / len(morning_vals) if morning_vals else 0
        avg_afternoon = sum(afternoon_vals) / len(afternoon_vals) if afternoon_vals else 0
        diff = abs(avg_morning - avg_afternoon)
        afternoon_lower = avg_afternoon < avg_morning
    else:
        avg_morning = avg_afternoon = 0
        diff = 0
        afternoon_lower = False

    insights = [
        f"• Overall attendance rate is {avg_att:.1f}%. {'Above average' if avg_att > 75 else 'Below expected threshold (75%). Immediate intervention recommended.'}",
        f"• {high_risk} students ({high_risk/total_students*100:.1f}%) are at HIGH RISK with attendance below 50%.",
        f"• {moderate_risk} students are at MODERATE RISK (50-75%). Targeted follow-up needed.",
        f"• {perfect_att} students have perfect attendance – consider recognition awards.",
        f"• Best attended day: {best_day} with {best_rate:.1f}% attendance.",
        f"• Worst attended day: {worst_day} with {worst_rate:.1f}% attendance.",
        f"• Afternoon sessions are {diff:.1f}% {'lower' if afternoon_lower else 'higher'} than morning sessions."
    ]
    for ins in insights:
        insights_sheet.cell(row=row, column=1, value=ins)
        row += 1

    insights_sheet.cell(row=row+1, column=1, value="RECOMMENDATIONS:")
    insights_sheet.cell(row=row+1, column=1).font = Font(bold=True)
    row += 2
    recos = [
        "• Send warning letters to high-risk students and schedule parent meetings.",
        "• Offer incentives for perfect attendance (certificates, extra credit).",
        "• Review schedule for low-attendance days – consider rescheduling important topics.",
        "• Send SMS/email reminders 1 hour before afternoon sessions.",
        "• Implement a buddy system for students with moderate risk."
    ]
    for rec in recos:
        insights_sheet.cell(row=row, column=1, value=rec)
        row += 1

    return workbook

# --- Main App ---
st.title("🏫 Weekly Attendance Automator Pro")
st.markdown("Upload biometric `.dat` logs and names file → Receive an **advanced analytics report** with charts and insights.")
st.divider()

col1, col2 = st.columns(2)
with col1:
    uploaded_files = st.file_uploader("📂 Biometric Logs (.dat)", type=["dat"], accept_multiple_files=True)
with col2:
    names_file = st.file_uploader("📝 Student Names (.csv or .xlsx)", type=["csv", "xlsx"])

exclude_unknown = st.checkbox("Exclude 'Unknown ID' from analysis", value=True)

if uploaded_files and names_file:
    all_dfs = []
    for f in uploaded_files:
        try:
            all_dfs.append(parse_dat_file(f))
        except Exception as e:
            st.warning(f"Error in {f.name}: {e}")

    if all_dfs:
        raw_df = pd.concat(all_dfs, ignore_index=True)
        raw_df['ID'] = raw_df['ID'].astype(str)

        if st.button("🚀 Generate Pro Analytics Report", use_container_width=True):
            with st.spinner("Processing and building advanced analytics..."):
                try:
                    # Load names
                    if names_file.name.endswith('.csv'):
                        names_df = pd.read_csv(names_file)
                    else:
                        names_df = pd.read_excel(names_file, engine='openpyxl')
                    if not {'Attendance_ID', 'Full_Name'}.issubset(names_df.columns):
                        st.error("Names file must have columns: 'Attendance_ID' and 'Full_Name'")
                        st.stop()
                    id_to_name = {str(row['Attendance_ID']): str(row['Full_Name']) for _, row in names_df.iterrows()}
                    raw_df['Full_Name'] = raw_df['ID'].map(id_to_name).fillna("Unknown ID")

                    # Session & grading
                    raw_df['Session'] = raw_df['Time'].apply(lambda x: 'Morning' if int(x.split(':')[0]) < 13 else 'Afternoon')
                    raw_df[['Status', 'Points']] = raw_df.apply(lambda row: pd.Series(grade_attendance(row['Time'], row['Session'])), axis=1)
                    first_scans = raw_df.sort_values('Time').groupby(['ID', 'Full_Name', 'Date', 'Session'], as_index=False).first()

                    # Detailed Grid (pivot)
                    grid_df = first_scans.pivot_table(
                        index=['ID', 'Full_Name'],
                        columns=['Date', 'Session'],
                        values='Status',
                        aggfunc='first',
                        fill_value='ABSENT'
                    )
                    grid_df.columns = pd.MultiIndex.from_tuples(sorted(grid_df.columns, key=lambda x: (x[0], x[1])))
                    grid_df = grid_df.sort_index(axis=1)

                    # Delegate Review
                    points_df = first_scans.groupby(['ID', 'Full_Name'])['Points'].sum().reset_index()
                    unique_dates = sorted(first_scans['Date'].unique())
                    total_possible = len(unique_dates) * 2 * 0.5
                    points_df['Total Points Earned'] = points_df['Points']
                    points_df['Total Possible'] = total_possible
                    points_df['Attendance (%)'] = ((points_df['Points'] / total_possible) * 100).round(1)
                    points_df['Attendance (%) String'] = points_df['Attendance (%)'].astype(str) + '%'

                    # --- Analytics DataFrames ---
                    # 1. Trend by date
                    daily_att = first_scans.groupby('Date')['Points'].sum() / (first_scans['ID'].nunique() * 1.0) * 100
                    trend_df = pd.DataFrame({'Date': daily_att.index, 'Attendance Rate (%)': daily_att.values.round(1)})

                    # 2. Risk Analysis
                    def risk_level(att):
                        if att < 50:
                            return "High Risk"
                        elif att < 75:
                            return "Moderate Risk"
                        else:
                            return "Low Risk"
                    points_df['Risk Level'] = points_df['Attendance (%)'].apply(risk_level)
                    risk_df = points_df[['Full_Name', 'Attendance (%)', 'Risk Level']].copy()
                    risk_df['Attendance (%)'] = risk_df['Attendance (%)'].astype(str) + '%'

                    # 3. Distribution bands
                    bins = [0, 50, 75, 90, 101]
                    labels = ['0-50%', '50-75%', '75-90%', '90-100%']
                    points_df['Band'] = pd.cut(points_df['Attendance (%)'], bins=bins, labels=labels, right=False)
                    dist_df = points_df['Band'].value_counts().reset_index()
                    dist_df.columns = ['Attendance Band', 'Number of Students']
                    dist_df = dist_df.sort_values('Attendance Band')

                    # 4. Session Heatmap
                    morning_att = first_scans[first_scans['Session'] == 'Morning'].groupby('Date')['Points'].sum() / (first_scans['ID'].nunique() * 0.5) * 100
                    afternoon_att = first_scans[first_scans['Session'] == 'Afternoon'].groupby('Date')['Points'].sum() / (first_scans['ID'].nunique() * 0.5) * 100
                    heatmap_df = pd.DataFrame({
                        'Date': unique_dates,
                        'Morning Attendance (%)': morning_att.reindex(unique_dates).fillna(0).round(1),
                        'Afternoon Attendance (%)': afternoon_att.reindex(unique_dates).fillna(0).round(1)
                    })

                    # Executive Summary
                    if exclude_unknown:
                        filtered = points_df[points_df['Full_Name'] != "Unknown ID"]
                        total_delegates = len(filtered)
                        avg_att = filtered['Attendance (%)'].mean()
                    else:
                        total_delegates = len(points_df)
                        avg_att = points_df['Attendance (%)'].mean()
                    summary_df = pd.DataFrame({
                        "Metric": ["Total Class Days Logged", "Total Active Delegates", "Average Overall Attendance Rate"],
                        "Value": [len(unique_dates), total_delegates, f"{avg_att:.1f}%"]
                    })

                    # Write all sheets using ExcelWriter
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        summary_df.to_excel(writer, sheet_name='Executive_Summary', index=False)
                        points_df[['ID', 'Full_Name', 'Total Points Earned', 'Total Possible', 'Attendance (%) String']].to_excel(
                            writer, sheet_name='Delegate_Review', index=False
                        )
                        grid_df.to_excel(writer, sheet_name='Detailed_Grid')
                        trend_df.to_excel(writer, sheet_name='Daily_Trend', index=False)
                        risk_df.to_excel(writer, sheet_name='Risk_Analysis_Table', index=False)
                        dist_df.to_excel(writer, sheet_name='Distribution_Bands', index=False)
                        heatmap_df.to_excel(writer, sheet_name='Session_Heatmap_Data', index=False)

                    # Load workbook and add charts & insights
                    wb = load_workbook(output)
                    wb = add_charts_and_insights(wb, summary_df, dist_df)

                    # Save to final buffer
                    final_io = io.BytesIO()
                    wb.save(final_io)
                    final_io.seek(0)

                    st.success("✅ Professional analytics report generated with embedded charts and insights!")
                    st.divider()

                    # Preview in Streamlit
                    st.write("### 📊 Executive Summary Preview")
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total Class Days", len(unique_dates))
                    c2.metric("Active Delegates", total_delegates)
                    c3.metric("Average Attendance", f"{avg_att:.1f}%")

                    with st.expander("🔍 View Key Insights"):
                        st.metric("Overall Attendance Rate", f"{avg_att:.1f}%", delta="Goal: 75%", delta_color="inverse" if avg_att < 75 else "normal")
                        st.write(f"**High Risk Students:** {len(risk_df[risk_df['Risk Level'] == 'High Risk'])}")
                        st.write(f"**Perfect Attendance:** {len(risk_df[risk_df['Risk Level'] == 'Low Risk'])}")
                        st.line_chart(trend_df.set_index('Date'))

                    st.download_button(
                        label="📥 Download Advanced Excel Report",
                        data=final_io,
                        file_name="Weekly_Attendance_Pro_Report.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error: {e}")
                    st.exception(e)
    else:
        st.error("No valid .dat files uploaded.")