# 🏫 Weekly Attendance Automator

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/)
[![Streamlit App](https://static.streamlit.io/badges/streamlit_badge_black_white.svg)](https://your-app-url.streamlit.app)

A streamlined, Python-powered web application built with **Streamlit** to automate the processing, grading, and reporting of biometric attendance logs. 

This tool eliminates the need for complex, brittle Excel formulas by handling all time-compression, duplicate filtering, and attendance grading internally, outputting a clean, professional, multi-sheet Excel report.

## ✨ Features

* **Intelligent Time Scaling:** Dynamically compresses raw biometric log times (e.g., a 3‑hour morning block or 90‑minute afternoon block) into standardised 30‑minute target windows.
* **Automated Grading Engine:** Automatically grades check‑in times as `PRESENT` or `LATE` based on configurable time thresholds, and assigns corresponding performance points.
* **Duplicate Protection:** Automatically detects and removes duplicate biometric scans to prevent grading errors or double penalties.
* **Interactive In‑App Analytics:**
  * **Delegate Review Table** – colour‑coded risk levels (High, Moderate, Low) with sortable columns.
  * **Student Detail View** – select any student to see their daily attendance trend chart.
* **Professional Excel Export:** Generates a ready‑to‑present `.xlsx` file containing multiple sheets:
  * **Executive Summary** – high‑level metrics and overall attendance rate, print‑ready with a generation timestamp.
  * **Delegate Review** – individual student performance with weekly breakdowns, risk flags, and a dropdown‑validated “Remarks” column.
  * **Detailed Grid** – complete calendar‑style matrix of every student’s daily Morning and Afternoon status.
  * **Risk Analysis** – risk‑level distribution with a pie chart.
  * **Daily Trend** – line chart of daily attendance rate over time.
  * **Session Heatmap** – stacked bar chart comparing Morning vs Afternoon performance.
  * **Distribution Bands** – bar chart showing how many delegates fall into each attendance band.
  * **Unknown IDs** – automatically flags biometric IDs not found in the names file.
  * **Automated Insights** – plain‑language recommendations based on risk levels and overall attendance.
* **User‑Friendly Configuration:** All time rules (Morning/Afternoon cut‑offs, compression windows) are adjustable directly from the sidebar.
* **Robust Validation:** Malformed lines are reported and skipped; duplicate Attendance_IDs are warned about.
* **Feedback Integration:** A “Send Feedback via Email” button opens your email client pre‑addressed to the developer.

## 🚀 Live Demo
You can access the live web application here: **[Link to your Streamlit App](https://attendanceautomator-gusfgsxavryibv3kw5ryyc.streamlit.app/)**

## 🛠️ How to Use the App

Using the application is simple and requires no coding knowledge:

1. **Configure Time Rules (Optional):**  
   Use the sidebar to adjust the “Present” cut‑off times for Morning and Afternoon sessions, enable/disable time compression, and set the target window.

2. **Upload Biometric Logs:**  
   Drag and drop one or multiple `.dat` files into the **“Biometric Logs”** area. A progress bar will show parsing status and any malformed lines are reported.

3. **Upload Names Dictionary:**  
   Drag and drop your `Names.csv` or `.xlsx` file (must contain `Attendance_ID` and `Full_Name` columns) into the **“Student Names”** area.

4. **Generate Report:**  
   Click the **“🚀 Generate Pro Analytics Report”** button. The app processes all scans, grades attendance, and builds the Excel file.

5. **Explore Results Interactively:**
   * **Executive Summary** – key metrics displayed as cards.
   * **Delegate Review** – interactive table with colour‑coded risk levels.
   * **Student Detail View** – dropdown to pick a student and see their daily attendance trend.
   * Expandable charts for daily trend, session comparison, and distribution bands.

6. **Download & Share:**  
   * **Excel Report** – full analytics workbook ready for management.
   * **Cleaned CSV** – daily summary CSV for further analysis.
   * **Feedback** – click the email button to send suggestions directly to the developer.

## 💻 Local Installation (For Developers)

If you wish to run this application locally on your machine, follow these steps:

1. **Clone the repository:**
   ```bash
   git clone [https://github.com/yourusername/Attendance_Automator.git](https://github.com/yourusername/Attendance_Automator.git)
   cd Attendance_Automator
   
## 🧰 Tech Stack

![Python](https://img.shields.io/badge/python-3670A0?style=for-the-badge&logo=python&logoColor=ffdd54)
![Streamlit](https://img.shields.io/badge/Streamlit-FF4B4B?style=for-the-badge&logo=Streamlit&logoColor=white)
![Pandas](https://img.shields.io/badge/pandas-%23150458.svg?style=for-the-badge&logo=pandas&logoColor=white)

## 🙏 Acknowledgments

This project was developed by the R&D Team at the **Centre for Marine And Offshore Technology Development (CMOTD)**. I acknowledge their support and resources in making this research and automation tool possible.

<br>
<img src="CMOTD_LOGO.jpg" alt="CMOTD LOGO" width="400">
