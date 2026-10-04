# Exam Hall Allocation & Attendance Sheet Generator
### Developed for Kaduna Polytechnic Examination Office (and all Academic Departments)

This application automates daily student hall allocations and attendance sheet generation for examinations. It reshuffles students daily using custom seeds, distributes them across designated exam halls according to exact venue capacities, and generates official attendance documents matching Kaduna Polytechnic's official template.

---

## 🚀 How to Run the Application

### Option A: 1-Click Launch (Recommended for Windows)
Simply double-click the **`run_app.bat`** file in this folder. It will start the server and open the interface directly in your default web browser.

### Option B: From Command Prompt or PowerShell
```powershell
streamlit run app.py
```
Then navigate to `http://localhost:8501` in your browser.

---

## 🌟 Key Features

1. **Customizable for Any Department & College**:
   - Default presets configured for Kaduna Polytechnic, College of Engineering, Computer Engineering, etc.
   - Any department can choose their name from the dropdown or type in custom details.
   - Configure Course Title, Course Code, Level, Session/Semester, Exam Date, Time In/Out, and Invigilators.

2. **Master Student List Import**:
   - **Upload Excel (.xlsx, .xls)** or **CSV**: Auto-detects `REG NO` and `NAME` columns, with manual column selector if custom headers are used.
   - **Download Excel Template**: One-click download button for a sample student template.
   - **Copy-Paste Input**: Directly paste text from portals or text files.
   - **Existing Documents**: Load students from prior `.docx` attendance sheets in the directory.

3. **Daily Reshuffling with Custom Seed**:
   - Enter any **Custom Seed** (e.g. `DAY1-CTE442`, `TUE29_MATH`, etc.) or click **Random Seed**.
   - Deterministic and 100% reproducible: entering the same seed reproduces the exact same seating order, while changing the seed gives a fresh, unbiased daily reshuffle.
   - Option to disable reshuffling and preserve the original order.

4. **Dynamic Venue Allocation**:
   - Add as many exam halls as needed and set the exact student count for each hall.
   - Quick-add from standard Kaduna Polytechnic venues (e.g. *Mining Theater-MPRE, Civil Room 101/102, Civil Transport Lab, Chem Eng Theater, Hall A/B/C-MPRE, Electronics Lab*).
   - Real-time balance counter shows whether allocated capacity matches the student total.
   - **Auto-Distribute** button to instantly balance capacities across all selected halls.

5. **Official Template & Door Notice Exports**:
   - **Attendance Register (.docx & .pdf)**: Contains all allocated halls separated by page breaks, formatted with the exact institutional header, invigilator lines, and student table (`S/N`, `REG NO`, `NAMES`, `BOOKLET NO`, `SIGN IN`, `SIGN OUT`).
   - **Venue Entrance Door Notices (.docx & .pdf)**: Generates high-visibility entrance posters for each hall displaying the Hall Name in a prominent header banner, Course Metadata, Candidate Count, Candidate Instructions, and a 2-column candidate roster (`S/N`, `REG NO`, `NAME`) so students can check their venue before entering.

---

## 📁 File Structure

- [app.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/app.py): The main Streamlit web application.
- [run_app.bat](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/run_app.bat): Windows batch file for 1-click startup.
- [core/models.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/core/models.py): Data models for students, venues, and exam configuration.
- [core/parser.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/core/parser.py): Importers for Excel, CSV, Word, and text paste.
- [core/shuffler.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/core/shuffler.py): Reshuffling engine with custom seed support.
- [core/docx_generator.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/core/docx_generator.py): Official Word attendance sheet generator.
- [core/pdf_generator.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/core/pdf_generator.py): Print-ready PDF generator.
- [config/presets.json](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/config/presets.json): Institutional presets for venues, departments, and sessions.
- [sample_students.xlsx](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/sample_students.xlsx): Ready-to-test sample student list.
- [tests/test_allocation.py](file:///c:/Users/user/Desktop/ATTENDANCE_CHECKER/tests/test_allocation.py): Automated test suite.
