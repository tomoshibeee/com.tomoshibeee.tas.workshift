# 📑 Feature Specification

This document describes the features of the Workshift Management Spreadsheet System.

---

## 1. Initialization
- **Input**: Month (e.g., September 2025)  
- **Process**: Initialize the system based on the specified month  
  - Generate calendar (dates)  
  - Prepare the monthly sheet  

---

## 2. Generate Worker Timetable Sheets
- **Input**: Worker list (IDs, names)  
- **Process**: Automatically create a timetable sheet for each worker  
- **Output**: Individual sheets per worker  

---

## 3. Edit Worker Timetable Sheets
- **Operation**: Managers or workers manually edit their timetable sheets  
- **Purpose**: Input shift preferences, make corrections, confirm schedules  
- **Output**: Confirmed timetables  

---

## 4. Export Workshift Diagrams

### 4-1. Daily Workshift Diagram (By Date)
- **Input**: Specific date (e.g., 2025/09/08)  
- **Output**: Generate a workshift diagram for that day on the daily sheet  

### 4-2. Monthly Workshift Diagrams (By Month)
- **Input**: Specific month (e.g., September 2025)  
- **Output**: Generate all daily workshift diagrams for the month  

### 4-3. Redraw Daily Workshift Diagram
- **Input**: Current or specified date  
- **Process**: Redraw the daily workshift diagram on the daily sheet using the latest data  
- **Output**: Updated workshift diagram  

---

## Workflow Summary
1. Initialize the system by specifying a month  
2. Generate timetable sheets for all workers  
3. Manually edit and confirm the timetables  
4. Export workshift diagrams (daily or monthly) and redraw as needed  
