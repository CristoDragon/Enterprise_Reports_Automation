# Type 3 Report

# Note
Confidential information has been replaced with ***** or XXXXX.

# Introduction
1. Convert SAS Type 3 Report job into python
2. Executing file: Type3_Report_Main.py
3. Table: 
    - PM_XXXXX.DLVR_BRAND@XXXXXP
    - PM_XXXXX_PREV.DLVR_BRAND@XXXXXP

4. Directory: \Corp\XXXXX\Weekly\Type3_Report_RPT
5. Template file: RXXXXX_temp.xlsx
6. Report file: RXXXXX - Type3 Rpt xxxx.xlsx
7. Control-M:
    - Job Type: AI RUNPYSCRPT
    - Job Name: T_XXXXX_PY
    - Parent Folder: T_XXXXX_WEEKLY/T_XXXXX_WEEKLY_PY


# General Flow
- Create connection to Oracle database
- Create an instance of class Type3_Report (has all the methods needed)
- Copy the template excel workbook
- For each category
    - For each attribute
        - Pull data from the database
        - Create a field sheet
        - Add summary data to main home sheet
    - Create a drill down sheet
    - Create a category home sheet
- Process details such as styles and formats
- Save and close the workbook