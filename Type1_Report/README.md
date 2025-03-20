# Type 1 Report

# Note
Confidential information has been replaced with ***** or XXXXX.

# Introduction
1. Automate the process of creating ABC summary reports (listed below)
2. Executing file: TOB_Type1_Report_A_Main.py
3. Table: 
    - XXXXX_PRD.PREP_XXXXX@*****
4. Directory & Report Files (13 reports in total):
    - \\Corp\XXXXX\Cigarettes\XXXXX\Python_Test
        - ABC_XXXXX_{period_code}.xlsx
        - ABC_XXXXX_{period_code}.xlsx
        - G360_XXXXX_{period_code}.xlsx
        - XXXXX_Summary_{period_code}.xlsx
    - \\Corp\XXXXX\Blu\XXXXX\Python_Test
        - ABC_XXXXX_{period_code}.xlsx
        - ABC_XXXXX_{period_code}.xlsx
        - G360_XXXXX_{period_code}.xlsx
    - \\Corp\XXXXX\Cigar\XXXXX\Python_Test
        - ABC_XXXXX_{period_code}.xlsx
        - ABC_XXXXX_{period_code}.xlsx
        - G360_XXXXX_{period_code}.xlsx
    - \\Corp\XXXXX\Otp\XXXXX\Python_Test
        - ABC_XXXXX_{period_code}.xlsx
        - ABC_XXXXX_{period_code}.xlsx
        - G360_XXXXX_{period_code}.xlsx
5. Template file
    - \\Corp\XXXXX\Python_Template
        - ABC_XXXXX_Template.xlsx
        - ABC_XXXXX_Template.xlsx
        - XXXXX_Template.xlsx
        - XXXXX_Template.xlsx
        - XXXXX_Template.xlsx
        - XXXXX_Template.xlsx
        - XXXXX_Summary_Template.xlsx
7. Control-M
    - Job Type: XXXXX
    - Job Name: XXXXX
    - Parent Folder: XXXXX\XXXXX


# General Flow
- Create a connection to Oracle database *****
- Pass in all the command line parameters to initialize an instance of the class TOB_ABC_Summary_Rpt
- Execute the run() method to run all the tasks
    - Create 3 reports for each category (Cig, Ecig, Cigar, OTP)
        - Create G360_XXXXX report -- 1 report
            - Get the data by executing dynamic query
            - Put the data into excel workbook
            - Format the workbook
        - Create WDC_XXXXX report -- 1 report
            - Get the data by executing dynamic query
            - Put the data into excel workbook
            - Format the workbook
        - Create G360 XXXXX_Report -- 1 report
            - Create curr_db sheet
            - Create prev_db sheet
                - Get the data from previous week's G360_XXXXX report
                - Rest of the process is the same as creating curr_db
            - Create Final_Comparison sheet
    - Create the Volume_Change_Summary report -- 1 report
        - Create a wdc sheet for each category (4 in total)
            - Get data from current week's WDC_XXXXX report
            - Put the data into excel worksheet
            - Format the worksheet
        - Create a final_comparison sheet for each category (4 in total)
            - Get data from current week's G360_XXXXX report
            - Construct some new columns
            - Put the data into excel worksheet
            - Format the worksheet
        - Create a summary sheet
            - Construct new columns one by one
            - Put the dataframe into excel worksheet
            - Format the worksheet
- Close and save all the workbooks