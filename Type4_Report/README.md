# Type 4 Report
This is an effort to automate the Type 4 Report generation using Python.

# Introduction
1. Automate the process of creating 11 type4 reports for Client_X (listed below)

2. Executing file
    - Type4_Report_Main.py

3. Tables
    - TXXX2P
        - DBA_USERS
        - SYS.USER$
        - DBA_ROLE_PRIVS
        - DBA_UTIL.DB_USERS
        - DBA_UTIL.MSA_SEC_APP_LOGON_TRACKER
        - Client_X_DDL_type4_V

    - TXXX1P
        - DBA_USERS
        - SYS.USER
        - DBA_ROLE_PRIVS
        - DBA_UTIL.MSA_SEC_APP_OS_USER
        - DBA_UTIL.MSA_SEC_APP_LOGON_TRACKER
        - Client_X_DDL_type4_V
        - DBA_UTIL.MSA_SEC_APP_LOGON_TRACKER
        - DBA_UTIL.MSA_SEC_APP_OS_USER
        - DBA_ROLE_PRIVS

    - IXXX1P
        - intadmin.admin_user_client
        - intadmin.admin_user
        - intadmin.admin_class
        - intadmin.admin_client
        - DDL_TRACKER

    - IXXXP
        - Client_X_ff_prd.PERSON

    - MySQL
        - wf82prd_ra.SMQUERY_8207
        - wf82prd_ra.SMFROMS_8207
        - wf82prd_ra.SMSESSIONS_8207

4. Directory & Report Files ("QX" represents which quarter, YY indicates which year):
    - \\Corp\XXXXX\Client_X_type4_20{YY}_Q{X}\Preview Reports for Review 
        (1) Client_X_type4_QXYY_Data_Correct_Active_Users_Rpt.xlsx
        (2) Client_X_type4_QXYY_TXXX2P_Trigger_File_Changes.xlsx
        (3) Client_X_type4_QXYY_XXXX_TXXX1P_Privileges.xlsx
        (4) Client_X_type4_QXYY_XXXX_TXXX1P_Structure.xlsx
        (5) Client_X_type4_QXYY_XXXX_TXXX1P_Users.xlsx
        (6) Client_X_type4_QXYY_XXXX_TXXX2P_Privileges.xlsx
        (7) Client_X_type4_QXYY_XXXX_TXXX2P_Structure.xlsx
        (8) Client_X_type4_QXYY_XXXX_TXXX2P_Users.xlsx
        (9) Client_X_type4_QXYY_DS_Access_Rpt.xlsx (external: control-M)
        (10) Client_X_type4_QXYY_Flag_Portal_Access_Report.xlsx.xlsx (external: Andy King)
        (11) Client_X_type4_QXYY_Program_Releases.xls (external: Wendy Boustead)

    - \\Corp\XXXXX\Client_X_type4_2024_Q2\FINAL_PUBLISH REPORTS - QXYY
        - The folder will be created manually to store all the final publish reports (reports generated above after manual inspection) as well as their pdf conversion files

5. Template file
    - \\Corp\XXXXX\Templates
        (1) Client_X_type4_Data_Correct_Active_Users_Rpt_Template.xlsx
        (2) Client_X_type4_TXXX2P_Trigger_File_Changes_Template.xlsx
        (3) Client_X_type4_XXXX_TXXX1P_Privileges_Template.xlsx
        (4) Client_X_type4_XXXX_TXXX1P_Structure_Template.xlsx
        (5) Client_X_type4_XXXX_TXXX1P_Users_Template.xlsx
        (6) Client_X_type4_XXXX_TXXX2P_Privileges_Template.xlsx
        (7) Client_X_type4_XXXX_TXXX2P_Structure_Template.xlsx
        (8) Client_X_type4_XXXX_TXXX2P_Users_Template.xlsx
        (9) Client_X_type4_DS_Access_Rpt_Template.xlsx
        (10) Client_X_type4_Flag Portal Access_Report_Template.xlsx
        (11) Client_X_type4_Program_Releases_Template.xls

6. Input File
    - \\Corp\XXXXX\Client_X_type4_2024_Q2\Report Queries
        (1) Client_X_Data Correct_Users_Query_IXXX1P.sql
        (2) Client_X_Measure_Def_Query_(Todds_Rewrite)_IXXX1P.sql
        (3) Client_X_TXXX1P_Privilege_Query.sql
        (4) Client_X_TXXX1P_Structure_Query.sql
        (5) Client_X_TXXX1P_User_Query_REWRITE_10_14_21_Jim_H.sql
        (6) Client_X_TXXX2P_Privilege_Query.sql
        (7) Client_X_TXXX2P_Structure_Query.sql
        (8) Client_X_TXXX2P_User_Query_REWRITE_07_17_19.sql

7. Control-M
    - Job Type:
    - Job Name:
    - Parent Folder:

8. Notes
    - Quarter definition
        - April 1st – data from Jan through March
        - July 1st – data from April through June
        - October 1st – data from July through September
        - January 1st – data from October through December
    - This program highlights terminal warning messages in yellow color
    - This program automatically adjust the column width of output excel files
    - Some of the input queries have been adjusted to make sure pd.read_sql() can execute them successfully
        - Removed the variables initialization at the front of the query
        - Removed some SQL Server specific clauses
        - Replaced the original start date and end date with "2024-01-01" and "2024-03-31" respectively in all input sql files, which serve as place holder to be dynamically replaced by the current quarter start date and end date
    - The program will not check rules if the dataframe in an type4 report is empty
    - This program will raise warnings if the execution of input query returns an empty dataframe (the programm will keep running), and "** No Records **" will be shown in the corresponding excel report
    - LDAP package can fill up most but not all the missing employee full names, which means there still need to be some manual work to fill the rest of missing names.

9. Preconditions & Post-conditions
    - The input file path must exist, otherwise warning will be raised in the terminal
    - The column names are hard-coded when implementing rules checking procedure for each type4 report, which requires the column name from template excel file and from the database remain the same
    - The template excel file must contain one and only one sheet

# General Workflow
- Create the first 8 reports
    - Load the template xlsx file from the template folder
    - Check if the number of sheets is one
    - Update the corresponding SQL query stored in a .sql file in the report queries folder
    - Execute the updated query to get the dataframe as result
    - Implement the checking rules for the report to filter the dataframe
    - Put the filtered dataframe to excel workbook
    - Set the style for the excel workbook
    - Close and save the changes in excel
- Create Client_X_type4_QXYY_DS_Access_Rpt.xlsx
- Create Client_X_type4_QXYY_Program_Releases.xls (if the input excel file exported from Power BI report exists)
- Create Client_X_type4_QXYY_Flag_Portal_Access_Report.xlsx.xlsx

