# Type2 Report

# Note
All the actual database, columns, tables, and client names have been replaced by XXXXX.

# Introduction 
1. Automate the process of creating a new subsystem for a client, including
    - Create unix directories and subdirectories
    - Replace the place holder with new client name in input SQL files
    - Create database users and grant privileges
    - Create SQL objects (tables, indexes, views, etc.)
    - Create INSERT statements
    - Create/update Control-M objects (folders, sub-folders, jobs, connection profiles, variables, etc.)
    - Deploy updated object configurations to Control-M

2. Executing file: NewSubsystem_Main.py
    - Class 1: `AutoSQL.py`
    - Class 2: `AutoUnix.py`
    - Class 3: `AutoCTM.py`

3. Executing file: NewSubsystem_Deploy.py
    - Class 1: `DeployCTM.py`

4. Tables:
    - XXXXX_PROD.project@DXXXXXP
    - XXXXX_PROD.XXXXX_enrollment@DXXXXXP
    - XXXXX_PRD.XXXXX_client@GXXXXXP
    - XXXXX_PRD.XXXXX_distributor@GXXXXXP
    - XXXXX_PRD.XXXXX_distributor@GXXXXXP
    - XXXXX_PRD.XXXXX_schema@GXXXXXP

5. Directory:
    - SQL input directory (holds all the input files)
        - DXXXXX
            - xyz_insert_PRD.sql
            - xyz_insert_TST.sql
        - GXXXXX
            - create_users.sql (database user creation file)
            - xyz_dsc_mpl.sql
            - xyz_insert.sql
            - xyz_lock_all.sql
            - xyz_lock_brand.sql
            - xyz_lock_hdesk.sql
            - xyz_lock_user.sql
            - xyz_unlock_all.sql
            - xyz_unlock_brand.sql
            - xyz_unlock_hdesk.sql
            - xyz_unlock_user.sql
            - XYZ-SUB-ODS-PRD-Indexes.sql
            - XYZ-SUB-ODS-PRD-Packages.sql
            - XYZ-SUB-ODS-PRD-Synonmyms.sql
            - XYZ-SUB-ODS-PRD-Tables.sql
            - XYZ-SUB-ODS-PRD-Triggers.sql
            - XYZ-SUB-ODS-PRD-Views.sql
            - XYZ-SUB-ODS-TST-Indexes.sql
            - XYZ-SUB-ODS-TST-Library.sql
            - XYZ-SUB-ODS-TST-MV.sql
            - XYZ-SUB-ODS-TST-Packages.sql
            - XYZ-SUB-ODS-TST-Procedure.sql
            - XYZ-SUB-ODS-TST-Sequences.sql
            - XYZ-SUB-ODS-TST-Synonyms.sql
            - XYZ-SUB-ODS-TST-Tables.sql
            - XYZ-SUB-ODS-TST-Types.sql
            - XYZ-SUB-ODS-TST-Views.sql
        - GCYM
            - create_users_1.sql
            - create_users_2.sql
    - SQL output directory (holds all the output files)
        - DXXXXX
            - xyz_insert_PRD.sql
            - xyz_insert_TST.sql
            - master_script_DXXXXX.sql (script running all SQL files)
        - GXXXXX
            - create_user_{database user}.sql (CREATE USER statements separated for each user)
            - grant_tables_{database user}.txt (table names that need to be granted privileges to)
            - {new_client_short_name}_dsc_mpl.sql
            - {new_client_short_name}_insert.sql
            - {new_client_short_name}_lock_all.sql
            - {new_client_short_name}_lock_brand.sql
            - {new_client_short_name}_lock_hdesk.sql
            - {new_client_short_name}_lock_user.sql
            - {new_client_short_name}_unlock_all.sql
            - {new_client_short_name}_unlock_brand.sql
            - {new_client_short_name}_unlock_hdesk.sql
            - {new_client_short_name}_unlock_user.sql
            - {new_client_short_name}-SUB-ODS-PRD-Indexes.sql
            - {new_client_short_name}-SUB-ODS-PRD-Packages.sql
            - {new_client_short_name}-SUB-ODS-PRD-Synonmyms.sql
            - {new_client_short_name}-SUB-ODS-PRD-Tables.sql
            - {new_client_short_name}-SUB-ODS-PRD-Triggers.sql
            - {new_client_short_name}-SUB-ODS-PRD-Views.sql
            - {new_client_short_name}-SUB-ODS-TST-Indexes.sql
            - {new_client_short_name}-SUB-ODS-TST-Library.sql
            - {new_client_short_name}-SUB-ODS-TST-MV.sql
            - {new_client_short_name}-SUB-ODS-TST-Packages.sql
            - {new_client_short_name}-SUB-ODS-TST-Procedure.sql
            - {new_client_short_name}-SUB-ODS-TST-Sequences.sql
            - {new_client_short_name}-SUB-ODS-TST-Synonyms.sql
            - {new_client_short_name}-SUB-ODS-TST-Tables.sql
            - {new_client_short_name}-SUB-ODS-TST-Types.sql
            - {new_client_short_name}-SUB-ODS-TST-Views.sql
            - master_script_GXXXXX.sql (script running all SQL files)
        - GCYM
            - create_user_{database user}.sql (CREATE USER statements separated for each user)
            - grant_tables_{database user}.txt (table names that need to be granted privileges to)
            - master_script_GCYM.sql (script running all SQL files)
    - Unix output directory
        - create_directory_{new_client_short_name}.txt (unix commands that create all the directories)
    
6. Control variables (stored in excel file for user to update)
    - New client short name
    - Data warehouse (1, 2, or BOTH)
    - FXXXXX (Y/N)
    - MXXXXX (Y/N)
    - Input directory
    - Output directory

7. Instruction for running the program
    - Clone the folder 'PYTHON SUBSYSTEM AUTOMATION' from repo
    - Insert control variables into command line
    - Change the path in NewSubsystem_Config.py
    - Copy and paste the production param into terminal and run
    - Check the output directory to see all the output files

8. Notes
    - Place holder is 'XYZ'
    - In create_users.sql
        - ';' should only be used to indicate the last SQL statement of each user group so that the program can separate users correctly
        - Should have double quotes around values
        - Assume granting the same privileges (INSERT, UPDATE, DELETE, SELECT) to all table names
    - In control variables
        - Should use the second row for new value
    - All input files should follow the same naming convention listed here

9. Deployment Information
    - Deployment location: `F:\XXXXX\XXXXX`
    - Control-M job location: `XXXXX_TST\XXXXX`



# General flow
- Create connections to Oracle database GXXXXXP and DXXXXXP
- Pass in all the command line parameters to initialize an instance of the class AutoSQL
- Execute the run() method to run all the tasks
    - Initialize instance variables based on the control file
        - Get new client short name
        - Get input directory
        - Get output directory
        - Get Y/N for fXXXXX (if it is enrolled for the new client)
        - Get Y/N for mXXXXX (if it is enrolled for the new client)
        - Get the file name of database user creation file
    - Initialize other instance variables
        - Query XXXXX_PRD.XXXXX_client to get
            - new client full name
            - new client oid
        - Query XXXXX_PROD.project to get
            - project oid
            - industry oid
            - file project oid
        - Query XXXXX_PRD.XXXXX_distributor to get
            - a list of dist_id for fXXXXX or/and mXXXXX
        - Query XXXXX_PRD.XXXXX_distributor to get
            - the current week code and end week for each dist_id
    - Update all the SQL files in the input directory
        - If the file name starts with the lowercase place holder('XYZ')
            - Replace the place holder with new client name in both file content and file name
            - Output them in the output directory
        - If the file name is database user creation file
            - Separate the input file into two files per user
                - grant_tables_{user}.txt
                - create_user_{user}.sql
            - Create a master script and add the file paths above into it
        - If the file name starts with the uppercase place holder('XYZ')
            - Remove all the DROP statements
            - Replace the place holder with new client short name
            - If the file is used to create 'Tables' or 'Indexes'
                - Use specific re pattern to filter
            - Output the updated file to the output directory
        - Adjust the execution order in the master script
    - Write the INSERT SQL queries to a file with name '{new_client_short_name}_update.sql'
- Pass in all the command line parameters to initialize an instance of the class AutoUnix
- Execute the run() method to run all the tasks
    - Initialize instance variables based on the control file
        - Get new client short name
        - Get output directory
    - Generate all the mkdir commands
    - Save those commands to 'create_directory_{new_client_short_name}.txt' in the specified output directory

