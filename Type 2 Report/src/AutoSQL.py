import os, warnings, re, traceback, random, string, json, shutil
import pandas as pd
from pandas import DataFrame
import pdr.handlers.Console_Handler as console
import src.config.API as API
import NewSubsystem_Config as config

# Description: This class is used to automatically update SQL file content and file names.


class AutoSQL:
    def __init__(self, connection_GXXXXXP, connection_DXXXXXP):
        # Initialize the connection to two Oracle database: GXXXXXP and DXXXXXP
        self.connection_GXXXXXP = connection_GXXXXXP
        self.connection_DXXXXXP = connection_DXXXXXP
        # Initialize input and output directory
        self.input_directory = ""
        self.input_directory_base = ""
        self.connection_profile_input_directory = ""
        self.output_directory = ""
        self.output_directory_base = ""
        self.connection_profile_output_directory = ""
        # place_holder is the client name in create_users.sql that needs to be replaced
        self.place_holder = "XYZ"
        self.place_holder_password = "*******"
        # Initialize the table names to be queryed
        self.schema1 = "MXXXXX_PROD"
        self.schema2 = "XXXXX_PRD"
        self.tb_xref_client = f"{self.schema2}.xref_client"
        self.tb_xref_distributor = f"{self.schema2}.xref_distributor"
        self.tb_helpdesk_distributor = f"{self.schema2}.helpdesk_distributor"
        self.tb_info_fact = f"{self.schema2}.info_fact_maint_schema"
        self.tb_project = f"{self.schema1}.project"
        self.tb_view_client_enrollment = f"{self.schema1}.view_client_enrollment"
        self.tb_transfer_info = f"{self.schema1}.transfer_info"
        # Initialize the information of the new client
        self.new_client_short_name = ""
        self.new_client_full_name = ""
        self.new_client_oid = None
        self.project_oid = None
        self.industry_oid = None
        self.file_project_id = None
        self.transfer_info_oid = None
        # Initialize the control variables
        self.database = ""
        self.data_warehouse = 0
        self.server = config.server.upper()
        self.farner = ""
        self.mclane = ""
        self.farner_dist_id = []
        self.mclane_dist_id = []
        self.user_lists = []
        self.col_dist_id = "DIST_ID"
        self.col_swk = "START_PERIOD_CODE"
        self.col_cwk = "END_PERIOD_CODE"
        # Initialize the current week code and end week code for each enrolled distributor ID for the new client
        self.swk_cwk = {}
        # Initialize password variables
        self.password_file = "xyz_password_list_SERVER.json"
        self.password_list = {}
        # Suppress the sqlalchemy database connection warning
        warnings.filterwarnings(
            "ignore", message="pandas only supports SQLAlchemy connectable"
        )


    @staticmethod
    def check_null_empty(value) -> bool:
        """
        Check if a value is None, empty, or contains no elements.

        Args:
            value (str, list, tuple): A value which can be a string, list, or tuple

        Returns:
            bool: True if the value is None, an empty string, or an empty collection, False otherwise
        """
        # Check if the value is None
        if value is None:
            return True
        # Check for string and use strip to remove any whitespace
        if isinstance(value, str):
            return value.strip() == ""
        # Check for list or tuple and if they are empty
        if isinstance(value, (list, tuple)):
            return len(value) == 0
        return False


    @staticmethod
    def check_missing_cols(df: DataFrame, col_names: list[str]):
        """Check if the DataFrame is missing any of the specified columns.

        Args:
            df (DataFrame): a DataFrame
            col_names (list[str]): a list of column names
        """
        missing_columns = [col for col in col_names if col not in df.columns]
        if missing_columns:
            raise ValueError(
                f"Missing columns in the dataframe: {', '.join(missing_columns)}"
            )


    def init_control_variables(self):
        try:
            # Initialize the new client short name, input directory, output directory, and enrollment status
            self.new_client_short_name = config.all_variables[0]
            self.input_directory = os.path.join(config.all_variables[4], "SQL")
            self.input_directory_base = self.input_directory
            self.connection_profile_input_directory = os.path.join(self.input_directory, "CONNECTION_PROFILES")
            self.output_directory = os.path.join(config.all_variables[5], self.new_client_short_name.upper(), self.server, "SQL")
            self.output_directory_base = self.output_directory
            self.connection_profile_output_directory = os.path.join(self.output_directory, "CONNECTION_PROFILES")
            self.data_warehouse = str(config.all_variables[1])
            if self.data_warehouse == "1" or self.data_warehouse == "2":
                self.data_warehouse = int(self.data_warehouse)
            elif self.data_warehouse.lower() == "both":
                self.data_warehouse = [1, 2]
            else:
                raise ValueError("Data warehouse must be 1, 2, or both")
            self.farner = config.all_variables[2]
            self.mclane = config.all_variables[3]
            self.password_file = self.password_file.replace("xyz", self.new_client_short_name.lower()).replace("SERVER", self.server)
            # Create the output directory if it does not exist
            os.makedirs(self.output_directory, exist_ok=True)
            # Clear the output directory
            [os.remove(os.path.join(self.output_directory, f)) for f in os.listdir(self.output_directory) if os.path.isfile(os.path.join(self.output_directory, f))]
        except Exception as e:
            console.log(f"An error occurred in init_control_variables(): {e}")
            raise e
        else:
            console.log(
                f"""Successfully read Control variables:
                New client short name: {self.new_client_short_name}
                Input directory: {self.input_directory}
                Output directory: {self.output_directory}"""
            )


    def _get_xref_client(self):
        try:
            # Initialize the SELECT query to get the new client information
            query = f"""SELECT * FROM {self.tb_xref_client}
            WHERE client_short_name='{self.new_client_short_name.upper()}'
            """
            # Execute the query and return the results as a DataFrame
            return pd.read_sql_query(query, self.connection_GXXXXXP)
        except Exception as e:
            console.log(f"An error occurred in get_xref_client(): {e}")
            raise e


    def _get_project(self):
        try:
            # Initialize the SELECT query to get the project information
            query = f"""SELECT * FROM {self.tb_project}
            WHERE project_short_name='{self.new_client_short_name.upper()}'
            """
            # Execute the query and return the results as a DataFrame
            return pd.read_sql_query(query, self.connection_DXXXXXP)
        except Exception as e:
            console.log(f"An error occurred in get_project(): {e}")
            raise e


    def _get_dist_id(self, distributor_name: str):
        try:
            # Initialize the SELECT query to get the distributor ID
            query = f"""SELECT DISTINCT {self.col_dist_id} FROM {self.tb_xref_distributor}
            WHERE dist_name LIKE '%{distributor_name.upper()}%'
            """
            # Execute the query and return the results as a DataFrame
            return pd.read_sql_query(query, self.connection_GXXXXXP)
        except Exception as e:
            console.log(
                f"An error occurred in get_dist_id() for {distributor_name.upper()}: {e}"
            )
            raise e


    def _get_swk_cwk(self, dist_id: int):
        try:
            # Initialize the SELECT query to get the current week code and end week code for each distributor ID
            query = f"""SELECT DISTINCT {self.col_swk}, {self.col_cwk} from {self.tb_helpdesk_distributor}
            WHERE xXXXXX = {self.new_client_oid} AND dist_id = {dist_id}
            """
            # Execute the query and return the results as a DataFrame
            return pd.read_sql_query(query, self.connection_GXXXXXP)
        except Exception as e:
            console.log(f"An error occurred in get_swk_cwk(): {e}")
            raise e
        

    def _get_transfer_info_oid(self):
        try:
            # Initialize the SELECT query to get the maximum TRANSFER_INFO_OID + 1 from the transfer_info table
            query = f"SELECT MAX(TRANSFER_INFO_OID)+1 FROM {self.tb_transfer_info}"
            # Execute the query and return the results as a single value
            return pd.read_sql_query(query, self.connection_DXXXXXP).iloc[0, 0]
        except Exception as e:
            console.log(f"An error occurred in get_transfer_info_oid(): {e}")
            raise e


    def _init_client_name(self):
        try:
            # Initialize the column names
            cols = ["CLIENT_NAME", "XXXXXX"]
            # Get the client data from the xref_client table
            df_client = self._get_xref_client()
            # Check if all the columns exist in the DataFrame
            AutoSQL.check_missing_cols(df_client, cols)
            # Access the first row of the DataFrame (assume only one row is returned)
            row_client = df_client.iloc[0]
            # Initialize new client full name and OID
            self.new_client_full_name = row_client[cols[0]]
            self.new_client_oid = row_client[cols[1]]
        except Exception as e:
            console.log(f"An error occurred in _init_client_name(): {e}")
            raise e
        else:
            console.log(
                f"""Successfully read '{self.tb_xref_client}' from GXXXXXP:
                New client full name: {self.new_client_full_name}
                New client OID: {self.new_client_oid}"""
            )


    def _init_oid(self):
        try:
            # Initialize the column names
            cols = ["PROJECT_OID", "INDUSTRY_OID", "FILE_PROJECT_ID"]
            # Get the project data from the project table
            df_project = self._get_project()
            # Check if all the columns exist in the DataFrame
            AutoSQL.check_missing_cols(df_project, cols)
            # Acess the first row of the DataFrame (assume only one row is returned)
            row_project = df_project.iloc[0]
            # Initialize project OID and industry OID
            self.project_oid = row_project[cols[0]]
            self.industry_oid = row_project[cols[1]]
            self.file_project_id = row_project[cols[2]]
            self.transfer_info_oid = self._get_transfer_info_oid()
        except Exception as e:
            console.log(f"An error occurred in _init_oid(): {e}")
            raise e
        else:
            console.log(
                f"""Successfully read '{self.tb_project}' from DXXXXXP:
                Project OID: {self.project_oid}
                Industry OID: {self.industry_oid}
                File Project ID: {self.file_project_id}"""
            )


    def _init_dist_id(self):
        try:
            # Get the distributor ID for FXXXXX Co and McLane if they are enrolled for the new client
            if self.farner.upper() == "Y":
                df_farner = self._get_dist_id("FXXXXX CO - CARROLL")
                # Convert the values in dist_id column to a list
                self.farner_dist_id = df_farner[self.col_dist_id].tolist()
            else:
                console.log(
                    "Skipped initializing dist_id for FXXXXX Co as it is not enrolled for the new client."
                )
            if self.mclane.upper() == "Y":
                df_mclane = self._get_dist_id("MCLANE")
                # Convert the values in dist_id column to a list
                self.mclane_dist_id = df_mclane[self.col_dist_id].tolist()
            else:
                console.log(
                    "Skipped initializing dist_id for McLane as it is not enrolled for the new client."
                )
        except Exception as e:
            console.log(f"An error occurred in _init_dist_id(): {e}")
            raise e
        else:
            console.log(
                f"""Successfully read XXXXX_PRD.xref_distributor from GXXXXXP:
                FXXXXX Co: {self.farner_dist_id}
                McLane: {self.mclane_dist_id}"""
            )


    def _init_week_code(self):
        try:
            # Iterate through the distributor ID for FXXXXX Co and McLane if they are not empty
            if not AutoSQL.check_null_empty(self.farner_dist_id):
                for dist_id in self.farner_dist_id:
                    df_swk_cwk = self._get_swk_cwk(dist_id)
                    # Check if all the columns exist in the DataFrame
                    AutoSQL.check_missing_cols(df_swk_cwk, [self.col_swk, self.col_cwk])
                    if not df_swk_cwk.empty:
                        # Access the first row of the DataFrame (assume only one row is returned)
                        row_swk_cwk = df_swk_cwk.iloc[0]
                        # Initialize the current week code and end week code for each distributor ID
                        self.swk_cwk[dist_id] = (
                            row_swk_cwk[self.col_swk],
                            row_swk_cwk[self.col_cwk],
                        )
                    else:
                        console.log(
                            f"No week code found in '{self.tb_helpdesk_distributor}' for distributor '{dist_id}'."
                        )
                        continue

            if not AutoSQL.check_null_empty(self.mclane_dist_id):
                for dist_id in self.mclane_dist_id:
                    df_swk_cwk = self._get_swk_cwk(dist_id)
                    # Check if all the columns exist in the DataFrame
                    AutoSQL.check_missing_cols(df_swk_cwk, [self.col_swk, self.col_cwk])
                    if not df_swk_cwk.empty:
                        # Access the first row of the DataFrame (assume only one row is returned)
                        row_swk_cwk = df_swk_cwk.iloc[0]
                        # Initialize the current week code and end week code for each distributor ID
                        self.swk_cwk[dist_id] = (
                            row_swk_cwk[self.col_swk],
                            row_swk_cwk[self.col_cwk],
                        )
                    else:
                        console.log(
                            f"No week code found in '{self.tb_helpdesk_distributor}' for distributor '{dist_id}'."
                        )
                        continue

        except Exception as e:
            console.log(f"An error occurred in _init_week_code(): {e}")
            raise e


    def init_client_info(self):
        """Initialize all instance variables by prompting the user for input values."""
        try:
            # Initialize the new client full name and oid
            self._init_client_name()
            # Initialize project_oid, industry_oid, and file_project_id
            self._init_oid()
            # Initialize the distributor ID for FXXXXX Co and McLane
            self._init_dist_id()
            # Initialize the start week code and end week code
            self._init_week_code()
        except Exception as e:
            console.log(f"An error occurred in init_client_info(): {e}")
            raise e
        else:
            console.log("All instance variables have been initialized successfully.")


    def update_sql(self, filename: str) -> tuple[str, str]:
        """Reads an SQL file, replaces the client name, and writes it to the output directory.

        Args:
            filename (str): name of SQL file to be read

        Returns:
            tuple[str, str]: updated content and file name
        """
        try:
            # Open and read the content of the SQL file
            with open(os.path.join(self.input_directory, filename), "r") as file:
                content = file.read()
        except IOError as e:
            console.log(f"Error reading file {filename}: {e}")
            # Skip this file and continue with the next one
            return
        # Replace client name and other values in file content
        new_content = content.replace(self.place_holder.upper(), self.new_client_short_name.upper())
        new_content = new_content.replace(self.place_holder.lower(), self.new_client_short_name.lower())
        new_content = new_content.replace("XXXXXX_VALUE", str(self.new_client_oid))
        new_content = new_content.replace("SXXXXXE_VALUE", f"'{self.new_client_short_name.upper()}_XXX_XXX_PRD'")
        new_content = new_content.replace("TRANSFER_INFO_OID_VALUE", str(self.transfer_info_oid))
        new_content = new_content.replace("PXXXXXVALUE", str(self.project_oid))
        new_content = new_content.replace("FXXXXXVALUE", str(self.file_project_id))
        # Replace client name in file name
        new_filename = filename.replace(
            self.place_holder.lower(), self.new_client_short_name
        )
        return (new_content, new_filename)


    def write_file(self, content: str, filename: str):
        """Write the updated content to a new file in the output directory

        Args:
            content (str): file content
            filename (str): file name
        """
        with open(os.path.join(self.output_directory, filename), "w") as file:
            file.write(content)
        console.log(
            f"Updated file written to {os.path.join(self.output_directory, filename)}"
        )
        
        
    def separate_users(self, content: str) -> dict[str, str]:
        """Separate the SQL content into user groups and store them in a dictionary.

        Args:
            content (str): SQL file content
            filename (str): name of the SQL file

        Returns:
            dict[str, str]: a dictionary of user names and their SQL commands
        """
        # Split the content by semicolon to separate user groups.
        # Note that the input file must only have semicolon at the end of last SQL statement of each user group.
        user_groups = content.split(';')
        # Initialize a dictionary to store users and their scripts
        users = {}
        # Regex to extract the username from the CREATE USER statement
        user_pattern = re.compile(r'CREATE USER "(.*?)"')
        # Iterate through each group of users
        for group in user_groups:
            # Ensure the group is not just whitespace
            if group.strip():
                # Find the username
                match = user_pattern.search(group)
                if match:
                    username = match.group(1)
                    # Initialize or append to the list of commands for the user
                    if username not in users:
                        users[username] = []
                    # Split the group into individual statements and clean them by
                    # removing double quotes and replace the password placeholder with a variable
                    statements = [stmt.replace('"', '').replace(self.place_holder_password, f'"{self.place_holder_password}"').strip() for stmt in group.split('\n') if stmt.strip()]
                    statements = [f"{stmt};" for stmt in statements]
                    # Extend the list of commands for the user
                    users[username].extend(statements)
                    # Generate new password and encrypt
                    encrypted_password = API.cipher.encrypt(self.generate_db_password().encode()).decode()
                    # Add the username and password to the password DataFrame
                    self.password_list[username] = {"Password": encrypted_password, "Database": f"{self.database}{self.server[0]}"}
        return users
        

    def create_user_files(self, users: dict[str, str]):
        """Create separate SQL files for each user with their respective commands.

        Args:
            users (dict[str, str]): a dictionary of user names and their SQL commands
        """
        # Create directory if it does not exist
        os.makedirs(self.output_directory, exist_ok=True)
        # Iterate through each user and their commands
        for user, commands in users.items():
            # Define file paths for the user's SQL files
            user_file_path = os.path.join(self.output_directory, f"create_user_{user}.sql")
            grant_file_path = os.path.join(self.output_directory, f"grant_tables_{user}.sql")
            # Open files for writing
            with open(user_file_path, 'w') as user_file, open(grant_file_path, 'w') as grant_file:
                for command in commands:
                    if "CREATE USER" in command:
                        insert = f"insert into dba_util.msa_sec_app_db_schema(db_schema, access_approver_email, db_schema_type, created, creator) values ('{user}', 'NONE','PROD', SYSDATE , 'PBHARGAVA');"
                        user_file.write(f"{command}\n")
                        user_file.write(f"{insert}\n")
                    else:
                        grant_file.write(f"{command}\n")
            console.log(f"'{user}': user and grant scripts have been created successfully.")
        

    def order_users(self, list_users: list) -> list[str]:
        """Reorder the list of users so that first_user is the first to be created.

        Args:
            list_users (list): a list of user names

        Returns:
            list[str]: a reordered list of user names
        """
        # Initialize the first user to be created
        first_user = f"{self.new_client_short_name.upper()}_XXX_XXX_PRD"
        # Reorder the list of users so that the first user is the first to be created
        list_users.remove(first_user)
        list_users.insert(0, first_user)
        # Return the reordered list of users
        return list_users
        

    def create_master_script(self, users: list, master_script: str):
        """Create a master script to run all the create user files."""
        try:
            # Initialize the master script content
            script = ""
            # Iterate through each user and add the file to master script
            for user in users:
                user_file = f"create_user_{user}.sql"
                user_path = os.path.join(self.output_directory, user_file)
                grant_file = f"grant_tables_{user}.sql"
                grant_path = os.path.join(self.output_directory, grant_file)

                self.user_lists.append(f"create_user_{user}.sql")
                self.user_lists.append(f"grant_tables_{user}.sql")

                if os.path.exists(user_path):
                    script += f'@@"{user_file}"\n'
                else:
                    console.log(f"Skipped file '{user_path}' since it does not exist.")
                if os.path.exists(grant_path):
                    script += f'@@"{grant_file}"\n'
                else:
                    console.log(f"Skipped file '{grant_path}' since it does not exist.")
            # Write the master script to a file
            self.write_file(script, master_script)
            # Reset the user lists
            self.user_lists = []
        except Exception as e:
            console.log(f"An error occurred in create_master_script(): {e}")
            raise e
        else:
            console.log(f"'{master_script}' has been created successfully.")


    def update_create_users(self, filename: str, master_script: str):
        """Update the SQL queries in filename so that all XYZ (client name place-holder) are replaced with the actual client name.

        Args:
            filename (str): name of SQL file to be read
            master_script (str): name of the master script file
        """
        try:
            # Open and read the content of the SQL file
            with open(os.path.join(self.input_directory, filename), "r") as file:
                content = file.read().strip()
            # Replace the client name place holder with the actual client name in file content
            new_content = content.replace(
                self.place_holder, self.new_client_short_name.upper()
            )
            # Create a dictionary to store users and their scripts
            users = self.separate_users(new_content)
            # Create separate SQL files for each user with their respective commands
            self.create_user_files(users)
            # Reorder the list of users
            if len(users) > 1:
                list_users = self.order_users(list(users.keys()))
            else:
                list_users = list(users.keys())
            # Create a master script to run all the create user files
            self.create_master_script(list_users, master_script)
        except IOError as e:
            console.log(f"Error reading file {filename}: {e}")
            raise e
        else:
            console.log(
                f"User creation scripts '{filename}' has been updated successfully."
            )
            

    def get_pattern_indexes(self) -> re.Pattern:
        """Get the pattern to filter out the CREATE INDEX statements from the SQL content.

        Returns:
            re.Pattern: a compiled regular expression pattern
        """
        # Keep only CREATE (UNIQUE) INDEX
        return re.compile(
            r'(CREATE\s+(?:UNIQUE\s+)?INDEX.*? ON .*?\(.*?\)\s+(?:NO)?LOGGING\s+TABLESPACE .*?)(?:\s+PCTFREE|;)',
            re.DOTALL
        )
        

    def get_pattern_tables(self) -> re.Pattern:
        """Get the pattern to filter out the CREATE TABLE statements from the SQL content.

        Returns:
            re.Pattern: a compiled regular expression pattern
        """
        # Keep only CREATE TABLE
        return re.compile(
            r'(CREATE\s+TABLE.*?)(?:\s+PCTUSED|;)',
            re.DOTALL
        )
        
            
    def update_sql_objects(self, filename: str, new_filename: str):
        """Update the objects creation scripts for two users (XXX_XXX_PRD, XXX_XXX_TST).

        Args:
            filename (str): name of SQL file to be read

        Raises:
            e: _description_
        """
        try:
            # Open and read the content of the SQL file
            with open(os.path.join(self.input_directory, filename), "r") as file:
                content = file.read().strip()
            # Remove all the 'DROP INDEX' and 'DROP TABLE' statements
            pattern1 = re.compile(r'DROP\s+\S+.*?;', re.IGNORECASE | re.DOTALL)
            content1 = re.sub(pattern1, '', content)
            # Replace old client name with new client name
            new_content = re.sub(self.place_holder, self.new_client_short_name.upper(), content1, flags=re.IGNORECASE)
            # Process index creation scripts specifically
            if 'Indexes' in filename or 'Tables' in filename:
                if 'Indexes' in filename:
                    pattern2 = self.get_pattern_indexes()
                else:
                    pattern2 = self.get_pattern_tables()
                # Find all the content filterned by the pattern
                new_content = pattern2.findall(new_content)
                # Ensure new_content is a list of strings, not tuples
                if new_content and isinstance(new_content[0], tuple):
                    new_content = [''.join(item) for item in new_content]
                # Convert the list of strings to a single string separated by newlines
                new_content_str = ';\n\n'.join(new_content) + ';' if new_content else ''                
                # Write the new content to the file
                self.write_file(new_content_str, new_filename)
            else:
                self.write_file(new_content, new_filename)
        except Exception as e:
            console.log(f"Error reading file {filename}: {e}")
            raise e
        else:
            console.log(f"Objects creation scripts '{filename}' has been updated successfully.")


    def add_master_script(self, list_sql_files: list[str], filename: str):
        """Add the list of sql object creation files to the master_script.

        Args:
            list_sql_files (list[str]): a list of sql object creation files
            filename (str): name of the master script file
        """
        try:
            # Open and read the content of the SQL file
            with open(os.path.join(self.output_directory, filename), "a") as file:
                for sql_file in list_sql_files:
                    file.write(f'@@"{sql_file}"\n')
        except Exception as e:
            console.log(f"Error reading file {filename}: {e}")
            raise e
        else:
            console.log(f"Objects creation scripts have been added to '{filename}' successfully.")
            

    def extract_file_list(self, master_script_path: str) -> list:
        """Extract the list of included files from the master script.

        Args:
            master_script_path (str): path to the master script file

        Returns:
           list: a list of included file paths
        """
        with open(master_script_path, 'r', encoding='utf-8') as file:
            content = file.read()
        # Regular expression to find all included file paths
        pattern = re.compile(r'@@(".*?"|\S+)')
        file_paths = pattern.findall(content)
        # Cleaning file paths by removing any possible quote marks
        file_paths = [path.strip('"') for path in file_paths]
        # Remove create users files from the list
        file_paths = [path for path in file_paths if "create_user" not in path]
        return file_paths


    def sort_files(self, file_list: list) -> tuple:
        """Sort the list of SQL files based on a custom sort key.
        The user with 'PRD' type will have higher priority than 'TST' type.
        The file type 'Tables' will have higher priority than 'Indexes' and then 'Views'.

        Args:
            file_list (list): a list of file names to be sorted

        Returns:
            tuple: a sorted list of file names
        """
        # Define a custom sort key
        def sort_key(filename: str):
            object_order = ["Library", "Types", "Tables", "Sequences", "MV", "Synonyms", "Views", "Triggers", "Indexes", "Packages", "Procedure"]
            # If the filename starts with the new client short name lower, it should be the first file
            if self.new_client_short_name.lower() in filename:
                return (5, 0, filename)
            # If the filename starts with 'create_users', it should be the first file
            if "create_user" in filename:
                return (1, 0, filename)
            # If the filename starts with 'grant_tables', it should be the third file
            elif "grant_tables" in filename:
                return (4, 0, filename)
            # If the filename starts with the new client short name upper and PRD, it should be the second file
            elif self.new_client_short_name.upper() in filename and "PRD" in filename:
                for obj in object_order:
                    if obj in filename:
                        return (2, object_order.index(obj) + 1, filename)
                return (2, 0, filename)
            # If the filename starts with the new client short name upper and TST, it should be the third file
            elif self.new_client_short_name.upper() in filename and "TST" in filename:
                for obj in object_order:
                    if obj in filename:
                        return (3, object_order.index(obj) + 1, filename)
                return (3, 0, filename)
            # Else, it should be the last file
            else:
                return (6, 0, filename)
        return sorted(file_list, key=sort_key)
    

    def adjust_order_master_script(self, file_list: list, master_script: str):
        """Adjust the execution order in the master_script based on the file_list.
        
        Args:
            file_list (list): a list of file names
            filename (str): name of the master script file
        """
        sorted_files = self.sort_files(file_list)
        for f in sorted_files:
            # Add set define off to the beginning of the file
            with open(os.path.join(self.output_directory, f), 'r') as file:
                data = file.read()
            data = "set define off;\n" + data
            with open(os.path.join(self.output_directory, f), 'w') as file:
                file.write(data)
        content = f"""-- {master_script}\n-- This script calls all other SQL scripts\nset define off\n"""
        # Generate script content with ordered @@ include commands
        for filename in sorted_files:
            content += f'@@"{filename}"\n'
        # Write to a master SQL script file
        self.write_file(content, master_script)
            

    def update_sql_files(self):
        try:
            """Update all the SQL files in the input directory and write them to the output directory."""
            # Update SQL for each database in input directory
            databases = [d for d in os.listdir(self.input_directory) if os.path.isdir(os.path.join(self.input_directory, d))]
            for database in databases:
                # Clear the output directory for the current database
                if os.path.exists(os.path.join(self.output_directory_base, database)):
                    shutil.rmtree(os.path.join(self.output_directory_base, database))
                # Skip if not related to data warehouse
                if isinstance(self.data_warehouse, int):
                    if database.startswith("GCYM") and str(self.data_warehouse) not in database:
                        continue
                # Skip the CONNECTION_PROFILES directory
                if database == "CONNECTION_PROFILES":
                    continue
                # Set the input and output directories for the current database
                self.database = database
                self.input_directory = os.path.join(self.input_directory_base, self.database)
                self.output_directory = os.path.join(self.output_directory_base, self.database)
                os.makedirs(self.output_directory, exist_ok=True)
                [os.remove(os.path.join(self.output_directory, f)) for f in os.listdir(self.output_directory) if os.path.isfile(os.path.join(self.output_directory, f))]
                # Set the master script file for the current database
                master_script_file = f"master_script_{self.database}.sql"
                # Initialize an empty list
                list_sql_files = []
                for filename in os.listdir(os.path.join(self.input_directory)):
                    # Choose all the sql files
                    if filename.endswith(".sql"):
                        if filename.startswith(self.place_holder.lower()):
                            # Skip file if not related to the server
                            if "TST" in filename or "PRD" in filename or "STG" in filename:
                                if self.server not in filename:
                                    continue
                            # Get the new SQL content and filename
                            new_content, new_filename = self.update_sql(filename)
                            # Write the updated content to a new file in the output directory
                            self.write_file(new_content, new_filename)
                            # Add the new filename to the list of sql object creation files
                            list_sql_files.append(new_filename)
                        elif filename.startswith("create_users"):
                            # Update the SQL queries that are used to create users
                            self.update_create_users(filename, master_script_file)
                        elif filename.startswith(self.place_holder):
                            # Get the new filename by replacing the place holder with new client short name
                            new_filename = filename.replace(self.place_holder, self.new_client_short_name.upper())
                            # Update the objects creation sql files for two users (XXX_XXX_PRD, XXX_XXX_TST)
                            self.update_sql_objects(filename, new_filename)
                            # Get a complete list of sql object creation files
                            list_sql_files.append(new_filename)
                    else:
                        console.log(f"Skipped '{filename}' when update_sql_files().")
                # Add the list of sql object creation files to the master_script
                self.add_master_script(list_sql_files, master_script_file)
                # Extract the list of included files from the master script
                list_files = self.extract_file_list(os.path.join(self.output_directory, master_script_file))
                # Adjust the execution order in the master script
                self.adjust_order_master_script(list_files, master_script_file)
        except Exception as e:
            console.log(f"An error occurred in update_sql_files(): {e}")
            raise e
        


    def create_insert_subsystem(self, tb_insert: str) -> str:
        """Create the SQL query to insert data into xref_subsystem into the database.
        
        Args:
            tb_insert (str): the table name to insert data into
            
        Returns:
            str: the SQL query to insert data into the table
        """
        query2 = f"""INSERT INTO {tb_insert}
                    (XXXXXX,SXXXXXE)
                    VALUES
                    ({self.new_client_oid},'{self.new_client_short_name.lower()}_XXX_XXX_prd');
                    """
        return query2


    def create_insert_subsystem_dist(self, dist_id: int, swk: int, cwk: int, tb_insert1: str, tb_insert2: str) -> str:
        """Create the SQL query to insert data into subsystem_week and xref_subsystem_distributor into the database.
        
        Args:
            dist_id (int): the distributor ID
            swk (int): the start week code
            cwk (int): the current week code
            tb_insert1 (str): the table name to insert data into subsystem_week
            tb_insert2 (str): the table name to insert data into xref_subsystem_distributor
            
        Returns:
            str: the SQL query to insert data into the tables
        """
        query1 = f"""INSERT INTO {tb_insert1}
                    (XXXXXX,DIST_ID,START_PERIOD,END_PERIOD)
                    VALUES
                    ({self.new_client_oid},{dist_id},{swk},{cwk});
                    """

        query2 = f"""INSERT INTO {tb_insert2}
                    (XXXXXX,DIST_ID)
                    values
                    ({self.new_client_oid},{dist_id});
                    """
        # Return the three INSERT SQL query
        return query1 + "\n" + query2


    def create_insert_fact_maint(self, tb_insert: str) -> str:
        """Create the SQL query to insert data into info_fact_maint_schema into the database.
        
        Args:
            tb_insert (str): the table name to insert data into
            
        Returns:
            str: the SQL query to insert data into the table
        """
        cols = "XXXXXX,WAREHOUSE_NUMBER,SXXXXXE"
        query1 = f"""INSERT INTO {tb_insert}
                    ({cols})
                    values
                    ({self.new_client_oid},0,'{self.new_client_short_name.upper()}_XXX_XXX_PRD');
                    """
        query2 = f"""INSERT INTO {tb_insert}
                    ({cols})
                    values
                    ({self.new_client_oid},4,'@TO_{self.new_client_short_name.upper()}_XXX_XXX_PR2.XXXXX.COM');
                    """
        # Return the two INSERT SQL query
        return query1 + "\n" + query2


    def write_insert_sql(self, filename: str):
        """Write the INSERT SQL queries to a file.

        Args:
            filename (str): name of the file to write the INSERT SQL queries to
        """
        # Create the SQL queries to insert data into the database
        insert1 = self.create_insert_transfer_into("transfer_info")
        insert2 = self.create_insert_subsystem("xref_subsystem")
        # Create the subsystem INSERT queries depending on the distributor enrollment status
        insert3 = ""
        for dist_id, week_code in self.swk_cwk.items():
            # Unpack the week code tuple
            swk, cwk = week_code
            # Create the INSERT queries for each distributor
            insert3 += self.create_insert_subsystem_dist(dist_id, swk, cwk, "subsystem_week", "xref_subsystem_distributor")
        insert4 = self.create_insert_fact_maint(self.tb_info_fact)
        # Concatenate all the INSERT queries together
        content = insert1 + "\n" + insert2 + "\n" + insert3 + "\n" + insert4
        with open(os.path.join(self.output_directory, filename), "w") as file:
            file.write(content)
        console.log(
            f"INSERT statements written to {os.path.join(self.output_directory, filename)}"
        )


    def update_connection_profiles(self):
        """
        Update all connection profile JSON files to match the new client short name.
        """
        try:
            # Copy connection profiles json files to the output directory
            shutil.copytree(self.connection_profile_input_directory, self.connection_profile_output_directory, dirs_exist_ok=True)

            # Get list of connection profile files
            connection_profile_files = [os.path.join(self.connection_profile_output_directory, f) for f in os.listdir(self.connection_profile_output_directory) if f.endswith(".json")]

            for connection_profile_file in connection_profile_files:
                # Skip and delete the connection profile file if it does not match the data warehouse
                if isinstance(self.data_warehouse, int):
                    digit_in_filename = re.findall(r"\d", connection_profile_file)
                    if digit_in_filename:
                        digit = int(digit_in_filename[0])
                        if digit != self.data_warehouse:
                            os.remove(connection_profile_file)
                            continue

                # Rename the connection profile file to match the new client short name
                connection_profile_def_file = connection_profile_file.replace(self.place_holder.upper(), self.new_client_short_name.upper())
                os.rename(connection_profile_file, connection_profile_def_file,)

                with open(connection_profile_def_file, "r") as f:
                    connection_profile_data = f.read()

                # Update connection profile data with new client short name
                connection_profile_data = connection_profile_data.replace(self.place_holder.lower(), self.new_client_short_name.lower())
                connection_profile_data = connection_profile_data.replace(self.place_holder.upper(), self.new_client_short_name.upper())

                # Save the updated connection profile data to the output directory
                with open(connection_profile_def_file, "w") as f:
                    f.write(connection_profile_data)

                console.log(f"Updated connection profile data for {connection_profile_def_file}")
        except Exception as e:
            console.log(f"An error occurred in update connection_profiles(): {e}")
            raise e
        else:
            console.log(f"Successfully updated and saved connection profile JSON files to the output directory: {self.connection_profile_output_directory}\n")


    def export_password_list(self):
        """
        Export the password data to JSON file.
        """
        # Define the password file path
        password_file = os.path.join(self.output_directory_base, self.password_file)
        
        # Write the password list to a JSON file
        with open(password_file, "w") as file:
            json.dump(self.password_list, file, indent=4)


    def get_password_list(self) -> dict:
        """
        Get the password list variable.
        """
        return self.password_list


    def run(self):
        """Run the AutoSQL program to update SQL files and write INSERT SQL queries to a file.

        Raises:
            e: An error occurred while running the program
        """
        try:
            # Initialize the instance variables
            self.init_control_variables()
            # Initialize client information based on table xref_client
            self.init_client_info()
            # Update all SQL files in the input directory and save them to the output directory
            self.update_sql_files()
            # Update the connection profiles
            self.update_connection_profiles()
            # Export the password list to an Excel file
            self.export_password_list()
        except Exception as e:
            console.log(f"An error occurred in run(): {e}\n{traceback.format_exc()}")
            raise e