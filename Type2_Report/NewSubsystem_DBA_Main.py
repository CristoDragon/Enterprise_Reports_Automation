import time, datetime, sys
import pdr.handlers.Warning_Handler as warn
import pdr.handlers.Console_Handler as console
import pdr.data.Connection as conn
import src.auto.AutoSQL as sql
import src.dply.DeploySQL as deploy
import NewSubsystem_Config as config

# Description: This is the main program to run the New Subsystem job.
# Will generate all CTM and Unix scripts for WLA user to run.


def main():
    current_time = datetime.datetime.now()
    # Logging the start of the program
    console.log("New Subsystem (DBA) Has Started on " + str(current_time))
    # Establish connection to database GCYF1P
    connection_GCYF1P = conn.oracle_connect(
        config.host[0],
        config.port,
        config.instance[0],
        config.username[0],
        config.password[0],
        msg=False,
    )
    # Establish connection to database DSC1P
    connection_DSC1P = conn.oracle_connect(
        config.host[1],
        config.port,
        config.instance[1],
        config.username[1],
        config.password[1],
        msg=False,
    )
    # Initialize an instance of AutoSQL
    sql_job = sql.AutoSQL(connection_GCYF1P, connection_DSC1P)
    # Run the AutoSQL job to create and update the SQL files in the output directory
    sql_job.run()
    # Get the password list
    password_list = sql_job.get_password_list()
    # Initialize an instance of DeploySQL
    deploy_sql = deploy.DeploySQL(password_list)
    # Run the AutoDeploy job to deploy all connection profiles to Control-M
    deploy_sql.run()



if __name__ == "__main__":
    warn.open_warning_handler(config.warnings_file)
    console.set_log(config.console_file)
    # Call the main method
    try:
        start = time.time()
        main()
        end = time.time()
        total = end - start
        console.log(
            "New Subsystem (DBA) Has COMPLETED SUCCESSFULLY ("
            + str("{:.2f}").format(total)
            + "s)"
        )
    except Exception as e:
        console.log(f"New Subsystem (DBA) Has ABORTED With WARNINGS/ERRORS: {e}")
        sys.exit(1)
    finally:
        warn.close()
