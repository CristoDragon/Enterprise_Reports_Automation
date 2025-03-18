import time
import pdr.handlers.Warning_Handler as warn
import pdr.handlers.Console_Handler as console
import pdr.data.Connection as conn
import src.TOB_ITG_VBA_CFG as config
import src.TOB_ITG_Summary_Rpt as itg
import datetime
import sys

# Author: Dragon Xu
# Date: 07/17/2024
# Description: This is the main program to run the Type 1 Report A job.

def main():
    current_time = datetime.datetime.now()
    # Logging the start of the program
    console.log("Type 1 Report A Job Has Started on " + str(current_time))
    # Establish connection to Oracle database
    connection = conn.oracle_connect(
        config.host,
        config.port,
        config.instance,
        config.username,
        config.password,
        msg=False,
    )
    # Initialize an instance of the class TOB_ITG_Summary_Rpt
    job = itg.TOB_ITG_Summary_Rpt(connection, config.table, config.report_id, config.template_file, config.output_file)
    # Run the Type 1 Report A job
    job.run()


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
            "Type 1 Report A Job Has COMPLETED SUCCESSFULLY ("
            + str("{:.2f}").format(total)
            + "s)"
        )
    except Exception as e:
        console.log(
            f"Type 1 Report A Job Has ABORTED With WARNINGS/ERRORS: {e}"
        )
        sys.exit(1)
    finally:
        warn.close()
