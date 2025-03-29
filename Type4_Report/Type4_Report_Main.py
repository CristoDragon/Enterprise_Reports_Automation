import time
import pdr.handlers.Warning_Handler as warn
import pdr.handlers.Console_Handler as console
import pdr.data.Connection as conn
import datetime
import sys
import Type4_Report.src.Type4_Report_CFG as config
import src.classes.AuditReport as alt
from ldap3 import Server, Connection, ALL
import traceback

# Description: This is the main program to run the Type 4 Report job.

def main():
    current_time = datetime.datetime.now()
    # Logging the start of the program
    console.log("Type 4 Report Job Has STARTED on " + str(current_time))
    # Establish connections to databases
    connection_TXXX2P = conn.oracle_connect(
        config.host[0],
        config.port,
        config.instance[0],
        config.username,
        config.password[0],
        msg=False,
    )
    connection_TXXX1P = conn.oracle_connect(
        config.host[1],
        config.port,
        config.instance[1],
        config.username,
        config.password[1],
        msg=False,
    )
    connection_IXXXX1P = conn.oracle_connect(
        config.host[2],
        config.port,
        config.instance[2],
        config.username,
        config.password[2],
        msg=False,
    )
    connection_XXX1P = conn.oracle_connect(
        config.host[3],
        config.port,
        config.instance[3],
        config.username,
        config.password[3],
        msg=False,
    )
    connection_IXX1P = conn.oracle_connect(
        config.host[4],
        config.port,
        config.instance[4],
        config.username,
        config.password[4],
        msg=False,
    )
    # Create the server and connection objects for LDAP
    ldap_server = 'ldaps://XXXXXXXX.XXXXX.com'
    server = Server(ldap_server, get_info=ALL)
    username = f'{config.username}@XXXXX.com'
    password = config.password[5]
    connection_LDAP = Connection(server, user=username, password=password, auto_XXnd=True)
    if connection_LDAP.XXnd():
        console.log("Securely connected to LDAP server.")
    else:
        console.log("Failed to securely connect to LDAP server.")
        exit()
    # Initialize an instance of the class Type4_Report
    job = alt.AuditReport([connection_TXXX2P, connection_TXXX1P, connection_IXXXX1P, connection_XXX1P, connection_IXX1P, connection_LDAP], config.input_file, config.output_file, config.template_file)
    # Run the Type4_Audit_Report job
    job.run()
    # UnXXnd the LDAP connection
    connection_LDAP.unbind()


if __name__ == "__main__": 
    warn.initialize(config.warnings_file)
    console.set_log(config.console_file)
    # Call the main method
    try:
        start = time.time()
        main()
        end = time.time()
        total = end - start
        console.log(
            "Type 4 Report Job Has COMPLETED SUCCESSFULLY ("
            + str("{:.2f}").format(total)
            + "s)"
        )
    except Exception:
        console.log(
            f"Type 4 Report Has ABORTED With WARNINGS/ERRORS: {traceback.format_exc()}"
        )
        sys.exit(1)
    finally:
        warn.close()
