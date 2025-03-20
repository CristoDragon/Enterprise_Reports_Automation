import time
import pdr.handlers.Warning_Handler as warn
import pdr.handlers.Console_Handler as console
import pdr.data.Connection as conn
import proj.TOB_ALT_SAS_CFG as config
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl import Workbook
import Type3_Report.src.Type3_Report as Type3_Rpt
from openpyxl import load_workbook

# Author: Dragon Xu
# Date: 06/04/2024
# Description: This main() program uses class Type3_Report to automate the generation of excel report for Type3_Report project,
# named as 'RXXXXX- Type 3 Report {period_code}.xlsx'.


def main():
    # Logging the start of the program
    console.log("RXXXXX - Type 3 Report Has Started\n")
    # Establish connection to Oracle database
    connection = conn.oracle_connect(
        config.host,
        config.port,
        config.instance,
        config.username,
        config.password,
        msg=False,
    )
    # Create an instance of the Type3_Report class
    dm = Type3_Rpt.Type3_Report(connection)
    # Create a workbook from the template excel file
    wb = load_workbook(config.template_file)
    # Initialize the main sheet
    ws_main = dm.create_main_sheet(wb, "Main Home")
    # Retrieve the data and put them into excel sheets
    run(dm, wb, ws_main)
    # Do some final processing after all the sheets have been created, including removing template sheets, updating field names, enabling links, and reordering sheets
    last_process(dm, wb)
    # Close and save the workbook with the formatted output file name
    close_wb(dm, wb)


def run(dm: Type3_Rpt, wb: Workbook, ws_main: Worksheet):
    """Retrieve the data and put them into excel sheets

    Args:
        dm (Type3_Rpt): an instance of the Type3_Report class
        wb (Workbook): the workbook object
        ws_main (Worksheet): the main sheet object
    """
    # Iterate through each category code (c1) and category name in the dictionary
    for c1, category_name in dm.dict_category.items():
        # Initialize variable c, which is the category code surrounded by % signs
        c = f"%{c1}%"
        # Initialize variable attr_count to 1 at the beginning of lopp for each category
        attr_count = 1
        # Reset the skip_rows_main_adjusted to the original skip_rows_main_sheet (8) for each category
        dm.skip_rows_main_adjusted = dm.skip_rows_main_sheet
        # Validate input data
        if c1 not in dm.dict_a1.keys():
            raise ValueError(f"No attributes defined for category {category_name[0]}")
        # For each category, iterate through each attribute. a1 is attribute name got directly from Oracle database, a2 is [attribute name + 'desc', more readable name]
        for a1, a2 in dm.dict_a1[c1].items():
            try:
                # Pull the merged and summary dataframes for each attribute in each category
                df_merged, df_summary = dm.pull_attributes(c, c1, a1, a2[0])
                # Generate the field sheet for each attribute in each category
                # Field sheet name for each attribute in each category, e.g. "Cig Fld 1"
                field_sheet_name = f"{category_name[1]} Fld {attr_count}"
                # Subset the merged dataframe to only include the first 6 columns because every category field report will have the same first 6 columns
                dm.create_field_sheet(
                    df_merged.iloc[:, :6],
                    wb,
                    field_sheet_name,
                    c1,
                    category_name[0],
                    a1,
                )
                # Increament attr_count by 1
                attr_count += 1
                # Put df_summary into the main sheet for each attribute in each category,
                # each time of loop it will put a row of data in certain column ranges into the main sheet
                dm.add_data_main_sheet(df_summary.iloc[:, 1:], ws_main, c1)
                # Increment the skip_rows_main_adjusted by 1
                dm.skip_rows_main_adjusted += 1
            except Exception as e:
                console.log(f"Error while processing data for {category_name} {a1}: {e}")

        # Generate the drill down data for each category
        do_drill_sheet(dm, wb, c, c1, category_name)
        # Generate the category home sheet for each category
        do_category_sheet(dm, wb, ws_main, c1, category_name)


def do_drill_sheet(dm: Type3_Rpt, wb: Workbook, c: str, c1: int, category_name: list):
    """Generate the drill down sheet for each category

    Args:
        dm (Type3_Rpt): an instance of the Type3_Report class
        wb (Workbook): the workbook object
        c (str): the category code surrounded by % signs
        c1 (int): the category code
        category_name (list): the category name list
    """
    # Generate drill down data for each category
    df_drilldown = dm.pull_drill_down_data(c, dm.dict_a1[c1])
    # Initialize sheet name for drill down report
    drill_sheet_name = f"{category_name[1]} Drill"
    # Create a drill sheet for each category in the workbook
    dm.create_drill_sheet(df_drilldown, wb, drill_sheet_name, category_name[0], c1)


def do_category_sheet(
    dm: Type3_Rpt, wb: Workbook, ws_main: Worksheet, c1: int, category_name: list
):
    """Generate the category home sheet for each category

    Args:
        dm (Type3_Rpt): an instance of the Type3_Report class
        wb (Workbook): the workbook object
        ws_main (Worksheet): the main sheet object
        c1 (int): the category code
        category_name (list): the category name list
    """
    # Initialize sheet name for drill down report
    category_sheet_name = f"{category_name[1]} Home"
    # Generate the category home sheet for each category
    dm.create_category_sheet(wb, category_sheet_name, ws_main, c1, category_name[0])


def last_process(dm: Type3_Rpt, wb: Workbook):
    """Do some final processing after all the sheets have been created, including removing template sheets, updating field names, enabling links, and reordering sheets

    Args:
        dm (Type3_Rpt): an instance of the Type3_Report class
        wb (Workbook): the workbook object
    """
    # Remove all the template sheets left in the workbook
    dm.remove_template_sheets(wb)
    # Update all the field sheet names
    dm.update_field_name(wb)
    # Enable the links between each sheet
    dm.enable_links(wb)
    # Reorder the sheets in the workbook
    dm.reorder_sheets(wb)


def close_wb(dm: Type3_Rpt, wb: Workbook):
    """Close and save the workbook with the formatted output file name

    Args:
        dm (Type3_Rpt): an instance of the Type3_Report class
        wb (Workbook): the workbook object
    """
    # Set the first sheet (Main Home) as the active sheet
    wb.active = 0
    # Get formatted output file name, make copy of template file
    output_file_name = config.output_file.format(period_code=dm.period)
    # Save the workbook
    wb.save(f"{output_file_name}")
    wb.close()


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
            "RXXXXX - Type 3 Report has COMPLETED SUCCESSFULLY ("
            + str("{:.2f}").format(total)
            + "s)"
        )
    except Exception as e:
        console.log(f"RXXXXX - Type 3 Report has ABORTED With WARNINGS/ERRORS: {e}")
    finally:
        warn.close()
