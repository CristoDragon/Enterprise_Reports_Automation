import pandas as pd
# Set the Matplotlib backend to 'Agg' before importing pyplot
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import os
import traceback
import PyPDF2
import io
import datetime
import textwrap
import pdr.handlers.Console_Handler as console
from typing import Union
import numpy as np

# Description: This class is used to convert excel file to pdf format.
# Author: Dragon Xu


class ExcelToPDF:
    def __init__(self, excel_path: Union[list, str], output_path: Union[list, str], header_index: dict = None):
        # Initialize a list of Excel file paths and output PDF file paths
        self.excel_path = excel_path if isinstance(excel_path, list) else [excel_path]
        self.output_path = output_path if isinstance(output_path, list) else [output_path]
        # e.g. headers={'Sheet1': 0, 'Sheet2': 2}
        self.header_index = header_index
        # Check if the input and output paths are valid
        self.check_file_exist(self.excel_path)
        self.check_file_exist([os.path.dirname(x) for x in self.output_path])
        # Get today's date in the format MM/DD/YYYY
        self.today = datetime.datetime.now().strftime("%m/%d/%Y")
        # Initialize constants for the PDF generation
        self.min_font_size = 7
        self.page_height = 4
        self.page_width = 11
        self.header_row_height_coef = 1.8 * self.min_font_size / 72 
        self.data_row_height_coef = 0.5 * self.min_font_size / 72
        self.header_row_height = self.header_row_height_coef * self.page_height
        self.data_row_height = self.data_row_height_coef * self.page_height
        self.top_margin = 0.3 * self.min_font_size / 72
        self.bottom_margin = self.top_margin
        self.left_margin = self.top_margin
        self.right_margin = self.top_margin
        self.character_width_in_points = 6
        # Initialize Excel metadata to store file name, sheet names, and page numbers
        self.excel_metadata = {}
        

    def check_null_empty(self, value) -> bool:
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
        

    def check_file_exist(self, file_path_list: Union[list, str]) -> None:
        """Check if the files in the list exist.

        Args:
            file_path_list (list): The list of file paths to check.

        Raises:
            FileNotFoundError: If the file is not found.
        """
        # Check if the file path list is None or empty
        if self.check_null_empty(file_path_list):
            raise ValueError("The file path list is None or empty.")
        # Check if the input file is a string
        if isinstance(file_path_list, str):
            # Print a warning message
            console.log(f"WARNING: The input file {file_path_list} is a string.")
            # Check if the file exists
            if not os.path.exists(file_path_list):
                raise FileNotFoundError(f"File str not found: {file_path_list}")
        else:
            for path in file_path_list:
                if not os.path.exists(path):
                    raise FileNotFoundError(f"File not found: {path}")


    def read_excel(self, excel_path: str) -> dict:
        """Read an Excel file and return a dictionary of DataFrames for each sheet.

        Args:
            excel_path (str): The path to the Excel file.

        Raises:
            Exception: _description_

        Returns:
            dict: A dictionary of DataFrames for each sheet in the Excel file.
        """
        
        try:
            # Read all sheets, if headers not specified, pandas uses the first row as headers by default
            if self.header_index:
                dict_dfs = {sheet: pd.read_excel(excel_path, sheet_name=sheet, header=self.header_index.get(sheet, 0))
                            for sheet in pd.ExcelFile(excel_path).sheet_names}
            else:
                dict_dfs = pd.read_excel(excel_path, sheet_name=None)
            # Check if each DataFrame is empty
            for sheet_name, df in dict_dfs.items():
                if df.empty:
                    console.log(f"Warning: DataFrame for sheet '{sheet_name}' is empty in {excel_path}")
            return dict_dfs
        except Exception as e:
            raise Exception(f"Failed to read Excel file: {e}")
        

    def wrap_columns(self, df: pd.DataFrame, available_width: int) -> list:
        """Wrap the column headers if they are too long to fit the available width.

        Args:
            df (pd.DataFrame): _description_
            available_width (int): _description_

        Returns:
            list: _description_
        """
        # Calculate the width for each column
        wrap_width = max(5, int(available_width / len(df.columns) / 0.1))
        # Wrap the column headers if they are too long
        return [textwrap.fill(str(col), width=wrap_width) for col in df.columns]
    

    def calculate_column_widths(self, df: pd.DataFrame) -> list:
        """Calculate the needed width for each column based on content.

        Args:
            df (pd.DataFrame): The DataFrame to analyze.

        Returns:
            list: A list of column widths.
        """
        # Assume character width as a constant, here it's a rough average character width in points
        char_width = self.min_font_size * 0.5 / 72  # Convert points to inches
        max_widths = []
        for column in df.columns:
            # Find the maximum length of the content in each column
            max_len = df[column].astype(str).map(len).max()
            # Calculate the width required to display the maximum length
            max_widths.append(max_len * char_width)
        return max_widths
        

    def set_style_col_headers(self, table, header_row_height_coef: float, data_row_height_coef: float):
        """Set the style for the column headers in the table for each page.

        Args:
            table (_type_): _description_
            header_row_height_coef (_type_): _description_
            data_row_height_coef (_type_): _description_
        """
        # Style the column header row (row 0)
        for (i, j), cell in table.get_celld().items():
            # Header row
            if i == 0:  
                cell.set_text_props(weight='bold')
                # Add border to header cells
                cell.set_edgecolor('black')  
                # Make the header border thicker
                cell.set_linewidth(1.2)  
                cell.set_height(header_row_height_coef)
            else:
                # Remove border from data cells
                cell.set_edgecolor('none')  
                # No border for data rows
                cell.set_linewidth(0)  
                cell.set_height(data_row_height_coef)


    def create_all_pages(self, dfs: dict, output_path: str, title: bool = False):
        """Create PDF pages from dataframes, ensuring headers are repeated on each page.

        Args:
            dfs (dict): a dictionary of DataFrames, where the key is the sheet name and the value is the DataFrame.
            output_path (str): the path to save the output PDF file.
            title (bool, optional): whether to print the sheet name for each pdf page. Defaults to False.
        """
        # Initialize single workbook metadata
        workbook_metadata = {}
        page_num = 0
        # Dynamically calculate the available height and width for the table
        available_height = self.page_height - self.header_row_height - self.top_margin - self.bottom_margin
        available_width = self.page_width - (self.left_margin + self.right_margin)
        with PdfPages(output_path) as pdf:
            for sheet_name, df in dfs.items():
                # Replace NaN values with empty string for better display
                df = df.fillna('')
                # Remove all newlines from the data
                df = df.replace('\n', ' ', regex=True)
                # Dynamically calculate how many data rows can fit on a page
                rows_per_page = int(available_height / self.data_row_height) + 24
                # Wrap column headers if they are too long
                column_widths = self.calculate_column_widths(df)
                total_column_width = sum(column_widths)
                wrapped_columns = self.wrap_columns(df, available_width)
                # Dynamically calculate the scaling factor for the table
                row_scale = min(1, available_height / (self.data_row_height * rows_per_page))
                col_scale = self.page_width / total_column_width if total_column_width > 0 else 1
                # Initialize worksheet page number list
                workbook_metadata[sheet_name] = [page_num]
                for start_row in range(0, len(df), rows_per_page):
                    # Get information about the current page data
                    end_row = min(start_row + rows_per_page, len(df))
                    num_rows = end_row - start_row
                    num_rows_not_used = rows_per_page - num_rows
                    # Create a new figure for each page
                    fig, ax = plt.subplots(figsize=(self.page_width, self.page_height)) 
                    # Turn off the axis
                    ax.axis('off')
                    # Adjust layout
                    plt.tight_layout() 
                    # Set the margins within the plot
                    plt.subplots_adjust(
                        top=1 - self.top_margin,
                        bottom=self.bottom_margin,
                        left=self.left_margin,
                        right=1 - self.right_margin
                    )
                    # Extend data by rows
                    original_data = df.iloc[start_row:end_row].values
                    num_columns = len(df.columns)
                    limit = 22
                    if num_rows_not_used > limit:
                        extra_rows = [np.array([""] * num_columns)] * (limit - num_rows)
                    else:
                        extra_rows = [np.array([""] * num_columns)] * 3
                    extended_data = np.vstack([original_data, extra_rows])
                    # Include column labels on each page (use wrapped columns)
                    table = ax.table(cellText=extended_data, colLabels=wrapped_columns, loc='center', cellLoc='center')
                    # Apply calculated column widths
                    for i, width in enumerate(column_widths):
                        table.auto_set_column_width(col=i)  # Turn off automatic column width
                        table._cells[(0, i)].set_width(width * col_scale)  # Set custom width for each column
                    # Adjust font size and scaling
                    table.auto_set_font_size(True)
                    # Adjust scaling to fit the layout better
                    table.scale(col_scale, row_scale)
                    # Set the style for the column headers
                    self.set_style_col_headers(table, self.header_row_height_coef, self.data_row_height_coef)
                    # Optionally add titles
                    if title:
                        plt.title(f"{sheet_name}")
                    pdf.savefig(fig, orientation='landscape', bbox_inches='tight')
                    plt.close(fig)
                    # Append the page number to the worksheet metadata
                    workbook_metadata[sheet_name].append(page_num)
                    # Increment the page number for the current worksheet
                    page_num += 1

        # Store the workbook metadata for the current Excel file
        self.excel_metadata[output_path] = workbook_metadata
  

    def add_footers_to_pdf(self, input_path: str, total_pages: int, font: str, font_size: int) -> None:
        """
        Adds a footer to each page of an existing PDF file.
        
        Args:
            input_path (str): Path to the existing PDF file.
            output_path (str): Path where the modified PDF file with footers will be saved.
            total_pages (int): The total number of pages in the PDF, used for "Page x of n" footer.
        """
        # Create a new PdfWriter object for outputting the modified PDF.
        output = PyPDF2.PdfWriter()
        # Open the existing PDF file.
        with open(input_path, "rb") as file:
            # Create a PdfReader object to read pages from the existing PDF.
            reader = PyPDF2.PdfReader(file)
            # Iterate through each page in the PDF.
            for page_number in range(total_pages):
                # Retrieve each page from the PDF reader.
                page = reader.pages[page_number]
                # Determine page width for dynamic positioning
                page_width = float(page.mediabox[2])
                # Create a byte stream to hold PDF commands for adding the footer.
                packet = io.BytesIO()
                # Create a canvas on which to draw the footer text.
                can = canvas.Canvas(packet, pagesize=(page_width, letter[1]))
                # Set the font and size for the footer text
                can.setFont(font, font_size)
                 # Calculate positions based on page width
                left_x = 8
                center_x = page_width / 2
                right_x = page_width - 8
                # Draw the footer text at the specified coordinates.
                # Left part: file name
                can.drawString(left_x, 20, str(os.path.basename(input_path)).split(".")[0])
                # Left part, below file name: sheet name
                if input_path in self.excel_metadata:
                    for sheet_name, page_num_list in self.excel_metadata[input_path].items():
                        if page_number in page_num_list:
                            can.drawString(left_x, 8, f"Sheet: {sheet_name}")
                            break
                # Center part: "Page x of n"
                can.drawCentredString(center_x, 8, f"Page {page_number + 1} of {total_pages}")
                # Right part: today's date
                can.drawRightString(right_x, 8, self.today)
                # Finalize the changes to the canvas.
                can.save()
                # Move back to the start of the StringIO buffer.
                packet.seek(0)
                # Read the footer content using a PdfReader from the buffer.
                new_pdf = PyPDF2.PdfReader(packet)
                # Merge the footer content onto the current page.
                page.merge_page(new_pdf.pages[0])
                # Add the modified page to the output PDF.
                output.add_page(page)
                
            # Write the output PDF to the specified output path.
            with open(input_path, "wb") as outputStream:
                output.write(outputStream)
                    

    def get_total_pdf_pages(self, pdf_path: str) -> int:
        """Get the total number of pages in a PDF file.

        Args:
            pdf_path (str): The path to the PDF file to analyze.

        Returns:
            int: The total number of pages in the PDF file.
        """
        with open(pdf_path, "rb") as file:
            reader = PyPDF2.PdfReader(file)
            return len(reader.pages)


    def convert_to_pdf(self, excel_path: str, output_path: str, title: bool = False):
        """Converts each sheet in the Excel file to a PDF report.

        Args:
            excel_path (str): the path to the Excel file to convert
            output_path (str): the path to save the output PDF file
            title (bool, optional): whether to print the sheet name for each pdf page. Defaults to False.

        Raises:
            Exception: if the conversion fails
        """
        try:
            # Read the Excel file as a dictionary of DataFrames
            dfs = self.read_excel(excel_path)
            # Create all the pdf pages for one sheet in the excel file
            self.create_all_pages(dfs, output_path, title)
            # Get the total number of pages in the pdf
            total_pages = self.get_total_pdf_pages(output_path)
            # Add footers to each page in the pdf
            self.add_footers_to_pdf(output_path, total_pages, "Helvetica", 8)
        except Exception as e:
            raise Exception(f"Failed to convert to PDF: {e}")


    def run_conversion(self):
        """Main method to run the conversion process."""
        try:
            # Iterate over each Excel file and output path
            for excel_path, pdf_path in zip(self.excel_path, self.output_path):
                # Start the conversion process
                self.convert_to_pdf(excel_path, pdf_path)
                console.log(f"Successfully converted {excel_path} to {pdf_path}")
        except Exception as e:
            console.log(f"Error during conversion: {e}\n{traceback.format_exc()}")
            traceback.print_exc()
            raise e