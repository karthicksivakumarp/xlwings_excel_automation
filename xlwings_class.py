import tkinter as tk
from tkinter import filedialog, messagebox
import xlwings as xw
from xlwings.constants import Direction
from xlwings.utils import rgb_to_int


class xlwings_driver:
    def __init__(self, workbook_name=None):
        # Initialize the xlwings_driver with workbook_name, active_sheet, and wb attributes
        self.active_sheet = None
        self.workbook_name = workbook_name
        self.wb = None  # Workbook reference

    def get_workbook_name_from_dialog(self):
        # Open a file dialog to get the name of an existing workbook
        root = tk.Tk()
        root.withdraw()  # Hide the main window
        self.workbook_name = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"),
                                                                                             ("All files", "*.*")])
        # Check if a workbook is selected, create a new xlwings Book, and save it with the provided name
        if not self.workbook_name:
            print("No workbook selected. Exiting.")
            exit()
        self.wb = xw.Book()
        print(f"Opened existing workbook: {self.workbook_name}")
        self.wb.save(self.workbook_name)
        root.destroy()

    def open_workbook(self):
        # Open an existing workbook or create a new one based on user choice
        if self.workbook_name is None:
            root = tk.Tk()
            root.withdraw()  # Hide the main window

            user_choice = messagebox.askquestion("Create or Open an Existing Excel",
                                                 "Click Yes to create a new workbook or "
                                                 "Click No to open an existing workbook")

            if user_choice == 'yes':
                self.create_new_workbook()
            else:
                self.get_workbook_name_from_dialog()
        else:
            # Try to open the workbook using xlwings.Book
            try:
                self.wb = xw.Book(self.workbook_name)
                print(f"Opened workbook: {self.workbook_name}")
            except Exception as e:
                # Handle the exception and retry
                print(f"Error opening workbook: {e}")
                self.open_workbook()

    def create_new_workbook(self):
        # Create a new workbook and save it with a specified name
        try:
            self.workbook_name = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                              filetypes=[("Excel files", "*.xlsx")])
            if not self.workbook_name:
                print("No workbook selected. Exiting.")
                exit()
            self.wb = xw.Book()
            print("Created a new workbook.")
            self.wb.save(self.workbook_name)
        except Exception as e:
            print(f"Error creating a new workbook: {e}")
            exit()

    def create_new_sheet(self, sheet_name):
        # Create a new sheet in the workbook
        try:
            self.wb.sheets.add(sheet_name)
            print(f"Created a new sheet: {sheet_name}")
        except Exception as e:
            print(f"Error creating a new sheet: {e}")

    def rename_sheet(self, old_name, new_name):
        # Rename an existing sheet in the workbook
        try:
            self.wb.sheets[old_name].name = new_name
            print(f"Renamed sheet from '{old_name}' to '{new_name}'")
        except Exception as e:
            print(f"Error renaming sheet: {e}")

    def list_sheet_names(self):
        # Get the names of all sheets in the workbook
        sheet_names = [sheet.name for sheet in self.wb.sheets]
        print(f"Sheet names: {sheet_names}")
        return sheet_names

    def select_sheet(self, sheet_name):
        # Select a specific sheet in the workbook
        try:
            self.active_sheet = self.wb.sheets[sheet_name]
            print(f"Selected sheet: {sheet_name}")
        except Exception as e:
            print(f"Error selecting sheet: {e}")

    def write_to_cell(self, cell_address, value):
        # Write a value to a specific cell in the active sheet
        try:
            self.active_sheet.range(cell_address).value = value
            print(f"Written '{value}' to cell {cell_address}")
        except Exception as e:
            print(f"Error writing to cell: {e}")

    def write_to_range(self, start_address, end_address, values):
        # Write values to a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            self.active_sheet.range(range_address).value = values
            print(f"Written values to range {range_address}")
        except Exception as e:
            print(f"Error writing to range: {e}")

    def read_from_cell(self, cell_address):
        # Read the value from a specific cell in the active sheet
        try:
            value = self.active_sheet.range(cell_address).value
            print(f"Read from cell {cell_address}: {value}")
            return value
        except Exception as e:
            print(f"Error reading from cell: {e}")

    def read_from_range(self, start_address, end_address):
        # Read values from a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            values = self.active_sheet.range(range_address).value
            print(f"Read values from range {range_address}: {values}")
            return values
        except Exception as e:
            print(f"Error reading from range: {e}")

    def set_bold(self, cell_address):
        # Set bold formatting for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Font.Bold = True
            print(f"Set bold formatting for cell {cell_address}")
        except Exception as e:
            print(f"Error setting bold formatting: {e}")

    def set_italic(self, cell_address):
        # Set italic formatting for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Font.Italic = True
            print(f"Set italic formatting for cell {cell_address}")
        except Exception as e:
            print(f"Error setting italic formatting: {e}")

    def set_font_size(self, cell_address, size):
        # Set font size for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Font.Size = size
            print(f"Set font size {size} for cell {cell_address}")
        except Exception as e:
            print(f"Error setting font size: {e}")

    def set_text_color(self, cell_address, color):
        # Set text color for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Font.Color = rgb_to_int(color)
            print(f"Set text color {color} for cell {cell_address}")
        except Exception as e:
            print(f"Error setting text color: {e}")

    def set_cell_color(self, cell_address, color):
        # Set cell color for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Interior.Color = rgb_to_int(color)
            print(f"Set cell color {color} for cell {cell_address}")
        except Exception as e:
            print(f"Error setting cell color: {e}")

    def set_text_direction(self, cell_address, direction):
        # Set text direction for a specific cell in the active sheet
        try:
            cell = self.active_sheet.range(cell_address)
            cell.api.Orientation = direction
            print(f"Set text direction {direction} for cell {cell_address}")
        except Exception as e:
            print(f"Error setting text direction: {e}")

    def set_bold_for_range(self, start_address, end_address):
        # Set bold formatting for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Font.Bold = True
            print(f"Set bold formatting for range {range_address}")
        except Exception as e:
            print(f"Error setting bold formatting for range: {e}")

    def set_italic_for_range(self, start_address, end_address):
        # Set italic formatting for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Font.Italic = True
            print(f"Set italic formatting for range {range_address}")
        except Exception as e:
            print(f"Error setting italic formatting for range: {e}")

    def set_font_size_for_range(self, start_address, end_address, size):
        # Set font size for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Font.Size = size
            print(f"Set font size {size} for range {range_address}")
        except Exception as e:
            print(f"Error setting font size for range: {e}")

    def set_text_color_for_range(self, start_address, end_address, color):
        # Set text color for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Font.Color = rgb_to_int(color)
            print(f"Set text color {color} for range {range_address}")
        except Exception as e:
            print(f"Error setting text color for range: {e}")

    def set_cell_color_for_range(self, start_address, end_address, color):
        # Set cell color for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Interior.Color = rgb_to_int(color)
            print(f"Set cell color {color} for range {range_address}")
        except Exception as e:
            print(f"Error setting cell color for range: {e}")

    def set_text_direction_for_range(self, start_address, end_address, direction):
        # Set text direction for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.Orientation = direction
            print(f"Set text direction {direction} for range {range_address}")
        except Exception as e:
            print(f"Error setting text direction for range: {e}")

    def get_last_filled_row(self, column):
        # Get the last filled row number in a specific column of the active sheet
        try:
            last_row = self.active_sheet.cells(self.active_sheet.cells.last_cell.row, column).end(Direction.xlUp).row
            print(f"Last filled row in column {column}: {last_row}")
            return last_row
        except Exception as e:
            print(f"Error getting last filled row: {e}")

    def get_last_filled_column(self, row):
        # Get the last filled column name in a specific row of the active sheet
        try:
            last_column = self.active_sheet.cells(row, self.active_sheet.cells.last_cell.column).end(
                Direction.xlToRight).column
            last_column_letter = xw.utils.col_name(last_column)
            print(f"Last filled column in row {row}: {last_column_letter}")
            return last_column_letter
        except Exception as e:
            print(f"Error getting last filled column: {e}")

    def wrap_text_for_range(self, start_address, end_address):
        # Enable text wrapping for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.WrapText = True
            print(f"Wrapped text for range {range_address}")
        except Exception as e:
            print(f"Error wrapping text for range: {e}")

    def autofit_columns_for_range(self, start_address, end_address):
        # Auto-size columns for a range of cells in the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.columns.autofit()
            print(f"Auto-sized columns for range {range_address}")
        except Exception as e:
            print(f"Error auto-sizing columns for range: {e}")

    def format_cells(self, start_address, end_address, number_format):
        # Apply a specific number format to cells in a range of the active sheet
        try:
            range_address = f"{start_address}:{end_address}"
            range_cell = self.active_sheet.range(range_address)
            range_cell.api.NumberFormat = number_format
            print(f"Formatted cells in range {range_address} with format: {number_format}")
        except Exception as e:
            print(f"Error formatting cells: {e}")

    def format_text_alignment(self, start_address, end_address, horizontal_alignment="center",
                              vertical_alignment="center"):
        try:
            # Construct the address string for the range
            range_address = f"{start_address}:{end_address}"
            # Get the range object using xlwings
            range_cell = self.active_sheet.range(range_address)

            # Convert alignment values to xlwings.constants values
            horizontal_alignment = getattr(xw.constants.HAlign, f"xlHAlign{horizontal_alignment.capitalize()}")
            vertical_alignment = getattr(xw.constants.VAlign, f"xlVAlign{vertical_alignment.capitalize()}")

            # Set the horizontal and vertical alignment for the range
            range_cell.api.HorizontalAlignment = horizontal_alignment
            range_cell.api.VerticalAlignment = vertical_alignment

            # Print a message indicating the successful formatting
            print(
                f"Formatted text alignment in range {range_address}: Horizontal - {horizontal_alignment}, "
                f"Vertical - {vertical_alignment}")
        except Exception as e:
            # Print an error message if an exception occurs during formatting
            print(f"Error formatting text alignment: {e}")

    def set_cell_style(self, start_address, end_address, style_name="Normal"):
        try:
            # Construct the address string for the range
            range_address = f"{start_address}:{end_address}"
            # Get the range object using xlwings
            range_cell = self.active_sheet.range(range_address)

            # Apply the specified style to the range
            range_cell.api.Style = style_name

            # Print a message indicating the successful application of cell style
            print(f"Set cell style in range {range_address}: {style_name}")
        except Exception as e:
            # Print an error message if an exception occurs during style setting
            print(f"Error setting cell style: {e}")

    def set_conditional_formatting_for_cell(self, cell_address, condition_type, operator, formula1, format_style):
        try:
            # Get the cell object using xlwings
            cell = self.active_sheet.range(cell_address)
            # Add conditional formatting to the cell
            cond_format = cell.api.FormatConditions.Add(Type=condition_type, Operator=operator, Formula1=formula1)
            # Set the font color for the conditional formatting
            cond_format.Font.Color = rgb_to_int(format_style)
            # Print a message indicating the successful application of conditional formatting
            print(f"Set conditional formatting for cell {cell_address}")
        except Exception as e:
            # Print an error message if an exception occurs during conditional formatting
            print(f"Error setting conditional formatting for cell: {e}")

    def set_conditional_formatting_for_range(self, start_address, end_address, condition_type, operator, formula1,
                                             format_style):
        try:
            # Get the range object using xlwings
            range_cell = self.active_sheet.range(f"{start_address}:{end_address}")
            # Add conditional formatting to the range
            cond_format = range_cell.api.FormatConditions.Add(Type=condition_type, Operator=operator, Formula1=formula1)
            # Set the font color for the conditional formatting
            cond_format.Font.Color = rgb_to_int(format_style)
            # Print a message indicating the successful application of conditional formatting
            print(f"Set conditional formatting for range {start_address}:{end_address}")
        except Exception as e:
            # Print an error message if an exception occurs during conditional formatting
            print(f"Error setting conditional formatting for range: {e}")

    def save_workbook(self):
        # Save the changes made to the workbook
        self.wb.save()
        # Print a message indicating that the workbook has been saved
        print(f"Workbook saved: {self.workbook_name}")

    def close_workbook(self):
        if self.wb is not None:
            # Close the workbook if it is not already closed
            self.wb.close()
            # Print a message indicating that the workbook has been closed
            print(f"Closed workbook: {self.workbook_name}")

