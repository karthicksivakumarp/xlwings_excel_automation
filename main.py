from xlwings_class import xlwings_driver


def test_xlwings_driver():
    # Create an instance of xlwings_driver
    xlwings_instance = xlwings_driver()

    # Test opening or creating a workbook
    xlwings_instance.open_workbook()

    # Test creating a new sheet
    xlwings_instance.create_new_sheet("TestSheet")

    # Test renaming a sheet
    xlwings_instance.rename_sheet("TestSheet", "RenamedSheet")

    # Test listing sheet names
    xlwings_instance.list_sheet_names()

    # Test selecting a sheet
    xlwings_instance.select_sheet("RenamedSheet")

    # Test writing to a cell
    xlwings_instance.write_to_cell("A1", "Hello, xlwings!")

    # Test writing to a range
    xlwings_instance.write_to_range("B1", "D1", [1, 2, 3])
    xlwings_instance.write_to_range("A4", "A7", [[4], [5], [6], [7]])

    # Test reading from a cell
    value = xlwings_instance.read_from_cell("A1")
    print("Read value from A1:", value)

    # Test reading from a range
    values = xlwings_instance.read_from_range("B1", "D1")
    print("Read values from B1:D1:", values)

    values = xlwings_instance.read_from_range("A4", "A7")
    print("Read values from A4:A7:", values)

    # Test formatting functions
    xlwings_instance.set_bold("A1")
    xlwings_instance.set_italic("A1")
    xlwings_instance.set_font_size("A1", 14)

    # Corrected text color and cell color settings
    xlwings_instance.set_text_color("A1", (255, 0, 0))  # Red color
    xlwings_instance.set_cell_color("A1", (255, 255, 0))  # Yellow color

    xlwings_instance.set_text_direction("A1", 45)

    # Test formatting for a range
    xlwings_instance.set_bold_for_range("B1", "D1")
    xlwings_instance.set_italic_for_range("B1", "D1")
    xlwings_instance.set_font_size_for_range("B1", "D1", 12)

    # Corrected text color and cell color settings for range
    xlwings_instance.set_text_color_for_range("B1", "D1", (0, 0, 255))  # Blue color
    xlwings_instance.set_cell_color_for_range("B1", "D1", (0, 255, 0))  # Green color

    xlwings_instance.set_text_direction_for_range("B1", "D1", -45)

    # Test getting last filled row and column
    last_row = xlwings_instance.get_last_filled_row(1)
    print("Last filled row in column 1:", last_row)

    # Test text wrapping and autofit columns
    xlwings_instance.wrap_text_for_range("A1", "D1")
    xlwings_instance.autofit_columns_for_range("A1", "D1")

    # Test number formatting and text alignment
    xlwings_instance.format_cells("B1", "D1", "0.00")

    # Text alignment settings
    xlwings_instance.format_text_alignment("B1", "D1", horizontal_alignment="Right", vertical_alignment="Top")

    # Test setting cell style
    xlwings_instance.write_to_range("A3", "D3", [100, 200, 300, 400])
    xlwings_instance.set_cell_style("A3", "D3", "Heading 1")

    # Test conditional formatting for a cell
    xlwings_instance.set_conditional_formatting_for_cell("B1", 1, 7, 10, (0, 0, 0))  # Format if greater than 10

    # Test conditional formatting for a range
    xlwings_instance.set_conditional_formatting_for_range("C1", "D1", 2, 5, "=B1", (255, 0, 0))  # Format if equal to B1

    # Test saving and closing the workbook
    xlwings_instance.save_workbook()
    xlwings_instance.close_workbook()


if __name__ == "__main__":
    test_xlwings_driver()
