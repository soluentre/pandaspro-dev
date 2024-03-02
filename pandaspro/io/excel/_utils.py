from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string


def index_cell(cellname: str):
    """
    This function converts an Excel cell name (e.g., 'A1') into its corresponding row and column indices using openpyxl's utility functions. It separates the alphabetic column identifier(s) and the numeric row identifier, then converts the column identifier to a numeric index.

    Parameters
    ----------
    cellname : str
        The name of the cell in Excel format (e.g., 'A1', 'B22'), where letters refer to the column and numbers refer to the row.

    Returns
    -------
    tuple
        A tuple containing two elements:
        - The first element is an integer representing the row number.
        - The second element is an integer representing the column number.

    Notes
    -----
    The function relies on openpyxl's `coordinate_from_string` to split the cell name into its letter and number components, and `column_index_from_string` to convert the column letter(s) to a numeric index. It assumes the input is a valid Excel cell reference.

    Examples
    --------
    >>> index_cell('C3')
    This would return (3, 3), indicating that the cell is in the 3rd row and 3rd column of the spreadsheet.
    """

    column_letter, row_number = coordinate_from_string(cellname)  # Separate letter and number
    column_number = column_index_from_string(column_letter)  # Convert letter to number
    return (row_number, column_number)

def get_cell_aside(cellname: str,
                   direction: str = 'right',
                   skipnum: int = 1):
    '''
    Extracting column letters and row numbers from the cell reference
    '''
    row, col = index_cell(cellname)
    if direction == 'right':
        result_row, result_col = row, col+1
    elif direction == 'left':
        result_row, result_col = row, col-1
    elif

    if result_row < 0 or result_col < 0:
        raise ValueError("Row 1 and Column A are the borders of an Excel Spreadsheet")
        result_col =
    col_index = column_index_from_string(col_letters)
    next_col_letter = get_column_letter(col_index + skipnum)
    return next_col_letter + row_numbers
