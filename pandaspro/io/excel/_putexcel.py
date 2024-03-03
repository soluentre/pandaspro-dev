from pandas import DataFrame
import pandas as pd
from pathlib import Path
import os
import xlwings as xw

from pandaspro import FramePro, resize


class FramexlWriter:
    '''
    key attributes are the xl cell ranges in String
    '''

    def __init__(self, frame: DataFrame | FramePro,
                 start_cell: str,
                 index: bool = False,
                 header: bool = True):
        header_row_count = len(frame.columns.levels) if isinstance(frame.columns, pd.MultiIndex) else 1
        index_column_count = len(frame.index.levels) if isinstance(frame.index, pd.MultiIndex) else 1

        if header == True and index == True:
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1] + index_column_count
            export_data = frame
        elif header == True and index == False:
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1]
            export_data = [frame.columns.tolist()] + frame.to_numpy().tolist()
        elif header == False and index == True:
            tr, tc = frame.shape[0], frame.shape[1] + index_column_count
            export_data = frame.reset_index().to_numpy().tolist()
        elif header == False and index == False:
            tr, tc = frame.shape[0], frame.shape[1]
            export_data = frame.to_numpy().tolist()

        self.frame = export_data
        self.start_cell = start_cell


class PutxlSet:
    def __init__(self,
                 workbook: str,
                 sheet_name: str,
                 noisily: bool = False):

        def _extract_filename_from_path(path):
            return Path(path).name

        def _get_open_workbook_by_name(name):
            """Return the open workbook by its name if exists, otherwise return None."""
            for app in xw.apps:
                for wb in app.books:
                    if wb.name == name:
                        return wb, app
            return None, None

        wb, app = _get_open_workbook_by_name(_extract_filename_from_path(workbook))  # Check if the file is already open
        if wb:
            if noisily:
                print(f"{workbook} is already open, closing ...")
            wb.save()
            wb.close()
            if not app.books:  # Check if the app has no more workbooks open; if true, then quit the app
                app.quit()
        elif noisily:
            print(f"Working on {workbook} now ...")

        if not os.path.exists(workbook):  # Check if the file already exists
            wb = xw.Book()  # If not, create a new Excel file
            wb.save(workbook)
        else:
            wb = xw.Book(workbook)

        current_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name in current_sheets:
            if sheetreplace:
                # Delete the existing sheet and create a new one
                new_sheet = wb.sheets.add(after=wb.sheets.count)
                wb.sheets[sheet_name].delete()
                new_sheet.range(start_cell).value = export_data
                new_sheet.name = sheet_name
            else:
                wb.sheets[sheet_name].range(start_cell).value = export_data
    def putxl(self, start_cell: str,
              frame,
              index: bool = False,
              header: bool = True):
        io = FramexlWriter(frame=frame,
                           start_cell=start_cell,
                           index=index,
                           header=header)
        self.wb.range(io.start_cell).value = io.frame


if __name__ == '__main__':
    import numpy as np

    # Define the countries
    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]

    # Generate random data for GDP (in trillion USD), Population (in millions), and GDP per Capita (in USD)
    np.random.seed(0)  # For reproducibility
    gdp = np.random.uniform(1, 20, size=len(countries))  # GDP in trillion USD
    population = np.random.uniform(10, 1400, size=len(countries))  # Population in millions
    gdp_per_capita = gdp * 1e12 / (population * 1e6)  # GDP per Capita in USD

    # Create the DataFrame
    df = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp.round(2),
        'Population (Millions)': population.round(1),
        'GDP per Capita (USD)': gdp_per_capita.round(2)
    })

    ps = PutxlSet('test.xlsx', 'Sheet1', noisily=True)
    ps.putxl('A1', df)
'''
putxl('A1', ...)
putxlsheet('A1', ..., sheet='Sheet2')
ps.switch('Sheet2')
putxl('A1', )
ps.switch('Dashboard')
'''
