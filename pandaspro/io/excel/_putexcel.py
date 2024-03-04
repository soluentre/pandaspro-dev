from pandas import DataFrame
import pandas as pd
from pathlib import Path
import os
import xlwings as xw

from pandaspro import FramePro
from pandaspro.io.excel._utils import CellPro


class FramexlWriter:

    def __init__(
            self,
            frame: DataFrame | FramePro,
            start_cell: str,
            index: bool = False,
            header: bool = True,
    ) -> None:
        cell = CellPro(start_cell)
        header_row_count = len(frame.columns.levels) if isinstance(frame.columns, pd.MultiIndex) else 1
        index_column_count = len(frame.index.levels) if isinstance(frame.index, pd.MultiIndex) else 1

        if header == True and index == True:
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1] + index_column_count
            export_data = frame
            range_index = cell.offset(header_row_count,0).resize(tr-header_row_count, index_column_count)
            range_indexnames = cell.resize(header_row_count, header_row_count)
            range_header = cell.offset(0, index_column_count).resize(header_row_count, tc-index_column_count)
        elif header == False and index == True:
            tr, tc = frame.shape[0], frame.shape[1] + index_column_count
            export_data = frame.reset_index().to_numpy().tolist()
            range_index = cell.resize(tr, index_column_count)
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            tr, tc = frame.shape[0], frame.shape[1]
            export_data = frame.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1]
            column_export = [list(lst) for lst in list(zip(*frame.columns.values))]
            export_data = column_export + frame.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cell.resize(header_row_count, tc)

        self.frame = export_data
        self.start_cell = start_cell
        self.tr = tr
        self.tc = tc
        self.range_data = cell.offset(header_row_count, index_column_count)\
                            .resize(tr - header_row_count, tc - index_column_count)
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_indexnames = range_indexnames.cell if range_indexnames !='N/A' else 'N/A'


class PutxlSet:

    def __init__(
            self,
            workbook: str,
            sheet_name: str,
            alwaysreplace: str = None,  # a global config that sets all the following actions to replace ...
            noisily: bool = False
    ) -> None:
        def _extract_filename_from_path(path):
            return Path(path).name

        def _get_open_workbook_by_name(name):
            # Return the open workbook by its name if exists, otherwise return None
            for app in xw.apps:
                for wb in app.books:
                    if wb.name == name:
                        return wb, app
            return None, None

        # App and Workbook declaration
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

        # Worksheet declaration
        current_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name in current_sheets:
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.add(after=wb.sheets.count)
            sheet.name = sheet_name

        self.workbook = workbook
        self.worksheet = sheet_name
        self.wb = wb
        self.ws = sheet
        self.globalreplace = alwaysreplace

    def putxl(
            self,
            start_cell: str,
            frame,
            index: bool = False,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = False,
            sheet_name: str = None,
            debug: bool = False
    ) -> None:
        io = FramexlWriter(frame=frame, start_cell=start_cell, index=index, header=header)
        replace_type = self.globalreplace if self.globalreplace else replace

        # If a sheet_name is specified at the very end, then override the current sheet
        if sheet_name and sheet_name != self.worksheet:
            if sheet_name in [sheet.name for sheet in self.wb.sheets]:
                ws = self.wb.sheets[sheet_name]
            else:
                ws = self.wb.sheets.add(after=self.wb.sheets.count)
                ws.name = sheet_name
        else:
            ws = self.ws

        # If sheetreplace or replace is specified, then delete the old sheet and create a new one
        if sheetreplace or replace_type == 'sheet':
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = ws.index
            original_name = ws.name
            total_count = self.wb.sheets.count
            if debug:
                print(f">>> Row 121: original index is {original_index}")
            ws.delete()
            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index - 1]])
                if debug:
                    print(f">>> Row 128: New sheet added after the sheet !'{_sheetmap[original_index - 1]}'")
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])
                if debug:
                    print(f">>> Row 132: New sheet added before the sheet !'{_sheetmap[original_index + 1]}'")
            new_sheet.name = original_name
            ws = new_sheet
            self.ws = ws

        ws.range(io.start_cell).value = io.frame

        if debug:
            print(f"\n>>> Cell Range Analysis")
            print(f" ----------------------")
            print(f">>> Total row: {io.tr}, Total column: {io.tc}")
            print(f">>> Range index: {io.range_index}, Range header: {io.range_header}, Range index names: {io.range_indexnames}\n")


if __name__ == '__main__':
    import pandas as pd
    import numpy as np

    # Define the countries
    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]

    # Generate random data for GDP (in trillion USD), Population (in millions), and GDP per Capita (in USD)
    np.random.seed(0)  # For reproducibility
    gdp = np.random.uniform(1, 20, size=len(countries))  # GDP in trillion USD
    population = np.random.uniform(10, 1400, size=len(countries))  # Population in millions
    gdp_per_capita = gdp * 1e12 / (population * 1e6)  # GDP per Capita in USD

    # Create the DataFrame
    df1 = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp.round(2),
        'Population (Millions)': population.round(1),
        'GDP per Capita (USD)': gdp_per_capita.round(2)
    })

    import pandas as pd

    # Re-create the initial DataFrame
    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]
    gdp = [11.43, 14.59, 12.45, 11.35, 9.05, 13.27, 9.31, 17.94, 19.31, 8.29]
    population = [1110.5, 745.2, 799.6, 1296.6, 108.7, 131.1, 38.1, 1167.3, 1091.6, 1219.3]
    gdp_per_capita = [x * 1e12 / y for x, y in zip(gdp, population)]
    df = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp,
        'Population (Millions)': population,
        'GDP per Capita (USD)': gdp_per_capita
    })

    # Convert index and column headers to MultiIndex

    # For the index, use a combination of 'Region' and 'Country'
    regions = ['North America', 'Asia', 'Asia', 'Europe', 'Asia', 'Europe', 'Europe', 'South America', 'Europe',
               'North America']
    index_multi = pd.MultiIndex.from_arrays([regions, df['Country']], names=['Region', 'Country'])

    # For the columns, create a MultiIndex with two levels: 'Indicator' and 'Measure'
    columns_multi = pd.MultiIndex.from_product([['Economic Indicators'], df.columns[1:]],
                                               names=['Indicator', 'Measure'])

    # Create a new DataFrame with MultiIndex for both rows and columns
    df = pd.DataFrame(df.values[:, 1:], index=index_multi, columns=columns_multi)

    ps = PutxlSet('test.xlsx', 'Sheet3', noisily=True)
    ps.putxl('A1', df1, index=True, header=True, sheetreplace=True, sheet_name='TT', debug=True)
    ps.putxl('A1', df1, index=True, header=False, sheetreplace=True, sheet_name='TF', debug=True)
    ps.putxl('A1', df1, index=False, header=True, sheetreplace=True, sheet_name='FT', debug=True)
    ps.putxl('A1', df1, index=False, header=False, sheetreplace=True, sheet_name='FF', debug=True)

'''
putxl('A1', ...)
putxlsheet('A1', ..., sheet='Sheet2')
ps.switch('Sheet2')
putxl('A1', )
ps.switch('Dashboard')
'''
