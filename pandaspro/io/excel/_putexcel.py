from pathlib import Path
import os
import xlwings as xw

from pandaspro.io.excel._framewriter import FramexlWriter


def is_range_filled(ws, range_str: str = None):
    if range_str is None:
        return False
    else:
        rng = ws.range(range_str)
        for cell in rng:
            if cell.value is not None and str(cell.value).strip() != '':
                return True
        return False


class PutxlSet:
    def __init__(
            self,
            workbook: str,
            sheet_name: str = None,
            alwaysreplace: str = None,  # a global config that sets all the following actions to replace ...
            noisily: bool = None
    ):
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
        if sheet_name is None:
            sheet_name = wb.sheets[0].name

        current_sheets = [sheet.name for sheet in wb.sheets]
        if sheet_name in current_sheets:
            sheet = wb.sheets[sheet_name]
        else:
            sheet = wb.sheets.add(after=wb.sheets.count)
            sheet.name = sheet_name

        # A small function to decide if sheet is blank ...
        def is_sheet_empty(sheet):
            used_range = sheet.used_range
            if used_range.shape == (1, 1) and not used_range.value:
                return True
            return False

        if 'Sheet1' in current_sheets and is_sheet_empty(wb.sheets['Sheet1']):
            wb.sheets['Sheet1'].delete()

        self.app = app
        self.workbook = workbook
        self.worksheet = sheet_name
        self.wb = wb
        self.ws = sheet
        self.globalreplace = alwaysreplace
        self.io = None

    def putxl(
            self,
            content,
            sheet_name: str = None,
            cell: str = 'A1',
            index: bool = False,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = False,
            font: str | tuple = None,
            font_name: str = None,
            font_size: int = None,
            font_color: str | tuple = None,
            italic: bool = False,
            bold: bool = False,
            underline: bool = False,
            strikeout: bool = False,
            align: str | list = None,
            merge: bool = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            check_para: bool = False,
            debug: bool = False,
    ) -> None:

        if hasattr(content, 'df'):
            content = content.df

        if not isinstance(content, str):
            for col in content.columns:
                content[col] = content[col].apply(lambda x: str(x) if isinstance(x, tuple) else x)

        from pandaspro.io.excel._xlwings import RangeOperator
        replace_type = self.globalreplace if self.globalreplace else replace

        # If a sheet_name is specified, then override the current sheet
        if sheet_name and sheet_name != self.worksheet:
            if sheet_name in [sheet.name for sheet in self.wb.sheets]:
                ws = self.wb.sheets[sheet_name]
            else:
                ws = self.wb.sheets.add(after=self.wb.sheets.count)
                ws.name = sheet_name
        else:
            ws = self.ws

        # Declare IO Object
        io = FramexlWriter(content=content, cell=cell, index=index, header=header)
        self.io = io

        # If sheetreplace or replace is specified, then delete the old sheet and create a new one
        if sheetreplace or replace_type == 'sheet':
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = ws.index
            original_name = ws.name
            total_count = self.wb.sheets.count
            if debug:
                print(f">>> Row 121: original index is {original_index}")
            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
                if debug:
                    print(f">>> Row 128: New sheet added after the sheet !'{_sheetmap[original_index]}'")
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])
                if debug:
                    print(f">>> Row 132: New sheet added before the sheet !'{_sheetmap[original_index + 1]}'")
            ws.delete()
            new_sheet.name = original_name
            ws = new_sheet
            self.ws = ws
        else:
            if not isinstance(io.content, str):
                if is_range_filled(self.ws, self.io.range_top_checker):
                    RangeOperator(ws.range(self.io.range_data)).format()
            # Add warning lines around the df if not replacing the sheet
            # io.cell.offset()

        # Export to target sheet
        ws.range(io.cell).value = io.content

        # Format the sheet
        if isinstance(io.content, str):
            RangeOperator(ws.range(io.cell)).format(
                font=font,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                italic=italic,
                bold=bold,
                underline=underline,
                strikeout=strikeout,
                align=align,
                merge=merge,
                border=border,
                fill=fill,
                fill_pattern=fill_pattern,

                fill_fg=fill_fg,
                fill_bg=fill_bg,
                check_para=check_para
            )

        # A small function to decide if sheet is blank ...
        def is_sheet_empty(sheet):
            used_range = sheet.used_range
            if used_range.shape == (1, 1) and not used_range.value:
                return True
            return False

        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if 'Sheet1' in current_sheets and is_sheet_empty(self.wb.sheets['Sheet1']):
            self.wb.sheets['Sheet1'].delete()

        self.wb.save()

        if debug:
            print(f"\n>>> Cell Range Analysis")
            print(f" ----------------------")
            print(f">>> Total row: {io.tr}, Total column: {io.tc}")
            print(
                f">>> Range index: {io.range_index}, Range header: {io.range_header}, Range index names: {io.range_indexnames}\n")

    def switchtab(self, sheet_name: str) -> None:
        """
        Switches to a specified sheet in the workbook.
        If the sheet does not exist, it creates a new one with the given name.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet to switch to or create.
        """
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if sheet_name in current_sheets:
            sheet = self.wb.sheets[sheet_name]
        else:
            sheet = self.wb.sheets.add(after=self.wb.sheets.count)
            sheet.name = sheet_name
        self.ws = sheet
        return


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

    # Re-create the initial DataFrame
    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]
    gdp = [11.43, 14.59, 12.45, 11.35, 9.05, 13.27, 9.31, 17.94, (1,2,3,4,5,6), 8.29]
    population = [1110.5, 745.2, 799.6, 1296.6, 108.7, 131.1, 38.1, 1167.3, 1091.6, 1219.3]
    df = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp,
        'Population (Millions)': population,
    })

    # Convert index and column headers to MultiIndex

    # For the index, use a combination of 'Region' and 'Country'
    regions = [(1,2), 'Asia', 'Asia', 'Europe', 'Asia', 'Europe', 'Europe', 'South America', 'Europe',
               'North America']
    index_multi = pd.MultiIndex.from_arrays([regions, df['Country']], names=['Region', 'Country'])

    # For the columns, create a MultiIndex with two levels: 'Indicator' and 'Measure'
    columns_multi = pd.MultiIndex.from_product([['Economic Indicators'], df.columns[1:]],
                                               names=['Indicator', 'Measure'])

    # Create a new DataFrame with MultiIndex for both rows and columns
    df = pd.DataFrame(df.values[:, 1:], index=index_multi, columns=columns_multi)

    ps = PutxlSet('sampledf.xlsx', 'Sheet3', noisily=True)
    ps.putxl(df, 'TT', 'A1', index=True, header=True, sheetreplace=True, debug=True)
    ps.putxl(df1, 'TF', 'A1', index=True, header=False, sheetreplace=True, debug=True)
    ps.putxl(df1, 'FT', 'A1', index=False, header=True, sheetreplace=True, debug=True)
    ps.putxl(df1, 'FF', 'A1', index=False, header=False, sheetreplace=True, debug=True)

    ps.putxl(df1, 'TF', 'A1', index=False, header=False, sheetreplace=True, debug=True)
    ps.putxl('SSSSS', 'TF', 'I2', index=False, header=False, sheetreplace=True, debug=True)
    #
    # ps.switchtab('new tab')
    # ps.putxl('A1', df1, sheetreplace=True)
    # ps.putxl('G1', df)

    # io = FramexlWriter(content=df, cell='M5', index=False, header=True)
    # ps.putxl('M5', df)
    # print(io.bottom_left_cell)
    # print(io.top_right_cell)
    # print(io.range_top_checker)
