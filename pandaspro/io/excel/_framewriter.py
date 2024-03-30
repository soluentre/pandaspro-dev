from pandaspro.core.stringfunc import parsewild
from pandaspro.io.excel._utils import CellPro, index_to_cell
import pandas as pd


class StringxlWriter:
    def __init__(
            self,
            content,
            cell: str,
    ) -> None:
        self.iotype = 'str'
        self.content = content
        self.cell = cell


class FramexlWriter:
    def __init__(
            self,
            content,
            cell: str,
            index: bool = False,
            header: bool = True,
            # column_list: list = None,
            # index_mask = None
    ) -> None:
        cellobj = CellPro(cell)
        header_row_count = len(content.columns.levels) if isinstance(content.columns, pd.MultiIndex) else 1
        index_column_count = len(content.index.levels) if isinstance(content.index, pd.MultiIndex) else 1

        dfmapstart = cellobj.offset(header_row_count, index_column_count)
        dfmap = content.copy()
        dfmap = dfmap.astype(str)

        # Create a cells Map
        i = 0
        for dfmap_index, row in dfmap.iterrows():
            j = 0
            for col in dfmap.columns:
                dfmap.loc[dfmap_index, col] = dfmapstart.offset(i, j).cell
                j += 1
            i += 1

        self.formatrange = "Please provide the column_list/index_mask to select a sub-range"

        # if column_list:
        #     if isinstance(column_list, str):
        #         column_list = parsewild(column_list, dfmap.columns)
        #     self.formatrange = dfmap[column_list]
        # if index_mask:
        #     self.formatrange = self.formatrange[index_mask]

        # Calculate the Ranges
        self.rawdata = content
        content = pd.DataFrame(content.to_dict())
        if header == True and index == True:
            self.export_type = 'htit'
            tr, tc = content.shape[0] + header_row_count, content.shape[1] + index_column_count
            export_data = content
            range_index = cellobj.offset(header_row_count, 0).resize(tr - header_row_count, index_column_count)
            range_indexnames = cellobj.resize(header_row_count, header_row_count)
            range_header = cellobj.offset(0, index_column_count).resize(header_row_count, tc - index_column_count)
        elif header == False and index == True:
            self.export_type = 'hfit'
            tr, tc = content.shape[0], content.shape[1] + index_column_count
            export_data = content.reset_index().to_numpy().tolist()
            range_index = cellobj.resize(tr, index_column_count)
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            self.export_type = 'hfif'
            tr, tc = content.shape[0], content.shape[1]
            export_data = content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            self.export_type = 'htif'
            tr, tc = content.shape[0] + header_row_count, content.shape[1]
            if isinstance(content.columns, pd.MultiIndex):
                column_export = [list(lst) for lst in list(zip(*content.columns.values))]
            else:
                column_export = [content.columns.to_list()]
            # noinspection PyTypeChecker
            export_data = column_export + content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cellobj.resize(header_row_count, tc)

        self.iotype = 'df'
        self.columns = self.rawdata.columns
        self.content = export_data
        self.cell = cell
        self.tr = tr
        self.tc = tc
        self.header_row_count = header_row_count
        self.index_column_count = index_column_count

        # data corners - cellpros
        self.start_cell = cellobj.offset(header_row_count, index_column_count)
        self.top_right_cell = cellobj.offset(0, self.tc - 1).cell
        self.bottom_left_cell = cellobj.offset(self.tr - 1, 0).cell
        self.end_cell = cellobj.offset(self.tr - 1, self.tc - 1).cell

        # ranges
        self.range_all = cell + ':' + self.end_cell
        self.range_data = self.start_cell.resize(tr - header_row_count, tc - index_column_count).cell
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_index_outer = CellPro(self.cell).resize(self.tr, self.index_column_count).cell
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_header_outer = CellPro(self.cell).resize(self.header_row_count, self.tc).cell
        self.range_indexnames = range_indexnames.cell if range_indexnames != 'N/A' else 'N/A'

        # format relevant
        self.cellmap = dfmap
        self.cols_index_merge = None

        # Special - Checker for sheetreplace
        self.range_top_empty_checker = CellPro(self.cell).offset(-1, 0).resize(1, self.tc).cell if CellPro(self.cell).cell_index[0] != 1 else None
        self.range_bottom_empty_checker = CellPro(self.bottom_left_cell).offset(1, 0).resize(1, self.tc).cell if CellPro(self.bottom_left_cell).cell_index[0] != 1 else None
        self.range_left_empty_checker = CellPro(self.cell).offset(0, -1).resize(self.tr, 1).cell if CellPro(self.cell).cell_index[0] != 1 else None
        self.range_right_empty_checker = CellPro(self.top_right_cell).offset(0, 1).resize(self.tr, 1).cell if CellPro(self.top_right_cell).cell_index[0] != 1 else None

    def _get_column_letter_by_indexname(self, levelname):
        col_count = list(self.rawdata.index.names).index(levelname)
        col_cell = CellPro(self.cell).offset(self.header_row_count, col_count)
        return col_cell

    def _get_column_letter_by_name(self, colname):
        col_count = list(self.columns).index(colname)
        col_cell = self.start_cell.offset(0, col_count)

        if self.export_type in ['htif', 'hfif']:
            col_cell = col_cell.offset(0, -self.index_column_count)

        return col_cell

    def _index_break(self, level: str = None):
        temp = self.rawdata.reset_index()

        def _count_consecutive_values(series):
            return series.groupby((series != series.shift()).cumsum()).size().tolist()

        return _count_consecutive_values(temp[level])

    def range_index_merge_inputs(
            self,
            level: str = None,
            columns: str | list = None
    ):
        result_dict = {}

        # Index Column
        merge_start_index = self._get_column_letter_by_indexname(level)
        for localid, rowspan in enumerate(self._index_break(level=level)):
            result_dict[f'indexlevel_{localid}_{rowspan}'] = merge_start_index.resize(rowspan, 1).cell
            merge_start_index = merge_start_index.offset(rowspan, 0)

        # Selected Columns
        if columns:
            self.cols_index_merge = columns if isinstance(columns, list) else parsewild(columns, self.columns)
            for index, col in enumerate(self.cols_index_merge):
                merge_start_each = self._get_column_letter_by_name(col)
                for localid, rowspan in enumerate(self._index_break(level=level)):
                    result_dict[f'col{index}_{localid}_{rowspan}'] = merge_start_each.resize(rowspan, 1).cell
                    merge_start_each = merge_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_hsections(self, level: str = None):
        if self.range_index is None:
            raise ValueError('index_sections method requires the input dataframe to have an index')
        else:
            result_dict = {'headers': CellPro(self.cell).resize(self.header_row_count, self.tc).cell}
            range_start_each = CellPro(self.cell).offset(self.header_row_count, 0)
            for localid, rowspan in enumerate(self._index_break(level=level)):
                result_dict[f'section_{localid}_{rowspan}'] = range_start_each.resize(rowspan, self.tc).cell
                range_start_each = range_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_selected_hsection(self, level: str = None, token: str = 'Total'):
        temp = self.rawdata.reset_index()

        def _find_occurrence_details(series, indexname):
            """
            This function finds the first occurrence of a specified token in a pandas Series,
            returns the index of its first appearance, and the count of its consecutive occurrences.
            """
            if indexname in series.values:
                first_occurrence_index = series[series == indexname].index[0]
                # Count the consecutive occurrences starting from the first occurrence index
                count = 1  # Start with 1 for the first occurrence
                for i in range(first_occurrence_index + 1, len(series)):
                    if series.iloc[i] == indexname:
                        count += 1
                    else:
                        break
                return first_occurrence_index, count
            else:
                return None, 0

        go_down_by, local_height = _find_occurrence_details(temp[level], token)
        result = self._get_column_letter_by_indexname(level).offset(go_down_by, 0).resize(local_height, self.tc).cell

        return result

    @property
    def range_index_levels(self):
        if self.range_index is None or not isinstance(self.rawdata.index, pd.MultiIndex):
            raise ValueError('index_levels method requires the input dataframe to be multi-index frame')
        else:
            result_dict = {}
            range_start_each = CellPro(self.cell)
            for each_index in self.rawdata.index.names:
                result_dict[f'index_{each_index}'] = range_start_each.resize(self.tr, 1).cell
                range_start_each = range_start_each.offset(0, 1)
            return result_dict

    def range_columnspan(self, start_col, stop_col):
        # Get the col indices and row index
        col_index1 = self._get_column_letter_by_name(start_col).cell_index[1]
        col_index2 = self._get_column_letter_by_name(stop_col).cell_index[1]
        row_index = self._get_column_letter_by_name(start_col).cell_index[0]

        # decide the top row cells with min/max - allow invert orders
        top_left_index = min(col_index1, col_index2)
        top_right_index = max(col_index1, col_index2)
        top_left = index_to_cell(row_index, top_left_index)
        top_right = index_to_cell(row_index, top_right_index)

        # Combine Range
        start_range = CellPro(top_left + ':' + top_right)

        return start_range.resize_h(self.tr - self.header_row_count).cell

    # def range_cdformat(self, colname, rules, applyto):
    #     a = CdFormat(self.rawdata, colname, rules, applyto)
    #     cdstart = self.start_cell
    #     print(cdstart)
    #     result = {
    #         '1': 'J3, J4, J7, J8, J9'
    #     }
    #     return result


if __name__ == '__main__':

    # a = FramexlWriter(sysuse_auto, 'G1', column_list='Country', index=True)
    # print(a.formatrange)
    #
    # paintdict = {
    #     'all': {
    #         'logic': 'grade == 1 and grade <2',
    #         'format': {
    #             'fill': '#FFF000',
    #             'font': 'bold 12'
    #         }
    #     }
    # }

    import wbhrdata as wb
    import xlwings as xw
    # from pandaspro import sysuse_auto

    ws = xw.Book('sampledf.xlsx').sheets['sob']
    data = wb.sob(region='AFE').pivot_table(index=['cmu_dept_major', 'cmu_dept'], values=['upi', 'age', 'yrs_in_assign', 'yrs_in_grade'], aggfunc='sum', margins_name='Total', margins=True)

    # core
    io = FramexlWriter(data, 'G1', index=True)
    ws.range('G1').value = io.content
    # , cols_index_merge = 'upi, age'

    # ws.range('G2, G3, G7:L10').font.color = '#FFF001'
    a1 = io.range_index_hsections(level='cmu_dept_major')
    a2 = io.range_index_outer
    a3 = io.range_index_levels
    a4 = io.range_columnspan('age', 'yrs_in_assign')
    a5 = io.range_index_merge_inputs('cmu_dept_major')
    a6 = io.range_index_selected_hsection('cmu_dept_major', 'PGs')

    # mpg_col = io.get_column_letter_by_name('mpg').cell
    # xw.apps.active.api.DisplayAlerts = False
    # ws.range('A4:A9').api.MergeCells = True
