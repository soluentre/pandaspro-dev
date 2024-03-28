from pandaspro.core.stringfunc import parsewild
from pandaspro.io.excel._utils import CellPro
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
            column_list: list = None,
            cols_index_merge: str | list = None,
            index_mask=None
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

        if column_list:
            if isinstance(column_list, str):
                column_list = parsewild(column_list, dfmap.columns)
            self.formatrange = dfmap[column_list]
        if index_mask:
            self.formatrange = self.formatrange[index_mask]

        # Calculate the Ranges
        if header == True and index == True:
            tr, tc = content.shape[0] + header_row_count, content.shape[1] + index_column_count
            export_data = content
            range_index = cellobj.offset(header_row_count, 0).resize(tr - header_row_count, index_column_count)
            range_indexnames = cellobj.resize(header_row_count, header_row_count)
            range_header = cellobj.offset(0, index_column_count).resize(header_row_count, tc - index_column_count)
        elif header == False and index == True:
            tr, tc = content.shape[0], content.shape[1] + index_column_count
            export_data = content.reset_index().to_numpy().tolist()
            range_index = cellobj.resize(tr, index_column_count)
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            tr, tc = content.shape[0], content.shape[1]
            export_data = content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            tr, tc = content.shape[0] + header_row_count, content.shape[1]
            if isinstance(content.columns, pd.MultiIndex):
                column_export = [list(lst) for lst in list(zip(*content.columns.values))]
            else:
                column_export = [content.columns.to_list()]
            export_data = column_export + content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cellobj.resize(header_row_count, tc)

        self.iotype = 'df'
        self.rawdata = content
        self.columns = self.rawdata.columns
        self.content = pd.DataFrame(export_data)
        self.cell = cell
        self.tr = tr
        self.tc = tc
        self.start_cell = cellobj.offset(header_row_count, index_column_count)
        self.top_right_cell = cellobj.offset(0, self.tc - 1).cell
        self.bottom_left_cell = cellobj.offset(self.tr - 1, 0).cell
        self.end_cell = cellobj.offset(self.tr - 1, self.tc - 1).cell
        self.range_all = cell + ':' + self.end_cell
        self.range_data = self.start_cell.resize(tr - header_row_count, tc - index_column_count).cell
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_indexnames = range_indexnames.cell if range_indexnames != 'N/A' else 'N/A'
        self.range_top_checker = CellPro(self.cell).offset(-1, 0).resize(1, self.tc).cell if CellPro(self.cell).index_cell()[0] != 1 else None
        self.cellmap = dfmap
        if cols_index_merge:
            self.cols_index_merge = cols_index_merge if isinstance(cols_index_merge, list) else parsewild(cols_index_merge, content)
        else:
            self.cols_index_merge = None

    def get_column_letter_by_name(self, colname):
        rowcount = list(self.columns).index(colname)
        col_cell = self.start_cell.offset(0, rowcount)
        return col_cell.cell

    def _index_break(self, level: str = None):
        temp = self.content.reset_index()

        def _count_consecutive_values(series):
            return series.groupby((series != series.shift()).cumsum()).size().tolist()

        return _count_consecutive_values(temp[level])

    def index_merge_inputs(self, level: str = None):
        result_dict = {}
        if self.cols_index_merge is None:
            raise ValueError('index_merge_inputs method requires cols_index_merge to be passed when constructing the FramexlWriter Object')
        else:
            for index, col in self.cols_index_merge:
                for rowspan in self._index_break(level=level):
                    result_dict[f'col{index}_{rowspan}'] = 1


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
    from pandaspro import sysuse_auto
    ws = xw.Book('sampledf.xlsx').sheets['sob']
    ws.range('G1').value = pd.DataFrame(sysuse_auto)

    ws = xw.Book('sampledf.xlsx').sheets['sob']
    d = wb.sob(region='AFE').pivot_table(index=['cmu_dept_major', 'cmu_dept'], values=['upi','age'], aggfunc='sum', margins_name='Total', margins=True)
    ws.range('G1').value = pd.DataFrame(d)
    a = FramexlWriter(d, 'G1', index=True, cols_index_merge='upi, age').get_column_letter_by_name('age')

    # xw.apps.active.api.DisplayAlerts = False
    # ws.range('A4:A9').api.MergeCells = True