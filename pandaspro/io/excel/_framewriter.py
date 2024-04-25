from pandaspro.core.stringfunc import parse_wild
from pandaspro.io.excel._cdformat import CdFormat
from pandaspro.core.tools.utils import df_with_index_for_mask
from pandaspro.io.excel._utils import CellPro, index_cell
import pandas as pd


class StringxlWriter:
    def __init__(
            self,
            content: str = None,
            cell: str = None,
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
    ) -> None:
        cellobj = CellPro(cell)
        header_row_count = len(content.columns.levels) if isinstance(content.columns, pd.MultiIndex) else 1
        index_column_count = len(content.index.levels) if isinstance(content.index, pd.MultiIndex) else 1

        # Calculate the Ranges
        self.rawdata = content
        content = pd.DataFrame(content)
        if header == True and index == True:
            self.export_type = 'htit'
            tr, tc = content.shape[0] + header_row_count, content.shape[1] + index_column_count
            xl_header_count, xl_index_count = header_row_count, index_column_count
            export_data = content
            range_index = cellobj.offset(header_row_count, 0).resize(tr - header_row_count, index_column_count)
            range_indexnames = cellobj.resize(header_row_count, header_row_count)
            range_header = cellobj.offset(0, index_column_count).resize(header_row_count, tc - index_column_count)
        elif header == False and index == True:
            self.export_type = 'hfit'
            tr, tc = content.shape[0], content.shape[1] + index_column_count
            xl_header_count, xl_index_count = 0, index_column_count
            export_data = content.reset_index().to_numpy().tolist()
            range_index = cellobj.resize(tr, index_column_count)
            range_indexnames = 'N/A'
            range_header = 'N/A'
        elif header == False and index == False:
            self.export_type = 'hfif'
            tr, tc = content.shape[0], content.shape[1]
            xl_header_count, xl_index_count = 0, 0
            export_data = content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = 'N/A'
        else:
            self.export_type = 'htif'
            tr, tc = content.shape[0] + header_row_count, content.shape[1]
            xl_header_count, xl_index_count = header_row_count, 0
            if isinstance(content.columns, pd.MultiIndex):
                column_export = [list(lst) for lst in list(zip(*content.columns.values))]
            else:
                column_export = [content.columns.to_list()]
            # noinspection PyTypeChecker
            export_data = column_export + content.to_numpy().tolist()
            range_index = 'N/A'
            range_indexnames = 'N/A'
            range_header = cellobj.resize(header_row_count, tc)

        # Calculate the Map
        dfmapstart = cellobj.offset(xl_header_count, 0)
        dfmap = df_with_index_for_mask(self.rawdata).copy()
        dfmap = dfmap.astype(str)

        for dfmap_index in range(len(dfmap)):
            for j, col in enumerate(dfmap.columns):
                dfmap.iloc[dfmap_index, j] = dfmapstart.offset(dfmap_index, j).cell

        self.iotype = 'df'
        self.columns_with_indexnames = self.rawdata.reset_index().columns
        self.columns = self.rawdata.columns
        self.content = export_data
        self.cell = cell
        self.index_bool = index
        self.header_bool = header
        self.tr = tr
        self.tc = tc
        self.header_row_count = header_row_count
        self.index_column_count = index_column_count

        # data corners - cellpros
        self.start_cellobj = cellobj.offset(xl_header_count, xl_index_count)
        self.start_cell = self.start_cellobj.cell
        self.top_right_cell = cellobj.offset(0, self.tc - 1).cell
        self.bottom_left_cell = cellobj.offset(self.tr - 1, 0).cell
        self.end_cell = cellobj.offset(self.tr - 1, self.tc - 1).cell

        # ranges
        self.range_all = cell + ':' + self.end_cell
        self.range_data = self.start_cellobj.resize(tr - header_row_count, tc - index_column_count).cell
        self.range_index = range_index.cell if range_index != 'N/A' else 'N/A'
        self.range_index_outer = CellPro(self.cell).resize(self.tr, self.index_column_count).cell
        self.range_header = range_header.cell if range_header != 'N/A' else 'N/A'
        self.range_header_outer = CellPro(self.cell).resize(self.header_row_count, self.tc).cell
        self.range_indexnames = range_indexnames.cell if range_indexnames != 'N/A' else 'N/A'

        # format relevant
        self.dfmap = dfmap
        self.cols_index_merge = None

        # Conditional Formatting
        self.cd_dfmap_1col = None
        self.cd_cellrange_1col = None

        # Special - Checker for sheetreplace
        self.range_top_empty_checker = CellPro(self.cell).offset(-1, 0).resize(1, self.tc).cell if CellPro(self.cell).cell_index[0] != 1 else None
        self.range_bottom_empty_checker = CellPro(self.bottom_left_cell).offset(1, 0).resize(1, self.tc).cell if CellPro(self.bottom_left_cell).cell_index[0] != 1 else None
        self.range_left_empty_checker = CellPro(self.cell).offset(0, -1).resize(self.tr, 1).cell if CellPro(self.cell).cell_index[1] != 1 else None
        self.range_right_empty_checker = CellPro(self.top_right_cell).offset(0, 1).resize(self.tr, 1).cell if CellPro(self.top_right_cell).cell_index[0] != 1 else None

    def get_column_letter_by_indexname(self, levelname):
        if not self.index_bool:
            raise ValueError('Cannot return a range with get_column_letter_by_indexname method when index = False is specified')

        col_count = list(self.rawdata.index.names).index(levelname)
        col_cell = CellPro(self.cell).offset(self.header_row_count, col_count)
        return col_cell

    def get_column_letter_by_name(self, colname):
        col_count = list(self.columns).index(colname)
        col_cell = self.start_cellobj.offset(0, col_count)
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
    ) -> dict:
        result_dict = {}

        # Index Column
        merge_start_index = self.get_column_letter_by_indexname(level)
        for localid, rowspan in enumerate(self._index_break(level=level)):
            result_dict[f'indexlevel_{localid}_{rowspan}'] = merge_start_index.resize(rowspan, 1).cell
            merge_start_index = merge_start_index.offset(rowspan, 0)

        # Selected Columns
        if columns:
            self.cols_index_merge = columns if isinstance(columns, list) else parse_wild(columns, self.columns)
            # print("framewriter cols_index_merge:", self.cols_index_merge)
            # print("columns:", columns, self.columns)
            for index, col in enumerate(self.cols_index_merge):
                merge_start_each = self.get_column_letter_by_name(col)
                for localid, rowspan in enumerate(self._index_break(level=level)):
                    result_dict[f'col{index}_{localid}_{rowspan}'] = merge_start_each.resize(rowspan, 1).cell
                    merge_start_each = merge_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_hsections(self, level: str = None) -> dict:
        if self.range_index is None:
            raise ValueError('index_sections method requires the input dataframe to have an index')
        else:
            result_dict = {'headers': CellPro(self.cell).resize(self.header_row_count, self.tc).cell}
            range_start_each = CellPro(self.cell).offset(self.header_row_count, 0)
            for localid, rowspan in enumerate(self._index_break(level=level)):
                result_dict[f'section_{localid}_{rowspan}'] = range_start_each.resize(rowspan, self.tc).cell
                range_start_each = range_start_each.offset(rowspan, 0)

        return result_dict

    def range_index_selected_hsection(self, level: str = None, token: str = 'Total') -> str:
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
        result = self.get_column_letter_by_indexname(level).offset(go_down_by, 0).resize(local_height, self.tc).cell

        return result

    ''' this is returning the whole level by level ranges in selection '''
    @property
    def range_index_levels(self) -> dict:
        result_dict = {}
        range_start_each = CellPro(self.cell)
        for each_index in self.rawdata.index.names:
            result_dict[f'index_{each_index}'] = range_start_each.resize(self.tr, 1).cell
            range_start_each = range_start_each.offset(0, 1)
        return result_dict

    def range_columns(self, c, header = False):
        if isinstance(c, str):
            clean_list = parse_wild(c, self.columns_with_indexnames)
        elif isinstance(c, list):
            clean_list = c
        else:
            raise ValueError('range_columns only accept str/list as inputs')

        result_list = []
        for colname in clean_list:
            if colname in self.columns:
                start_range = self.get_column_letter_by_name(colname)
            elif colname in self.rawdata.index.names:
                start_range = self.get_column_letter_by_indexname(colname)
            else:
                raise ValueError(f'Searching name <<{colname}>> is not in column nor index.names')

            below_range = start_range.resize_h(self.tr - self.header_row_count).cell

            # noinspection PySimplifyBooleanCheck
            if header == True:
                below_range = CellPro(below_range).offset(-self.header_row_count, 0).resize_h(self.tr).cell
            if header == 'only':
                below_range = CellPro(below_range).offset(-self.header_row_count, 0).resize_h(self.header_row_count).cell
            result_list.append(below_range)

        return ', '.join(result_list)

    def range_cspan(self, s = None, e = None, c = None, header = False):
        # Declaring starting and ending columns
        if s and e:
            col_index1 = self.get_column_letter_by_name(s).cell_index[1]
            col_index2 = self.get_column_letter_by_name(e).cell_index[1]
            row_index = self.get_column_letter_by_name(s).cell_index[0]

            # Decide the top row cells with min/max - allow invert orders
            top_left_index = min(col_index1, col_index2)
            top_right_index = max(col_index1, col_index2)
            top_left = index_cell(row_index, top_left_index)
            top_right = index_cell(row_index, top_right_index)
            start_range = CellPro(top_left + ':' + top_right)

        # Declaring only 1 column
        elif c:  # Para C: declare column only
            selected_column = self.get_column_letter_by_name(c)
            start_range = selected_column

        else:
            raise ValueError('At least 1 set of Paras: (1) s+e or (2) c must be declared ')

        final = start_range.resize_h(self.tr - self.header_row_count).cell
        # noinspection PySimplifyBooleanCheck
        if header == True:
            final = CellPro(final).offset(-self.header_row_count, 0).resize_h(self.tr).cell
        if header == 'only':
            final = CellPro(final).offset(-self.header_row_count, 0).resize_h(self.header_row_count).cell

        return final

    def range_cdformat(
            self,
            column,
            rules = None,
            applyto = 'self'
    ):
        mycd = CdFormat(
            df=self.rawdata,
            column=column,
            cd_rules=rules,
            applyto=applyto
        )
        # print(mycd.df.columns, mycd.df_with_index.columns, mycd.column)
        if mycd.col_not_exist:
            cd_cellrange_1col = {'void_rule': {'cellrange': 'no cells', 'format': ''}}
        else:
            apply_columns = mycd.apply
            cd_dfmap_1col = {}
            for key, mask_rule in mycd.rules_mask.items():
                cd_dfmap_1col[key] = {}
                cd_dfmap_1col[key]['dfmap'] = self.dfmap[mask_rule['mask']][apply_columns]
                cd_dfmap_1col[key]['format'] = mask_rule['format']
            self.cd_dfmap_1col = cd_dfmap_1col

            def _df_to_mystring(df):
                lcarray = df.values.flatten()
                long_string = ','.join([str(value) for value in lcarray])
                return long_string

            cd_cellrange_1col = {}
            for key, mask_rule in mycd.rules_mask.items():
                cd_cellrange_1col[key] = {}
                temp_dfmap = self.dfmap[mask_rule['mask']][apply_columns]
                cd_cellrange_1col[key]['cellrange'] = _df_to_mystring(temp_dfmap)
                cd_cellrange_1col[key]['format'] = mask_rule['format']
            self.cd_cellrange_1col = cd_cellrange_1col

        '''
        should be something like ...
        {
            "AFWDE": {
                "cellrange": "B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2", 
                "format": "blue"
            },
            "AFWVP": {
                "cellrange": "B3,C3,D3,E3,F3,G3,H3,I3,J3,K3,L3,M3", 
                "format": "orange"
            },
        }
        '''
        return cd_cellrange_1col


class cpdFramexl:
    def __init__(self, name, **kwargs):
        self.name = name
        self.paras = kwargs
