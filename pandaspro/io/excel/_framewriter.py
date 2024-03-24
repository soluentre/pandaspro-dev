from pandaspro.core.stringfunc import parsewild
from pandaspro.io.excel._utils import CellPro
import pandas as pd



class FramexlWriter:

    def __init__(
            self,
            content,
            cell: str,
            index: bool = False,
            header: bool = True,
            column_list: list = None,
            index_mask = None
    ) -> None:
        if isinstance(content, str):
            self.content = content
            self.cell = cell
            self.tr = None
            self.tc = None
            self.range_index = None
            self.range_header = None
            self.range_indexnames = None

        else:
            cellobj = CellPro(cell)
            header_row_count = len(content.columns.levels) if isinstance(content.columns, pd.MultiIndex) else 1
            index_column_count = len(content.index.levels) if isinstance(content.index, pd.MultiIndex) else 1

            dfmapstart = cellobj.offset(header_row_count, index_column_count)
            dfmap = content.copy()

            # Create a cells Map
            i = 0
            for index, row in dfmap.iterrows():
                j = 0
                for col in dfmap.columns:
                    dfmap.loc[index, col] = dfmapstart.offset(i, j).cell
                    j += 1
                i += 1

            self.formatrange = "Please provide the column_list/index_mask to select a sub-range"
            self.valuerange
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

            self.content = export_data
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

if __name__ == '__main__':
    import pandas as pd
    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]
    gdp = [11.43, 14.59, 12.45, 11.35, 9.05, 13.27, 9.31, 17.94, (1,2,3,4,5,6), 8.29]
    population = [1110.5, 745.2, 799.6, 1296.6, 108.7, 131.1, 38.1, 1167.3, 1091.6, 1219.3]
    df = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp,
        'Population (Millions)': population,
    })

    # print(FramexlWriter(df, 'A1', index=True).range_index)

    import xlwings as xw
    ws = xw.Book('test.xlsx').sheets['FF']
    ws.range('G1').value = df
    a = FramexlWriter(df, 'G1', column_list='Country')
    print(a.formatrange)

    paintdict = {
        'all': {
            'logic': 'grade == 1 and grade <2',
            'format': {
                'fill': '#FFF000',
                'font': 'bold 12'
            }
        }
    }

