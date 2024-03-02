from pandas import DataFrame
from pandaspro import FramePro
import pandas as pd

class PutxlSet:

    pass


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
            export_data = frame
            paintstart = _cell_to_the_right(start_cell, index_column_count)
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1] + index_column_count
        elif header == True and index == False:
            export_data = [frame.columns.tolist()] + frame.to_numpy().tolist()
            paintstart = start_cell
            tr, tc = frame.shape[0] + header_row_count, frame.shape[1]
        elif header == False and index == True:
            export_data = frame.reset_index().to_numpy().tolist()
            paintstart = _cell_to_the_right(start_cell, index_column_count)
            tr, tc = frame.shape[0], frame.shape[1] + index_column_count
        elif header == False and index == False:
            export_data = frame.to_numpy().tolist()
            paintstart = start_cell
            tr, tc = frame.shape[0], frame.shape[1]