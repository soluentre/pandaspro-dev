import maya
import pandas as pd
from datetime import datetime


class DatePro:
    def __init__(self, date):
        self.original_date = date
        if isinstance(date, pd.Timestamp):
            date_maya = maya.parse(str(date))
            self.datetype = 'pd.Timestamp'
        elif isinstance(date, datetime):
            date_maya = maya.MayaDT.from_datetime(date)
            self.datetype = 'datetime'
        elif isinstance(date, str):
            datestr = maya.parse(date)
            self.datetype = 'str'
        else:
            raise ValueError('Invalid type for date passed, only support [pd.Timestamp, str] objects for this version')

        self.datestr

