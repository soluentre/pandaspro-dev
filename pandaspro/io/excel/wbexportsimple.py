import xlwings as xw
from pandaspro.io.excel._putexcel import PutxlSet
from datetime import datetime


class WorkbookExportSimplifier:
    last_declared_wb = None

    @classmethod
    def declare_workbook(
            cls,
            file: str = f'sw_Export_Default_Template_{datetime.now().strftime("%b %d, %Y")}.xlsx',
            sheet_name: str = 'Sheet1',
            alwaysreplace: str = None,
            noisily: bool = None
    ):
        setworkbook = PutxlSet(
            workbook=file,
            sheet_name=sheet_name,
            alwaysreplace=alwaysreplace,
            noisily=noisily
        )
        cls.last_declared_wb = setworkbook
        print(f"Declared workbook: {setworkbook.workbook}")

    @classmethod
    def get_last_declared_workbook(cls):
        if cls.last_declared_wb is None:
            raise ValueError("No workbook has been declared.")
        return cls.last_declared_wb