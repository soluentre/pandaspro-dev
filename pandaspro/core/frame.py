import pandas as pd
import pandaspro
from functools import partial

from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.varnames import varnames
from pandaspro.core.tools.inlist import inlist
from pandaspro.io.excel._utils import lowervarlist
from pandaspro.io.excel._putexcel import PutxlSet


__pp_default_export_global_wb = None


def excel_d(file: str = 'Exported Results.xlsx', noisily = None):
    global __pp_default_export_global_wb
    __pp_default_export_global_wb = PutxlSet(file, noisily=noisily)


def return_same_type_decor(func):
    def wrapper(self, *args, **kwargs):
        result = func(*args, **kwargs)
        if isinstance(result, pd.DataFrame):
            return self.__class__(data=result)
        return result

    return wrapper


class FramePro(pd.DataFrame):
    def __init__(self, data=None, *args, **kwargs):
        super().__init__(data, *args, **kwargs)
        self.namemap = "This attribute displays the original names when importing data using 'readpro' method in io.excel._base module, and currently is not activated"

        for attr_name in dir(pd.DataFrame):
            if callable(getattr(pd.DataFrame, attr_name, None)) and not attr_name.startswith("_"):
                if attr_name not in ['sparse']:
                    try:
                        setattr(self, attr_name, partial(return_same_type_decor(getattr(self, attr_name)), self))
                    except AttributeError:
                        pass

    @property
    def d(self):
        return pd.DataFrame(self)

    @property
    def varnames(self):
        return varnames(self)

    def tab(self, name: str, d: str = 'brief', m: bool = False, sort: str = 'index', ascending: bool = True):
        return FramePro(tab(self, name, d, m, sort, ascending))

    def dfilter(self, inputdict: dict = None, debug: bool = False):
        return FramePro(dfilter(self, inputdict, debug))

    def inlist(self, colname: str, *args, engine: str = 'b', inplace: bool = False, invert: bool = False, debug: bool = False):
        result = inlist(
            self,
            colname,
            *args,
            engine=engine,
            inplace=inplace,
            invert=invert,
            debug=debug,
        )
        if debug:
            print(type(result))
        if engine == 'm':
            return result
        else:
            return FramePro(result)

    def lowervarlist(self, engine='columns', inplace=False):
        if engine == 'data':
            return FramePro(lowervarlist(self, engine, inplace=inplace))
        return lowervarlist(self, engine, inplace=inplace)

    def merge(*args, **kwargs):
        result = super().merge(*args, **kwargs, indicator=True)
        print(result.tab('_merge'))
        return FramePro(result)

    def excel_e(
            self,
            sheet_name: str = 'Sheet1',
            start_cell: str = 'A1',
            index: bool = False,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = False,
            format: str | list = None,
            rowformat: dict = None,
            colformat: dict = None,
            override: bool = None,
    ):
        wb = __pp_default_export_global_wb
        wb.putxl(
            frame=self,
            sheet_name=sheet_name,
            start_cell=start_cell,
            index=index,
            header=header,
            replace=replace,
            sheetreplace=sheetreplace
        )

        if override:
            return wb

    tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__
    dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    inlist.__doc__ = pandaspro.core.tools.inlist.__doc__
    varnames.__doc__ = pandaspro.core.tools.varnames.varnames.__doc__
    lowervarlist.__doc__ = pandaspro.io.excel._utils.lowervarlist.__doc__

    # Overwriting original methods
    def rename(self, columns=None, *args, **kwargs):
        return FramePro(super().rename(columns=columns, *args, **kwargs))


if __name__ == '__main__':
    a = FramePro({'a': [1,2,3],'b':[3,4,5]})
    b = a.inlist('a',1, engine = 'm')
    e = a.inlist('a', 1, engine='b')