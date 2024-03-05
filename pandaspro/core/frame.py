import pandas as pd
import pandaspro
from functools import partial
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.varnames import varnames
from pandaspro.core.tools.inlist import inlist
from pandaspro.io.excel._base import lowervarlist


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
    def data(self):
        return pd.DataFrame(self)

    @property
    def varnames(self):
        return varnames(self)

    def tab(self, name: str, d: str = 'brief', m: bool = False, sort: str = 'index', ascending: bool = True):
        return tab(self, name, d, m, sort, ascending)

    def dfilter(self, inputdict: dict = None, debug: bool = False):
        return dfilter(self, inputdict, debug)

    def inlist(self, colname: str, *args, engine: str = 'b', inplace: bool = False, invert: bool = False, debug: bool = False):
        return inlist(
            data=self,
            colname=colname,
            *args,
            engine=engine,
            inplace=inplace,
            invert=invert,
            debug=debug
        )

    def lowervarlist(self, engine='columns', inplace=False):
        return lowervarlist(self, engine, inplace=inplace)

    def merge(*args, **kwargs):
        result = super().merge(*args, **kwargs, indicator=True)
        print(result.tab('_merge'))
        return result

    tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__
    dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    inlist.__doc__ = pandaspro.core.tools.inlist.__doc__
    varnames.__doc__ = pandaspro.core.tools.varnames.varnames.__doc__
    lowervarlist.__doc__ = pandaspro.io.excel._base.lowervarlist.__doc__


if __name__ == '__main__':
    pass
