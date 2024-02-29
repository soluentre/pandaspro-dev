import pandas as pd
import pandaspro
from functools import partial
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.varnames import varnames

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

        for attr_name in dir(pd.DataFrame):
            if callable(getattr(pd.DataFrame, attr_name, None)) and not attr_name.startswith("_") and attr_name not in ['sparse']:
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

    def tab(self, name, d='brief', m=False, sort='index', ascending=True):
        return tab(self, name, d, m, sort, ascending)

    def dfilter(self, input, debug):
        return dfilter(self, input, debug)

    tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__
    dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    varnames.__doc__ = pandaspro.core.tools.varnames.varnames.__doc__
    # def inlist(self):
    #     pass