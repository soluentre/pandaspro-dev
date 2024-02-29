import pandas as pd
import pandaspro.core.tools.tab
from functools import partial
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.tab import tab


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

        ## 1. change from number of cols to number of rows
        ## 2. add the index number to the left
        ## 3. adjust the looking using css codes

        if not isinstance(self, pd.DataFrame):
            print('Please declare one dataframe')
        else:
            names = self.columns.to_list()
            num_cols = 5

            num_rows = -(-len(names) // num_cols)
            self = [['' for i in range(num_cols)] for j in range(num_rows)]

            for k, name in enumerate(names):
                row = k % num_rows
                col = k // num_rows
                self[row][col] = name
            out = pd.DataFrame(self).style.hide(axis=0).hide(axis=1).set_table_styles([
                {'selector': 'td',
                 'props': 'padding: 10px; text-align: center; font-weight: regular; background: lightgoldenrodyellow; border: 1px dotted black;'},
                {'selector': 'tr', 'props': 'width: 100% !important'},
                {'selector': 'table', 'props': 'width: 100% !important'}
            ], overwrite=True)
        return out

    def tab(self, name, d='brief', m=False, sort='index', ascending=True):
        return tab(self, name, d, m, sort, ascending)
    tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__

    def dfilter(self, input, debug):
        return dfilter(self, input, debug)
    dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    # def inlist(self):
    #     pass