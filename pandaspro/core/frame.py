import numpy as np
import pandas as pd
import pandaspro
from functools import partial

from pandaspro.core.stringfunc import parsewild
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.varnames import varnames
from pandaspro.core.tools.inlist import inlist
from pandaspro.io.excel._utils import lowervarlist
from pandaspro.io.excel._putexcel import PutxlSet
from pandaspro.io.excel.wbexportsimple import WorkbookExportSimplifier


class FramePro(pd.DataFrame):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.namemap = "This attribute displays the original names when importing data using 'readpro' method in io.excel._base module, and currently is not activated"

        # for attr_name in dir(pd.DataFrame):
        #     if callable(getattr(pd.DataFrame, attr_name, None)) and not attr_name.startswith("_"):
        #         if attr_name not in ['sparse']:
        #             try:
        #                 setattr(self, attr_name, partial(return_same_type_decor(getattr(self, attr_name)), self))
        #             except AttributeError:
        #                 pass

    @property
    def _constructor(self):
        return FramePro

    @property
    def df(self):
        return pd.DataFrame(self)

    @property
    def varnames(self):
        return varnames(self)

    def tab(self, name: str, d: str = 'brief', m: bool = False, sort: str = 'index', ascending: bool = True):
        return self._constructor(tab(self, name, d, m, sort, ascending))

    def dfilter(self, inputdict: dict = None, debug: bool = False):
        return self._constructor(dfilter(self, inputdict, debug))

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
            return self._constructor(result)

    def lowervarlist(self, engine='columns', inplace=False):
        if engine == 'data':
            return self._constructor(lowervarlist(self, engine, inplace=inplace))
        return lowervarlist(self, engine, inplace=inplace)

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
        declaredwb = WorkbookExportSimplifier.get_last_declared_workbook()
        declaredwb.putxl(
            frame=self.df,
            sheet_name=sheet_name,
            start_cell=start_cell,
            index=index,
            header=header,
            replace=replace,
            sheetreplace=sheetreplace
        )

        if override:
            return declaredwb

    def cvar(self, promptstring):
        if self.empty:
            print('Nothing to check/browse in an empty dataframe')
            return []
        else:
            return parsewild(promptstring, self.columns)

    def br(self, promptstring):
        if self.cvar(promptstring) == []:
            print('Nothing to check/browse in an empty dataframe')
            return self
        else:
            return self[self.cvar(promptstring)]

    def insert_blank(self, locator_dict=None, how='after', num_rows=1):
        if locator_dict is None:
            if how == 'first':
                insert_positions = [0] * num_rows
            elif how == 'last':
                insert_positions = [self.index[-1] + i + 1 for i in range(num_rows)]
            else:
                print("Invalid 'how' parameter value when 'locator_dict' is None.")
                return self
        else:
            condition = pd.Series([True] * len(self))
            for col, value in locator_dict.items():
                if col in self.columns:
                    condition &= (self[col] == value)
                else:
                    print(f"Column '{col}' does not exist in the Frame.")
                    return self

            indices = self.index[condition].tolist()
            insert_positions = []
            for i in indices:
                if how == 'before':
                    insert_positions.extend([i - j - 1 for j in range(num_rows)])
                else:  # 'after'
                    insert_positions.extend([i + j + 1 for j in range(num_rows)])

        blank_rows = pd.DataFrame(np.nan, index=insert_positions, columns=self.columns)
        result = self._constructor(pd.concat([self, blank_rows]).reset_index(drop=True))

        return result

    def add_total(
            self,
            total_label_column,
            label: str = 'Total',
            sum_columns: str = '_all'
    ):
        total_row = {col: np.nan for col in self.columns}
        total_row[total_label_column] = label

        if sum_columns == '_all':
            sum_columns = self.select_dtypes(include=[np.number]).columns.tolist()
        elif isinstance(sum_columns, (str, int)):
            sum_columns = [sum_columns]

        for col in sum_columns:
            if col in self.columns:
                total_sum = self[col].sum(min_count=1)  # 使用min_count=1确保全为np.nan时结果为0
                total_row[col] = total_sum if not pd.isna(total_sum) else 0

        total_df = pd.DataFrame([total_row], columns=self.columns)
        result = self._constructor(pd.concat([self, total_df], ignore_index=True))

        return result

    tab.__doc__ = pandaspro.core.tools.tab.tab.__doc__
    dfilter.__doc__ = pandaspro.core.tools.dfilter.dfilter.__doc__
    inlist.__doc__ = pandaspro.core.tools.inlist.__doc__
    varnames.__doc__ = pandaspro.core.tools.varnames.varnames.__doc__
    lowervarlist.__doc__ = pandaspro.io.excel._utils.lowervarlist.__doc__

    # Overwriting original methods
    def merge(self, *args, **kwargs):
        update = kwargs.pop('update', None)  # Extract the 'update' parameter and remove it from kwargs
        result = super().merge(*args, **kwargs, indicator=True)

        if update == 'missing':
            for col in result.columns:
                if '_x' in col and col.replace('_x', '_y') in result.columns:
                    # Update only if the left column has missing values
                    result[col] = result[col].fillna(result[col.replace('_x', '_y')])
            # Drop the columns from the right DataFrame
            result = result.drop(columns=[col for col in result.columns if '_y' in col])

        elif update == 'all':
            for col in result.columns:
                if '_x' in col and col.replace('_x', '_y') in result.columns:
                    # Update the left column with values from the right column
                    result[col] = result[col.replace('_x', '_y')]
            # Drop the columns from the right DataFrame
            result = result.drop(columns=[col for col in result.columns if '_y' in col])

        result.columns = [col.replace('_x', '') for col in result.columns]
        print(result.tab('_merge'))
        return self._constructor(result)

    def rename(self, columns=None, *args, **kwargs):
        return self._constructor(super().rename(columns=columns, *args, **kwargs))


pd.DataFrame.excel_e = FramePro.excel_e


if __name__ == '__main__':
    a = c({'a': [1, 2, 3], 'b':[3, 4, 5]})
    b = a.inlist('a',1, engine = 'm')
    e = a.inlist('a', 1, engine='b')