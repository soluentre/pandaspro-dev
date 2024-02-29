import pandas as pd
from functools import partial

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

    def tab(self, name, d='brief', m=False, sort='index', asce=True):
        sort_dict = {
            f'{name}': f'{name}',
            'index': f'{name}',
            'percent': 'Percent'
        }
        if m == 'missing' or m == True:
            df = self[name].value_counts(dropna=False).sort_index().to_frame()
        else:
            df = self[name].value_counts().sort_index().to_frame()

        if d == 'brief':
            # Sort
            if sort == 'index':
                df = df.sort_index(ascending=asce)
            else:
                df = df.sort_values(sort_dict[sort], ascending=asce)
            return df

        elif d == 'detail':
            # Calculate Percent and Cumulative Percent
            df = df.reset_index()
            df['Percent'] = (df['count'] / df['count'].sum() * 100).round(2)

            # Sort
            df = df.sort_values(sort_dict[sort], ascending=asce)
            df['Cum.'] = df['Percent'].cumsum().round(2)

            # Create a Total row
            total_row = pd.Series({
                name: 'Total',
                'count': df['count'].sum(),
                'Percent': 100.00,
                'Cum.': ''
            })

            # Concatenate the Total row to the DataFrame
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
            return df

        elif d == 'export':
            df = df.reset_index()
            df['Percent'] = (df['count'] / df['count'].sum()).round(3)
            total_row = pd.Series({
                name: 'Total',
                'count': df['count'].sum(),
                'Percent': 1
            })

            # Sort
            df = df.sort_values(sort_dict[sort], ascending=asce)

            # Concatenate the Total row to the DataFrame
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
            df.columns = [name, 'Count', 'Percent']
            return df

    # def inlist(self):
    #     pass