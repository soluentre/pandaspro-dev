import pandas as pd
from functools import partial

def return_same_type_decor(func):
    def wrapper(self, *args, **kwargs):
        result = func(*args, **kwargs)
        if isinstance(result, pd.DataFrame):
            return self.__class__(data=result)
        return result
    return wrapper

class swFrame(pd.DataFrame):
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
                'index': 'Total',
                name: df[name].sum(),
                'Percent': 1
            })

            # Sort
            df = df.sort_values(sort_dict[sort], ascending=asce)

            # Concatenate the Total row to the DataFrame
            df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
            df.columns = [name, 'Count', 'Percent']
            return df

    def inlist(self):
        pass