import pandas as pd
from pandaspro.core.tools.toolObject import toolObject
from pandaspro.core.stringfunc import parse_wild

mytools = toolObject()


def df_with_index_for_mask(df):
    if df.index.name is not None:
        rename_index = {item: f'__myindex_{str(i)}' for i, item in enumerate(df.index.names)}
        rename_index_back = {f'__myindex_{str(i)}': item for i, item in enumerate(df.index.names)}
        index_preparing = df.reset_index()
        index_wiring = index_preparing.rename(columns=rename_index)
        for column in df.index.names:
            index_wiring[column] = index_preparing[column]
        index_wiring = index_wiring.set_index(list(rename_index.values()))
        index_wiring.index.names = [rename_index_back.get(name) for name in index_wiring.index.names]
        reorder_columns = list(df.index.names) + list(df.columns)
        index_wiring = index_wiring[reorder_columns]

        return index_wiring
    else:
        return df


class CdFormat:
    def __init__(
            self,
            df,
            column: str,
            cd_rules: dict,
            applyto: str | list = 'self'
    ):
        self.df = df
        self.column = column
        self.cd_rules = cd_rules
        self.rules_mask = None
        self.locate = None
        self.df_with_index = df_with_index_for_mask(self.df)
        self.col_not_exist = None if self.column in self.df_with_index.columns else True

        def _apply_decide(local_input):
            if local_input == 'self':
                return [self.column]
            elif local_input == 'inner':
                return self.df.columns
            elif local_input == 'all':
                return self.df_with_index.columns
            elif isinstance(local_input, str):
                return parse_wild(applyto, self.df.columns)
            elif isinstance(local_input, list):
                return local_input
            else:
                raise TypeError('Unexpected type of applyto parameter, only str/list being accepted')

        self.apply = _apply_decide(applyto)
        if self.column in self.df_with_index.columns:
            self.rules_mask = self._configure_rules_mask()

    def _configure_rules_mask(self):
        result = {}

        for rulename, value in self.cd_rules.items():
            result[rulename] = {}

            if self.column in self.df_with_index.columns:
                pass
            else:
                raise ValueError('Invalid column name specified.')

            # Mask and Format Prompt
            if isinstance(value, str):
                result[rulename]['mask'] = mytools.inlist(self.df_with_index, self.column, rulename, engine='m')
                result[rulename]['format'] = value

            # If rule is given as a list, only range filtering can satisfy
            elif isinstance(value, list) and len(value) == 3 and isinstance(value[0], range):
                result[rulename]['mask'] = self.df[self.column].between(value[0].start, value[0].stop,
                                                                        inclusive=value[1])
                result[rulename]['format'] = value[2]

            # ... other types of lists will trigger an error
            elif isinstance(value, list):
                raise ValueError(
                    "Simple Conditional Formatting Object can only take 2-element lists when filtering a range")

            # If rule is given as a dictionary
            elif isinstance(value, dict) and 'r' in value.keys() and 'f' in value.keys():
                if isinstance(value['r'], list):
                    engine = value['r'][0]
                    myargs = [item for item in value['r'][1:] if not isinstance(item, dict)]
                    mykwargs_list = [item for item in value['r'] if isinstance(item, dict)]
                    if len(mykwargs_list) > 1:
                        raise ValueError("In 'r' you can only use 1 dictionary for kwargs of the filter engine method")
                    else:
                        if len(mykwargs_list) == 1:
                            mykwargs = mykwargs_list[0]
                        else:
                            mykwargs = {}
                    method_call = getattr(mytools, engine)
                    result[rulename]['mask'] = method_call(self.df, self.column, *myargs, **mykwargs, engine='m')
                    result[rulename]['format'] = value['f']

                elif isinstance(value['r'], pd.Series):
                    result[rulename]['mask'] = value['r']
                    result[rulename]['format'] = value['f']

                else:
                    raise ValueError("Please pass a list object when declaring 'r' in certain rule")

            else:
                raise ValueError('Edit your rules dictionary prompt, use helpfile(cd_format) to view valid examples')

        return result
