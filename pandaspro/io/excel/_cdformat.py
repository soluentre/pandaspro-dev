import pandas as pd
from pandaspro.core.tools.toolObject import toolObject
from pandaspro.core.stringfunc import parse_wild

mytools = toolObject()


class CdFormat:
    def __init__(self,
                 df,
                 column: str,
                 cd_rules: dict,
                 applyto: str | list = 'self'):
        self.df = df
        self.column = column
        self.cd_rules = cd_rules
        self.rules_mask = None
        self.locate = None

        def _apply_decide(local_input):
            if local_input == 'self':
                return [column]
            elif local_input == 'all':
                return self.df.columns
            elif isinstance(local_input, str):
                return parse_wild(applyto, self.df.columns)
            elif isinstance(local_input, list):
                return local_input
            else:
                raise TypeError('Unexpected type of applyto parameter, only str/list being accepted')

        self.apply = _apply_decide(applyto)
        self.rules_mask = self._configure_rules_mask()

    def _configure_rules_mask(self):
        result = {}

        for rulename, value in self.cd_rules.items():
            result[rulename] = {}

            # Mask and Format Prompt
            if isinstance(value, str):
                if self.column in self.df.columns:
                    result[rulename]['mask'] = mytools.inlist(self.df, self.column, rulename, engine='m')
                elif self.column in self.df.index.names:
                    result[rulename]['mask'] = self.df.index.get_level_values(self.column) == rulename
                else:
                    raise ValueError('Invalid column name specified.')
                result[rulename]['format'] = value

            else:
                # If rule is given as a list, only range filtering can satisfy
                if isinstance(value, list) and len(value) == 3 and isinstance(value[0], range):
                    result[rulename]['mask'] = self.df[self.column].between(value[0].start, value[0].stop, inclusive=value[1])
                    result[rulename]['format'] = value[2]

                # ... other types of lists will trigger an error
                elif isinstance(value, list):
                    raise ValueError("Simple Conditional Formatting Object can only take 2-element lists when filtering a range")

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
                    raise ValueError('Edit your rules dictionary prompt')

        return result

