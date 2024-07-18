import pandas as pd
from pandaspro.core.tools.toolObject import toolObject
from pandaspro.core.stringfunc import parse_wild
from pandaspro.core.tools.utils import df_with_index_for_mask
from pandaspro.utils.cpdLogger import cpdLogger

mytools = toolObject()


@cpdLogger
class CdFormat:
    def __init__(
            self,
            df,
            column: str,
            cd_rules: dict,
            applyto: str | list = 'self',
    ):
        self.df = df
        self.column = column
        self.cd_rules = cd_rules
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

        # self.logger.debug_section_spec_start("Creating CdFormat Instance")

    def get_rules_mask(self):
        if self.column in self.df_with_index.columns:
            self.rules_mask = self._configure_rules_mask()
            return self.rules_mask
        else:
            return {}
        # self.logger.debug_section_spec_end()

    '''
    Example of the rules parameter

    rules = {
        'rule1': {
            'r': ['inlist', ....],
            'f': ....
        },
        'rule2': {
            'r': pd.Series,
            'f': ...
        },
    }
    '''
    def _configure_rules_mask(self):
        result = {}
        self.debug_section_spec_start('Creating CdFormat Class')
        self.logger.debug('+++ Created result dict as blank {}')

        for rulename, value in self.cd_rules.items():
            self.logger.debug(f'+++ [key - rulename]: **{rulename}**, [value - value]: **{value}**')
            result[rulename] = {}

            if self.column not in self.df_with_index.columns:
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
