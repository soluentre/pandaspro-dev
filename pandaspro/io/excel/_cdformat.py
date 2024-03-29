import pandas as pd
from pandaspro.core.tools.toolObject import toolObject
from pandaspro.core.stringfunc import parsewild

mytools = toolObject()

class CdFormat:
    def __init__(self,
                 df,
                 col: str,
                 rules: dict,
                 applyto: str | list = 'self'):
        self.df = df
        self.col = col
        self.rules = rules
        self.rules_mask = None

        def _apply_decide(input):
            if input == 'self':
                return [col]
            elif input == 'all':
                return self.df.columns
            elif isinstance(input, str):
                return parsewild(applyto, self.df.columns)
            else:
                return input

        self.apply = _apply_decide(applyto)

    def configure_rules_mask(self):
        self.rules_mask = {}

        for rulename, value in self.rules.items():
            self.rules_mask[rulename] = {}

            if not isinstance(rulename, str):
                raise ValueError("Simple Conditional Formatting can only accept str inputs in rules dictionary's keys")

            elif isinstance(rulename, str) and isinstance(value, str):
                self.rules_mask[rulename]['mask'] = mytools.inlist(self.df, self.col, rulename, engine='m')

            else:
                # If rule is given as a list, only range filtering can satisfy
                if isinstance(value, list) and len(value) == 3 and isinstance(value[0], range):
                    self.rules_mask[rulename]['mask'] = self.df[self.col].between(value[0].start, value[0].stop, inclusive=value[1])

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
                            mykwargs = mykwargs_list[0]
                        method_call = getattr(mytools, engine)
                        self.rules_mask[rulename]['mask'] = method_call(self.df, self.col, *myargs, **mykwargs, engine='m')

                    elif isinstance(value['r'], pd.Series):
                        self.rules_mask[rulename]['mask'] = value['r']

                    else:
                        raise ValueError("Please pass a list object when declaring 'r' in certain rule")

                else:
                    raise ValueError('Edit your rules dictionary prompt')

        return self.rules_mask

    def locate_cells(self):
        self.locate = {}
        for key, value in self.configure_rules_mask().items():
            subrange = self.self.df[value['mask']][self.apply]
            self.locate['cells'] = subrange




if __name__ == '__main__':
    from pandaspro import sysuse_auto
    rules = {
        'USA': '#FFF000 bold',
        'China': '#e63e31',
        'rule1': [range(0,9), 'both', '#e63141'],
        'rule2': {
            'r': ['inlist',1,2,3, {'invert':True}],
            'f': 'bold'
        }
    }

    a = sysuse_auto.head(5)
    mask1 = a.inlist('make', 'AMC Concord', engine='m')

    myrule = {
        'AMC Concord': '#FFF000',
        'rule1': {'r': mask1, 'f': 'bold'}
    }

    myformat = CdFormat(a, 'make', rules=myrule)
    dict1 = myformat.configure_rules_mask()

    # FramePro(a).inlist('make', ('AMC Concord'))
