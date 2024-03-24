from pandaspro import FramePro
from pandaspro import sysuse_auto, sysuse_countries


class CdFormat:
    def __init__(self, col: str, rules: dict, apply: str = 'self'):
        self.col = col
        self.rules = rules
        self.rules_mask = None
        self.apply = apply

    def configure_rules_mask(self, df):
        self.rules_mask = {}

        for rulename, value in self.rules.items():
            self.rules_mask[rulename] = {}

            if not isinstance(rulename, str):
                raise ValueError("Simple Conditional Formatting can only accept str inputs in rules dictionary's keys")

            elif isinstance(rulename, str) and isinstance(value, str):
                self.rules_mask[rulename]['mask'] = FramePro(df).inlist(self.col, rulename)

            else:
                # If rule is given as a list, only range filtering can satisfy
                if isinstance(value, list) and len(value) == 2 and isinstance(value[0], range):
                    self.rules_mask[rulename]['mask'] = df[self.col].between(value[0].start, value[0].stop, inclusive="left")

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
                        method_call = getattr(FramePro(df), engine)
                        self.rules_mask[rulename]['mask'] = method_call(self.col, *myargs, **mykwargs, engine='m')
                    else:
                        raise ValueError("Please pass a list object when declaring 'r' in certain rule")

        return self.rules_mask

    def locate_cells(self, df):
        pass


if __name__ == '__main__':
    # rules = {
    #     'USA': '#FFF000',
    #     'China': '#e63e31',
    #     'rule1': {'r': range(0,9), 'f': 'bold'},
    #     'rule2': {'r': ['inlist',1,2,3, {'invert':True}], 'f': 'bold'}
    # }

    myrule = {
        'rule1': {'r': ['inlist', 'AMC Concord', 'AMC Pacer', {'invert':True}], 'f': 'bold'}
    }
    a = sysuse_auto
    myformat = CdFormat('make', rules=myrule)
    dict1 = myformat.configure_rules_mask(a)

    # FramePro(a).inlist('make', ('AMC Concord'))