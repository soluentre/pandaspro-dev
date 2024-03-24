from pandaspro import FramePro
from pandaspro import sysuse_auto, sysuse_countries


class CdFormat:
    def __init__(self, df, col: str, rules: dict, apply: str = 'self'):
        self.df = df
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
                    self.rules_mask[rulename]['mask'] = df[col].between(value[0].start, value[0].stop, inclusive="left")

                # ... other types of lists will trigger an error
                elif isinstance(value, list):
                    raise ValueError("Simple Conditional Formatting Object can only take 2-element lists when filtering a range")

                # If rule is given as a dictionary
                elif isinstance(value, dict) and 'r' in value.keys() and 'f' in value.keys():
                    if isinstance(value['r'], list):
                        engine = value['r'][0]
                        myargs_list = [item for item in value['r'][1:]]
                        mykwargs_list = [item for item in value['r'] if isinstance(item, dict)]
                        if len(mykwargs_list) > 1:
                            raise ValueError("In 'r' you can only use 1 dictionary for kwargs of the filter engine method")
                        else:
                            mykwargs = mykwargs_list[0]


                    else:
                        raise ValueError("Please pass a list object when declaring 'r' in certain rule")


    def locate_cells(self, df):
        pass


if __name__ == '__main__':
    col = 'Country'
    rules = {
        'USA': '#FFF000',
        'China': '#e63e31',
        'rule1': {'r': range(0,9), 'f': 'bold'},
        'rule2': {'r': ['inlist',1,2,3, {'invert':True}], 'f': 'bold'}
    }

    # myformat = CdFormat()
    a = sysuse_auto
    key = range(0,3800)
    b = sysuse_auto['price'].between(key.start, key.stop, inclusive="left")

    def sum(*args):
        print(*args)

    sum(*[1,2,3])

    FramePro(a).inlist('make', ('AMC Concord'))

#
#
# abc = {
#     ({'a':'c'}): 1
# }
