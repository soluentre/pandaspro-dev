from pandaspro import FramePro

class CdFormat:
    def __init__(self, col: str, rules: dict, apply: str = 'self'):
        self.col = col
        self.rules = rules
        self.rules_mask = None
        self.apply = apply

    def configure_rules_mask(self, df):
        self.rules_mask = {}
        for key in self.rules.keys():
            if isinstance(key, str):
                self.rules_mask[key] = FramePro(df).inlist(self.col, key)


    def locate_cells(self, df):

