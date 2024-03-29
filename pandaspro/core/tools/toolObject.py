from pandaspro.core.tools.tab import tab
from pandaspro.core.tools.dfilter import dfilter
from pandaspro.core.tools.inlist import inlist
from pandaspro.core.tools.inrange import inrange


class toolObject:
    def __init__(self):
        self.inlist = inlist
        self.inrange = inrange
        self.tab = tab
        self.dfilter = dfilter