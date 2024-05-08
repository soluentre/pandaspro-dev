from pandaspro.io.excel.putexcel import PutxlSet
from pandaspro.sampledf.sampledf import wbuse_pivot


path = './pandaspro/test/sampledf.xlsx'

ps = PutxlSet(path)
ps.putxl(
    wbuse_pivot,
    sheet_name='newtab3',
    cell='B2',
    index=True,
    design='wbblue',
    tab_color='blue'
)