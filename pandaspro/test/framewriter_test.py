from pandaspro.io.excel._framewriter import FramexlWriter, StringxlWriter
from pandaspro.sampledf.sampledf import wbuse_pivot

path = '.pandaspro/test/sampledf.xlsx'

f = FramexlWriter(wbuse_pivot, 'B2', index=True, header=False)
s = StringxlWriter('B3')
# print(f.range_columns('cmu_dept_major', header=True))
print(f.dfmap)
# print(f.range_cdformat(
#     'cmu_dept',
#     {
#         'AFWDE': 'blue',
#         'AFWVP': 'orange'
#     },
#     applyto='all'
# ))
