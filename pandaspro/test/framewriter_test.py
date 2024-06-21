from pandaspro.io.excel._framewriter import FramexlWriter, StringxlWriter
from pandaspro.sampledf.sampledf import wbuse_pivot

path = '.pandaspro/test/sampledf.xlsx'

f = FramexlWriter(wbuse_pivot, 'B2', index=True, header=True)
s = StringxlWriter(cell='B3:B10')
# print(f.range_columns('cmu_dept_major', header=True))
# print(f.dfmap)
# print(f.range_cdformat(
#     'cmu_dept',
#     {
#         'AFWDE': 'blue',
#         'AFWVP': 'orange'
#     },
#     applyto='all'
# )

print(f.range_columns('GD', header=True))
print(f.get_column_letter_by_name('GD').cell)
print(f.start_cellobj.cell)

print(f.range_index_selected_hsection(level='cmu_dept_major', token='Total'))