from pandaspro.io.excel._framewriter import FramexlWriter
from pandaspro.sampledf.sampledf import wbuse_pivot

path = './sampledf.xlsx'

f = FramexlWriter(wbuse_pivot, 'B2', index=True, header=False)
print(f.range_columns('cmu_dept_major', header=True))
print(f.tr)