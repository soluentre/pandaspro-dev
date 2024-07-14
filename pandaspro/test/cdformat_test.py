from pandaspro.io.excel.cdformat import CdFormat
from pandaspro.sampledf.sampledf import wbuse_pivot

path = '.pandaspro/tests/sampledf.xlsx'

f = CdFormat(
    wbuse_pivot,
    'cmu_dept',
    cd_rules= {
        'AFWDE': 'blue'
    },
    applyto='self'
)
p
print(f._configure_rules_mask())
print(f.apply)
