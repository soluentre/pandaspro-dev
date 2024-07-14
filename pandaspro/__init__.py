from pandaspro.email.api import (
    emailfetcher,
    create_mail_class
)

from pandaspro.core.api import (
    bdate,
    dfilter,
    FramePro,
    tab,
    str2list,
    wildcardread,
    parse_method,
    parse_wild,
    df_with_index_for_mask,
    create_column_color_dict,
    csort,
)

from pandaspro.io.api import (
    CellPro,
    index_cell,
    cell_index,
    resize,
    offset,
    getrange,
    PutxlSet,
    pwread,
    WorkbookExportSimplifier,
    fw
)
from pandaspro.io.excel.base import lowervarlist

from pandaspro.sampledf.api import (
    sysuse_countries,
    sysuse_auto,
    wbuse_pivot
)

excel_d = WorkbookExportSimplifier().declare_workbook