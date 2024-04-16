from wbhrdata import wbuse_pivot as d
from pandaspro import df_with_index_for_mask

data = df_with_index_for_mask(d)
orderlist = list(data['cmu_dept'].dropna().unique())

unique_values = set(data['cmu_dept'].dropna().unique())
provided_values = set(orderlist)
missing_values = list(unique_values - provided_values)
orderlist.extend(missing_values)

print(d.csort('cmu_dept', ['HAW', 'SAW']))