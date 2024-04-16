import pandas as pd
from pandaspro.core.tools.utils import df_with_index_for_mask


def csort(
        data,
        column,
        orderlist=None,
        value=None,
        before=None,
        after=None,
        inplace=False
):
    """
    Sorts the DataFrame by the given column according to a custom or dynamically generated order.
    Automatically completes the orderlist with all unique values in the column if it's partially provided.
    Optionally, positions rows with a specified value before or after another value.

    :param data: DataFrame to sort
    :param column: Column name on which to sort
    :param orderlist: List defining the custom order, dynamically completed if partially provided
    :param value: The value to reposition (optional)
    :param before: The value before which the specified value should be placed (optional)
    :param after: The value after which the specified value should be placed (optional)
    :param inplace: Whether to sort the DataFrame in place (default: False)

    :return: Sorted DataFrame or None if sorted in place
    """
    if column in data.columns:
        pass
    elif column in data.index.names:
        data = df_with_index_for_mask(data)
    else:
        raise ValueError(f'Column {column} not found in either the dataframe nor the index namelist')

    if orderlist is None:
        orderlist = list(data[column].dropna().unique())
    else:
        provided_values = set(orderlist)
        provided_reorder = [x for x in data[column].dropna().unique() if x in provided_values]
        missing_reorder = [x for x in data[column].dropna().unique() if x not in provided_values]
        full_orderlist = provided_reorder + missing_reorder
        orderlist = full_orderlist

    # Reorder the list if value and before/after are specified
    if value and (before or after):
        if before and (value in orderlist) and (before in orderlist):
            orderlist.remove(value)
            before_index = orderlist.index(before)
            orderlist.insert(before_index, value)
        elif after and (value in orderlist) and (after in orderlist):
            orderlist.remove(value)
            after_index = orderlist.index(after) + 1
            orderlist.insert(after_index, value)

    cat_type = pd.CategoricalDtype(categories=orderlist, ordered=True)
    data['__cpd_sort'] = data[column].astype(cat_type)

    if inplace:
        data.sort_values(by='__cpd_sort', inplace=True)
        if set(data.index.names) <= set(data.columns):
            data.drop(list(data.index.names) + ['__cpd_sort'], axis=1, inplace=True)
    else:
        result = data.sort_values(by='__cpd_sort')
        if set(result.index.names) <= set(result.columns):
            result.drop(list(data.index.names) + ['__cpd_sort'], axis=1, inplace=True)
        return result, orderlist
