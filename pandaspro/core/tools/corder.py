import pandas as pd

def corder(
        data,
        column: str | list,
        before=None,
        after=None,
):
    """
    Reorder the DataFrame columns by positioning specified columns before or after the corresponding columns.

    :param data: DataFrame to reorder
    :param column: Column name on which to change position, will be switched to the beginning of the DataFrame if param before and after are not specified
    :param before: The column name before which the specified column should be placed (optional)
    :param after: The column name after which the specified column should be placed (optional)

    :return: Reordered DataFrame or None if reordered in place
    """
    if column in data.columns:
        pass
    else:
        raise ValueError(f'Column {column} not in the dataframe')

    if isinstance(column, str):
        cols = [i.strip() for i in column.split(';')]
    elif isinstance(column, list):
        cols = column

    old_order = list(data.columns)
    if before:
        index = old_order.index(before)
        new_order = old_order[:index] + cols + old_order[index:]
    elif after:
        index = old_order.index(after)
        new_order = old_order[:index+1] + cols + old_order[index+1:]
    else:
        new_order = cols + old_order

    return data[new_order]
