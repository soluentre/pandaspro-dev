import pandas as pd
import numpy as np


def inlist(
    data,
    colname: str,
    *args,
    engine: str = 'b',
    inplace: bool = False,
    invert: bool = False,
    rename: str = None,
    debug: bool = False,
):
    """
    Filters a DataFrame based on whether values in a specified column are in a given list. Supports various
    operation types including filtering, masking, and creating a new indicator column.

    Parameters
    ----------
    data : DataFrame
        The DataFrame to operate on.
    colname : str
        The name of the column to check values against the list.
    *args : list or elements
        The list of values to check against or multiple arguments forming the list.
    engine : str, optional
        The operation type:
        'b' for creating a copy of boolean indexing (default)
        'r' for row filtering
        'm' for mask
        'c' for adding a new column.
    inplace : bool, optional
        If True and engine is 'r', filters the DataFrame in place. Defaults to False.
    invert : bool, optional
        If True, inverts the condition to select rows not in the list. Defaults to False.
    debug : bool, optional
        If True, prints debugging information. Defaults to False.

    Returns
    -------
    DataFrame or Series or None
        The output depends on the engine parameter.
        It may return a filtered DataFrame, a boolean Series (mask), or None if inplace=True.

    Examples
    --------
    >>> df = pd.DataFrame({'A': [1, 2, 3, 4, 5]})
    >>> inlist(df, 'A', 2, 3, engine='b')
    Filters `df` to include only rows where column 'A' contains 2 or 3.

    >>> inlist(df, 'A', [1, 2], engine='r', inplace=True)
    Modifies `df` in place, keeping only rows where column 'A' contains 1 or 2.

    >>> mask = inlist(df, 'A', 4, engine='m')
    Creates a boolean mask for rows where column 'A' contains 4.

    >>> df = inlist(df, 'A', 5, engine='c', invert=True)
    Adds a new column '_inlist' to `df`, marking with 1 the rows where column 'A' does not contain 5.
    """
    bool_list = []
    for arg in args:
        if isinstance(arg, (int, float, np.nan)):
            bool_list.append(arg)
        elif isinstance(arg, list):
            bool_list.extend(arg)
        else:
            raise ValueError("position parameters must be either str or a list")

    if debug:
        print(bool_list)

    # Update the input var when inplace == True or engine == r:
    if engine == 'r':
        if debug:
            print("type r code executed ..., trimming the original dataframe")
        if not invert:
            data.drop(data[~data[colname].isin(bool_list)].index, inplace=True)
        else:
            data.drop(data[data[colname].isin(bool_list)].index, inplace=True)

    elif engine == 'b':
        if debug:
            print("type b code executed ..., creating a tailored new dataframe, original frame remain untouched")

        if inplace:
            if not invert:
                data.drop(data[~data[colname].isin(bool_list)].index, inplace=True)
            else:
                data.drop(data[data[colname].isin(bool_list)].index, inplace=True)
        else:
            return data[data[colname].isin(bool_list)] if invert == False else data[~(data[colname].isin(bool_list))]

    elif engine == 'm':
        if debug:
            print("type m code executed ..., creating a mask")
        return data[colname].isin(bool_list) if invert == False else ~(data[colname].isin(bool_list))

    elif engine == 'c':
        if debug:
            print("type c code executed ...")

        new_name = rename if rename else '_inlist'
        if inplace:
            if not invert:
                data.loc[data[colname].isin(bool_list), new_name] = 1
                data.loc[~data[colname].isin(bool_list), new_name] = 0
            else:
                data.loc[~(data[colname].isin(bool_list)), new_name] = 1
                data.loc[data[colname].isin(bool_list), new_name] = 0
        else:
            df = data.copy()
            if not invert:
                df.loc[data[colname].isin(bool_list), new_name] = 1
                df.loc[~data[colname].isin(bool_list), new_name] = 0
            else:
                df.loc[~(data[colname].isin(bool_list)), new_name] = 1
                df.loc[~data[colname].isin(bool_list), new_name] = 0
            return df
    else:
        print('Unsupported type')


