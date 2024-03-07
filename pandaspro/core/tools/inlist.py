import pandas as pd


def inlist(
    data,
    colname: str,
    engine: str = 'b',
    inplace: bool = False,
    invert: bool = False,
    debug: bool = False,
    *args
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
        'b' for boolean indexing (default)
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
    data = pd.DataFrame(data)
    boolist = args[0] if isinstance(args[0], list) else list(args)
    if debug:
        print(boolist)

    # Update the input var when inplace == True or engine == r:
    if engine == 'r' or True == inplace:
        if debug:
            print("type r code executed ..., trimming the original dataframe")
        if not invert:
            data.drop(data[~data[colname].isin(boolist)].index, inplace=True)
        else:
            data.drop(data[data[colname].isin(boolist)].index, inplace=True)
    elif engine == 'b':
        if debug:
            print("type b code executed ..., creating a tailored dataframe, original frame remain untouched")
        return data[data[colname].isin(boolist)] if invert == False else data[~(data[colname].isin(boolist))]

    elif engine == 'm':
        if debug:
            print("type m code executed ..., creating a mask")
        return data[colname].isin(boolist) if invert == False else ~(data[colname].isin(boolist))

    elif engine == 'c':
        if debug:
            print("type c code executed ...")
        if not invert:
            data.loc[data[colname].isin(boolist), '_inlist'] = 1
        else:
            data.loc[~(data[colname].isin(boolist)), '_inlist'] = 0
        return data
    else:
        print('Unsupported type')
