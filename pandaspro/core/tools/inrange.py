import pandas as pd


def inrange(
    data,
    colname: str,
    start,
    stop,
    inclusive='left',
    engine: str = 'b',
    inplace: bool = False,
    invert: bool = False,
    debug: bool = False,
):

    data = pd.DataFrame(data)
    if debug:
        print('start: ', start, ';stop: ', stop, ';inclusive: ', inclusive)

    # Update the input var when inplace == True or engine == r:
    if engine == 'r' or True == inplace:
        if debug:
            print("type r code executed ..., trimming the original dataframe")
        if not invert:
            data.drop(data[~data[colname].between(start, stop, inclusive=inclusive)].index, inplace=True)
        else:
            data.drop(data[data[colname].between(start, stop, inclusive=inclusive)].index, inplace=True)
    elif engine == 'b':
        if debug:
            print("type b code executed ..., creating a tailored dataframe, original frame remain untouched")
        return data[data[colname].between(start, stop, inclusive=inclusive)] if invert == False else data[~(data[colname].between(start, stop, inclusive=inclusive))]

    elif engine == 'm':
        if debug:
            print("type m code executed ..., creating a mask")
        return data[colname].between(start, stop, inclusive=inclusive) if invert == False else ~(data[colname].between(start, stop, inclusive=inclusive))

    elif engine == 'c':
        if debug:
            print("type c code executed ...")
        if not invert:
            data.loc[data[colname].between(start, stop, inclusive=inclusive), '_inrange'] = 1
        else:
            data.loc[~(data[colname].between(start, stop, inclusive=inclusive)), '_inrange'] = 0
        return data
    else:
        print('Unsupported type')


if __name__ == '__main__':
    import numpy as np
    from pandaspro import sysuse_auto
    auto = sysuse_auto
    a = auto.inrange('price', -np.inf, 4000, inclusive='right').df
