import numpy as np


def strpos(
        data,
        colname: str,
        substring: str,
        engine: str = 'b',
        inplace: bool = False,
        invert: bool = False,
        rename: str = None,
        debug: bool = False,
):
    if debug:
        print('substring: ', substring, '; column: ', colname)

    data[colname] = data[colname].replace(np.nan, '')

    if engine == 'r':
        if debug:
            print("type r code executed ..., trimming the original dataframe")
        if not invert:
            data.drop(data[~data[colname].str.contains(substring)].index, inplace=True)
        else:
            data.drop(data[data[colname].str.contains(substring)].index, inplace=True)

    elif engine == 'b':
        if debug:
            print("type b code executed ..., creating a tailored dataframe, original frame remain untouched")
        return data[data[colname].str.contains(substring)] if invert == False else data[
            ~(data[colname].str.contains(substring))]

    elif engine == 'm':
        if debug:
            print("type m code executed ..., creating a mask")
        return data[colname].str.contains(substring) if invert == False else ~(
            data[colname].str.contains(substring))

    elif engine == 'c':
        if debug:
            print("type c code executed ...")

        new_name = rename if rename else '_strpos'
        if inplace:
            if not invert:
                data.loc[data[colname].str.contains(substring), new_name] = 1
                data.loc[~data[colname].str.contains(substring), new_name] = 0
            else:
                data.loc[~(data[colname].str.contains(substring)), new_name] = 1
                data.loc[data[colname].str.contains(substring), new_name] = 0
        else:
            df = data.copy()
            if not invert:
                df.loc[data[colname].str.contains(substring), new_name] = 1
                df.loc[~data[colname].str.contains(substring), new_name] = 0
            else:
                df.loc[~(data[colname].str.contains(substring)), new_name] = 1
                df.loc[~data[colname].str.contains(substring), new_name] = 0
            return df

    else:
        print('Unsupported engine type')
