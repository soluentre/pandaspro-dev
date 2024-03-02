import pandas as pd

def dfilter(data, inputdict: dict, debug: bool=False):
    """
    dfilter is a method that filters the dataframe according to an input dictionary
    In the dictionary, there's a format key which states the combine logic for multiple masks
    And each other key will contain the mask logic following the same structure:

        (column_name, [~]method_key, corresponding list/str object filters)

    Parameters
    ----------
    data : DataFrame
        dataframe object to be filtered
    inputdict : dictionary
        a dictionary follows the required syntax:
    debug : bool, optional and default =False
        if True, print intermediate results along the process

    Returns
    ----------
    ret
        a filtered dataframe object

    Dictionary Syntax
    ----------
    inputdict = {
        'format': "[(] [!] nickname1 [& | !] [nickname2] ... [)],

        'nickname1': (column_name1, [~]method_key1, list/str object),
        'nickname2': (column_name2, [~]method_key1, list/str object),
        ....
    }

    method_key can take values:
    1. isin :  uses the pandas.Series.isin
            receives a list
    2. contains : uses the pandas.Series.str.lower().str.contains
            receives a string
    3. number : no method called
            receives a string with operator and value
    4. na : uses the pandas.Series.isna() method
            the third element in tuple can be omitted or anything, no impact at all

    Example
    ----------
    Below provides an example:
    sample = {
        'format': "(A & B & C & D) & (1 & 2)",

        'A': ('dept', 'isin', ['Africa']),
        'B': ('emp_status', 'isin', ['Active']),
        'C': ('appt_type', 'isin', ['OPEN', 'TERM']),
        'D': ('grade', '~isin', []),

        1: ('title', 'contains', 'program leader|sector leader'),
        2: ('yrs_in_assign', 'number', '>=2'),
        }

    """
    if not isinstance(data, pd.DataFrame):
        print('Please declare one dataframe.')
        return
    df = data.copy()
    filterstring = inputdict['format']
    if debug == True:
        print("Initial Logic Line Read")
        print("---------------------")
        print(filterstring)
        print("")
    if filterstring == "":
        print("Please enter a valid format logic, it cannot be empty.")
        return
    filters = {}  # Use a dictionary to store the filters

    # Evaluate the criteria and save to dicts
    for key in list(inputdict.keys())[1:]:
        cdict = {
            'isin': {
                'func': '.isin',
                'inbracket': f"({inputdict[key][2]})"
            },
            'contains': {
                'func': '.str.lower().str.contains',
                'inbracket': f"('{inputdict[key][2]}', regex=True)"
            },
            'number': {
                'func': '',
                'inbracket': f"{inputdict[key][2]}"
            },
            'na': {
                'func': '.isna',
                'inbracket': f"()"
            }
        }
        usemethod = inputdict[key][1]

        if "~" in usemethod:
            usemethod = usemethod.replace("~", "")
            filter_code = f"~(df['{inputdict[key][0]}']{cdict[usemethod]['func']}{cdict[usemethod]['inbracket']})"
        else:
            filter_code = f"df['{inputdict[key][0]}']{cdict[usemethod]['func']}{cdict[usemethod]['inbracket']}"
        filters[f"filter{key}"] = eval(filter_code)
        if debug == True:
            print(f"item {key}: {filter_code}")
            print("")

    # Combine filterstring
    for i in inputdict.keys():
        filterstring = filterstring.replace(f"{i}", f"filters['filter{i}']")
    if debug == True:
        print("")
        print("Final Line to be Evaluated")
        print("---------------------")
        print(filterstring)
    final = eval(f"df[{filterstring}]")
    return final