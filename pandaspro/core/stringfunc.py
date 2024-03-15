import re

def wildcardread(stringlist, varkey):
    '''
    This is the wildcard reader function which can parse containing-wildcard varnames into meaningful list of varnames
    For example: mak* can return the list of ["make1", "make2", "make3"] which can be further used to slice dataframes

    :param stringlist: a list of vars with wildcards
    :param varkey: a variable key with wildcard in it to match one or more variables
    :return:
    '''
    if '-' in varkey:
        crange = re.split(r'\s*-\s*', varkey)
        element1 = crange[0]
        element2 = crange[1]
        if element1 not in stringlist or element2 not in stringlist:
            print('Invalid column name')
            return None
        if stringlist.index(element1) > stringlist.index(element2):
            element1, element2 = element2, element1
        return stringlist[stringlist.index(element1): stringlist.index(element2) + 1]

    else:
        pattern = re.escape(varkey)
        pattern = '^' + pattern.replace(r'\*', '.*').replace('\?', '.') + '$'
        regex = re.compile(pattern)
        matching_strings = [s for s in stringlist if regex.match(s)]
        return matching_strings


def str2list(inputstring: str) -> str:
    '''
    This function is used to turn a string of vars to a list object
    Python can not automatically parse list of vars as written in a string separated by space, like "make price mpg rep78" as comparing to Stata
    And this function will serve as the parser to separate the string with spaces into var/var with wildcard sections

    :param inputstring: the key input a string with many varnames separated by X number of spaces
    Note: you can use three types of wildcard: * ? -, as supported with the wildcardread function

    :return: a list of varnames
    '''
    pattern = r'\w+\s*-\s*\w+'
    match = re.findall(pattern, inputstring)
    if not match:
        newlist = inputstring.split()
    else:
        for index, item in enumerate(match):
            inputstring = inputstring.replace(item, '__' + str(index) + '__')
        aloneitem = inputstring.split()
        for index, item in enumerate(match):
            newlist = [item if s == '__' + str(index) + '__' else s for s in aloneitem]
    return newlist


def parsewild(checklist: list, promptstring: str):
    '''
    This function will return the searched varnames from a python dataframe according to the prompt string

    :param checklist: list
    :param promptstring: for example: "name* title*"
    :return: a list of available varnames
    '''
    varlist = []
    result_list = []
    for varkey in str2list(promptstring):
        varlist += wildcardread(checklist, varkey)
    [result_list.append(x) for x in varlist if x not in result_list]
    return result_list


def getvarlist(df, promptstring):
    return parsewild(df.columns, promptstring)