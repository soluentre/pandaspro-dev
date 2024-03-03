import re


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
