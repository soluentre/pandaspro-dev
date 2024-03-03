from datetime import datetime


def now(format='B, Y', override=False, info=False):
    if info == True:
        print('''
By default, the function will seperate  ... 
        ''')

    format_mapping = {
        'BdY': '%B %d, %Y',
        'bdY': '%b %d, %Y'
    }

    def refine_format(format):
        for letter in ['b', 'B', 'Y', 'm', 'd']:
            format = format.replace(letter, '%' + letter)
        return format

    if override == False:
        format = format_mapping[format] if format in format_mapping.keys() else ''.join(
            filter(lambda x: x.isalpha() or x == '%', format))

        return datetime.now().strftime(format)
    elif override == True:
        return datetime.now().strftime(format)


print(now())


def currentdate(display='textDMY'):
    if display == 'num':
        return datetime.now().strftime("%Y-%m-%d").replace('-', '')
    if display == 'textDMY':
        return datetime.now().strftime("%d") + " " + month_mapping[
            int(datetime.now().strftime("%Y-%m").split('-')[1])] + " " + datetime.now().strftime("%Y-%m").split('-')[0]
