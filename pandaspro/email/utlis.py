import re


def replace_with_dict(text, dictionary):
    for key, value in dictionary.items():
        pattern = re.compile(re.escape(key))
        text = pattern.sub(value, text)
    return text
