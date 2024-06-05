from termcolor import colored

def calculate_similarity(row1, row2, match_and_weights, debug=False):
    """
    This function is used to return the similarity index when comparing two dataframes with selected columns
    the selected columns mapping is provided using the match_and_weights para

    :param row1: the row of the data to be checked
    :param row2: source databases row
    :param match_and_weights: column name mapping and weight setting

    Sample:
    columns_mapping = {
        'name': {'col':'name_full', 'weight': 0.7},
        'title': {'col':'title', 'weight': 0.1},
        'grade': {'col':'grade', 'weight': 0.1},
        'nationality': {'col':'nationality', 'weight': 0.1}
    }

    :return: similarity score, the higher the better the match is
    """
    try:
        from fuzzywuzzy import fuzz
    except ImportError:
        raise ImportError("Please install 'fuzzywuzzy' package to enable this method, or you may use 'pip install your_package_name[fuzzy]' to install all the dependencies required the first time you install pandaspro")

    total_similarity = 0.0

    for small_col, data in match_and_weights.items():
        large_col = str(data['col'])
        weight = data['weight']
        total_similarity += (fuzz.token_sort_ratio(row1[large_col], row2[small_col]) / 100.0) * weight
    return total_similarity


def search2df(data_small=None, data_large=None, dictionary=None, key=None, mapsample=None, threshold=0.9, show=True, debug=False):
    """
    This function used the calculate_similarity as listed above to create checking dev-reports and generate the key column
    in the smaller (to be checked) dataframe according to the larger (source) dataframe.

    :param data_small: smaller, or to be checked dataframe
    :param data_large: larger, or source dataframe
    :param dictionary: mapping - this is to build relationship between two dataframes
    :param key: the id column in the source dataframe that is to be generated in the smaller one

    :param threshold: set the threshold to map the key column into the to-be-checked dataframe
    :param show: if True, then display the mapping result, by default set to true

    :return: the updated smaller dataframe with key
    """
    finaldf = data_small.copy()
    if mapsample:
        sample = {
            'name': {'col': 'name_full', 'weight': 0.7},
            'title': {'col': 'title', 'weight': 0.1},
            'grade': {'col': 'grade', 'weight': 0.1},
            'nationality': {'col': 'nationality', 'weight': 0.1}
        }
        return sample

    count = 1
    for idx_small, row_small in data_small.iterrows():
        if show:
            print(f">>>> Count {count}/{len(data_small)}: \n")
            print("Data on Left:")
            print("=======================")
            print(row_small[list(dictionary.keys())])

        found_match = False
        for idx_large, row_large in data_large.iterrows():
            similarity = calculate_similarity(row_large, row_small, dictionary, debug=debug)

            if similarity > threshold:
                if show:
                    print("\nData on Right:")
                    print("=======================")
                    print(row_large[[col['col'] for col in dictionary.values()]])
                    print("")
                finaldf.at[idx_small, key] = row_large[key]
                found_match = True
                break

        if show:
            if not found_match:
                print(
                    colored("\n###########################################\n [!] Searched but no results for this item\n###########################################\n", 'red')
                )
            print("----------------------------------------------------------------------------------")
            print("")
        count += 1
    return finaldf