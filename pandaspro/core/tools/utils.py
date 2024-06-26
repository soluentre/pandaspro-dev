def df_with_index_for_mask(df, force: bool = False):
    if df.index.names[0] is not None or force:
        # Assign a name if not multiple index
        if len(df.index.names) == 1 and df.index.names[0] is None:
            df.index.names = ['_temp_index_sw_assigned']

        # Process
        rename_index = {item: f'__myindex_{str(i)}' for i, item in enumerate(df.index.names)}
        rename_index_back = {f'__myindex_{str(i)}': item for i, item in enumerate(df.index.names)}
        index_preparing = df.reset_index()
        index_wiring = index_preparing.rename(columns=rename_index)

        for column in df.index.names:
            index_wiring[column] = index_preparing[column]
        index_wiring = index_wiring.set_index(list(rename_index.values()))
        index_wiring.index.names = [rename_index_back.get(name) for name in index_wiring.index.names]
        reorder_columns = list(df.index.names) + list(df.columns)
        index_wiring = index_wiring[reorder_columns]

        return index_wiring
    else:
        return df


def create_column_color_dict(df, column, colorlist):
    data = df.reset_index()
    dct = {}
    for i, value in enumerate(data[column].unique()):
        dct[value] = colorlist[i % len(colorlist)]

    return dct
