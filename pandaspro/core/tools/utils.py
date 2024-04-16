import pandas as pd


def df_with_index_for_mask(df):
    if df.index.names[0] is not None:
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
