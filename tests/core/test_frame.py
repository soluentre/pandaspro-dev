from pandaspro.core.frame import FramePro


def test_framepro_initialization():
    data = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df = FramePro(data)
    assert isinstance(df, FramePro)
    assert df.shape == (3, 2)


def test_tab_method():
    data = {'A': [1, 2, 3], 'B': [4, 5, 6]}
    df = FramePro(data)
    result = df.tab('A', 'detail')
    assert isinstance(result, FramePro)
    assert 'A' in result.columns


def test_add_total_method():
    data = {'Category': ['A', 'B', 'C'], 'Value': [100, 200, 300]}
    df = FramePro(data)
    df_total = df.add_total(total_label_column='Category', sum_columns='Value')
    assert df_total.iloc[-1]['Category'] == 'Total'
    assert df_total.iloc[-1]['Value'] == 600
