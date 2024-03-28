'''
现在已知framewrite可以获取index列对应的区域，例如单列index，可以获取G2:G11（参考framewriter下方测试）
然后就是需要得到一个列表，是按取值划分的长度，代码如下：
例如获取了【6，2，1，4】就可用offset结合resize对G2开始进行单元格计算，获取到所有的区域


'''

import pandas as pd
import numpy as np

# Create a DataFrame with 15 rows
# First column is just a range from 1 to 15
# Second column is randomly chosen letters from A, B, C, D
df = pd.DataFrame({
    'Col1': range(1, 10),
    'Col2': np.random.choice(['A', 'B', 'C', 'D'], 9)
})

# Sort the DataFrame by the second column
df_sorted = df.sort_values(by='Col2')

df_sorted

def count_consecutive_values(series):
    return series.groupby((series != series.shift()).cumsum()).size().tolist()

# Apply the function to the second column of the sorted DataFrame
consecutive_counts = count_consecutive_values(df_sorted['Col2'])

consecutive_counts


