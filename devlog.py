'''
现在已知framewrite可以获取index列对应的区域，例如单列index，可以获取G2:G11（参考framewriter下方测试）
然后就是需要得到一个列表，是按取值划分的长度，代码如下：
例如获取了【6，2，1，4】就可用offset结合resize对G2开始进行单元格计算，获取到所有的区域

3月27日（明早记录）
1. 给五月推工作
信息 + 技术 + 方法，信息是需要大家自己搜集的，我可以提供技术，提供写作方法，以及提供获取信息的工具
进一步总结

2. hr权力大

3. 老板大晚上要活，只能说好用的工具没开发完啊

4. 雪琪的pivot更好用

5. 初审终于基本通过了

6. bank医保刷新认知


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


