'''
let's create a pivot pro:
pivot_table will return a pd.DataFrame object, let's first inherit the DataFrame, then make a child class:
    1. the child class will recognize the header, index, Row total and Column total
    2. we need to create subtotal rows and columns according to the header/index -> we are going to play with multi_index
    3. we also need to create total rows and columns quickly
    4. we need be able to claim a dataframe as this class easily: classname(df, '0-y-y-0)

'''