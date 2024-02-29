import pandas as pd

def tab(data, name, d='brief', m=False, sort='index', ascending=True):
    sort_dict = {
        f'{name}': f'{name}',
        'index': f'{name}',
        'percent': 'Percent'
    }
    if m == 'missing' or m == True:
        df = data[name].value_counts(dropna=False).sort_index().to_frame()
    else:
        df = data[name].value_counts().sort_index().to_frame()

    if d == 'brief':
        # Sort
        if sort == 'index':
            df = df.sort_index(ascending=ascending)
        else:
            df = df.sort_values(sort_dict[sort], ascending=ascending)
        return df

    elif d == 'detail':
        # Calculate Percent and Cumulative Percent
        df = df.reset_index()
        df['Percent'] = (df['count'] / df['count'].sum() * 100).round(2)

        # Sort
        df = df.sort_values(sort_dict[sort], ascending=ascending)
        df['Cum.'] = df['Percent'].cumsum().round(2)

        # Create a Total row
        total_row = pd.Series({
            name: 'Total',
            'count': df['count'].sum(),
            'Percent': 100.00,
            'Cum.': ''
        })

        # Concatenate the Total row to the DataFrame
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
        return df

    elif d == 'export':
        df = df.reset_index()
        df['Percent'] = (df['count'] / df['count'].sum()).round(3)
        total_row = pd.Series({
            name: 'Total',
            'count': df['count'].sum(),
            'Percent': 1
        })

        # Sort
        df = df.sort_values(sort_dict[sort], ascending=ascending)

        # Concatenate the Total row to the DataFrame
        df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
        df.columns = [name, 'Count', 'Percent']
        return df