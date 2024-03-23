from pandaspro.io.excel._framewriter import FramexlWriter
from pandaspro.io.excel._xlwings import RangeOperator
import xlwings as xw


class CustomStyle:
    def __init__(
            self,
            ws,
            framewriter: FramexlWriter,
    ):
        self.ws = ws
        self.frame = framewriter

        self.allstyle = RangeOperator(self.ws.range(self.frame.range_all))
        self.indexstyle = RangeOperator(self.ws.range(self.frame.range_index))

    def style1(self):
        self.allstyle.format(border='all, 3')

    def style2(self):
        self.allstyle.format(border='inner, 1')
        self.allstyle.format(border='outer, 3')




if __name__ == '__main__':


    import pandas as pd

    countries = ["USA", "China", "Japan", "Germany", "India", "UK", "France", "Brazil", "Italy", "Canada"]
    gdp = [11.43, 14.59, 12.45, 11.35, 9.05, 13.27, 9.31, 17.94, (1,2,3,4,5,6), 8.29]
    population = [1110.5, 745.2, 799.6, 1296.6, 108.7, 131.1, 38.1, 1167.3, 1091.6, 1219.3]
    df = pd.DataFrame({
        'Country': countries,
        'GDP (Trillion USD)': gdp,
        'Population (Millions)': population,
    })

    ws = xw.Book('test.xlsx').sheets['FF']
    ws.range('G1').value = df
    a = FramexlWriter(df, 'G1', index=True)

    b = CustomStyle(ws, a)
    print('end')44

    ## Create a CdFormat
    myformat = CdFormat('grade', applyrange='self', rows={'GA':'#FFF00'})
    myformat2 = CdFormat('salary', applrange='age', rows={range(0,100):'bold'})
    a.putxl(df, 'sheet', 'G1', cdformat=[myformat, myformat2])

    cdFormat('grade', {}, ['age','name'])

    filter1  = df[df['grade']>df['hgrade2']]
    myformat = (cdFormat(conplex)
    myformat.add('#FFF000', index_mask=filter1, applyrange=['grade','grade2'])

    dfmap[filter1]