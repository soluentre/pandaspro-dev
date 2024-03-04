import xlwings as xw
import re


class RangeOperator:
    _alignment_map = {
        'hcenter': ['h', xw.constants.HAlign.xlHAlignCenter],
        'center_across_selection': ['h', xw.constants.HAlign.xlHAlignCenterAcrossSelection],
        'hdistributed': ['h', xw.constants.HAlign.xlHAlignDistributed],
        'fill': ['h', xw.constants.HAlign.xlHAlignFill],
        'general': ['h', xw.constants.HAlign.xlHAlignGeneral],
        'hjustify': ['h', xw.constants.HAlign.xlHAlignJustify],
        'left': ['h', xw.constants.HAlign.xlHAlignLeft],
        'right': ['h', xw.constants.HAlign.xlHAlignRight],
        'bottom': ['v', xw.constants.VAlign.xlVAlignBottom],
        'vcenter': ['v', xw.constants.VAlign.xlVAlignCenter],
        'vdistributed': ['v', xw.constants.VAlign.xlVAlignDistributed],
        'vjustify': ['v', xw.constants.VAlign.xlVAlignJustify],
        'top': ['v', xw.constants.VAlign.xlVAlignTop],
    }

    def __init__(self, xwrange: xw.Range) -> None:
        self.range = xwrange

    def format(
            self,
            font: str | tuple = None,
            font_name: str = None,
            font_size: str = None,
            font_color: str | tuple = None,
            italic: bool = False,
            bold: bool = False,
            underline: bool = False,
            strikeout: bool = False,
            align: str | list = None,
    ) -> None:
        if font:
            if isinstance(font, tuple):
                self.range.font.color = font
            elif isinstance(font, str):
                for item in font.split():
                    if re.match(item, '#\d{6}'):

        # 'color' attribute can follow formats: (1) string of hex starts with '#'; (2) tuple of RGB
        self.range.font.name = font_name
        self.range.font.size = font_size
        self.range.font.color = font_color
        self.range.font.italic = italic
        self.range.font.bold = bold
        self.range.api.Font.Underline = underline
        self.range.api.Font.Strikethrough = strikeout

        if align:
            def _alignfunc(alignkey):
                if align in ['center', 'justify', 'distributed']:
                    self.range.api.VerticalAlignment = self._alignment_map['v' + alignkey][1]
                    self.range.api.HorizontalAlignment = self._alignment_map['h' + alignkey][1]
                elif self._alignment_map[alignkey][0] == 'v':
                    self.range.api.VerticalAlignment = self._alignment_map[alignkey][1]
                elif self._alignment_map[alignkey][0] == 'h':
                    self.range.api.HorizontalAlignment = self._alignment_map[alignkey][1]
                elif align not in self._alignment_map.keys():
                    raise ValueError(f'Alignment {alignkey} is not supported')
                return
            if isinstance(align, str):
                for item in align.split():
                    _alignfunc(item)
            elif isinstance(align, list):
                for item in align:
                    _alignfunc(item)

        return

    def clear(self):
        self.range.clear()


if __name__ == '__main__':
    wb = xw.Book('test.xlsx')
    sheet = wb.sheets[0]  # Reference to the first sheet

    # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
    my_range = sheet.range("C1:D2")

    # Step 3: Create an object of the RangeOperator class with the specified range
    a = RangeOperator(my_range)
    a.format(font_color='#FFFF00')
    print(a.range)
