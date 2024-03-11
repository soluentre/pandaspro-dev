import xlwings as xw
import re


def _extract_tuple(s):
    pattern = r'\((\d+,\s*\d+,\s*\d+)\)'
    matches = list(re.finditer(pattern, s))
    if len(matches) == 0:
        return None, s.strip()
    elif len(matches) == 1:
        match = matches[0]
        tuple_str = match.group(1)
        color_tuple = tuple(map(int, tuple_str.split(',')))
        remaining_str = s[:match.start()] + s[match.end():]
        return color_tuple, remaining_str.strip()
    else:
        raise ValueError(f"Multiple tuples found in '{s}'")


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
        self.xwrange = xwrange

    def format(
            self,
            font: str | tuple | list = None,
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
                self.xwrange.font.color = font
            elif isinstance(font, int):
                self.xwrange.font.size = font
            elif isinstance(font, str) and re.fullmatch(r'#[0-9A-Fa-f]{6}', font):
                self.xwrange.font.name = font
            elif isinstance(font, str):
                color, remaining = _extract_tuple(font)
                if color:
                    self.xwrange.font.color = color
                for item in remaining.split(','):
                    item = item.strip()
                    if isinstance(item, int):
                        self.xwrange.font.size = item
                    elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                        self.xwrange.font.color = item
                    else:
                        self.xwrange.font.name = item

        if font_name:
            self.xwrange.font.name = font_name

        if font_size is not None:
            self.xwrange.font.size = font_size

        if font_color:
            self.xwrange.font.color = font_color

        if italic is not None:
            self.xwrange.font.italic = italic

        if bold is not None:
            self.xwrange.font.bold = bold

        if underline is not None:
            self.xwrange.api.Font.Underline = underline

        if strikeout is not None:
            self.xwrange.api.Font.Strikethrough = strikeout

        if align:
            def _alignfunc(alignkey):
                if align in ['center', 'justify', 'distributed']:
                    self.xwrange.api.VerticalAlignment = self._alignment_map['v' + alignkey][1]
                    self.xwrange.api.HorizontalAlignment = self._alignment_map['h' + alignkey][1]
                elif self._alignment_map[alignkey][0] == 'v':
                    self.xwrange.api.VerticalAlignment = self._alignment_map[alignkey][1]
                elif self._alignment_map[alignkey][0] == 'h':
                    self.xwrange.api.HorizontalAlignment = self._alignment_map[alignkey][1]
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
        self.xwrange.clear()


if __name__ == '__main__':
    wb = xw.Book('test.xlsx')
    sheet = wb.sheets[0]  # Reference to the first sheet

    # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
    my_range = sheet.range("C1:D2")

    # Step 3: Create an object of the RangeOperator class with the specified range
    a = RangeOperator(my_range)
    a.format(font='Calibri, 14, #FFFF00', align='center')    # print(a.range)
    # print(_extract_tuple('12 (4,255,67)'))
    # a.range.font.color = '#FF0000'