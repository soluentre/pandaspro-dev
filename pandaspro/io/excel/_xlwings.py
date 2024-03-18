import pandas as pd
import xlwings as xw
import re

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

_fpattern_map = {
    'none': 0,  # xlNone
    'solid': 1,  # xlSolid
    'gray50': 2,  # xlGray50
    'gray75': 3,  # xlGray75
    'gray25': 4,  # xlGray25
    'horstripe': 5,  # xlHorizontalStripe
    'verstripe': 6,  # xlVerticalStripe
    'diagstripe': 8,  # xlDiagonalDown
    'revdiagstripe': 7,  # xlDiagonalUp
    'diagcrosshatch': 9,  # xlDiagonalCrosshatch
    'thinhorstripe': 11,  # xlThinHorizontalStripe
    'thinverstripe': 12,  # xlThinVerticalStripe
    'thindiagstripe': 14,  # xlThinDiagonalDown
    'thinrevdiagstripe': 13,  # xlThinDiagonalUp
    'thinhorcrosshatch': 15,  # xlThinHorizontalCrosshatch
    'thindiagcrosshatch': 16,  # xlThinDiagonalCrosshatch
    'thickdiagcrosshatch': 10,  # xlThickDiagonalCrosshatch
    'gray12p5': 17,  # xlGray12.5
    'gray6p25': 18  # xlGray6.25
}

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


def _is_number(s: str):
    pattern = re.compile(r'^[-+]?(\d+(\.\d*)?|\.\d+)([eE][-+]?\d+)?$')
    return bool(pattern.match(s))


def hex_to_int(hex: str):
    hex = hex.lstrip('#')
    red = int(hex[:2], 16)
    green = int(hex[2:4], 16)
    blue = int(hex[4:], 16)
    return red | (green << 8) | (blue << 16)


class RangeOperator:

    def __init__(self, xwrange: xw.Range) -> None:
        self.xwrange = xwrange

    def format(
            self,
            font: str | tuple | list = None,
            font_name: str = None,
            font_size: str = None,
            font_color: str | tuple = None,
            italic: bool = None,
            bold: bool = None,
            underline: bool = None,
            strikeout: bool = None,
            align: str | list = None,
            merge: bool = None,
            fill: str | tuple | list  = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None
    ) -> None:

        # Font Attributes
        ##################################
        if font:
            if isinstance(font, tuple):
                self.xwrange.font.color = font
            elif isinstance(font, (int, float)):
                self.xwrange.font.size = font
            elif isinstance(font, str):
                color, remaining = _extract_tuple(font)
                if color:
                    self.xwrange.font.color = color
                for item in remaining.split(','):
                    item = item.strip()
                    if _is_number(item):
                        self.xwrange.font.size = item
                    elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                        self.xwrange.font.color = item
                    elif item == 'bold':
                        self.xwrange.font.bold = True
                    elif item == 'italic':
                        self.xwrange.font.italic = True
                    elif item == 'underline':
                        self.xwrange.api.Font.Underline = True
                    elif item == 'strikeout':
                        self.xwrange.api.Font.Strikethrough = True
                    else:
                        self.xwrange.font.name = item
            elif isinstance(font, list):
                for item in font:
                    if isinstance(item, tuple):
                        self.xwrange.font.color = item
                    elif isinstance(item, (int, float)):
                        self.xwrange.font.size = item
                    elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                        self.xwrange.font.color = item
                    elif isinstance(item, str) and item == 'bold':
                        self.xwrange.font.bold = True
                    elif isinstance(item, str) and item == 'italic':
                        self.xwrange.font.italic = True
                    elif isinstance(item, str) and item == 'underline':
                        self.xwrange.api.Font.Underline = True
                    elif isinstance(item, str) and item == 'strikeout':
                        self.xwrange.api.Font.Strikethrough = True
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

        # Align Attributes
        ##################################
        if align:
            def _alignfunc(alignkey):
                if align in ['center', 'justify', 'distributed']:
                    self.xwrange.api.VerticalAlignment = _alignment_map['v' + alignkey][1]
                    self.xwrange.api.HorizontalAlignment = _alignment_map['h' + alignkey][1]
                elif _alignment_map[alignkey][0] == 'v':
                    self.xwrange.api.VerticalAlignment = _alignment_map[alignkey][1]
                elif _alignment_map[alignkey][0] == 'h':
                    self.xwrange.api.HorizontalAlignment = _alignment_map[alignkey][1]
                elif align not in _alignment_map.keys():
                    raise ValueError(f'Alignment {alignkey} is not supported')
                return

            if isinstance(align, str):
                for item in align.split(','):
                    item = item.strip()
                    _alignfunc(item)
            elif isinstance(align, list):
                for item in align:
                    _alignfunc(item)

        # Merge Attributes
        ##################################
        if merge:
            xw.apps.active.api.DisplayAlerts = False
            self.xwrange.api.MergeCells = merge
            xw.apps.active.api.DisplayAlerts = True

        elif not merge:
            if self.xwrange.api.MergeCells:
                self.xwrange.unmerge()

        # Border Attributes
        ##################################


        # Fill Attributes
        ##################################
        if fill:
            if isinstance(fill, list):
                if (len(fill) == 1) or (len(fill) == 2 and 'solid' in fill):
                    for item in fill:
                        if isinstance(item, tuple):
                            self.xwrange.api.Interior.Color = xw.utils.rgb_to_int(item)
                        elif item in list(_fpattern_map.keys()):
                            self.xwrange.api.Interior.Pattern = _fpattern_map[item]
                        elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                            self.xwrange.api.Interior.Color = hex_to_int(item)
                elif len(fill) == 2 and 'solid' not in fill:
                    for item in fill:
                        if isinstance(item, tuple):
                            self.xwrange.api.Interior.PatternColor = xw.utils.rgb_to_int(item)
                        elif item in list(_fpattern_map.keys()):
                            self.xwrange.api.Interior.Pattern = _fpattern_map[item]
                        elif re.fullmatch(r'#[0-9A-Fa-f]{6}', item):
                            self.xwrange.api.Interior.PatternColor = hex_to_int(item)
                else:
                    raise ValueError("Can only accept 2 parameters (one for pattern and one for color) at most when passing a list object to 'fill'.")


            elif isinstance(fill, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill)
                self.xwrange.api.Interior.Color = foreground_color_int
            elif isinstance(fill, str):
                patternkeys = '(' + '|'.join(_fpattern_map.keys()) + ')'
                compiled_patternkeys = re.compile(patternkeys, re.IGNORECASE)
                patternlist = re.findall(compiled_patternkeys, fill)

                colorrule = r'(#\w{6}|\(\s*(25[0-5]|2[0-4]\d|1\d{2}|[1-9]\d|\d)\s*,\s*(25[0-5]|2[0-4]\d|1\d{2}|[1-9]\d|\d)\s*,\s*(25[0-5]|2[0-4]\d|1\d{2}|[1-9]\d|\d)\s*\))'
                compiled_colorrule = re.complie(colorrule, re.IGNORECASE)
                colorlist = re.findall(compiled_colorrule, fill)

                for item in patternlist + colorlist:
                    leftover = fill.replace(item, '').replace(',', '').strip()

                if len(leftover) > 0:
                    raise ValueError('Incorrect pattern or color specified, please check')
                elif len(patternlist) > 1 or len(colorlist) > 1:
                    raise ValueError('Can not specify more than one color or more than one pattern, respectively.')
                elif len(patternlist) == 1:
                    self.xwrange.api.Interior.Pattern = patternlist[0]
                elif len(colorlist) == 1:
                    if '#' in colorlist[0]:
                        colorint = hex_to_int(colorlist[0])
                    else:
                        colortuple = tuple(map(int, colorlist[0].replace('(','').replace(')','').split(',')))
                        colorint = xw.utils.rgb_to_int(colortuple)
                    self.xwrange.api.Interior.PatternColor = colorint
                    if patternlist[0] == 'solid':
                        self.xwrange.api.Interior.Color = colorint



        if fill_pattern:
            self.xwrange.api.Interior.Pattern = _fpattern_map[fill_pattern]

        if isinstance(fill_fg, tuple):
            foreground_color_int = xw.utils.rgb_to_int(fill_fg)
            self.xwrange.api.Interior.PatternColor = foreground_color_int
        elif isinstance(fill_fg, str):
            fill_fg = fill_fg.lstrip('#')
            red = int(fill_fg[:2], 16)
            green = int(fill_fg[2:4], 16)
            blue = int(fill_fg[4:], 16)
            foreground_color_int = red | (green << 8) | (blue << 16)
            self.xwrange.api.Interior.PatternColor = foreground_color_int

        if isinstance(fill_bg, tuple):
            background_color_int = xw.utils.rgb_to_int(fill_bg)
            self.xwrange.api.Interior.Color = background_color_int
        elif isinstance(fill_bg, str):
            fill_bg = fill_bg.lstrip('#')
            red = int(fill_bg[:2], 16)
            green = int(fill_bg[2:4], 16)
            blue = int(fill_bg[4:], 16)
            background_color_int = red | (green << 8) | (blue << 16)
            self.xwrange.api.Interior.Color = background_color_int

        return

    def clear(self):
        self.xwrange.clear()


if __name__ == '__main__':
    wb = xw.Book(r'C:\Users\xli7\Desktop\try.xlsx')
    sheet = wb.sheets[0]  # Reference to the first sheet

    # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
    my_range = sheet.range('E15')

    # Step 3: Create an object of the RangeOperator class with the specified range
    a = RangeOperator(my_range)
    a.format(font_color='FFFF00', align='center', merge=False, fill_pattern='solid')