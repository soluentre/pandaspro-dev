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

_border_side_map = {
    'none': None,
    'inner': None,
    'outer': None,
    'all': None,
    'left': 7,
    'top': 8,
    'bottom': 9,
    'right': 10,
    'inner_vert': 11,
    'inner_hor': 12,
    'down_diagonal': 5,
    'up_diagonal': 6
}

_border_style_map = {
    'continue': 1,
    'dash': 2,
    'dot': 3,
    'dash_dot': 4,
    'dash_dot_dot': 5,
    'slant_dash': 6,
    'thick_dash': 8,
    'double': 9,
    'thick_dash_dot_dot': 11
}

_border_weight_map = {
    'thin': 2,
    'thick': 3,
    'thicker': 4
}

_border_custom = {
    'none': None,
    'all_thin': ['all', 'continue', 'thin', '#000000'],
    'all_thick': ['all', 'continue', 'thick', '#000000'],
    'inner_thin': ['inner', 'continue', 'thin', '#000000'],
    'inner_thick': ['inner', 'continue', 'thick', '#000000'],
    'outer_thin': ['outer', 'continue', 'thin', '#000000'],
    'outer_thick': ['outer', 'continue', 'thick', '#000000']
}

_cpdpuxl_color_map = {
    "darkred": "#C00000",
    "red": "#FF0000",
    "orange": "#FFC000",
    "yellow": "#FFFF00",
    "lightgreen": "#92D050",
    "green": "#00B050",
    "lightblue": "#00B0F0",
    "blue": "#0070C0",
    "darkblue": "#002060",
    "purple": "#7030A0",
    "grey": "#808080",
    "grey25": "#BFBFBF",
    "white": "#FFFFFF",
    "msbluegray": "#44546A",
    "msblue": "#4472C4",
    "msorange": "#ED7D31",
    "msgray": "#A5A5A5",
    "msyellow": "#FFC000",
    "msdarkblue": "#5B9BD5",
    "msgreen": "#70AD47",
    "bluegray80": "#D6DCE4",
    "blue80": "#D9E1F2",
    "orange80": "#FCE4D6",
    "gray80": "#EDEDED",
    "yellow80": "#FFF2CC",
    "darkblue80": "#DDEBF7",
    "green80": "#E2EFDA",
    "bluegray60": "#ACB9CA",
    "blue60": "#B4C6E7",
    "orange60": "#F8CBAD",
    "gray60": "#DBDBDB",
    "yellow60": "#FFE699",
    "darkblue60": "#BDD7EE",
    "green60": "#C6E0B4",
    "bluegrayd25": "#333F4F",
    "blued25": "#2F75B5",
    "oranged25": "#C65911",
    "grayd25": "#7B7B7B",
    "yellowd25": "#BF8F00",
    "darkblued25": "#305496",
    "greend25": "#548235",
    "bluegrayd50": "#222B35",
    "blued50": "#1F4E78",
    "oranged50": "#833C0C",
    "grayd50": "#525252",
    "yellowd50": "#806000",
    "darkblued50": "#203764",
    "greend50": "#375623"
}


def print_cell_attributes(file, sheet_name, lcrange):
    lcsheet = xw.Book(file).sheets[sheet_name]
    color_range = lcsheet.range(lcrange)

    cell_colors = {}
    for cell in color_range:
        cell_address = cell.address
        rgb_int = int(cell.api.Font.Color)
        red = rgb_int % 256
        green = (rgb_int // 256) % 256
        blue = (rgb_int // 256 ** 2) % 256
        hex_color = f"#{red:02X}{green:02X}{blue:02X}"
        cell_colors[cell_address] = hex_color

    for address, color in cell_colors.items():
        print(f"Cell {address} has color {color}")


def extract_tuple(s):
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


def _is_valid_hex_color(s):
    pattern = r'^#[0-9A-F]{6}$'
    return bool(re.match(pattern, s, re.IGNORECASE))


def _is_valid_rgb(rgb):
    if not isinstance(rgb, (list, tuple)) or len(rgb) != 3:
        return False
    return all(isinstance(n, int) and 0 <= n <= 255 for n in rgb)


def color_to_int(color: str | tuple):
    if isinstance(color, str) and _is_valid_hex_color(color):
        local_hex = color.lstrip('#')
        red = int(local_hex[:2], 16)
        green = int(local_hex[2:4], 16)
        blue = int(local_hex[4:], 16)
        return red | (green << 8) | (blue << 16)
    elif isinstance(color, tuple) and _is_valid_rgb(color):
        return xw.utils.rgb_to_int(color)
    else:
        return None


def list_str_w_color(mystr: str):
    color, remaining = extract_tuple(mystr)
    result = [i for i in remaining.replace(' ', '').split(',') if i != '']
    if color is not None:
        result = result + [color]

    return result


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
            wrap: bool = None,
            width = None,
            height = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            appendix: bool = False
    ) -> None:

        if appendix:
            print('Please choose one value from the corresponding parameter: \n'
                  f'align: {list(_alignment_map.keys())}; \n'
                  f'fill_pattern: {list(_fpattern_map.keys())};\n'
                  f'border_custom: {list(_border_custom.keys())};\n')

        # Font Attributes
        ##################################
        if font:
            if isinstance(font, tuple):
                self.xwrange.font.color = font
            elif isinstance(font, (int, float)):
                self.xwrange.font.size = font
            elif isinstance(font, str):
                color, remaining = extract_tuple(font)
                if color:
                    self.xwrange.font.color = color
                for item in remaining.split(','):
                    item = item.strip()
                    # noinspection RegExpSimplifiable
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
                    elif re.fullmatch(r'#[\dA-Fa-f]{6}', item):
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
        def _alignfunc(alignkey):
            if alignkey in ['center', 'justify', 'distributed']:
                self.xwrange.api.VerticalAlignment = _alignment_map['v' + alignkey][1]
                self.xwrange.api.HorizontalAlignment = _alignment_map['h' + alignkey][1]
            elif _alignment_map[alignkey][0] == 'v':
                self.xwrange.api.VerticalAlignment = _alignment_map[alignkey][1]
            elif _alignment_map[alignkey][0] == 'h':
                self.xwrange.api.HorizontalAlignment = _alignment_map[alignkey][1]
            elif align not in _alignment_map.keys():
                raise ValueError(f'Alignment {alignkey} is not supported')
            return

        if align:
            if isinstance(align, str):
                for item in align.split(','):
                    item = item.strip()
                    _alignfunc(item)
            elif isinstance(align, list):
                for item in align:
                    _alignfunc(item)

        # Merge and Wrap Attributes
        ##################################
        if merge:
            xw.apps.active.api.DisplayAlerts = False
            self.xwrange.api.MergeCells = merge
            _alignfunc('center')
            xw.apps.active.api.DisplayAlerts = True

        # noinspection PySimplifyBooleanCheck
        if merge == False:
            if self.xwrange.api.MergeCells:
                self.xwrange.unmerge()

        if wrap is not None:
            self.xwrange.api.WrapText = wrap

        # Width and Height Attributes
        ##################################
        '''
        default width and height for a excel cell is 8.54 (around ...) and 14.6
        '''
        if width:
            self.xwrange.api.EntireColumn.ColumnWidth = width

        if height:
            self.xwrange.api.RowHeight = height

        # Border Attributes
        ##################################
        if border:

            if isinstance(border, str) and border.strip() == 'none':
                for i in range(1, 12):
                    self.xwrange.api.Borders(i).LineStyle = 0

            if isinstance(border, str) and border.strip() in list(_border_custom.keys()):
                border_para = _border_custom[border.strip()]

            elif isinstance(border, str) and border.strip() != 'none':
                border_para = list_str_w_color(border)

            elif isinstance(border, list):
                border_para = [i.strip() for i in border]

            else:
                raise ValueError(
                    'Invalid boarder specification, please use check_para=True to see the valid lists.')

            def find_border_side(mylist):
                result = []
                for local_item in mylist:
                    if isinstance(local_item, str) and local_item in list(_border_side_map.keys()):
                        result.append(local_item)
                return result

            def find_border_style(mylist):
                result = []
                for local_item in mylist:
                    if isinstance(local_item, str) and local_item in list(_border_style_map.keys()):
                        result.append(local_item)
                return result

            def find_border_weight(mylist):
                result = []
                for local_item in mylist:
                    if isinstance(local_item, str) and local_item in list(_border_weight_map.keys()):
                        result.append(local_item)
                return result

            def find_border_color(mylist):
                result = []
                for local_item in mylist:
                    if isinstance(local_item, str) and _is_valid_hex_color(local_item):
                        result.append(local_item)
                    elif isinstance(local_item, (list, tuple)) and _is_valid_rgb(local_item):
                        result.append(local_item)
                return result

            # Parse the list and get the Pattern and Color Lists (should be only 1 or none)
            sidelist = find_border_side(border_para)
            stylelist = find_border_style(border_para)
            weightlist = find_border_weight(border_para)
            colorlist = find_border_color(border_para)
            leftover = [item for item in border_para if item not in sidelist + stylelist + weightlist + colorlist]
            if any(len(lst) > 1 for lst in [sidelist, stylelist, weightlist, colorlist]) or len(leftover) > 0:
                raise ValueError(
                    'Invalid input. At most 1 side, 1 style, 1 weight and 1 color can be specified.')

            # Create patter and color parameter
            border_side = sidelist[0] if len(sidelist) == 1 else None
            border_style = stylelist[0] if len(stylelist) == 1 else 'continue'
            border_weight = _border_weight_map[weightlist[0]] if len(weightlist) == 1 else 'thin'
            border_color = color_to_int(colorlist[0]) if len(colorlist) == 1 else '#000000'

            if border_side == 'none':
                for i in range(1, 12):
                    self.xwrange.api.Borders(i).LineStyle = 0

            elif border_side == 'all':
                self.xwrange.api.Borders.LineStyle = _border_style_map[border_style]
                self.xwrange.api.Borders.Weight = border_weight
                self.xwrange.api.Borders.Color = border_color

            elif border_side == 'inner':
                self.xwrange.api.Borders(11).LineStyle = _border_style_map[border_style]
                self.xwrange.api.Borders(11).Weight = border_weight
                self.xwrange.api.Borders(11).Color = border_color
                self.xwrange.api.Borders(12).LineStyle = _border_style_map[border_style]
                self.xwrange.api.Borders(12).Weight = border_weight
                self.xwrange.api.Borders(12).Color = border_color

            elif border_side == 'outer':
                for i in range(7, 11):
                    self.xwrange.api.Borders(i).LineStyle = _border_style_map[border_style]
                    self.xwrange.api.Borders(i).Weight = border_weight
                    self.xwrange.api.Borders(i).Color = border_color

            elif border_side in _border_side_map.keys():
                self.xwrange.api.Borders(_border_side_map[border_side]).LineStyle = _border_style_map[border_style]
                self.xwrange.api.Borders(_border_side_map[border_side]).Weight = border_weight
                self.xwrange.api.Borders(_border_side_map[border_side]).Color = border_color

        # Fill Attributes
        ##################################
        if fill:
            def fill_with_mylist(fill_list):
                def find_pattern(mylist):
                    result = []
                    for local_item in mylist:
                        if isinstance(local_item, (tuple, list, str)) and local_item.lower() in _fpattern_map.keys():
                            result.append(local_item)
                    return result

                def find_colors(mylist):
                    result = []
                    for local_item in mylist:
                        if isinstance(local_item, str) and _is_valid_hex_color(local_item):
                            result.append(local_item)
                        elif isinstance(local_item, (list, tuple)) and _is_valid_rgb(local_item):
                            result.append(local_item)
                    return result

                # Parse the list and get the Pattern and Color Lists (should be only 1 or none)
                patternlist_fill = find_pattern(fill_list)
                colorlist_fill = find_colors(fill_list)
                leftover_fill = [term for term in fill_list if term not in patternlist_fill + colorlist_fill]
                if len(leftover_fill) > 0 or len(patternlist_fill) > 1 or len(colorlist_fill) > 1:
                    raise ValueError(
                        'Invalid input. Please check if pattern or color are specified correctly. At most 1 color and 1 pattern')

                # Create patter and color parameter
                parse_fill_pattern = patternlist_fill[0] if len(patternlist_fill) == 1 else None
                parse_fill_color = colorlist_fill[0] if len(colorlist_fill) == 1 else None

                if parse_fill_pattern:
                    self.xwrange.api.Interior.Pattern = _fpattern_map[parse_fill_pattern.lower()]

                if parse_fill_color:
                    if parse_fill_pattern == 'solid' or parse_fill_pattern is None:
                        self.xwrange.api.Interior.Color = color_to_int(parse_fill_color)
                    else:
                        self.xwrange.api.Interior.PatternColor = color_to_int(parse_fill_color)

            if isinstance(fill, list):
                fill_with_mylist(fill)

            elif isinstance(fill, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill)
                self.xwrange.api.Interior.Color = foreground_color_int

            elif isinstance(fill, str):
                cleanlist = list_str_w_color(fill)
                fill_with_mylist(cleanlist)

        if fill_pattern:
            self.xwrange.api.Interior.Pattern = _fpattern_map[fill_pattern.lower()]

        if fill_fg:
            if isinstance(fill_fg, tuple):
                foreground_color_int = xw.utils.rgb_to_int(fill_fg)
                self.xwrange.api.Interior.PatternColor = foreground_color_int
            elif isinstance(fill_fg, str):
                self.xwrange.api.Interior.PatternColor = color_to_int(fill_fg)

        if fill_bg:
            if isinstance(fill_bg, tuple):
                background_color_int = xw.utils.rgb_to_int(fill_bg)
                self.xwrange.api.Interior.Color = background_color_int
            elif isinstance(fill_bg, str):
                self.xwrange.api.Interior.Color = color_to_int(fill_bg)

        return

    def clear(self):
        self.xwrange.clear()


class cpdStyle:
    def __init__(self, **kwargs):
        self.format_dict = kwargs

    def __hash__(self):
        # Convert the format_dict to a tuple of items, which is hashable, and then hash it
        # Note: This assumes that all values in the dictionary are also hashable
        return hash(tuple(sorted(self.format_dict.items())))

    def __eq__(self, other):
        # Check if the other object is an instance of cpdStyle and if their format_dicts are equal
        return isinstance(other, cpdStyle) and self.format_dict == other.format_dict


def parse_format_rule(rule):
    if isinstance(rule, cpdStyle):
        return rule.format_dict

    elif not isinstance(rule, str):
        raise ValueError('format prompt key word must be str')

    promptlist = [prompt.strip() for prompt in rule.split(',')]
    return_dict = {}

    def _parse_str_format_key(prompt):
        result = {}
        keysmatch = {
            'italic': {'italic': True},
            'noitalic': {'italic': False},
            'bold': {'bold': True},
            'nobold': {'bold': False},
            'underline': {'underline': True},
            'nounderline': {'underline': False},
            'strikeout': {'strikeout': True},
            'nostrikeout': {'strikeout': False},
            'merge': {'merge': True},
            'nomerge': {'merge': False},
            'wrap': {'wrap': True},
            'nowrap': {'wrap': False},
        }
        patterns = {
            r'font_name=(.*)': ['font_name', lambda local_match: local_match.group(1)],
            r'font_size=(.*)': ['font_size', lambda local_match: float(local_match.group(1))],
            r'font_color=(.*)': ['font_color', lambda local_match: local_match.group(1)],
            r'align=(.*)': ['align', lambda local_match: local_match.group(1)],
            r'width=(.*)': ['width', lambda local_match: float(local_match.group(1))],
            r'height=(.*)': ['height', lambda local_match: float(local_match.group(1))],
            r'border=(.*)': ['border', lambda local_match: local_match.group(1)],
            r'(#[A-Z0-9]{6})': ['fill', lambda local_match: local_match.group(1)],
            r'fill=(.*)': ['fill', lambda local_match: local_match.group(1)],
        }

        if prompt in keysmatch.keys():
            result.update(keysmatch['prompt'])

        if prompt in _cpdpuxl_color_map.keys():
            lc_hex = _cpdpuxl_color_map[prompt]
            result.update({'fill': lc_hex})

        for pattern, value in patterns.items():
            match = re.fullmatch(pattern, prompt)
            if match:
                append_dict = {value[0]: value[1](match).replace('"', '').replace('\'', '')}
                result.update(append_dict)

        return result

    for term in promptlist:
        return_dict.update(_parse_str_format_key(term))

    return return_dict

if __name__ == '__main__':
    file = r'C:\Users\wb539289\OneDrive - WBG\Desktop\temp.xlsx'
    print_cell_attributes(file, 'Sheet1', 'A1:A14')
