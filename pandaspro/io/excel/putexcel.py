import re
import textwrap
# import textwrap
from pathlib import Path
import os

import pandas
import pandas as pd
import xlwings as xw
from pandaspro.core.stringfunc import parse_method, str2list
from pandaspro.io.excel._framewriter import FramexlWriter, StringxlWriter, cpdFramexl, CellxlWriter
from pandaspro.io.excel._utils import cell_range_combine, CellPro
from pandaspro.io.excel._xlwings import RangeOperator, parse_format_rule, color_to_int
from pandaspro.utils.cpd_logger import cpd_logger


def is_range_filled(ws, range_str: str = None):
    if range_str is None:
        return False
    else:
        rng = ws.range(range_str)
        for cell in rng:
            if cell.value is not None and str(cell.value).strip() != '':
                return True
        return False


def is_sheet_empty(sheet):
    used_range = sheet.used_range
    if used_range.shape == (1, 1) and not used_range.value:
        return True
    return False


@cpd_logger
class PutxlSet:
    def __init__(
            self,
            workbook: str,
            sheet_name: str = None,
            alwaysreplace: str = None,  # a global config that sets all the following actions to replace ...
            noisily: bool = None,
    ):
        def _extract_filename_from_path(path):
            return Path(path).name

        def _get_open_workbook_by_name(name):
            # Return the open workbook by its name if exists, otherwise return None
            for curr_app in xw.apps:
                for curr_wb in curr_app.books:
                    if curr_wb.name == name:
                        return curr_wb, curr_app
            return None, None

        # App and Workbook declaration
        open_wb, app = _get_open_workbook_by_name(_extract_filename_from_path(workbook))  # Check if the file is already open
        if open_wb:
            if noisily:
                print(f"{workbook} is already open, closing ...")
            open_wb.save()
            open_wb.close()
            if not app.books:  # Check if the app has no more workbooks open; if true, then quit the app
                app.quit()
        elif noisily:
            print(f"Working on {workbook} now ...")

        if not os.path.exists(workbook):  # Check if the file already exists
            open_wb = xw.Book()  # If not, create a new Excel file
            open_wb.save(workbook)
        else:
            open_wb = xw.Book(workbook)

        # Worksheet declaration
        if sheet_name is None:
            sheet_name = open_wb.sheets[0].name

        current_sheets = [sheet.name for sheet in open_wb.sheets]
        if sheet_name in current_sheets:
            sheet = open_wb.sheets[sheet_name]
        else:
            sheet = open_wb.sheets.add(after=open_wb.sheets.count)
            sheet.name = sheet_name

        if 'Sheet1' in current_sheets and is_sheet_empty(open_wb.sheets['Sheet1']) and sheet_name != 'Sheet1':
            open_wb.sheets['Sheet1'].delete()

        self.open_wb, self.app = _get_open_workbook_by_name(_extract_filename_from_path(workbook))  # Check if the file is already open
        self.workbook = workbook
        self.wb = open_wb
        self.ws = sheet
        self.alwaysreplace = alwaysreplace
        self.io = None
        self.curr_cell = None

    # noinspection PyMethodMayBeStatic
    def helpfile(self, para='all'):
        cd_file = """
        cd_format: the main function to add format to core export data ranges (exc. headers and indices)
        This parameter will take a dictionary which allows only three keys (and applyto maybe omitted)
        (refer to the module _cdformat on the class design: _cdformat >> _framewriter.range_cdformat >> _putexcel.PutxlSet.putxl)

        key 1: column = indicating the conditional formatting columns
        key 2: rules = a dictionary with formatting rules (only based on the column above, like inlist, value equals to, etc.)
        key 3: applyto = where to apply, whether column itself or the whole dataframe, or, several selected columns     
        (default = self)

        >>> ... cd_format={'column': 'age', 'rules': {...}}
        >>> ... cd_format={'column': 'grade', 'rules': {'GA':'#FF0000'}, 'applyto': 'self'}
        
        For rules, it should be a dictionary:
        (1) the key can be values in the selected, then the value should be parable format
        (2) the key can also be a rule token, then the value should be a dictionary with r and f
        
        >>> ... {'GA': '#FF0000'}
        >>> ... {   
                    'rule1':  
                        {
                            'r': ['GA', 'GB', 'GC'],
                            'f': 'blue'
                        }，
                    'rule1':
                        {
                            'r': mask,
                            'f': 'green'
                        }
                }
        """
        if para == 'cd_format':
            print(cd_file)

    def putxl(
            self,
            content,
            sheet_name: str = None,
            cell: str = 'A1',
            index: bool = True,
            header: bool = True,
            replace: str = None,
            sheetreplace: bool = None,
            replace_warning: bool = False,
            tab_color: str | tuple = None,

            # Section. String Format
            width=None,
            height=None,
            font: str | tuple = None,
            font_name: str = None,
            font_size: int = None,
            font_color: str | tuple = None,
            italic: bool = None,
            bold: bool = None,
            underline: bool = None,
            strikeout: bool = None,
            number_format: str = None,
            align: str | list = None,
            merge: bool = None,
            wrap: bool = None,
            border: str | list = None,
            fill: str | tuple | list = None,
            fill_pattern: str = None,
            fill_fg: str | tuple = None,
            fill_bg: str | tuple = None,
            color_scale: str = None,
            gridlines: bool = None,
            appendix: bool = False,

            # Section. special/personalize format
            index_merge: dict = None,
            header_wrap: bool = None,
            design: str = None,
            df_style: str | list = None,
            df_format: dict = None,
            cd_style: str | list = None,
            cd_format: list | dict = None,
            config: dict = None,
            mode: str = None,
            log: bool = True,
            debug: str | bool = 'critical',
            debug_file: str | bool = None,
    ) -> None:
        if debug or debug_file:
            self.reconfigure_logger(debug=debug or self.debug, debug_file=debug_file or self.debug_file)

        self.logger.info("")
        self.logger.info(">" * 30)
        self.logger.info(">>>>>>> LOG FOR PUTXL  <<<<<<<")
        self.logger.info(">" * 30)
        self.logger.info(f"> CONTENT: {content if isinstance(content, str) else 'DataFrame with Size of: ' + str(content.shape)}")
        self.logger.info(f"> SHEET_NAME: {sheet_name}")
        self.logger.info(f"> CELL: {cell}")
        self.logger.info("> LOG ACTIVATED - INFO LEVEL")
        self.logger.debug("> LOG ACTIVATED - DEBUG LEVEL")

        # Pre-Cleaning: (1) transfer FramePro to dataframe; (2) change tuple cells to str
        ################################

        # For Framepro objects
        if hasattr(content, 'df'):
            content = content.df

        # For FramexlWriter Objects
        if isinstance(content, FramexlWriter):
            cell = content.start_cell
            index = content.index_bool
            header = content.header_bool
            content = content.content

        # If content's columns is reachable
        if hasattr(content, 'columns'):
            for col in content.columns:
                content[col] = content[col].apply(lambda x: str(x) if isinstance(x, tuple) else x)

        # Sheetreplace? If a sheet_name is specified, then override the current sheet
        ################################
        replace_type = self.alwaysreplace if self.alwaysreplace else replace

        if sheet_name and sheet_name != self.ws.name:
            if sheet_name in [sheet.name for sheet in self.wb.sheets]:
                self.ws = self.wb.sheets[sheet_name]
            else:
                self.ws = self.wb.sheets.add(after=self.wb.sheets.count)
                self.ws.name = sheet_name

        # If sheetreplace or replace is specified, then delete the old sheet and create a new one
        ################################
        if sheetreplace or replace_type == 'sheet':
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = self.ws.index
            original_name = self.ws.name
            total_count = self.wb.sheets.count
            self.info_section_lv1("SECTION: sheetreplace or replace_type")
            self.logger.info(f"Replacing sheet **!'{self.ws.name}'**: [sheetreplace] is declared as **True**, [alwaysreplace] for PutxlSet is declared as **{self.alwaysreplace}**")
            self.logger.info(f"In the workbook, total sheets number is **{total_count}**, while original index is **{original_index}**")

            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
                self.logger.info(f"Sheet <is> the last sheet, new sheet added after the sheet **!'{_sheetmap[original_index]}'**")
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])
                self.logger.info(f"Sheet <is not> the last sheet, new sheet added before the sheet **!'{_sheetmap[original_index + 1]}'**")

            self.ws.delete()
            new_sheet.name = original_name
            self.ws = new_sheet

        # Declare IO Object
        ################################
        self.info_section_lv1("SECTION: content (i.e. IO object) declaration")
        if isinstance(content, str):
            self.logger.info(f"Validation 1: **{content}** is passed as a valid str type object")
            self.logger.info(f"Validation 2: **{content}** value will lead to [CellPro(content).valid] taking the value of **{CellPro(content).valid}**")

            if CellPro(content).valid and mode != 'text':
                io = CellxlWriter(cell=content)
                self.logger.info(f"Passed <Cell>: updating **{content}** format")

            else:
                io = StringxlWriter(text=content, cell=cell)
                # Note: start_cell is named intentional to be consistent with DF mode and may refer to a cell range
                self.logger.info(f"Passed <Text>: filling **{io.content}** into **{io.range_cell}** plus any other format settings ... ")

            RangeOperator(self.ws.range(io.range_cell)).format(
                width=width,
                height=height,
                font=font,
                font_name=font_name,
                font_size=font_size,
                font_color=font_color,
                italic=italic,
                bold=bold,
                underline=underline,
                strikeout=strikeout,
                number_format=number_format,
                align=align,
                merge=merge,
                wrap=wrap,
                border=border,
                fill=fill,
                fill_pattern=fill_pattern,
                fill_fg=fill_fg,
                fill_bg=fill_bg,
                color_scale=color_scale,
                gridlines=gridlines,
                appendix=appendix,
                debug=debug
            )
            self.io = io
            self.ws.range(io.range_cell).value = io.content
            self.curr_cell = CellPro(io.range_cell).offset(1, 0).cell

        elif isinstance(content, pandas.DataFrame):
            io = FramexlWriter(frame=content, cell=cell, index=index, header=header)
            self.ws.range(io.start_cell).value = io.content
            self.io = io
            self.curr_cell = CellPro(io.bottom_left_cell).offset(1, 0).cell

        else:
            raise ValueError(f'Invalid type for parameter [content] as {type(content)} is passed, only takes either str (for cell/text to fill in) or dataframe-like objects.')

        # Format the sheet (Shelley, Li)
        ################################
        '''
        Extra Format (not in the group of format parameters): highlight area in existing-content excel
        This is embedded and will be triggered automatically if not replacing sheet 
        '''
        if replace_warning:
            match_dict = {
                'top': self.io.range_top_empty_checker,
                'bottom': self.io.range_bottom_empty_checker,
                'left': self.io.range_left_empty_checker,
                'right': self.io.range_right_empty_checker
            }
            for direction in list(match_dict.keys()):
                if is_range_filled(self.ws, match_dict[direction]):
                    RangeOperator(self.ws.range(self.io.range_all)).format(border=[direction, 'thicker', '#FF0000'], debug=debug)

        if tab_color:
            self.info_section_lv1("SECTION: tab_color")
            paint_tab = color_to_int(tab_color)
            self.logger.info(f"Setting sheet <{self.ws.name}> tab color to **{tab_color}**, the value was transformed into int **{paint_tab}**")
            self.ws.api.Tab.Color = paint_tab

        if design:
            self.info_section_lv1("SECTION: design")
            self.logger.info("The design argument passed with look up values from the dictionary in excel_table_mydesign.py file in the pandaspro package. Both pre-defined style and cd rules can be passed through 1 design")
            self.logger.info("A str is expected to be used as the lookup key")

            '''
            SPECIAL DESIGN: _index as suffix for design arguemnt:
            -----------------------------------------------------------
            For index_merge, add the _index to the selected design like: wbblue_index(indexname, columnnames)
            This will add index_merge(level=..., columns=...) to the style keys

            For example
            >>> wbblue_index(PGs) === index_merge(level=PGs)
            '''
            from pandaspro.user_config.excel_table_mydesign import excel_export_mydesign as local_design
            if re.fullmatch(r'(.*)_index\(([^,]+),?\s*(.*)\)', design):
                match = re.fullmatch(r'(.*)_index\(([^,]+),?\s*(.*)\)', design)
                design = match.group(1)
                index_key = match.group(2)
                index_columns = match.group(3)
                design_style = local_design[design]['style'] + f"; index_merge({index_key},{index_columns})"
                self.info_section_lv2("Sub-section: _index as suffix for design argument")
                self.logger.info(f"Recognized **{design}**, with extra df_style of **{local_design[design]['style']}** and added **index_merge({index_key}, {index_columns})** ")
            else:
                design_style = local_design[design]['style']
                self.logger.info(f"Recognized **{design}**, with extra df_style of **{design_style}**")

            design_config = local_design[design]['config']
            design_config_shorten_version = {key: design_config[key] for key in list(design_config.keys())[:3]}
            self.logger.info(f"Recognized **{design}**, with extra config of **{design_config_shorten_version}**")
            self.logger.debug(f"Full-length design_config is **{design_config}**")

            design_cd = local_design[design]['cd']
            self.logger.info(f"Recognized **{design}**, with extra style of **{design_cd}**")

            message_warning_design = "Note that the design will not override, but instead added to the df_style, cd_style and config arguments you passed. And it will take effect before df_style, cd_style, ... which further means it could be overwritten by customized claimed arguments"
            wrapped_message = textwrap.fill(message_warning_design, width=120)
            self.logger.info(wrapped_message)
            if df_style:
                df_style = ";".join([design_style, df_style])
            else:
                df_style = design_style

            if config:
                config = config.update(design_config)
            else:
                config = design_config

            if cd_style:
                cd_style = ";".join([design_cd, cd_style])
            else:
                cd_style = design_cd

        '''
        For config para, the accepted dict must use column/index name as keys
        The direct value follow each column/index name must be a dictionary, 
        and there must be readable keys in it.
    
        Currently support: 
        1. width
        2. number_format
    
        For example:
        >>> {
        >>>     'staff id': {'width': 24, 'color': '#00FFFF'},
        >>>     'age': {'width': 15}
        >>>     'salary': {'width': 30, 'haligh': 'left'}
        >>> }
        '''
        if config:
            self.info_section_lv1("SECTION: config")
            self.logger.info(f"Config is taking the value of **{config}**")
            for name, setting in config.items():
                format_update = {k: v for k, v in setting.items() if not pd.isna(v)}
                self.logger.debug(f"Adjusting [{name}]: from config file read format setting: **{format_update}**")
                if name in io.columns_with_indexnames:
                    if debug:
                        print(self.ws.range(io.range_columns(name, header=True)))
                    RangeOperator(self.ws.range(io.range_columns(name, header=True))).format(
                        **format_update,
                        debug=debug
                    )

        '''
        For index_merge para, the accepted dict only accepts two keys:
        1. level: for which level of the index to be set as merge benchmark
        2. columns: for which columns should apply the merge according to the benchmark index
        
        columns can either be a list or a str, and power-wildcard is embedded when using str:
        >>> ['grade', 'staff_id', 'age']
        >>> '* Total' 
        # this will match all columns in the dataframe ends with Total
        
        Example: {'level': 'cmu_dept', 'columns': '*Total'}
        '''
        if index_merge:
            if debug:
                print("================================================")
                print(index_merge)
            for key, local_range in io.range_index_merge_inputs(**index_merge).items():
                if debug:
                    print(key, local_range)
                RangeOperator(self.ws.range(local_range)).format(merge=True, wrap=True, debug=debug)

        if header_wrap:
            RangeOperator(self.ws.range(io.range_header)).format(wrap=True, debug=debug)

        # Format with defined rules using a Dictionary
        def apply_df_format(localinput_format):
            if debug:
                print("Applying Df Format: ")
                print(localinput_format)
            for rule, rangeinput in localinput_format.items():
                # Parse the format to a dictionary, passed to the .format for RangeOperator
                # parse_format_rule is taken from _xlwings module
                format_kwargs = parse_format_rule(rule)
                if debug:
                    print("================================================")
                    print(format_kwargs)

                # Declare range as list/cpdFramexl Object
                def _declare_ranges(local_input):
                    if isinstance(local_input, str):
                        parsedlist = [local_input]
                        if debug:
                            print(parsedlist)
                        cpdframexl_dict = None
                    elif isinstance(local_input, list):
                        parsedlist = local_input
                        cpdframexl_dict = None
                    elif isinstance(local_input, cpdFramexl):
                        parsedlist = None
                        cpdframexl_dict = getattr(io, 'range' + local_input.name)(**local_input.paras)
                    else:
                        raise ValueError('Unsupported type in df_format dictionary values')
                    return parsedlist, cpdframexl_dict

                ioranges, dict_from_cpdframexl = _declare_ranges(rangeinput)
                if debug:
                    print("------------------------------")
                    print("ioranges and dict")
                    print(ioranges)
                    print(dict_from_cpdframexl)

                if ioranges:
                    for each_range in ioranges:
                        if debug:
                            print(">>>>>>>>")
                            print("IO Ranges - Each Range", each_range, type(each_range))
                        # Parse the input string as method name + kwargs
                        range_affix, method_kwargs = parse_method(each_range)[0], parse_method(each_range)[1]
                        if debug:
                            print("Parsing the methods:", range_affix, method_kwargs)
                        attr_method = getattr(io, 'range_' + range_affix)
                        if callable(attr_method):
                            range_cells = attr_method(**method_kwargs)
                        else:
                            range_cells = attr_method

                        if isinstance(range_cells, dict):
                            for range_key, range_content in range_cells.items():
                                if debug:
                                    print("d_format Dictionary Reading This Range", range_content, "as", f'"{range_affix}"')
                                RangeOperator(self.ws.range(range_content)).format(**format_kwargs, debug=debug)
                        elif isinstance(range_cells, str) and range_cells != '':
                            if debug:
                                print("d_format Dictionary Reading This Range", range_cells, "as", range_affix)
                            RangeOperator(self.ws.range(range_cells)).format(**format_kwargs, debug=debug)
                        elif range_cells == '':
                            print('Empty Range Cells: ', range_cells)
                        else:
                            print('Invalid Range Cells')

                if dict_from_cpdframexl:
                    for range_key, range_content in dict_from_cpdframexl.items():
                        RangeOperator(self.ws.range(range_content)).format(**format_kwargs, debug=debug)
            if debug:
                print("")
                print("--------- End of Apply Format ---------")
                print("")
        '''
        style: the main parameter to add pre-defined format to core export data ranges (exc. headers and indices)
        use style_sheets command to view pre-defined formats
        '''
        if df_style:
            from pandaspro.user_config.style_sheets import style_sheets

            # First parse string to lists
            if isinstance(df_style, str):
                loop_list = str2list(df_style)
            elif isinstance(df_style, list):
                loop_list = df_style
            else:
                raise ValueError('Invalid object for style parameter, only str or list accepted')

            # Reorder the items in loop
            checked_dict = {}
            for element in loop_list:
                if element in style_sheets:
                    checked_dict[element] = element
                elif re.match(r'index_merge\(([^,]+),?\s*(.*)\)', element):
                    checked_dict['index_merge'] = element
                else:
                    raise ValueError(f'Specified style {element} not in style sheets')

            checked_list = []
            for key in style_sheets.keys():
                if key in checked_dict.keys():
                    checked_list.append(checked_dict[key])

            # Loop and apply style by checking the style py module
            for each_style in checked_list:
                match = re.match(r'index_merge\(([^,]+),?\s*(.*)\)', each_style)
                if match:
                    index_name = match.group(1)
                    columns = match.group(2) if match.group(2) != '' else 'None'
                    content_border = style_sheets['index_merge']['border=outer_thick']
                    content_border[1] = content_border[1].replace('__index__', index_name)
                    style_sheets['index_merge']['merge'] = style_sheets['index_merge']['merge'].replace(
                        '__index__', index_name).replace('__columns__', columns)
                    apply_style = style_sheets['index_merge']
                else:
                    apply_style = style_sheets[each_style]

                apply_df_format(apply_style)

        '''
        df_format: the main function to add format to ranges
        This parameter will take a dictionary which uses:
        (1) format prompt key words as the keys
        (2) a list of range key words, which may be just a str term (attribute) ... 
            or a cpdFramexl object 

        >>> ... df_format={'msblue80': 'header'}
        >>> ... df_format={'msblue80': cpdFramexl(name='index_merge_inputs', level='cmu_dept_major', columns=['age', 'salary']}
        >>> ... df_format={'blued25; font_color=white': 'columns(['a','b'], header=only)'}

        NOTE! You must specify the kwargs' paras when declaring, like name=, c=, level=, otherwise will be error
        '''
        if df_format:
            apply_df_format(df_format)

        # Conditional Format (1 column based)
        # This function will always check the type of the argument that is passed to this parameter
        # If a list type is detected, then use loop to loop through the list and call the cd_paint function many times
        # So whether dictionary or a list of dictionaries, the format has to comply with standard cpd cd dict format
        # .. which you may refer to the comments before "if cd" line
        def apply_cd_format(input_cd):
            def cd_paint(lcinput):
                if debug:
                    print("================================================")
                    print("Applying Cd Format: ")
                    print(lcinput)
                cleaned_rules = io.range_cdformat(**lcinput)

                # Work with the cleaned_rules to adjust the cell formats in Excel with RangeOperator
                for rulename, lc_content in cleaned_rules.items():
                    cellrange = lc_content['cellrange']
                    cd_format_rule = lc_content['format']
                    if debug:
                        print("Cd format name and content >>>")
                        print(rulename, lc_content)

                    if cellrange == 'no cells':
                        return
                    else:
                        # Parse the cd_format_rule to a dictionary, as **kwargs to be passed to the .format for RangeOperator
                        # parse_format_rule is taken from _xlwings module
                        cd_format_kwargs = parse_format_rule(cd_format_rule)
                        if debug:
                            print("Cd format kwargs >>>")
                            print(cd_format_kwargs)

                        if cellrange == '':
                            pass
                        elif len(cellrange) <= 30:
                            RangeOperator(self.ws.range(cellrange)).format(debug=debug, **cd_format_kwargs)
                        else:
                            # Here is the combine function
                            '''
                            cell_range_combine method from _utils
                            takes a list and returns a dict (from 1 dimension to 2 dimensions)
                            
                            Previously like:
                            'B2,C2,D2,E2,F2,G2,H2,I2,J2,K2,L2,M2,O2,B3' 
                            
                            After combine will be:
                            {2: ['B2:M2', 'O2:O2'], 3: ['B3:B3']}
                            '''
                            cellrange_dict = cell_range_combine(cellrange.split(','))
                            if debug:
                                print(cellrange_dict)
                            for range_list in cellrange_dict.values():
                                for combined_range in range_list:
                                    RangeOperator(self.ws.range(combined_range)).format(debug=debug, **cd_format_kwargs)

            # Decide if cd_format is a dict or not
            if isinstance(input_cd, dict):
                cd_paint(input_cd)

            if isinstance(input_cd, list):
                for rule in input_cd:
                    cd_paint(rule)

        '''
        cd: the main parameter to add pre-defined conditional formatting to core export data ranges (exc. headers and indices)
        use cd_sheets command to view pre-defined formats
        '''
        if cd_style:
            from pandaspro.user_config.cd_sheets import cd_sheets

            # First parse string to lists
            if isinstance(cd_style, str):
                loop_list = str2list(cd_style)
            elif isinstance(cd_style, list):
                loop_list = cd_style
            else:
                raise ValueError('Invalid object for cd parameter, only str or list accepted')

            # Loop and apply cd by checking the cd py module
            for each_cd in loop_list:
                apply_cd = cd_sheets[each_cd]
                if isinstance(apply_cd, list):
                    for each_cd_sub in apply_cd:
                        apply_cd_format(each_cd_sub)
                elif isinstance(apply_cd, dict):
                    apply_df_format(apply_cd)

        '''
        cd_format: the main function to add format to core export data ranges (exc. headers and indices)
        This parameter will take a dictionary which allows only three keys (and applyto maybe omitted)
        (refer to the module _cdformat on the class design: _cdformat >> _framewriter.range_cdformat >> _putexcel.PutxlSet.putxl)

        key 1: column = indicating the conditional formatting columns
        key 2: rules = a dictionary with formatting rules (only based on the column above, like inlist, value equals to, etc.)
        key 3: applyto = where to apply, whether column itself or the whole dataframe, or, several selected columns     
        (default = self)

        >>> ... cd_format={'column': 'age', 'rules': {...}}
        >>> ... cd_format={'column': 'grade', 'rules': {'GA':'#FF0000'}, 'applyto': 'self'}
        >>> ... cd_format={'column': 'grade', 'rules': {'rule1':{'r':...(pd.Series), 'f':...}}, 'applyto': 'self'}
        '''
        if cd_format:
            apply_cd_format(cd_format)

        # Remove Sheet1 if blank and exists (the Default tab) ...
        ################################
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if 'Sheet1' in current_sheets and is_sheet_empty(self.wb.sheets['Sheet1']):
            self.wb.sheets['Sheet1'].delete()

        self.wb.save()

        # Print Export Success Message to Console ...
        ################################
        export_notice_name = self.wb.name
        export_notice_name = export_notice_name.replace('.xlsx', '')[0:35] + ' (...) .xlsx' if len(export_notice_name) > 36 else export_notice_name

        if isinstance(content, str):
            if CellPro(content).valid and mode != 'text':
                print(f"Cell range {content} successfully updated in <<{export_notice_name}>>, worksheet <<{self.ws.name}>> with declared format")
            else:
                print(f"Text {content} successfully filled in <<{export_notice_name}>>, worksheet <<{self.ws.name}>> in cell {cell}")

        elif isinstance(content, pandas.DataFrame):
            print(f"Frame with size {content.shape} successfully exported to <<{export_notice_name}>>, worksheet <<{self.ws.name}>> at cell {cell}")
        # for else, an error should already been thrown in the previous content/io declaration stage

    def tab(self, sheet_name: str, sheetreplace: bool = False, debug: bool = False) -> None:
        """
        Switches to a specified sheet in the workbook.
        If the sheet does not exist, it creates a new one with the given name.

        Parameters
        ----------
        sheet_name : str
            The name of the sheet to switch to or create.
        sheetreplace: bool
            If true, replace the content in the sheet
        debug:
            For developers
        """
        current_sheets = [sheet.name for sheet in self.wb.sheets]
        if sheet_name in current_sheets:
            sheet = self.wb.sheets[sheet_name]
        else:
            sheet = self.wb.sheets.add(after=self.wb.sheets.count)
            sheet.name = sheet_name
        self.ws = sheet

        # If sheetreplace is specified, then delete the old sheet and create a new one
        ################################
        if sheetreplace:
            _sheetmap = {sheet.index: sheet.name for sheet in self.wb.sheets}
            original_index = self.ws.index
            original_name = self.ws.name
            total_count = self.wb.sheets.count

            if original_index == total_count:
                new_sheet = self.wb.sheets.add(after=self.wb.sheets[_sheetmap[original_index]])
            else:
                new_sheet = self.wb.sheets.add(before=self.wb.sheets[_sheetmap[original_index + 1]])

            self.ws.delete()
            new_sheet.name = original_name
            self.ws = new_sheet

        return

    def close(self):
        self.open_wb.close()
        self.app.quit()


if __name__ == '__main__':
    import wbhrdata as wb
    import pandaspro as cpd
    debuglevel = 'info'
    r = wb.impact(analysis_year='FY24', sob_version='2024-05-31', mgr_anchor_version='2023-06-30')
    ps = cpd.PutxlSet('delete_impact_table.xlsx')
    ps.putxl('AFW', cell='A4', font_size=12, bold=True, sheetreplace=True)
    ps.putxl(
        r.table_region('AFW'),
        cell='A5', index=False,
        design='wbblue',
        df_format={
            'font_size=12': 'all',
            'number_format=0.0': 'cspan(s="Overall Rating Average", e="Results")'
        },
        debug=debuglevel
    )