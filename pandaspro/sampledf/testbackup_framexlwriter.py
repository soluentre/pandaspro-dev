# wb = xw.Book('sampledf.xlsx')
# sheet = wb.sheets[0]  # Reference to the first sheet
#
# # Step 2: Specify the range you want to work with in Excel, e.g., "A1:B2"
# # my_range = sheet.range("H2:I4")
# my_range = sheet.range("F8:I11")
#
# # Step 3: Create an object of the RangeOperator class with the specified range
# # a = RangeOperator(my_range)
# # a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border='outer, thicker')    # print(a.range)
# # a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border=['inner', 'thin'])    # print(a.range)
# # a.format(width=20, height=15)    # print(a.range)
#
# # my_range = sheet.range("A1:B12")
# # a = RangeOperator(my_range)
# # style = cpdStyle(font=['bold', 'strikeout', 12.5, (0,0,0)])
# # a.format(**style.format_dict)
#
# print_cell_attributes('sampledf.xlsx', 'Sheet3', 'A1:A34')
# print(parse_format_rule('red, font_size=12'))
# a = RangeOperator(my_range)
# a.format(font=['bold', 'strikeout', 12.5, (0, 0, 0)], fill='horstripe, (0,255,0)',
#          border='thicker, inner, #FF00FF')  # print(a.range)
# # a.format(font=['bold', 'strikeout', 12.5, (0,0,0)], border=['inner', 'thin'])    # print(a.range)