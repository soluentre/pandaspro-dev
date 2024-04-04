style_sheets = {
    'black': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'black; font_color=white': 'header_outer'
    },
    'blue': {
        'border=inner_thin; align=center': 'all',
        'border=outer_thick': 'all',
        'blue80; font_color=black; wrap': 'header_outer'
    },
    'index_merge': {
        'merge': 'index_merge_inputs(level=__index__, columns=__columns__)',
        'border=outer_thick': [
            'index_levels',
            'index_hsections(level=__index__)'
        ]
    }
}
