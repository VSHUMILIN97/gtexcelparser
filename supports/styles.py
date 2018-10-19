import xlwt
from .colors import *


def set_style(colour, splitter=None, header=None):
    """ Set style for the current cell

    Args:
        colour: cell colour
        splitter: condition for the special cell style
        header: condition for the special cell style

    Returns:
        style: style object for xlwt
    """
    style = xlwt.XFStyle()
    pattern = xlwt.Pattern()
    alignment = xlwt.Alignment()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    alignment.horz = xlwt.Alignment.HORZ_CENTER
    alignment.vert = xlwt.Alignment.VERT_CENTER
    pattern.pattern_fore_colour = xlwt.Style.colour_map[colour]
    # Cell borders
    borders = xlwt.Borders()
    borders.left = 1
    borders.right = 1
    borders.top = 1
    borders.bottom = 1
    style.alignment = alignment
    style.borders = borders
    style.pattern = pattern
    if header:
        style.font.bold = True
        style.font.colour_index = xlwt.Style.colour_map[WHITE_COLOR]
    else:
        pass
    if splitter:
        pass
    else:
        style.num_format_str = r'#,##0.00'
    return style


PALE_BLUE_STYLE = set_style(PALE_BLUE_COLOR)
WHITE_STYLE = set_style(WHITE_COLOR)
LIGHT_GREEN_STYLE = set_style(LIGHT_GREEN_COLOR)
PALE_BLUE_STYLE_WO_COMMAS = set_style(PALE_BLUE_COLOR, splitter=True)
WHITE_STYLE_WO_COMMAS = set_style(WHITE_COLOR, splitter=True)
