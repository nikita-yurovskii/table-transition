import docx
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Cm, Emu
from docx.enum.table import WD_ROW_HEIGHT_RULE
import math
import itertools

def EMU_to_inch(value):
    return round(value / 914400, 4)
def set_cell_widths_in_table(table, width_map, mr, mc):
    """
    Sets the width of each cell in a Word table according to a width map.

    Args:
        table (docx.table.Table): The Word table object.
        width_map (dict): Dictionary where the key is a tuple (row, col) of cell coordinates (starting from 0),
                           and the value is the width in Excel format (e.g., "A", "AA", "AZ").
    """
    table.autofit = False
    table.allow_autofit = False

    width_map = dict(itertools.islice(width_map.items(), len(table.columns)))
    set_col_widths(table, width_map)


def set_cell_height_in_table(table, heights, mr, mc):
    print("_+_=-=-=_+-+_+_+HEIGHTS,", heights)
    """
    Sets the width of each cell in a Word table according to a width map.

    Args:
        table (docx.table.Table): The Word table object.
        width_map (dict): Dictionary where the key is a tuple (row, col) of cell coordinates (starting from 0),
                           and the value is the width in Excel format (e.g., "A", "AA", "AZ").
    """
    table.autofit = False
    table.allow_autofit = False

    height_map = dict(itertools.islice(heights.items(), len(table.columns)))


    set_col_heights(table, height_map)
def set_col_widths(table, width):
    widths = excel_width_to_inches(width)
    for row in table.rows:
        for idx, width in enumerate(widths):
            row.cells[idx].width = Inches(width)
            #print(idx,width)


def set_col_heights(table, heights):
    heights = excel_height_to_inches(heights)
    for row in table.rows:
        for idx, height in enumerate(heights):
            row.cells[idx].width = Inches(height)
            print(idx,height)

def excel_width_to_inches(width_map):
    """
    Converts an Excel numeric column width to inches.

    This conversion depends on the default font used in Excel (usually Calibri 11).
    The formula below is empirical and needs to be calibrated.

    Args:
        excel_width (float or int): The numeric column width from Excel.

    Returns:
        float: Width in inches.
    """


    sum = 0
    a = []
    for (row, col), excel_width in width_map.items():
        sum+=excel_width
    for (row, col), excel_width in width_map.items():
        a.append(excel_width/sum)

    for i in range(len(a)):
        a[i] *= ((6.728))


    return a



def excel_height_to_inches(height_map):
    """
    Converts an Excel numeric column width to inches.

    This conversion depends on the default font used in Excel (usually Calibri 11).
    The formula below is empirical and needs to be calibrated.

    Args:
        excel_width (float or int): The numeric column width from Excel.

    Returns:
        float: Width in inches.
    """

    # Empirical formula for conversion.  Requires calibration.
    #  The crucial constant here is based on the default font (Calibri 11)
    #  and the average width of characters in that font.  This *will* need tuning.

    a = []
    for (row, col), excel_width in height_map.items():
        a.append(excel_width)


    return a