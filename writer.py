import argparse
import os
import sys
from operator import attrgetter

from supports import *
from parser import parse_data


# TODO REFACTOR THIS FOR THE LOVE OF GOD
def return_parsed_args():
    parser = argparse.ArgumentParser(description='Process new columns.')
    parser.add_argument('dispatch_low', metavar='N', type=int, nargs='+',
                    help='an integer for the parser')
    parser.add_argument('dispatch_high', metavar='M', type=int, nargs='+',
                    help='an integer for the parser')
    args = parser.parse_args()
    return args


def setup_new_sheet():
    ws.row(0).height_mismatch = True
    ws.row(0).height = int(2*260)
    for cell_index, cell_header in enumerate(ROW_HEADER_LIST):
        ws.col(cell_index).width = int((len(cell_header) + 5) * 260)
        ws.write(0, cell_index, cell_header, set_style(SKY_BLUE_COLOR, header=True))

RAW_CONST = 1
DISPATCH_INDEXES = set()

if __name__ == '__main__':
    parsed_args = return_parsed_args()
    result_of_parser_py = sorted(parse_data(parsed_args.dispatch_low[0], parsed_args.dispatch_high[0]), key=attrgetter('dispatch_num'))
    # First row is the Main index
    dispatch_idx = result_of_parser_py[0].dispatch_num
    wb = XLWTWorkbook()
    ws = wb.add_sheet(f'S{int(dispatch_idx)}')

    setup_new_sheet()
    for idx, gold_tiger in enumerate(result_of_parser_py):
        DISPATCH_INDEXES.add(gold_tiger.dispatch_num)
        if gold_tiger.dispatch_num != dispatch_idx:
            RAW_CONST = 1
            dispatch_idx = gold_tiger.dispatch_num
            ws = wb.add_sheet(f'S{int(dispatch_idx)}')
            setup_new_sheet()
        for inner_index, g_t_data in enumerate(gold_tiger):
            if RAW_CONST % 2 != 0:
                try:
                    ws.col(inner_index).width = int((len(g_t_data) + 4)*260)
                except TypeError:
                    pass
                if isinstance(g_t_data, float):
                    ws.write(RAW_CONST, inner_index, float(g_t_data), PALE_BLUE_STYLE)
                elif isinstance(g_t_data, int):
                    ws.write(RAW_CONST, inner_index, int(g_t_data), PALE_BLUE_STYLE_WO_COMMAS)
                else:
                    ws.write(RAW_CONST, inner_index, g_t_data, PALE_BLUE_STYLE)
            else:
                try:
                    ws.col(inner_index).width = int((len(g_t_data) + 4)*260)
                except TypeError:
                    pass
                if isinstance(g_t_data, float):
                    ws.write(RAW_CONST, inner_index, float(g_t_data), WHITE_STYLE)
                elif isinstance(g_t_data, int):
                    ws.write(RAW_CONST, inner_index, int(g_t_data), WHITE_STYLE_WO_COMMAS)
                else:
                    ws.write(RAW_CONST, inner_index, g_t_data, WHITE_STYLE)
        if gold_tiger.bag_ipk > 13:
            ws.write(RAW_CONST, len(gold_tiger), '')
        else:
            ws.write(RAW_CONST, len(gold_tiger), '', LIGHT_GREEN_STYLE)
        RAW_CONST += 1

    wb.worksheets.sort(key=lambda name: int(name.name.replace('S', '')))
    directory = ''
    while not os.path.isdir(directory):
        directory = input('Provide directory to store the file: ')
    if sys.platform == 'win32':
        directory = fr'{directory}\\'
    else:
        directory = f'{directory}/'
    wb.save(
        f'{directory}КОК Депеши S{parsed_args.dispatch_low[0]}-{parsed_args.dispatch_high[0]}.xls'
    )
