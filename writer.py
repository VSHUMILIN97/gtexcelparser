import os
import sys

from supports import *
from parser import parse_data


def setup_new_sheet(ws):
    ws.row(0).height_mismatch = True
    ws.row(0).height = int(2*260)
    for cell_index, cell_header in enumerate(ROW_HEADER_LIST):
        ws.col(cell_index).width = int((len(cell_header) + 5) * 260)
        ws.write(0, cell_index, cell_header, set_style(SKY_BLUE_COLOR, header=True))
    del cell_header, cell_index


PARCEL_INDEXES = {}

if __name__ == '__main__':

    parcel_destination_list = '/'
    while not isinstance(parcel_destination_list, list):
        parcel_destination_list = (
            (input('Parcel destination via comma(,): ')
                .replace(' ', '').upper().split(',')
             )
        )

    # Make it predictable and append sort for MIN MAX
    for parcel_city in parcel_destination_list:
        edge_indexes = []
        while not len(edge_indexes) > 1:
            edge_indexes = input(
                (f'Input lowest and highest index for {parcel_city} '
                 'dispatches in the following order - min, max: ')
            ).replace(' ', '').split(',')

        PARCEL_INDEXES[parcel_city] = edge_indexes

    result_of_parser_py = parse_data(PARCEL_INDEXES)

    for parcel_city in result_of_parser_py:
        # First row is the Main index
        RAW_CONST = 1
        dispatch_idx = result_of_parser_py[parcel_city][0].dispatch_num
        wb = XLWTWorkbook()
        ws = wb.add_sheet(f'S{int(dispatch_idx)}')

        setup_new_sheet(ws)
        for idx, gold_tiger in enumerate(result_of_parser_py[parcel_city]):
            if gold_tiger.dispatch_num != dispatch_idx:
                RAW_CONST = 1
                dispatch_idx = gold_tiger.dispatch_num
                ws = wb.add_sheet(f'S{int(dispatch_idx)}')
                setup_new_sheet(ws)
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
            (f'{directory}КОК Депеши '
             f'S{result_of_parser_py[parcel_city][0].dispatch_num}-'
             f'{result_of_parser_py[parcel_city][-1].dispatch_num}.xls'
             )
        )
        del wb, ws, idx, gold_tiger, inner_index, g_t_data, dispatch_idx
