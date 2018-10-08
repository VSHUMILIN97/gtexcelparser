import os
import collections
import copy

import xlrd

__all__ = (
    'GoldTiger',
    'parse_data'
)

GoldTiger = collections.namedtuple(
    'GoldTiger',
    [
        'dispatch_num',
        'UPU_bag',
        'UPU_box',
        'bag_weight',
        'items_in_bag',
        'bag_ipk'
    ]
)
# Current order of fields
# {'Corridor': 0, 'Clevy WareHouse Code': 1, 'Clevy Shipment No': 2,
# 'Clevy Shipment Date': 3, 'AWB': 4, 'Offloading airport': 5,
# 'Dispatch No': 6, 'Clevy Batch Number': 7, 'Clevy Box Number': 8,
# 'UPU Box Number': 9, 'UPU Bag number': 10, 'Bag Gross Weight (in Kg)': 11,
# 'Number of items in Bag': 12, 'Bag IPK': 13, 'Is bag IPK <= 13 ?': 14}


def parse_data(dispatch_indexes_dict):
    clean_data_for_return = []

    header_indexes = {}

    for city in dispatch_indexes_dict:
        current_city = dispatch_indexes_dict[city]
        dispatch_indexes_dict[city] = (
            [x for x in range(int(current_city[0]), int(current_city[1])+1)]
        )
    excel_file = '/'
    while not os.path.isfile(excel_file):
        excel_file = input('Provide path to excel: ')

    workbook = xlrd.open_workbook(excel_file)

    sheet = workbook.sheet_by_index(0)

    for index, name in enumerate(sheet.row_values(0)):
        header_indexes[name.replace('\n', '')] = index
    for city in dispatch_indexes_dict:
        print(city, dispatch_indexes_dict[city])
        for dispatch_index in dispatch_indexes_dict[city]:
            for row in range(1, sheet.nrows):
                data_in_row = sheet.row_values(row)
                if data_in_row[6] == dispatch_index and data_in_row[5] == city:
                    UPU_bag_number = data_in_row[10]
                    UPU_box_number = data_in_row[9]
                    bag_gross_weight = data_in_row[11]
                    number_of_items_in_bag = data_in_row[12]
                    bag_ipk = data_in_row[13]
                    clean_data_for_return.append(GoldTiger(
                        dispatch_num=int(dispatch_index),
                        UPU_bag=int(UPU_bag_number),
                        UPU_box=UPU_box_number,
                        bag_weight=bag_gross_weight,
                        items_in_bag=int(number_of_items_in_bag),
                        bag_ipk=bag_ipk
                    ))
        dispatch_indexes_dict[city] = copy.deepcopy(clean_data_for_return)
        clean_data_for_return.clear()
    return dispatch_indexes_dict
