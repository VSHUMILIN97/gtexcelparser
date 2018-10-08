import os
import collections

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


def parse_data(new_dispatch_index, new_high_dispatch_index):
    excel = '/'
    while not os.path.isfile(excel):
        excel = input('Provide path to excel: ')

    workbook = xlrd.open_workbook(excel)

    sheet = workbook.sheet_by_index(0)
    clean_data = []
    for row in range(1, sheet.nrows):
        parsed_data = sheet.row_values(row)
        dispatch_index = parsed_data[6]
        if parsed_data[5] != 'EKA':
            continue
        try:
            if int(dispatch_index) > new_high_dispatch_index:
                continue
            if int(dispatch_index) < new_dispatch_index < int(sheet.row_values(row - 1)[6]):
                iterator = 0
                fake_row = row
                while True:
                    fake_parser = sheet.row_values(fake_row)
                    if fake_parser[5] != 'EKA':
                        continue
                    if int(fake_parser[6]) != int(sheet.row_values(fake_row + 1)[6]):
                        break
                    else:
                        UPU_bag_number = fake_parser[10]
                        UPU_box_number = fake_parser[9]
                        bag_gross_weight = fake_parser[11]
                        number_of_items_in_bag = fake_parser[12]
                        bag_ipk = fake_parser[13]
                        clean_data.append(GoldTiger(
                            dispatch_num=int(dispatch_index),
                            UPU_bag=int(UPU_bag_number),
                            UPU_box=UPU_box_number,
                            bag_weight=bag_gross_weight,
                            items_in_bag=int(number_of_items_in_bag),
                            bag_ipk=bag_ipk
                        ))
                    iterator += 1
                    fake_row += 1
                row += iterator
                pass
            elif int(dispatch_index) < new_dispatch_index:
                continue
        except ValueError:
            continue
        UPU_bag_number = parsed_data[10]
        UPU_box_number = parsed_data[9]
        bag_gross_weight = parsed_data[11]
        number_of_items_in_bag = parsed_data[12]
        bag_ipk = parsed_data[13]
        clean_data.append(GoldTiger(
            dispatch_num=int(dispatch_index),
            UPU_bag=int(UPU_bag_number),
            UPU_box=UPU_box_number,
            bag_weight=bag_gross_weight,
            items_in_bag=int(number_of_items_in_bag),
            bag_ipk=bag_ipk
        ))

    return clean_data
