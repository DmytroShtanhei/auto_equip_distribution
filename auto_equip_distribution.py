"""
Script
for processing "Договір" and "Групування" .xlsx source files
and creating "Рознарядка" .xlsx file
"""

from openpyxl import load_workbook
import utils
from named_styles import header_style, data_style
# from operator import itemgetter
import locale
locale.setlocale(locale.LC_ALL, "")

distribution_wb = load_workbook(filename='Рознарядка.xlsx')
contract_ws = distribution_wb['Договір']

original_grouping_wb = load_workbook(filename='Групування.xlsx')
original_grouping_ws = original_grouping_wb.active

# Copy the original grouping_ws to distribution_wb and prepare
if 'Групування' in distribution_wb:
    del distribution_wb['Групування']

grouping_copied_ws = distribution_wb.create_sheet('Групування')
utils.copy_table(source_ws=original_grouping_ws, target_ws=grouping_copied_ws)
utils.prepare_grouping_table(grouping_ws_to_be_prepared=grouping_copied_ws)
utils.style_table_in_worksheet(workbook=distribution_wb,
                               worksheet=grouping_copied_ws,
                               custom_header_style=header_style,
                               custom_data_style=data_style,
                               max_header_row=4,
                               )

if 'Рознарядка' in distribution_wb:
    del distribution_wb['Рознарядка']

distribution_ws = distribution_wb.create_sheet('Рознарядка')

# Get list of positions from Contract table
position_list = utils.get_position_list(contract_ws)
# Get list of LVU from Grouping table
lvu_list = utils.get_lvu_list(grouping_copied_ws)

# Get distribution data list
distribution_data_list = utils.get_distribution_data_list(position_list,
                                                          grouping_copied_ws,
                                                          )

# get_distribution_full_list
distribution_full_list = utils.get_distribution_full_list(position_list,
                                                          lvu_list,
                                                          distribution_data_list,
                                                          )

# Replace LVU codes in distribution_full_list with LVU names
utils.replace_lvu_codes_with_names(distribution_full_list)

# Sort rows by LVU
# distribution_full_list_sorted_by_lvu = sorted(distribution_full_list, key=itemgetter(0))
distribution_full_list_sorted_by_lvu = sorted(distribution_full_list, key=lambda s: (locale.strxfrm(s[0])))
# print(distribution_full_list_sorted_by_lvu)

# Create header and first column (numbers in order) for distribution spreadsheet
utils.init_table_in_distribution_ws(position_list,
                                    lvu_list,
                                    distribution_ws,
                                    )

# Populate distribution_ws with distribution data from distribution_full_list
utils.populate_table_in_distribution_ws(distribution_ws,
                                        lvu_list,
                                        position_list,
                                        distribution_full_list_sorted_by_lvu,
                                        )

# Style the table in distribution spreadsheet
utils.style_table_in_worksheet(workbook=distribution_wb,
                               worksheet=distribution_ws,
                               custom_header_style=header_style,
                               custom_data_style=data_style,
                               )

# Save distribution workbook
distribution_wb.save(f'Рознарядка.xlsx')
