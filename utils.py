"""
Utility functions
for processing "Договір" and "Групування" .xlsx source files
and creating "Рознарядка" .xlsx file
"""


def copy_table(source_ws, target_ws):
    """Copy table from Source Spreadsheet to Target Spreadsheet of Another Workbook"""
    # Get source table as a list
    source_table_as_list = []
    for row in source_ws.rows:
        row_as_list = []
        for cell in row:
            row_as_list.append(cell.value)
        source_table_as_list.append(row_as_list)

    # Paste table as a list to target worksheet
    for row in source_table_as_list:
        target_ws.append(row)


def prepare_grouping_table(grouping_ws_to_be_prepared):
    """Code for automated preparing of copied grouping spreadsheet."""
    # Spread names of position inside each group of formerly merged cells in the position colon.
    for col in grouping_ws_to_be_prepared.iter_cols(min_col=13, max_col=13, min_row=5):
        previous_cell_val = 'Не визначено'
        for cell in col:
            if cell.value is None:
                cell.value = previous_cell_val
            previous_cell_val = cell.value


def get_positions_n_units_list(contract_ws):
    """Get position list from contract spreadsheet"""
    # contract_position_n_units_list = []
    # for col in contract_ws.iter_cols(min_row=3, max_col=1, values_only=True):
    #     for val in col:
    #         if val is not None:
    #             contract_position_n_units_list.append(val)

    contract_positions_n_units_list = []
    for row in contract_ws.rows:
        index_row = row[0].row
        if index_row >= 3:
            position_n_unit = [row[0].value, row[14].value]
            contract_positions_n_units_list.append(position_n_unit)

    return contract_positions_n_units_list


def get_lvu_list(grouping_ws):
    """Get LVU list from distribution spreadsheet"""
    distribution_lvu_list = []
    for col in grouping_ws.iter_cols(min_row=5, min_col=8, max_col=8, values_only=True):
        for val in col:
            if val is not None:
                distribution_lvu_list.append(val)

    return sorted(list(set(distribution_lvu_list)))


def get_lvu_list_for_position(grouping_ws, position):
    """Get list of LVU which ordered given position."""
    lvu_list_for_pos = []
    for row in grouping_ws.iter_rows(min_row=5, min_col=2, max_col=13, values_only=True):
        if row[11] == position:
            lvu_list_for_pos.append(row[6])

    return sorted(list(set(lvu_list_for_pos)))


def get_distribution_data_list(positions_n_units_list, grouping_ws):
    """Create list with raw data of sum per position per lvu: lvu | position | sum"""
    distribution_data_list = []
    for item in positions_n_units_list:
        lvu_list_for_position = get_lvu_list_for_position(grouping_ws=grouping_ws, position=item[0])
        for curr_lvu in lvu_list_for_position:
            dist_data_row = []
            curr_sum = 0
            for row in grouping_ws.iter_rows(min_row=5, min_col=2, max_col=13, values_only=True):
                if row[11] == item[0] and row[6] == curr_lvu:
                    curr_sum += row[2]
            # print(f'{curr_lvu} {item[0]} {curr_sum}')
            dist_data_row.append(curr_lvu)
            dist_data_row.append(item[0])
            dist_data_row.append(curr_sum)
            # print(dist_data_row)
            distribution_data_list.append(dist_data_row)

    return distribution_data_list


def get_sum_from_distribution_data_list(distribution_data_list, lvu, position):
    """Get needed sum from raw data list for given LVU and position"""
    # print(distribution_data_list)
    for item in distribution_data_list:
        # print(item)
        if lvu == item[0] and position == item[1]:
            return item[2]

    return 0


def get_distribution_full_list(positions_n_units_list, lvu_list, distribution_data_list):
    """
    Create distribution data full list in form of asking table:
    lvu | sum_for_pos_1 | ... | sum_for_pos_n
    """
    distribution_full_list = []
    for curr_lvu in lvu_list:
        curr_lvu_list = [curr_lvu]
        for item in positions_n_units_list:
            sum_for_lvu_for_pos = get_sum_from_distribution_data_list(distribution_data_list, curr_lvu, item[0])
            curr_lvu_list.append(sum_for_lvu_for_pos)
        distribution_full_list.append(curr_lvu_list)

    return distribution_full_list


def replace_lvu_codes_with_names(distribution_full_list):
    """Replace codes in distribution_full_list with names from lvu_names_dict"""
    # LVU names dictionary
    lvu_names_dict = {
        7001: 'ТОВ «ОГТСУ» апарат',
        7102: 'Харківське ЛВУМГ',
        7103: 'Запорізьке ЛВУМГ',
        7104: 'Запорізьке ЛВУМГ (Криворізьке)',
        7105: 'Миколаївське ЛВУМГ',
        7106: 'Краматорське ЛВУМГ',
        7107: 'Краматорське ЛВУМГ (Сєвєродонецьке)',
        7202: 'Кременчуцьке ЛВУМГ',
        7203: 'Золотоніське ЛВУМГ',
        7204: 'Золотоніське ЛВУМГ (Барське)',
        7205: 'Миколаївське ЛВУМГ(Одеське)',
        7302: 'Сумське ЛВУМГ',
        7304: 'Лубенське ЛВУМГ',
        7305: 'Боярське ЛВУМГ',
        7306: 'Бердичівське ЛВУМГ',
        7402: 'Богородчанське ЛВУМГ',
        7403: 'Богородчанське ЛВУМГ (Долинське)',
        7404: 'Закарпатське ЛВУМГ',
        7405: 'Бібрське ЛВУМГ',
        7406: 'Бібрське ЛВУМГ (Рівненське)',
        7502: 'ОРУ',
    }

    for curr_lvu_list in distribution_full_list:
        curr_lvu_code = curr_lvu_list[0]
        curr_lvu_list[0] = lvu_names_dict[curr_lvu_code]


def init_table_in_distribution_ws(positions_n_units_list, lvu_list, distribution_ws):
    """Create header for distribution spreadsheet in form of asked table:
    lvu | sum_for_pos_1 | ... | sum_for_pos_n
    """
    # Initialise header list
    distribution_header_list = [
        '№ п-п',
        'Назва ЛВУМГ (ЛВУМГ що замовляло)',
    ]
    # Append positions to header list
    for item in positions_n_units_list:
        distribution_header_list.append(item[0])

    # Populate distribution_ws header row
    for row in distribution_ws.iter_rows(max_row=1, max_col=len(distribution_header_list)):
        p = 0
        for cell in row:
            cell.value = distribution_header_list[p]
            p += 1

    # Populate Number in order column
    for col in distribution_ws.iter_cols(max_col=1, min_row=2, max_row=len(lvu_list) + 1):
        num_in_order = 1
        for cell in col:
            cell.value = num_in_order
            num_in_order += 1


def populate_table_in_distribution_ws(distribution_ws, lvu_list, positions_n_units_list, distribution_full_list):
    """Populate distribution_ws with distribution data from distribution_full_list"""
    curr_lvu_list_index = 0
    for row in distribution_ws.iter_rows(min_row=2,
                                         max_row=len(lvu_list)+1,
                                         min_col=2,
                                         max_col=len(positions_n_units_list) + 2,
                                         ):
        curr_item_in_lvu_list_index = 0
        for cell in row:
            value = distribution_full_list[curr_lvu_list_index][curr_item_in_lvu_list_index]
            if value == 0:
                cell.value = None
            else:
                cell.value = value
            curr_item_in_lvu_list_index += 1
        curr_lvu_list_index += 1


def auto_adjust_col_width(worksheet):
    """Change columns width according to columns value length"""
    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter  # Get the column name
        for cell in col:
            try:  # Necessary to avoid error on empty cells
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width


def style_table_in_worksheet(workbook, worksheet, custom_header_style, custom_data_style, max_header_row=1):
    """Style the table in a given worksheet"""
    # Register custom Named Styles in the Workbook if they are not registered yet.
    if custom_header_style.name not in workbook.named_styles:
        workbook.add_named_style(custom_header_style)
    if custom_data_style.name not in workbook.named_styles:
        workbook.add_named_style(custom_data_style)

    # Add Named Styles to cells
    for row in worksheet.rows:
        row_index = row[0].row
        if row_index <= max_header_row:
            for cell in row:
                cell.style = custom_header_style.name
        else:
            for cell in row:
                cell.style = custom_data_style.name
                # if cell.column > 2:
                #     cell.number_format = '0.000'

    # Change column width
    worksheet.column_dimensions['B'].width = 37 + 2
