import xlwt
import xlrd


def write_headers_in_new_sheets(first_new_sheet, second_new_sheet):
    headers = [
        sheet_inventory.cell_value(HEADERS_ROW, col)
        for col in range(sheet_inventory.ncols)
    ]
    for index, value in enumerate(headers):
        first_new_sheet.write(HEADERS_ROW, index, value)
        second_new_sheet.write(HEADERS_ROW, index, value)
    first_new_sheet.write(1, 0, sheet_inventory.cell_value(1, 0))
    second_new_sheet.write(1, 0, sheet_inventory.cell_value(1, 0))


def extract_columns_from_inventory():
    product_columns = []
    quantity_columns = []
    location_columns = []
    for inventory_row in range(sheet_inventory.nrows):
        if inventory_row != HEADERS_ROW:
            product_columns.append(
                sheet_inventory.cell_value(
                    inventory_row, INVENTORY_PRODUCTS_COLUMN
                )
            )
            quantity_columns.append(
                sheet_inventory.cell_value(
                    inventory_row, INVENTORY_QUANTITIES_COLUMN
                )
            )
            location_columns.append(
                sheet_inventory.cell_value(
                    inventory_row, INVENTORY_LOCATIONS_COLUMN
                )
            )
    return {
        'products': product_columns,
        'quantities': quantity_columns,
        'locations': location_columns
    }


def extract_columns_from_stocks():
    stocks_ids_columns = []
    stocks_name_columns = []
    for stock_row in range(sheet_stocks.nrows):
        if stock_row != HEADERS_ROW:
            stocks_ids_columns.append(
                sheet_stocks.cell_value(stock_row, STOCK_ID_COLUMN))
            stocks_name_columns.append(
                sheet_stocks.cell_value(stock_row, STOCK_NAME_COLUMN))
    return {
        'names': stocks_name_columns,
        'ids': stocks_ids_columns
    }


def replace_stock_names_for_ids():
    stock_columns = extract_columns_from_stocks()
    stocks_ids = []
    for row in range(sheet_inventory.nrows):
        if row != HEADERS_ROW:
            stock_to_search = sheet_inventory.cell_value(
                row, INVENTORY_LOCATIONS_COLUMN
            )
            if isinstance(stock_to_search, (int, float)):
                stock_to_search = str(int(stock_to_search))
            index_of_stock = stock_columns['names'].index(stock_to_search)
            id_of_stock = stock_columns['ids'][index_of_stock]
            stocks_ids.append(id_of_stock)
    return stocks_ids


def split_list_in_half(list_to_split):
    length = len(list_to_split)
    middle_index = length // 2
    return {
        'first_half': list_to_split[:middle_index],
        'second_half': list_to_split[middle_index:]
    }


def get_columns_to_strings(inventory, stocks):
    if not SPLIT_DOCUMENT:
        return {
            'products': {
                'first': ','.join(map(str, inventory['products'])),
                'second': ''
            },
            'quantities': {
                'first': ','.join(map(str, inventory['quantities'])),
                'second': ''
            },
            'locations': {
                'first': ','.join(map(str, stocks)),
                'second': ''
            }
        }
    products = split_list_in_half(inventory['products'])
    quantities = split_list_in_half(inventory['quantities'])
    locations = split_list_in_half(stocks)
    return {
        'products': {
            'first': ','.join(map(str, products['first_half'])),
            'second': ','.join(map(str, products['second_half']))
        },
        'quantities': {
            'first': ','.join(map(str, quantities['first_half'])),
            'second': ','.join(map(str, quantities['second_half']))
        },
        'locations': {
            'first': ','.join(map(str, locations['first_half'])),
            'second': ','.join(map(str, locations['second_half']))
        }
    }


def write_in_new_document(data):
    new_sheet_inventory1.write(
        1, INVENTORY_PRODUCTS_COLUMN, data['products']['first'])
    new_sheet_inventory1.write(
        1, INVENTORY_QUANTITIES_COLUMN, data['quantities']['first'])
    new_sheet_inventory1.write(
        1, INVENTORY_LOCATIONS_COLUMN, data['locations']['first'])
    if SPLIT_DOCUMENT:
        new_sheet_inventory2.write(
            1, INVENTORY_PRODUCTS_COLUMN, data['products']['second'])
        new_sheet_inventory2.write(
            1, INVENTORY_QUANTITIES_COLUMN, data['quantities']['second'])
        new_sheet_inventory2.write(
            1, INVENTORY_LOCATIONS_COLUMN, data['locations']['second'])
    write_workbook.save('initial.xls')


if __name__ == "__main__":
    HEADERS_ROW = 0
    STOCK_FILE_NAME = 'stock.location.xls'
    STOCK_ID_COLUMN = 0
    STOCK_NAME_COLUMN = 1
    INVENTORY_FILE_NAME = 'inventory.xls'
    INVENTORY_PRODUCTS_COLUMN = 1
    INVENTORY_QUANTITIES_COLUMN = 2
    INVENTORY_LOCATIONS_COLUMN = 3
    SPLIT_DOCUMENT = False
    MAX_NUMBER_OF_PRODUCTS_PER_SHEET = 100

    workbook_stocks = xlrd.open_workbook(STOCK_FILE_NAME)
    sheet_stocks = workbook_stocks.sheet_by_index(0)
    workbook_inventory = xlrd.open_workbook(INVENTORY_FILE_NAME)
    sheet_inventory = workbook_inventory.sheet_by_index(0)

    write_workbook = xlwt.Workbook()
    new_sheet_inventory1 = write_workbook.add_sheet('inventory1')
    new_sheet_inventory2 = write_workbook.add_sheet('inventory2')

    write_headers_in_new_sheets(new_sheet_inventory1, new_sheet_inventory2)
    inventory_columns = extract_columns_from_inventory()
    if len(inventory_columns['products']) > MAX_NUMBER_OF_PRODUCTS_PER_SHEET:
        SPLIT_DOCUMENT = True
    stock_ids = replace_stock_names_for_ids()
    new_columns_data = get_columns_to_strings(inventory_columns, stock_ids)
    write_in_new_document(new_columns_data)
    print("initial.xls successfully created!")
