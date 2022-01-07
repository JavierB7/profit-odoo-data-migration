import xlwt
import xlrd

PRODUCTS_SHEET = 7
HEADERS_ROW = 0
ATTRIBUTES_COLUMN = 8
VALUES_ATTRIBUTES_COLUMN = 9
PRODUCT_CODE_COLUMN = 10

workbook_data = xlrd.open_workbook('data.xls')
sheet_products = workbook_data.sheet_by_index(PRODUCTS_SHEET)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet('test')

headers = [sheet_products.cell_value(0, col) for col in range(sheet_products.ncols)]


def write_headers_in_sheet(sheet):
    for cell_index, cell_value in enumerate(headers):
        if cell_index < 8:
            sheet.write(HEADERS_ROW, cell_index, cell_value)
        if cell_index > PRODUCT_CODE_COLUMN:
            sheet.write(HEADERS_ROW, PRODUCT_CODE_COLUMN, cell_value)
    sheet.write(HEADERS_ROW, ATTRIBUTES_COLUMN, "Atributos del producto / Atributo")
    sheet.write(HEADERS_ROW, VALUES_ATTRIBUTES_COLUMN, "Atributos del producto / Valores")


custom_row = 1
custom_sheet = new_sheet
write_headers_in_sheet(custom_sheet)
for row in range(sheet_products.nrows):
    if row == HEADERS_ROW:
        continue
    if row % 700 == 0:
        custom_row = 1
        custom_sheet = write_workbook.add_sheet("test%d" % row)
        write_headers_in_sheet(custom_sheet)

    products_row = []
    attributes_row = []
    for col in range(sheet_products.ncols):
        products_row.append(sheet_products.cell_value(row, col))
    for index, value in enumerate(products_row):
        if index < ATTRIBUTES_COLUMN:
            custom_sheet.write(custom_row, index, value)
        if index > PRODUCT_CODE_COLUMN:
            custom_sheet.write(custom_row, PRODUCT_CODE_COLUMN, value)
    model = products_row[8]
    if model is not '':
        custom_sheet.write(custom_row, ATTRIBUTES_COLUMN, "Modelo")
        custom_sheet.write(custom_row, VALUES_ATTRIBUTES_COLUMN, model)
        custom_row = custom_row + 1
    warranty = products_row[9]
    if warranty is not '':
        custom_sheet.write(custom_row, ATTRIBUTES_COLUMN, "Garantia")
        custom_sheet.write(custom_row, VALUES_ATTRIBUTES_COLUMN, warranty)
        custom_row = custom_row + 1
    brand = products_row[10]
    if brand is not '':
        custom_sheet.write(custom_row, ATTRIBUTES_COLUMN, "Marca")
        custom_sheet.write(custom_row, VALUES_ATTRIBUTES_COLUMN, brand)
        custom_row = custom_row + 1


write_workbook.save('products.xls')

