import xlwt
import xlrd

# This script is to process the information of the product and generate another xls file with their attributes
# in the format that Odoo expects. The new file is divided into several sheets for the upload process in Odoo.

DATA_FILE_NAME = "MASI_FALTANTES_3.xls"
PRODUCT_FILE_NAME = "files/products.xls"
PRODUCTS_SHEET_NUMBER = 0
HEADERS_ROW = 0
MODEL_COLUMN = 7
BRAND_COLUMN = 8
ATTRIBUTES_COLUMN = 23
VALUES_ATTRIBUTES_COLUMN = 24
ATTRIBUTES_COLUMN_HEADER = "Atributos del producto / Atributo"
VALUES_ATTRIBUTES_COLUMN_HEADER = "Atributos del producto / Valores"

workbook_data = xlrd.open_workbook(DATA_FILE_NAME)
sheet_products = workbook_data.sheet_by_index(PRODUCTS_SHEET_NUMBER)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet("test0")

headers = [sheet_products.cell_value(0, col) for col in range(sheet_products.ncols)]

def write_headers_in_sheet(sheet):
    for cell_index, cell_value in enumerate(headers):
        if cell_index not in [MODEL_COLUMN, BRAND_COLUMN]:
            sheet.write(HEADERS_ROW, cell_index, cell_value)
    sheet.write(HEADERS_ROW, ATTRIBUTES_COLUMN, ATTRIBUTES_COLUMN_HEADER)
    sheet.write(HEADERS_ROW, VALUES_ATTRIBUTES_COLUMN, VALUES_ATTRIBUTES_COLUMN_HEADER)


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
        if index not in [MODEL_COLUMN, BRAND_COLUMN]:
            custom_sheet.write(custom_row, index, value)
    model = products_row[MODEL_COLUMN]
    if model is not '':
        custom_sheet.write(custom_row, ATTRIBUTES_COLUMN, "Modelo")
        custom_sheet.write(custom_row, VALUES_ATTRIBUTES_COLUMN, model)
        custom_row = custom_row + 1
    brand = products_row[BRAND_COLUMN]
    if brand is not '':
        custom_sheet.write(custom_row, ATTRIBUTES_COLUMN, "Marca")
        custom_sheet.write(custom_row, VALUES_ATTRIBUTES_COLUMN, brand)
        custom_row = custom_row + 1


write_workbook.save(PRODUCT_FILE_NAME)
print(f"File {PRODUCT_FILE_NAME} created!")
