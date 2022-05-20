import xlwt
import xlrd

# This script is to process the information of the product and generate another xls file with their attributes
# in the format that Odoo expects. The new file is divided into several sheets for the upload process in Odoo.

DATA_FILE_NAME = "prod.xls"
PRODUCT_FILE_NAME = "products_capacitors.xls"
PRODUCTS_SHEET_NUMBER = 0
HEADERS_ROW = 0
MODEL_COLUMN = 8
WARRANTY_COLUMN = 9
BRAND_COLUMN = 10
ATTRIBUTES_COLUMN = 15
VALUES_ATTRIBUTES_COLUMN = 16
ATTRIBUTES_COLUMN_HEADER = "Atributos del producto / Atributo"
VALUES_ATTRIBUTES_COLUMN_HEADER = "Atributos del producto / Valores"

workbook_data = xlrd.open_workbook(DATA_FILE_NAME)
sheet_products = workbook_data.sheet_by_index(PRODUCTS_SHEET_NUMBER)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet("test0")

headers = [sheet_products.cell_value(0, col) for col in range(sheet_products.ncols)]

def write_headers_in_sheet(sheet):
    for cell_index, cell_value in enumerate(headers):
        if cell_index not in [MODEL_COLUMN, WARRANTY_COLUMN, BRAND_COLUMN]:
            sheet.write(HEADERS_ROW, cell_index, cell_value)
    sheet.write(HEADERS_ROW, ATTRIBUTES_COLUMN, ATTRIBUTES_COLUMN_HEADER)
    sheet.write(HEADERS_ROW, VALUES_ATTRIBUTES_COLUMN, VALUES_ATTRIBUTES_COLUMN_HEADER)
