import xlwt
import xlrd

# This file is for create a file with the product codes and their image URL
COMPRESORES_URL = "https://www.compresoresservicios.com/img/pictures/"
HEADERS_ROW = 0
URL_COLUMN = 1
CODES_COLUMN = 0
NEW_IMAGES_COLUMN = 1

workbook_data = xlrd.open_workbook('pimages_old.xls')
sheet_products = workbook_data.sheet_by_index(0)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet("codes_with_image")

new_sheet.write(0, 0, "default_code")
new_sheet.write(0, 1, "image_1920")

def extract_code_from_url(url):
    return url.replace(COMPRESORES_URL, "").replace(".png", "")

custom_row = 0
for row in range(sheet_products.nrows):
    if row == HEADERS_ROW:
        continue
    url = sheet_products.cell_value(row, URL_COLUMN)
    if not url:
        continue
    custom_row = custom_row + 1
    code = extract_code_from_url(url)
    new_sheet.write(custom_row, CODES_COLUMN, code)
    new_sheet.write(custom_row, NEW_IMAGES_COLUMN, url)

write_workbook.save('code_and_url.xls')
