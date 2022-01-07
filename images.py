import xlwt
import xlrd
import requests

COMPRESORES_URL = "https://www.compresoresservicios.com/img/pictures/"
HEADERS_ROW = 0
IDS_COLUMN = 0
CODES_COLUMN = 1
NEW_IMAGES_COLUMN = 1


workbook_data = xlrd.open_workbook('product.template.xls')
sheet_products = workbook_data.sheet_by_index(0)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet('products')


def get_image_url(code):
    return f'{COMPRESORES_URL}{str(code)}.png'


def code_have_image(code):
    url = get_image_url(code)
    try:
        res = requests.get(url, timeout=5)
        if res.status_code != requests.codes.ok:
            return False
    except requests.exceptions.ConnectionError as e:
        return False
    except requests.exceptions.Timeout as e:
        return False
    return True


new_sheet.write(0, 0, "id")
new_sheet.write(0, 1, "image_1920")
for row in range(sheet_products.nrows):
    if row == HEADERS_ROW:
        continue
    id = sheet_products.cell_value(row, IDS_COLUMN)
    product_code = sheet_products.cell_value(row, CODES_COLUMN)
    have_image = code_have_image(str(product_code))
    if have_image:
        url = get_image_url(str(product_code))
        new_sheet.write(row, IDS_COLUMN, id)
        new_sheet.write(row, NEW_IMAGES_COLUMN, url)

write_workbook.save('images.xls')