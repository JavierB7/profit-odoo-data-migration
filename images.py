import xlwt
import xlrd
import requests

COMPRESORES_URL = "https://www.compresoresservicios.com/img/pictures/"
HEADERS_ROW = 0
IDS_COLUMN = 0
CODES_COLUMN = 1
NEW_IMAGES_COLUMN = 1
IMAGES_FILE_NAME = "images.xls"


workbook_data = xlrd.open_workbook('product.template.xls')
sheet_products = workbook_data.sheet_by_index(0)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet('products')


def get_image_url(code):
    return f'{COMPRESORES_URL}{str(code)}.png'


def get_list_of_codes_with_image():
    code_and_url_data = xlrd.open_workbook("code_and_url.xls")
    sheet_cd = code_and_url_data.sheet_by_index(0)
    return [sheet_cd.cell_value(sheet_row, 0) for sheet_row in range(sheet_cd.nrows)]


def code_image_in_web(code):
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


def match_code_with_image(code):
    codes = get_list_of_codes_with_image()
    return code in codes


def code_has_image(code, in_web):
    if in_web:
        return code_image_in_web(code)
    else:
        return match_code_with_image(code)


new_sheet.write(0, 0, "id")
new_sheet.write(0, 1, "image_1920")
for row in range(sheet_products.nrows):
    if row == HEADERS_ROW:
        continue
    id = sheet_products.cell_value(row, IDS_COLUMN)
    product_code = sheet_products.cell_value(row, CODES_COLUMN)
    has_image = code_has_image(str(product_code), False)
    if has_image:
        url = get_image_url(str(product_code))
        new_sheet.write(row, IDS_COLUMN, id)
        new_sheet.write(row, NEW_IMAGES_COLUMN, url)

write_workbook.save(IMAGES_FILE_NAME)
print(f"File {IMAGES_FILE_NAME} created!")
