import xlwt
import xlrd
import requests

COMPRESORES_URL = "https://www.compresoresservicios.com/img/pictures/"

workbook_data = xlrd.open_workbook('codes.xls')
sheet_codes = workbook_data.sheet_by_index(0)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet('codes')


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


cont = 0
cont_products = 0
for row in range(sheet_codes.nrows):
    if row == 0:
        continue
    code = sheet_codes.cell_value(row, 0)
    if isinstance(code, (int, float)):
        code = str(int(code))
    have_image = code_have_image(str(code))
    if have_image:
        new_sheet.write(row, 0, str(code))
        cont = cont + 1
    else:
        cont_products = cont_products + 1
print(cont)
print(cont_products)

write_workbook.save('productswithcodesnew.xls')