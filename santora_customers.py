import xlwt
import xlrd

# This script is to process the information of the customers and
# generate another xls file with their fields in the format that Odoo expects.

HEADERS_ROW = 0
KNRD_COL = 0
DATA_FILE_NAME = "Bestellungen T 42_KW 42.xls"
CUSTOMERS_FILE_NAME = "data.xls"
CUSTOMERS_SHEET_NUMBER = 0

workbook_data = xlrd.open_workbook(DATA_FILE_NAME)
old_data_sheet = workbook_data.sheet_by_index(CUSTOMERS_SHEET_NUMBER)

write_workbook = xlwt.Workbook()
new_sheet = write_workbook.add_sheet("customers")

HEADERS_MAP = {
    "Name1": "name",
    "Strasse": "street",
    "PLZ": "zip",
    "Ort": "city",
    "Telefon1": "mobile",
    "Telefon2": "phone",
    "Web": "Website Link",
    "Email1": "email",
}
OLD_HEADERS_COLUMNS = {}
NEW_HEADERS_COLUMNS = {
    "name": 0,
    "street": 1,
    "zip": 2,
    "city": 3,
    "mobile": 4,
    "phone": 5,
    "email": 6,
    "Website Link": 7,
    "Contact/Name": 8,
    "Contact/Mobile": 9,
    "Contact/Address Type": 10,
}


def write_headers_in_new_sheets(old_sheet, n_sheet):
    headers = [
        old_sheet.cell_value(HEADERS_ROW, col)
        for col in range(old_sheet.ncols)
    ]
    col = 0
    for index, value in enumerate(headers):
        if value in HEADERS_MAP:
            value = HEADERS_MAP[value]
            n_sheet.write(HEADERS_ROW, col, value)
            col = col + 1
    n_sheet.write(HEADERS_ROW, col, "Contact/Name")
    col = col + 1
    n_sheet.write(HEADERS_ROW, col, "Contact/Mobile")
    col = col + 1
    n_sheet.write(HEADERS_ROW, col, "Contact/Address Type")


def save_headers_column_number(old_sheet):
    for col in range(old_sheet.ncols):
        OLD_HEADERS_COLUMNS[old_sheet.cell_value(HEADERS_ROW, col)] = col


def write_values(o_sheet, n_sheet):
    clients_knr = []
    write_row = 0
    for row in range(o_sheet.nrows):
        if row == HEADERS_ROW:
            continue
        if o_sheet.cell_value(row, KNRD_COL) not in clients_knr:
            write_row = write_row + 1
            clients_knr.append(o_sheet.cell_value(row, KNRD_COL))
        else:
            continue
        name = ""
        has_contact_1 = False
        has_contact_2 = False
        has_contact_3 = False
        for col in range(o_sheet.ncols):
            header_value = o_sheet.cell_value(HEADERS_ROW, col)
            value = o_sheet.cell_value(row, col)
            if header_value == "Name1" or header_value == "Name2":
                name = name + " " + value if name else name + value
            if header_value == "Strasse" and value:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["street"], value.strip())
            if header_value == "PLZ" and value:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["zip"], value)
            if header_value == "Ort" and value:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["city"], value.strip())
            if header_value == "Email1" and value:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["email"], value.strip())
            if header_value == "Web" and value:
                print(value)
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["Website Link"], value.strip())

            # Adding child contacts
            if header_value == "Kontakt1" and value:
                has_contact_1 = True
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["Contact/Name"], value.strip())
                phone_1 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon1"])
                if phone_1:
                    n_sheet.write(write_row, NEW_HEADERS_COLUMNS["Contact/Mobile"], phone_1)
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["Contact/Address Type"], "contact")
            if header_value == "Kontakt2" and value:
                has_contact_2 = True
                aux_row = write_row
                if has_contact_1:
                    aux_row = aux_row + 1
                n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Name"], value.strip())
                phone_2 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon2"])
                if phone_2:
                    n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Mobile"], phone_2)
                n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Address Type"], "contact")
            if header_value == "Kontakt3" and value:
                has_contact_3 = True
                aux_row = write_row
                if has_contact_1:
                    aux_row = aux_row + 1
                if has_contact_2:
                    aux_row = aux_row + 1
                n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Name"], value.strip())
                phone_3 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon3"])
                if phone_3:
                    n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Mobile"], phone_3)
                n_sheet.write(aux_row, NEW_HEADERS_COLUMNS["Contact/Address Type"], "contact")

        if name:
            n_sheet.write(write_row, NEW_HEADERS_COLUMNS["name"], name.strip())
        if has_contact_1 and has_contact_2:
            write_row = write_row + 1
        elif has_contact_1 and has_contact_3:
            write_row = write_row + 1
        elif has_contact_2 and has_contact_3:
            write_row = write_row + 1
        elif has_contact_1 and has_contact_2 and has_contact_3:
            write_row = write_row + 2

        if not has_contact_1 and not has_contact_2 and not has_contact_3:
            phone_1 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon1"])
            phone_2 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon2"])
            phone_3 = o_sheet.cell_value(row, OLD_HEADERS_COLUMNS["Telefon3"])

            if phone_1 and phone_2:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_1)
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["phone"], phone_2)
            elif phone_2 and phone_3:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_2)
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["phone"], phone_3)
            elif phone_1 and phone_3:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_1)
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["phone"], phone_3)

            if phone_1 and not phone_2 and not phone_3:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_1)
            elif phone_2 and not phone_1 and not phone_3:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_2)
            elif phone_3 and not phone_1 and not phone_2:
                n_sheet.write(write_row, NEW_HEADERS_COLUMNS["mobile"], phone_3)


write_headers_in_new_sheets(old_data_sheet, new_sheet)
save_headers_column_number(old_data_sheet)
write_values(old_data_sheet, new_sheet)
write_workbook.save(CUSTOMERS_FILE_NAME)
