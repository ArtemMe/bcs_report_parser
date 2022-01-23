import xlrd
from zipfile import ZipFile
from rarfile import RarFile
import xlsxwriter
import os

TABLE_INSTRUMENTS_TYPE = ["Акция", "Пай", "АДР"] # добавить АДР пример FIVE_ADR
TABLE_INSTRUMENTS_COLUMN = ["", "Дата", "Номер", "Время", "Куплено, шт", "Цена", "Сумма платежа", "Продано, шт", "Цена продажи",
                     "Сумма выручки", "Валюта", "Валюта платежа", "Дата соверш.", "Время соверш.", "Тип сделки",
                     "Оплата (факт)", "Поставка (факт)", "Место сделки"]
SECOND_COLUMN = 1
INTRUMENTS_TRANSLATION = {
    "HK_486": "RUAL",
    "FXWO_RM": "FXWO",
    "NVTK_02": "NVTK",
    "ROSN": "MCX:ROSN",
    "FXCN_RM": "FXCN",
    "BABA_US": "BABA",
    "FIVE_ADR": "FIVE",
    "RU000A1000F9": "SBGB",
    "FXRU.MRG": "FXRU",
}
IGNORING_INSTRUMENTS = ["GAZP2", "SKFL_01"]


def get_table_value(table_row, column, sheet):
    return sheet.row(table_row)[column].value


def find_deals(sheet):
    DEALS_INDEX = 1
    DEALS_TITLE = "2.1. Сделки:"
    for rx in range(sheet.nrows):
        if sheet.row(rx)[DEALS_INDEX].value == DEALS_TITLE:
            print(sheet.row(rx))
            return rx


def parse_instruments(sheet, start_row):
    if start_row is None:
        return {}
    instrument_map = {}
    index = start_row + 2
    new_section = "3. Активы:"

    while index < sheet.nrows:
        instrument_type = get_table_value(index, SECOND_COLUMN, sheet)
        if instrument_type == new_section:
            break

        if instrument_type in TABLE_INSTRUMENTS_TYPE:
            ticket_pair = parse_instruments_table(sheet, index + 1, instrument_type)
            instrument_map[instrument_type] = ticket_pair[0]
            index = ticket_pair[1]
        else:
            index = index + 1
            continue
    return instrument_map


def parse_instruments_table(sheet, start_row, instrument_type):
    index = start_row + 2
    currency_string = "Валюта цены"
    empty_string = ""
    repo = "в т.ч. по репо:"
    ticket_map = {}
    while index < sheet.nrows:
        if get_table_value(index, SECOND_COLUMN, sheet) == empty_string:
            break

        if get_table_value(index, SECOND_COLUMN, sheet).find(currency_string) != -1:
            index = index + 2
            continue

        if get_table_value(index, SECOND_COLUMN, sheet).find(repo) != -1:
            index = index + 1
            continue

        ticket_name = get_table_value(index, SECOND_COLUMN, sheet)
        row_list = parse_instrument_deals(sheet, index, instrument_type, ticket_name)
        ticket_map[ticket_name] = row_list
        index = index + len(row_list) + 2
    return ticket_map, index


def parse_instrument_deals(sheet, start_row, instrument_type, ticket_name):
    row_list = []
    end_table_string = "Итого по"
    for rx in range(start_row + 1, sheet.nrows):
        if sh.row(rx)[TABLE_INSTRUMENTS_COLUMN.index("Дата")].value.find(end_table_string) != -1:
            break
        row = InstrumentRow(
            get_ticket_name(ticket_name),
            sh.row(rx)[TABLE_INSTRUMENTS_COLUMN.index("Дата")].value,
            get_price(sh.row(rx)),
            get_deals_type(sh.row(rx)),
            sh.row(rx)[TABLE_INSTRUMENTS_COLUMN.index("Валюта")].value,
            instrument_type,
            get_amount(sh.row(rx))
        )

        row_list.append(row)
    return row_list


def get_ticket_name(ticket_name):
    if INTRUMENTS_TRANSLATION.get(ticket_name) is not None:
        return INTRUMENTS_TRANSLATION.get(ticket_name)
    return ticket_name


def get_deals_type(table_row):
    if table_row[TABLE_INSTRUMENTS_COLUMN.index("Куплено, шт")].value == "":
        return DealsType.BUY
    else:
        return DealsType.SELL


def get_price(table_row):
    if table_row[TABLE_INSTRUMENTS_COLUMN.index("Куплено, шт")].value == "":
        return table_row[TABLE_INSTRUMENTS_COLUMN.index("Цена продажи")].value
    else:
        return table_row[TABLE_INSTRUMENTS_COLUMN.index("Цена")].value


def get_amount(table_row):
    if table_row[TABLE_INSTRUMENTS_COLUMN.index("Куплено, шт")].value == "":
        return -1 * table_row[TABLE_INSTRUMENTS_COLUMN.index("Продано, шт")].value
    else:
        return table_row[TABLE_INSTRUMENTS_COLUMN.index("Куплено, шт")].value


def unzip_files(path, path_to_extract):
    zf = ZipFile(path, 'r')
    zf.extractall(path_to_extract)
    zf.close()


def unrar_files(path, path_to_extract):
    with RarFile(path) as rf:
        rf.extractall(path=path_to_extract)


def find_zip_archive(path, path_to_extract):
    entries = os.listdir(path)
    if not os.path.isdir(path_to_extract):
        os.mkdir(path_to_extract, 7777)
    for zip_archive in entries:
        if zip_archive.find("zip") != -1:
            unzip_files(path+"/"+zip_archive, path_to_extract)
        if zip_archive.find("rar") != -1:
            unrar_files(path+"/"+zip_archive, path_to_extract)


def write_row_to_xls(worksheet, row, instrument_row):
    col = 0

    worksheet.write(row, col, instrument_row.ticket_name)
    worksheet.write(row, col+1, instrument_row.price)
    worksheet.write(row, col+2, instrument_row.amount)
    worksheet.write(row, col+4, instrument_row.date)
    worksheet.write(row, col+5, instrument_row.currency)
    worksheet.write(row, col+6, instrument_row.instrument_type)


def instrument_filter(instrument_row):
    if instrument_row.ticket_name in IGNORING_INSTRUMENTS:
        return False
    return True


class DealsType:
    SELL = 0
    BUY = 1


class InstrumentRow(object):

    def __init__(self, ticket_name, date, price, deals_type, currency, instrument_type, amount):
        self.ticket_name = ticket_name
        self.date = date
        self.price = price
        self.deals_type = deals_type
        self.currency = currency
        self.instrument_type = instrument_type
        self.amount = amount
        pass


if __name__ == '__main__':
    dir_with_deals_arhive = 'report_bcs'
    dir_with_unpacked_files = 'result'

    find_zip_archive(dir_with_deals_arhive, dir_with_unpacked_files)

    entries = os.listdir(dir_with_unpacked_files)
    row = 1
    result_report_xsl_name = "common_report.xlsx"
    workbook = xlsxwriter.Workbook(result_report_xsl_name)
    worksheet = workbook.add_worksheet()

    for report in entries:
        if report.find("xls") != -1:
            book = xlrd.open_workbook(dir_with_unpacked_files+"/"+report)
            sh = book.sheet_by_index(0)

            start_row = find_deals(sh)
            print(start_row)
            parsed_instruments = parse_instruments(sh, start_row)

            for instruments in parsed_instruments.values():
                for deals in instruments.values():
                    for instrumet_row in deals:
                        if instrument_filter(instrumet_row):
                            write_row_to_xls(worksheet, row, instrumet_row)
                            row = row + 1

    workbook.close()


# install

# pip install pandas
# pip install openpyxl
# pip install xlwt
# from zipfile import ZipFile
# pip install pyunpack
# pip install patool
# pip install xlsxwriter

#you have to install unrar programm to use patool!!
# brew install carlocab/personal/unrar
# **************
# Example fo xlrd
# **************

#     book = xlrd.open_workbook(report_file_2)
#     print("The number of worksheets is {0}".format(book.nsheets))
#     print("Worksheet name(s): {0}".format(book.sheet_names()))
#     sh = book.sheet_by_index(0)
#     print("{0} {1} {2}".format(sh.name, sh.nrows, sh.ncols))
#     print("Cell D30 is {0}".format(sh.cell_value(rowx=29, colx=3)))