import xlrd
from zipfile import ZipFile
from rarfile import RarFile
import xlsxwriter
import os
import re
from datetime import datetime

TABLE_INSTRUMENTS_COLUMN = ["", "Дата", "Номер", "Время", "Куплено, шт", "Цена", "Сумма платежа", "Продано, шт", "Цена продажи",
                     "Сумма выручки", "Валюта", "Валюта платежа", "Дата соверш.", "Время соверш.", "Тип сделки",
                     "Оплата (факт)", "Поставка (факт)", "Место сделки"]
SECOND_COLUMN = 1


def get_table_value(table_row, column, sheet):
    return sheet.row(table_row)[column].value


def find_assets_section(sheet):
    DEALS_INDEX = 1
    DEALS_TITLE = "3. Активы:"
    for rx in range(sheet.nrows):
        if sheet.row(rx)[DEALS_INDEX].value == DEALS_TITLE:
            print(sheet.row(rx))
            return rx


def find_period_date(sheet):
    DEALS_INDEX = 1
    DEALS_TITLE = "Период:"
    for rx in range(sheet.nrows):
        if sheet.row(rx)[DEALS_INDEX].value == DEALS_TITLE:
            print(sheet.row(rx))
            return rx


def parse_period_date(sheet, start_row):
    if start_row is None:
        return {}

    period_date = get_table_value(start_row, 5, sheet)

    return parse_period_date_str(period_date)


# example "с 01.06.2021 по 30.06.2021"
def parse_period_date_str(period_date):
    periods = re.findall('(\d\d\.\d\d\.\d\d\d\d)', period_date)
    return periods[0], periods[1]


def find_general_agreement(sheet):
    DEALS_INDEX = 1
    DEALS_TITLE = "Генеральное соглашение:"
    for rx in range(sheet.nrows):
        if sheet.row(rx)[DEALS_INDEX].value == DEALS_TITLE:
            print(sheet.row(rx))
            return rx


def parse_general_agreement(sheet, start_row):
    if start_row is None:
        return ''
    period_date = get_table_value(start_row, 5, sheet)
    return re.split("\s", period_date)[0]


def parse_start_end_portofolio_cost(sheet, start_row):
    if start_row is None:
        return {}
    portfolio_cost = {}
    index = start_row + 2

    while index < sheet.nrows:
        column_value = get_table_value(index, SECOND_COLUMN, sheet)
        if column_value == "Стоимость портфеля (руб.):":
            start_period = get_table_value(index, 7, sheet)
            end_period = get_table_value(index, 11, sheet)
            return PortfolioCost(start_period, end_period)
        index = index + 1

    return portfolio_cost


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


def write_row_to_xls(worksheet, row, portfolio_return):
    col = 0

    worksheet.write(row, col, portfolio_return.start_period_cost)
    worksheet.write(row, col+1, portfolio_return.end_period_cost)
    worksheet.write(row, col + 2, portfolio_return.ratio)
    worksheet.write(row, col + 3, portfolio_return.start_period_date)
    worksheet.write(row, col + 4, portfolio_return.end_period_date)
    worksheet.write(row, col + 5, portfolio_return.general_agreement)


def mergeBills(portfolio_return_list):
    portfolio_map = {}
    for portfolio_return_row in portfolio_return_list:
        list = portfolio_map.get(portfolio_return_row.start_period_date)
        if list is None:
            portfolio_map[portfolio_return_row.start_period_date] = [portfolio_return_row]
            continue
        list.append(portfolio_return_row)
        portfolio_map[portfolio_return_row.start_period_date] = list

    portfolio_return_merged_list = []
    for portfolio_list in portfolio_map.values():
        first_account = portfolio_list[0]

        if len(portfolio_list) > 1:
            second_account = portfolio_list[1]
            start_period_cost = first_account.start_period_cost + second_account.start_period_cost
            end_period_cost = first_account.end_period_cost + second_account.end_period_cost
            ratio = (end_period_cost - start_period_cost)*100/start_period_cost

            portfolio_return_merged_list.append(
                PortfolioReturnRow(
                    start_period_cost,
                    end_period_cost,
                    ratio,
                    first_account.start_period_date,
                    first_account.end_period_date,
                    first_account.general_agreement
                )
            )
            continue

        portfolio_return_merged_list.append(first_account)

    return portfolio_return_merged_list


class PortfolioCost(object):
    def __init__(self, start_period_cost, end_period_cost):
        self.start_period_cost = int(start_period_cost)
        self.end_period_cost = int(end_period_cost)

    def getRatio(self):
        if self.start_period_cost == 0:
            return 0
        return (self.end_period_cost-self.start_period_cost)*100/self.start_period_cost


class PortfolioReturnRow(object):
    def __init__(
            self, start_period_cost, end_period_cost, ratio, start_period_date, end_period_date, general_agreement):
        self.start_period_cost = float(start_period_cost)
        self.end_period_cost = float(end_period_cost)
        self.start_period_date = start_period_date
        self.end_period_date = end_period_date
        self.ratio = float(ratio)
        self.general_agreement = general_agreement

    def getRatio(self):
        if self.start_period_cost == 0:
            return 100
        return (self.end_period_cost-self.start_period_cost)*100/self.start_period_cost


if __name__ == '__main__':
    dir_with_deals_arhive = 'report_bcs'
    dir_with_unpacked_files = 'result'

    find_zip_archive(dir_with_deals_arhive, dir_with_unpacked_files)

    entries = os.listdir(dir_with_unpacked_files)

    portfolio_return_list = []

    for report in entries:
        if report.find("xls") != -1:
            book = xlrd.open_workbook(dir_with_unpacked_files+"/"+report)
            sh = book.sheet_by_index(0)

            period_date_row = find_period_date(sh)
            start_period, end_period = parse_period_date(sh, period_date_row)

            general_agreement_row = find_general_agreement(sh)
            general_agreement = parse_general_agreement(sh, general_agreement_row)

            assets_section_row = find_assets_section(sh)

            portfolio_cost = parse_start_end_portofolio_cost(sh, assets_section_row)

            portfolio_return_row = PortfolioReturnRow(
                portfolio_cost.start_period_cost,
                portfolio_cost.end_period_cost,
                portfolio_cost.getRatio(),
                start_period,
                end_period,
                general_agreement
            )

            portfolio_return_list.append(portfolio_return_row)

    result_report_xsl_name = "analytics_report.xlsx"
    workbook = xlsxwriter.Workbook(result_report_xsl_name)
    worksheet = workbook.add_worksheet()

    row = 1
    col = 0
    worksheet.write(row, col, "Start period cost")
    worksheet.write(row, col + 1, "End period cost")
    worksheet.write(row, col + 2, "Ratio")
    worksheet.write(row, col + 3, "Start date")
    worksheet.write(row, col + 4, "End date")
    worksheet.write(row, col + 4, "General agreement")

    portfolio_return_merged_list = mergeBills(portfolio_return_list)
    date_pattern = '%d.%m.%Y'
    portfolio_return_merged_list.sort(key=lambda x:  datetime.strptime(x.start_period_date, date_pattern), reverse=False)

    for portfolio_return in portfolio_return_merged_list:
        write_row_to_xls(worksheet, row, portfolio_return)
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