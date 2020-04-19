# _*_ coding:utf-8 _*_
# author: secdongle
# time: 2020/4/19 10:29
# file: main.py
import xlrd
import argparse
import os


def write_data(line_str):
    with open("ths.txt", 'a', encoding="gbk") as f:
        f.write(line_str + "\n")


def change_format(stock_id, stock_concepts):

    new_row_str = ""
    if stock_id.endswith(".SZ"):
        new_row_str = "0|" + stock_id.replace(".SZ", "") + "|" + stock_concepts + "|0.000"
    elif stock_id.endswith(".SH"):
        new_row_str = "1|" + stock_id.replace(".SH", "") + "|" + stock_concepts + "|0.000"
    return new_row_str


def read_xls(file_path):
    work_book = xlrd.open_workbook(file_path)
    sheets = work_book.sheet_names()
    work_sheet = work_book.sheet_by_name(sheets[0])
    num = 0
    for i in range(0, work_sheet.nrows):
        row = work_sheet.row(i)
        stock_id = str(row[0]).replace("text:", "").replace('\'', "")
        stock_concepts = str(row[2]).replace("text:", "").replace('\'', "").replace(";", " ")
        format_str = change_format(stock_id, stock_concepts)
        print(format_str)
        if format_str:
            write_data(format_str)
            num = num + 1
    print("Write {0} items.".format(num))


def main():
    parser = argparse.ArgumentParser(description="Convert i问财 xls file data to specific format txt file")
    parser.add_argument('-f', metavar="FilePath", type=str, default="", help="set xls file path")
    args = parser.parse_args()
    if os.path.exists("ths.txt"):
        os.remove("ths.txt")
    read_xls(args.f)


if __name__ == '__main__':
    main()
