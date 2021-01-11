from openpyxl import load_workbook

work_book = load_workbook('../GetData/申请代码.xlsx')

# 查找长度有问题的申请代码，手动更正
def clean(begin_year, end_year):
    work_sheet_list = [work_book[str(year)] for year in range(begin_year, end_year + 1)]
    for year, work_sheet in enumerate(work_sheet_list):
        col_code = work_sheet['A']
        print(year+begin_year)
        for i, c in enumerate(col_code):
            if len(c.value) != 3 and len(c.value) != 5 and len(c.value) != 7:
                print(i, c, c.value, len(c.value))


if __name__ == '__main__':
    clean(2008, 2020)