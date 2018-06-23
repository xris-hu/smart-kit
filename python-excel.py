from xlrd import open_workbook
import xlrd
import xlwt
import sys
import os
import datetime
import time

def readsheet(s, row_count=-1, col_cout=-1):
    nrows = s.nrows
    ncols = s.ncols

    row_index = 0
    while row_index < nrows:
        yield [s.cell(row_index, col).value for col in range(ncols)]
        row_index += 1

def ParseExcel(file):
    if not os.path.exists(file):
        print("file: ", file, " not exist.")
        sys.exit()
    d = {}
    list = []
    wb = open_workbook(file)
    count = 0

    for s in wb.sheets():
        for row in readsheet(s, -1, -1):
            if count == 0:
                count = count +1
                continue

            collectTime = xlrd.xldate.xldate_as_datetime(row[1], 0)

            if collectTime < datetime.datetime(1990, 1, 1, 1, 0, 0, 0):
                continue

            date = collectTime.strftime('%Y-%m-%d')
            if row[0] not in d:
                list.append(row[0])
                month = {}
                day = {}
                day["morning"] = collectTime
                day["night"] = collectTime
                month[date] = day
                d[row[0]] = month
            else:
                month = d[row[0]]
                if date not in month:
                    day = {}
                    day["morning"] = collectTime
                    day["night"] = collectTime
                    month[date] = day
                    d[row[0]] = month
                else:
                    if collectTime > month[date]["night"] and collectTime > month[date]["morning"] :
                        month[date]["night"] = collectTime
                    elif collectTime < month[date]["morning"] and collectTime < month[date]["night"]:
                        month[date]["morning"] = collectTime
                    d[row[0]] = month

    # f = open("员工考勤表.txt", "w")
    #
    #
    # for name in list:
    #     for date in sorted(d[name].keys()):
    #         print("姓名: ", name, " 日期: ", date, "早晨: ", d[name][date]["morning"])
    #         morning = "姓名: " + str(name) + " 日期: " + str(date) + " 早晨: " +  str(d[name][date]["morning"])
    #         f.write(morning)
    #         f.write('\n')
    #
    #         print("姓名: ", name, " 日期: ", date, "晚上: ", d[name][date]["night"])
    #         night = "姓名: " + str(name) + " 日期: " + str(date) + " 晚上: " +  str(d[name][date]["night"])
    #         f.write(night)
    #         f.write('\n')
    #
    #
    # f.close()

    file = xlwt.Workbook('utf-8')
    table = file.add_sheet('考勤')

    index = 0
    for name in list:
        for date in sorted(d[name].keys()):
            table.write(index, 0, "姓名")
            table.write(index, 1, str(name))
            table.write(index, 2, "日期")
            table.write(index, 3, str(date))
            table.write(index, 4, "早晨")
            table.write(index, 5,  str(d[name][date]["morning"]))

            table.write(index + 1 , 0, "姓名")
            table.write(index + 1, 1, str(name))
            table.write(index + 1, 2, "日期")
            table.write(index + 1, 3, str(date))
            table.write(index + 1, 4, "晚上")
            table.write(index + 1, 5, str(d[name][date]["night"]))

            index = index +2

    file.save('考勤表.xls')


if __name__=="__main__":
    if len(sys.argv) < 2:
        print("Please input excel file path:")
        sys.exit()

    ParseExcel(sys.argv[1])
