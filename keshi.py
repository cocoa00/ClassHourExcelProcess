from sortedcontainers import SortedSet
import datetime
import xlrd

def getWeekend(datanumber):
    dt = xlrd.xldate_as_datetime(datanumber, 0)
    weekday_num = dt.weekday()
    weekdays = {0: '星期一', 1: '星期二', 2: '星期三', 3: '星期四', 4: '星期五', 5: '星期六',6: '星期日'}
    return weekdays[weekday_num]

def keshi(classTimeObject, personSheet, wb):
    ws = wb.create_sheet('课时', 1)
    # 基础地方
    ws.cell(row=4, column=2).value = "姓名"
    # 查找存在的所有日期和人
    sTime = SortedSet()
    for i in classTimeObject:
        if i.classtype == 2:
            sTime.add(i.time)

    # 一二列名字
    indexRow = 5
    for per in personSheet:
        ws.cell(row=indexRow, column=2).value = per
        ws.cell(row=indexRow, column=1).value = indexRow - 4
        indexRow += 1

    indexCol = 3
    for tm in sTime:
        ws.cell(row=2, column=indexCol).value = getWeekend(tm)
        # dt = xlrd.xldate_as_datetime(tm, 0)
        # tm.number_format = "m月d日"
        ws.cell(row=3, column=indexCol).value = tm
        ws.cell(row=3, column=indexCol).number_format = "m月d日"
        ws.cell(row=4, column=indexCol).value = "上课"

        dt = dict()
        for i in classTimeObject:
            if i.time == tm and i.classtype == 2:
                if dt.get(i.name) == None:
                    if i.ismore == False:
                        dt[i.name] = 1
                    else:
                        dt[i.name] = 1.5
                else:
                    if i.ismore == False:
                        dt[i.name] = dt[i.name] + 1
                    else:
                        dt[i.name] = dt[i.name] + 1.5
        # 遍历dict
        for j in dt:
            flag = None
            # j:名字、dt[j]:次数

            # 获取j在表格中的行数
            for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=2, max_col=2):
                for cell in row:
                    if cell.value == j:
                        flag = cell.row
                        break
            if flag != None:
                ws.cell(row=flag, column=indexCol).value = dt[j]
        indexCol += 1

    # 处理第一行
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=indexCol)
    ws['A1'] = "上课统计明细"
    ws.merge_cells(start_row=2, start_column=1, end_row=4, end_column=1)
    ws['A2'] = "序号"

    # 生成合计
    ws.merge_cells(start_row=2, start_column=indexCol, end_row=4, end_column=indexCol)
    ws.cell(row=2, column=indexCol).value = "合计"

    for row in ws.iter_rows(min_row=5, max_row=ws.max_row, min_col=3, max_col=indexCol - 1):
        cnt = 0;
        nowrow = 0
        for cell in row:
            nowrow = cell.row
            if cell.value is None:
                continue
            cnt += cell.value
        ws.cell(row=nowrow, column=indexCol).value = cnt