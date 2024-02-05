from sortedcontainers import SortedSet
def zaodusheet(classTimeObject, wb):
    ws = wb.create_sheet('早读统计表', 2)
    # 查找存在的所有日期
    sTime = SortedSet()
    sClassName = SortedSet()
    for i in classTimeObject:
        if i.classtype == 1:
            sTime.add(i.time)
            sClassName.add(i.classname)

    # 第一列班级名字
    indexRow = 2
    ws.cell(row=1, column=1).value = "班级"
    for cla in sClassName:
        ws.cell(row=indexRow, column=1).value = cla
        indexRow += 1

    # 第一行日期
    indexCol = 2
    for tm in sTime:
        ws.cell(row=1, column=indexCol).value = tm
        ws.cell(row=1, column=indexCol).number_format = "m月d日"
        indexCol += 2

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=2, max_col=ws.max_column + 1):
        for cell in row:
            row = cell.row
            col = cell.column
            tclassname = ws.cell(row=row, column=1).value
            ttime = ws.cell(row=1, column=col).value
            # 当前纵目录日期是空白，代表填节数
            if ttime == None and ws.cell(row=row, column=col - 1).value != None:
                cell.value = 1
            else:
                for i in classTimeObject:
                    if i.classname == tclassname and i.time == ttime:
                        cell.value = i.name
                        break

