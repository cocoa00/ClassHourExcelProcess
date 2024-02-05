from sortedcontainers import SortedSet
def keshisheet(classTimeObject, wb):
    ws = wb.create_sheet('课时统计表', 0)
    # 查找存在的所有日期
    sTime = SortedSet()
    sPerson = SortedSet()
    for i in classTimeObject:
        if i.classtype == 2:
            sTime.add(i.time)
            sPerson.add(i.name)

    indexCol = 1
    for tm in sTime:
        ws.cell(row=1, column=indexCol).value = tm
        ws.cell(row=1, column=indexCol).number_format = "m月d日"
        indexCol += 1
        ws.cell(row=1, column=indexCol).value = "值"
        indexCol += 1
        ws.cell(row=1, column=indexCol).value = "出现次数"
        indexCol += 1

        indexRow = 2
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
            ws.cell(row=indexRow, column=indexCol - 2).value = j
            ws.cell(row=indexRow, column=indexCol - 1).value = dt[j]
            indexRow += 1
