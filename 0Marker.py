import xlwings as xw
import os
import shutil

checks = [
    {"key": 77004.66667, "rows": [3], "method":"float"},
    {"key": 64403.75, "rows": [4], "method":"float"},
    {"key": 56151.8, "rows": [5], "method":"float"},
    {"key": 54649.42, "rows": [13], "method":"float"},
    {"key": 68183.27596, "rows": [14], "method":"float"},
    {"key": 73524.83024, "rows": [15], "method":"float"},
    {"key": 18827.71852, "rows": range(3,11), "method":"float"},
    {"key": 607023673.9, "rows": range(3,11), "method":"float"},
    {"key": 16546.48, "rows": range(13,21), "method":"float"},
    {"key": 496663402.8, "rows": range(13,21), "method":"float"},
    {"key": 933.26, "rows": [36,35, 34, 33], "method":"float"},
    {"key": 81700.52083, "rows": [32, 33, 34], "method":"float"},
    {"key": 86042.5516226558, "rows": [23,24,26,27,28,29,30,31], "method":"float"},
    {"key":  60061.18, "rows": [23,24,26,27,28,29,30,31], "method":"float"},
    {"key":  31747.97, "rows": [23,24,26,27,28,29,30,31], "method":"float"},
    {"key": 41908.04167, "rows": [22], "method":"float"},
    {"key": 0.478452326, "rows": [23], "method":"float"},
    {"key": 1.3058985, "rows": [24], "method":"float"},
    {"key": 1.297900444, "rows": [25], "method":"float"},
    {"key": 45211.78, "rows": [26], "method":"float"},
    {"key": 45490.39203, "rows": [26], "method":"float"},
    {"key": 40634.8878, "rows": [27], "method":"float"},
    {"key": 40386.02, "rows": [27], "method":"float"},
    {"key": 52740.03898, "rows": [28], "method":"float"},
    {'key': "trend", "rows": [29, 33, 30], "method":"formula"},
    {'key': "ANOVA", "rows": [29, 33, 30], "method":"keywordsAnd"},
    {'key': "0.5 best", "rows": [19], "method":"keywordsAnd"},
    {'key': "3 best", "rows": [9], "method":"keywordsAnd"},
    {'key': "regression best", "rows": [35], "method":"keywordsAnd"},
    {'key':"quickly accraute fluctua histor sensitive actual smooth lag short long noise","rows":[7,8,17,18],"method":"keywordsAny"}
]

tmpmakr = 'C:\\Users\\dieze\\OneDrive - Brock University\\TA\\1P97\\2021w\\assignment\\Assignment2 Mark Template.xlsx'

def findkey(sht, check):
    key = check['key']
    hit = False
    rows = check['rows']
    if all([mbsht.cells(6,r).value for r in rows]):
        return
    for i in range(sht.api.UsedRange.Row, sht.api.UsedRange.Rows.Count+1):
        for j in range(sht.api.UsedRange.Column, sht.api.UsedRange.Columns.Count+1):
            val = sht.cells(i, j).value
            if check['method'] == 'float':
                try:
                    if round(float(val), 2) == round(key, 2):
                        hit = True
                except Exception as e:
                    # print(e)
                    pass
            elif check['method'] == 'formula':
                formula = sht.cells(i, j).formula
                if key.lower() in formula.lower():
                    hit = True
            elif check['method'] == 'keywordsAnd':
                if isinstance(val, str) and all([k.lower() in val.lower() for k in key.split(' ')]):
                    hit = True
            elif check['method'] == 'keywordsAny':
                if isinstance(val, str) and any([k.lower() in val.lower() for k in key.split(' ')]):
                    hit = True
            if hit:
                for row in check['rows']:
                    mbsht.cells(row, 6).value = mbsht.cells(row, 5).value
                    print('find answer ', key, 'add ',
                          mbsht.cells(row, 5).value)
                return True

# def finddummy(sht):
#     founds = []
#     for i in range(sht.api.UsedRange.Row, sht.api.UsedRange.Rows.Count+1):
#         for j in range(sht.api.UsedRange.Column, sht.api.UsedRange.Columns.Count+1):
#             val = sht.cells(i, j).value
#             for mon in ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','M2','M3','M4','M5','M6','M7','M8','M9','M10','M11','M12']:
#                 if isinstance(val, str) and mon.lower() in val.lower():
#                     founds.append(mon)
#     founds = set(founds)
#     found = len(founds)
#     if found==11:
#         mbsht.cells(32, 6).value = mbsht.cells(32, 5).value
#         print('find dummy 11')
#         print('dummies ', founds)
#         return True
#     if found == 12:
#         mbsht.cells(32, 6).value = mbsht.cells(32, 5).value
#         print('find dummy 12')
#         print('dummies ', founds)
#         return True
#     return False

def finddummy(sht):
    key = 59042
    for i in range(sht.api.UsedRange.Row, sht.api.UsedRange.Rows.Count+1):
        for j in range(sht.api.UsedRange.Column, sht.api.UsedRange.Columns.Count+1):
            val = sht.cells(i, j).value
            try:
                if int(val) == key:
                    if sht.cells(i,j).expand('right').size>11:
                        print('found dummy')
                        mbsht.cells(32, 6).value = mbsht.cells(32, 5).value
                        return True
            except:
                pass
    return False

def checkCollength(sht):
    key = 59042
    for i in range(sht.api.UsedRange.Row, sht.api.UsedRange.Rows.Count+1):
        for j in range(sht.api.UsedRange.Column, sht.api.UsedRange.Columns.Count+1):
            val = sht.cells(i, j).value
            try:
                if int(val) == key:
                    if sht.cells(i,j).expand('down').size <48:
                        print('not in one column')
                        mbsht.cells(3, 7).value = 'put all data from 4 years into one column'
                        mbsht.cells(6, 6).value=1
                    return True
            except:
                pass
    return False


filenames = [filename for filename in os.listdir() if not filename.endswith('Mark.xlsx') and filename.endswith('xlsx')]
scores = {}
for num, filename in enumerate(filenames):
    print(num)
    # if num<24:
    #     continue
    # filename = 'ab19ep.xlsx'
    markfilename = filename.replace('.xlsx', 'Mark.xlsx')
    if not os.path.exists(markfilename):
        shutil.copy(tmpmakr,markfilename)
    mb = xw.Book(markfilename)
    mbsht = mb.sheets['Sheet1']
    print('init', mbsht.range('F38').value)
    if mbsht.range('F38').value:
        print('already marked')
        for r in [7,8,17,18]:
            if mbsht.range(r,6).value == 0:
                mbsht.range(r,6).value = 1
        for r in [29,30]:
            sc = mbsht.range(r,6).value
            mbsht.cells(r,6).value = round(sc/2 + sc/2*(sum(mbsht.range('F22:F28').value)/sum(mbsht.range('E22:E28').value)),0)
        if mbsht.range('F32').value+mbsht.range('F34').value == 8:
            mbsht.cells(33,6).value = 8
        elif mbsht.cells(33,6).value >0:
                mbsht.cells(33,6).value = round(4 + 4/2*(mbsht.range('F32').value+mbsht.range('F34').value)/8,0)
        if mbsht.cells(36,6).value == 7:
            mbsht.cells(35,6).value ==3
        mb.save()
        mb.app.quit()
        continue
    wb = xw.Book(filename)
    mbsht.range('F38').value = '=SUM(F3:F37)'


    dummy = False
    colength = False
    for check in checks:
        print('checking ',check['key'])
        result = False
        breaked = False
        for sht in wb.sheets:
            if not dummy:
                dummy = finddummy(sht)
            if not colength:
                colength = checkCollength(sht)
            result = findkey(sht, check)
            if result:
                breaked = True
                break
        if not breaked:
            dummy = True
            colength = True

    plot1 = mbsht.range(6,6).value
    if not plot1:
        mbsht.range(6,6).value = '=sum(F3:F5)'

    plot2 = mbsht.range(16,6).value
    if not plot2:
        mbsht.range(16,6).value = '=sum(F13:F15)'

    for r in range(3,37):
        if mbsht.range(r,5).value and not mbsht.range(r,6).value:
            mbsht.range(r,6).value = 0
    for r in [29,30]:
        if mbsht.cells(r,6).value == 0:
            continue
        sc = mbsht.range(r,6).value
        mbsht.cells(r,6).value = round(2.5 + 2.5*(sum(mbsht.range('F22:F28').value)/sum(mbsht.range('E22:E28').value)),0)
    if mbsht.cells(33,6).value > 0:
        mbsht.cells(33,6).value = round(4 + 4/2*(mbsht.range('F32').value+mbsht.range('F34').value)/8,0)


    mb.api.RefreshAll()
    total = mbsht.range('F38').value
    scores[filename] =total
    print(filename, 'total ',total)
    mb.save()
    mb.app.quit()
    wb.app.quit()


# filename = 'ac18ur.xlsx'
# markfilename = filename.replace('.xlsx', 'Mark.xlsx')
# mb = xw.Book(markfilename)
# mbsht = mb.sheets['Sheet1']
# wb = xw.Book(filename)
# sht = wb.sheets['D']
# check = checks[0]
# findkey(sht,check)

# sht.range('c2').expand('right').size

# val = sht.range('c55').value

