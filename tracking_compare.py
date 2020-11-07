import xlrd, xlwt

########################################################################################

def checkKeyValue(war_old,war_new):
    v_check = 0
    if war_old == war_new:
        return 1
    else:
        return 0

def writeRecord2Sheet(recordNumber, st, typeOfRec):
    columnNumber = 0
    current_st = st
    for columnNumber in range (len(lista_new[recordNumber])):
        if recordNumber == 1 and columnNumber>7:
            current_st = style4
        elif recordNumber > 1 and columnNumber == 0:
            current_st = style5
        elif columnNumber in (13,14,20,21,27,28,32,34,35,41,42):
            current_st = style3
        else:
            if typeOfRec == 0:
                current_st = st
            else:
                if (lista_new[recordNumber][columnNumber] == lista_old[d][columnNumber]):
                    current_st = st
                else:
                    current_st = style6
        ws.write(recordNumber, columnNumber, lista_new[recordNumber][columnNumber], current_st)

########################################################################################

# files opening and reading "tracking" sheet
try:
    wb_old = xlrd.open_workbook('C:/Users/przemyslaw.lelewski/Documents/Python/R65Tracking_old.xlsx')
    wb_new = xlrd.open_workbook('C:/Users/przemyslaw.lelewski/Documents/Python/R65Tracking_new.xlsx')
except:
    print("I can't find a file: R65Tracking_new.xlsx")

sheet_old = wb_old.sheet_by_name('Tracking')
sheet_new = wb_new.sheet_by_name('Tracking')
lista_old = []
lista_new = []

style0 = xlwt.easyxf('font: color-index black; borders: top thin, right thin; pattern: pattern solid,fore-colour yellow')
style1 = xlwt.easyxf('font: color-index black; borders: top thin, right thin; pattern: pattern solid,fore-colour light_green')
style2 = xlwt.easyxf('font: color-index black; borders: top thin, right thin')
style3 = xlwt.easyxf('font: color-index black, bold on; borders: top thin, right thin; align: wrap yes,vert centre,horiz center; pattern: pattern solid,fore-colour gray25')
style4 = xlwt.easyxf('font: color-index black; borders: top thin, right thin; align: rotation 90')
style4.num_format_str = 'dd/mm/yyyy'
style5 = xlwt.easyxf('font: color-index white; borders: top thin, right thin')
style6 = xlwt.easyxf('font: color-index black; borders: top thin, right thin; pattern: pattern solid,fore-colour orange')

wb = xlwt.Workbook()
ws = wb.add_sheet('Tracking')

# records loading to lists
for a in range(sheet_old.nrows):
    lista_old.append(sheet_old.row_values(a,0))

for b in range(sheet_new.nrows):
    lista_new.append(sheet_new.row_values(b,0))

# compare both list to find differences
for c in range(len(lista_new)):
    war_odnaleziona = 0
    for d in range(len(lista_old)):
        ff = 0
        ff = checkKeyValue(lista_new[c][0],lista_old[d][0])

        if ff == 1:
            war_odnaleziona = 1
            print ('Compare new to old record for value: ' + lista_new[c][0])
            if (lista_new[c] == lista_old[d]):
                writeRecord2Sheet(c,style2,war_odnaleziona)
            else:
                writeRecord2Sheet(c,style0,war_odnaleziona)

    if war_odnaleziona == 0:
        print (lista_new[c][0] + ' is a new record on the list')
        writeRecord2Sheet(c,style1,war_odnaleziona)
try:
    wb.save('C:/Users/przemyslaw.lelewski/Documents/Python/R65Tracking_report.xls')
except:
    print('I can\'t save the file - please check if you havn\'t already opened')