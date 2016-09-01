import openpyxl
wb = openpyxl.load_workbook('OS CHECK FILE 082516 COMBINED.xlsx', data_only=True)
sheetNames = wb.get_sheet_names()
alphabet = list('abcdefghijklmnopqrstuvwxyz'.upper())
writeWB = openpyxl.Workbook()
writeSheet = writeWB.active
numericCounter = 1
alphabetCounter = 0
for name in sheetNames:
    currentSheet = wb.get_sheet_by_name(name)
    for row in currentSheet:
        for cellObj in row:
            writeSheet[alphabet[alphabetCounter] + str(counter)] = cellObj.value
            alphabetCounter += 1
        numericCounter +=1
        alphabetCounter = 0
writeWB.save('export.xlsx')