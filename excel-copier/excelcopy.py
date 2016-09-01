import openpyxl
wb = openpyxl.load_workbook('OS CHECK FILE 082516 COMBINED.xlsx', data_only=True)
sheetNames = wb.get_sheet_names()
alphabet = list('abcdefghijklmnopqrstuvwxyz'.upper())
writeWB = openpyxl.Workbook()
writeSheet = writeWB.active
counter = 1
alphabetCounter = 0
for name in sheetNames:
    currentSheet = wb.get_sheet_by_name(name)
    for row in currentSheet:
        for cellObj in row:
            #print(cellObj.coordinate, cellObj.value)
            #print(name)
            writeSheet[alphabet[alphabetCounter] + str(counter)] = cellObj.value
            #print( writeSheet[alphabet[alphabetCounter] + str(counter)].value)
            alphabetCounter += 1
        #print('End of Row')
        counter +=1
        alphabetCounter = 0
writeWB.save('export.xlsx')
