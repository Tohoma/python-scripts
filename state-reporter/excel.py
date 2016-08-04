import openpyxl
import re


wb =  openpyxl.load_workbook('2016_1099.xlsx')
wbSave = openpyxl.Workbook()
wbSave.get_sheet_names()
outputSheet = wbSave.active
sheet = wb.get_sheet_by_name('Sheet1')
file = open("JUL2017.txt","r")
output = open("newpeople.txt","w")
pagebegin = re.compile("LMAP7LST")
SSN = re.compile("(\d{3}-\d{2}-\d{4})")
address = re.compile("\d+\s([a-zA-Z]\s?\.?\s?)+\d*")
attnaddress=re.compile("\s{2,}\d+\s([a-zA-Z]\s?)+\d*")
apt = re.compile("((?<=Unit\s)\d+)|((?<=Suite\s)\d+)|((?<=#)(\d+)?(\w)?)|((?<=ste\s)\d+)|((?<=apt\s)\d+)|((?<=Box\s)\d+)",flags=re.IGNORECASE)
city = re.compile("([a-zA-Z]+)(\s[a-zA-Z]{3,})?")
zipcode = re.compile("(\d{5}(\-\d+)?)")
total = re.compile("YTD\sTotal\:")
nameCombo = re.compile("(\w+)(\,\s)?(\w+)?(\s\w)?(\s\w+)?")
#Purposely not including anchor to remove the $ sign.
cash = re.compile("([0-9]{1,3}((\,[0-9]{3})+)?(\.[0-9]{1,2})?)")
space = re.compile("^\s*$")
ADR = re.compile("(\d*)")
streetNameRegex = re.compile("([a-zA-Z]\s?\.?\s?)+\d*")
ssnList = []
item = 0

def pageCheck(file,line):
    while space.match(line):
        line = file.next()
    #print("The line passed into page check is " + line)
    pagebeginstr = pagebegin.search(line)
    if pagebeginstr:
        print("New Page!")
        for i in range(7):
            line = file.next()
        while space.match(line) or attnaddress.match(line):
            print(line)
            line = file.next()
        print("The line being passed out is " + line)
        return line
    else: return line


for ssn in sheet.columns[3]:
    ssnList.append(ssn.value)
    print (ssn.value)



for line in file:
    #Skip white space
    while space.match(line):
        line = file.next()
    line = pageCheck(file,line)
    #Name validation
    name = line
    print("The name is " + name)
    line = pageCheck(file,line)
    ssnLine = file.next()
    ssnLine = pageCheck(file,ssnLine)
    foundSSN = SSN.search(ssnLine)
    foundAddress = address.search(ssnLine)
    foundApt = apt.search(ssnLine)
    if foundSSN:
        personSSN = foundSSN.group(0)
        print(personSSN)
    else: 
        personSSN = False
        print("Invalid SSN")
    if foundAddress:
        personFoundAddress = foundAddress.group(0)
        print("The address is " + foundAddress.group(0))
    else:
        personFoundAddress = "ADDRESS REGEX DID NOT MATCH"
    if foundApt:
        personApt = foundApt.group(0)
        print ("The apartment number is " + personApt)
    else:
        personApt = ""
    cityLine=file.next()
    line = cityLine
    line = pageCheck(file,line)
    foundCityGroup = city.findall(line)
    if foundCityGroup:
        personCity = ('').join(foundCityGroup[0])
        if foundCityGroup[1]:
            personState = ('').join(foundCityGroup[1])
        else:
            personState = "Regex error"
        print("CITY: "+('').join(foundCityGroup[0]))
        print("STATE: "+('').join(foundCityGroup[1]))
        print("The line is " + line)
    print("The line being passed to zipcode is " + line) 
    foundZipcode = zipcode.findall(line)
    print("The regex object should be below")
    #print(foundZipcode)
    if foundZipcode[1]:
        personZipcode = foundZipcode[1][0]
        print("ZIPCODE: "+personZipcode)
    else:
        print("Could not find regex obj")
        personZipcode = "Regex error"
    currentline = file.next()
    #Finding the YTD Total section
    while not(total.search(currentline)):
        currentline = file.next()
    cashCheck = cash.findall(currentline)
    #print(cashCheck)
    if cashCheck:
        cashTotal = float(cash.search(currentline).group(0).replace(',',''))
        if cashCheck[0][3]:
            change = cashCheck[0][3].replace('.','')
        else:
            change = "00"
        print("The change is " + change)
        print(cashTotal)
    else:
        cashTotal = "Regex error"
   
    
    
    if personSSN and personSSN not in ssnList and cashTotal > 600:
        firstName = nameCombo.search(name).group(3)
        lastName = nameCombo.search(name).group(1)
        adr = ADR.search(personFoundAddress).group(0)
        streetName = streetNameRegex.search(personFoundAddress).group(0)
        if firstName is None:
            firstName = "DID NOT MATCH REGEX"
        if lastName is None:
            lastName = "DID NOT MATCH REGEX"
        if adr is None:
            adr = ""
        
        item +=1
        print("Add to worksheet")
        output.write("Number: "+str(item) + "\n")
        output.write(name+"\n")
        output.write("FIRSTNAME: " + firstName + " LASTNAME: " + lastName + "\n")
        output.write(personSSN + "\n")
        output.write(personFoundAddress + "\n")
        output.write("ADR: " + adr + "\n")
        output.write("Street Name " + streetName + "\n")
        output.write(personCity + "\n")
        output.write(personState + "\n")
        output.write(personZipcode + "\n")
        output.write(str(cashTotal) + "\n") 
        output.write("-"*20)
        output.write("\n")
        contentList = [firstName.upper(),' ',lastName.upper(),personSSN,adr,streetName,personApt,personCity,personState,personZipcode,str(int(cashTotal)),str(change) ,' ',]
        counter = 0
        for cellRow in outputSheet["A"+str(item):"M"+str(item)]:
            for cell in cellRow:
                cell.value = contentList[counter]
                #print(counter)
                counter += 1
    print("-------------------------------")
    


wbSave.save("export.xlsx")
output.close()

print("Done")



