import win32com.client
import subprocess
import time
import re
import csv

path = "C:/Users/Tadeu/Documents/"
qttFile = "Quantidade_ATU.txt"
pkgFile = "Pacotes_ATU.txt" #DESTRUIR ARQUIVOS AO FINAL DO SCRIPT
newPkgFile = "newPkgFile.txt"
dateCastFile = "dateCastFile.txt"
dateFileHMP = "DateFileHMP.txt"
dateFilePRD = "DateFilePRD.txt"
elemFile = "elemFile.txt"
sheet = "executePackage.csv"
sheet2 = "deniePackage.csv"
rows = 15
rows2 = 9
margin = 5
margin2 = 11

Ex = subprocess.Popen("C:\Program Files (x86)\Attachmate\E!E2K\Sessions\Rede2.edp", shell=True)
time.sleep(10)
SYSTEM = "UNKNOWN"
system = win32com.client.Dispatch('EXTRA.System')
session = system.ActiveSession
screen = session.Screen

def cleanFiles():
    open(path + dateCastFile, 'w').close()
    open(path + elemFile, 'w').close()
    open(path + dateFileHMP, 'w').close()
    open(path + dateFilePRD, 'w').close()
    open(path + qttFile, 'w').close()
    open(path + pkgFile, 'w').close()
    open(path + newPkgFile, 'w').close()

def write(screen,row,col,text):
    screen.row = row
    screen.col = col
    screen.SendKeys(text)

def navigate():
    write(screen, 17, 38, "xxx")
    screen.SendKeys('<Enter>')
    
    time.sleep(1)
    write(screen, 2, 1, "user")
    screen.SendKeys('<Enter>')
    time.sleep(1)
    write(screen, 8, 20, "Password")
    screen.SendKeys('<Enter>')
    time.sleep(3)
    screen.SendKeys('<Pf3>')
    
    time.sleep(1)
    write(screen, 3, 16, "tso qedit")
    screen.SendKeys('<Enter><Enter>')
    write(screen, 2, 15, "p")
    time.sleep(1)
    screen.SendKeys('<Enter>')
    write(screen, 2, 15, "1")
    screen.SendKeys('<Tab>')
    write(screen, 14, 26, "pacotes*")
    screen.SendKeys('<Tab>')
    write(screen, 19, 22, "n")
    screen.SendKeys('<Tab>')
    write(screen, 19, 45, "n")
    screen.SendKeys('<Tab><Tab><Tab>')
    write(screen, 20, 45, "n")
    screen.SendKeys('<Tab><Tab>')
    write(screen, 21, 22, "n")
    screen.SendKeys('<Tab>')
    write(screen, 21, 45, "n")
    screen.SendKeys('<Enter>')
    time.sleep(1)

def storeTotalPackage():
    x1, y1, x2, y2 = 1, 1, 1, 80
    qtt = screen.Area(x1,y1,x2,y2)
    qtt = str(qtt)
    string = "NO PACKAGES SELECTED"
    if string in qtt:
        print("Nao ha pacotes para executar")
        exit(0)
    else:
        print("Ha pacotes para executar")
        packageNum = qtt.split(" of ")[1]
        packageNum = int(packageNum)
    storePackageName(packageNum)

def storePackageName(packageNum):
    x1, y1, x2, y2 = 6, 4, 20, 19
    if packageNum <= rows:
        x2 = (packageNum % rows) + margin
        pkgs = screen.Area(x1,y1,x2,y2)
        print(pkgs, file=open(path + pkgFile, "w"))
   
        with open(path + pkgFile, 'r') as f:
            contents = f.read().replace(' ', '\n')

        with open(path + newPkgFile, 'w') as f:
            f.write(contents.replace('\n\n', '\n'))
    else:
        pkgs = screen.Area(x1,y1,x2,y2)
        print(pkgs, file=open(path + pkgFile, "w"))
        realNum = (packageNum - rows) / rows
        pagination = round(realNum)
       
        if pagination - realNum < 0: # if 0, realNum is multiple of rows.
            pagination += 1
        i = 1
        while i <= pagination:
            screen.SendKeys('<Pf8>')
            time.sleep(0.25)
            if i == pagination:
                if packageNum % rows != 0:
                    x2 = (packageNum % rows) + margin
                pkgs = screen.Area(x1,y1,x2,y2)
                print(pkgs, file=open(path + pkgFile, "a"))
                break
            print(pkgs, file=open(path + pkgFile, "a"))
            i += 1
        with open(path + pkgFile, 'r') as f:
            contents = f.read().replace(' ', '\n')
           
        with open(path + newPkgFile, 'w') as f:
            f.write(contents.replace('\n\n', '\n'))
    file = open(path + newPkgFile)
    packageList = file.read().splitlines()
    file.close()
    time.sleep(0.5)
    screen.SendKeys('<Pf3>')
    time.sleep(0.5)
    storeElement(packageList)

def storeElement(packageList):
    for i in packageList:
        write(screen, 2, 15, "1")
        screen.SendKeys('<Tab>')
        screen.SendKeys(i)
        screen.SendKeys('<Enter>')
        time.sleep(0.5)

        x1, y1, x2, y2 = 18, 1, 18, 80
        rc = screen.Area(x1,y1,x2,y2)
        rc = str(rc)
        string = "CAST"
        if string not in rc:
            print("Erro, linha errada para consulta")
            exit(1)
        else:
            x1, y1, x2, y2 = 18, 25, 18, 37
            castDate = screen.Area(x1,y1,x2,y2)
            castDate = str(castDate)
            newTimestampCast = dictionary(castDate)
            print(newTimestampCast, file=open(path + dateCastFile, "a"))
            write(screen, 22, 15, "s")
            screen.SendKeys('<Enter>')
            time.sleep(0.5)

        x1, y1, x2, y2 = 4, 1, 4, 80
        rc = screen.Area(x1,y1,x2,y2)
        rc = str(rc)
        string = "TRANSFER ELEMENT"
        if string not in rc:
            print("Erro, linha errada para consulta")
            exit(1)
        else:
            pattern = rc[20:].replace("'", "")
            pattern = pattern.replace(" ", "")
            print(pattern, file=open(path + elemFile, "a"))
            screen.SendKeys('<Pf3><Pf3>')
            time.sleep(0.5)
    screen.SendKeys('<Pf3>') #Para voltar a tela "CA Endevor SCM Quick Edit 18.1.00"
    file = open(path + elemFile)
    elementList = file.read().splitlines()
    getElementDate(elementList)

def getElementDate(elementList):
    for i in elementList:
        screen.SendKeys('<Tab><Tab>')
        write(screen, 12, 18, "HMP")
        screen.SendKeys('<Tab><Tab>')
        write(screen, 13, 18, "PRD")
        screen.SendKeys('<Tab>')
        write(screen, 13, 67, "Y")
        screen.SendKeys('<Tab>')
        write(screen, 14, 18, i[:3])
        screen.SendKeys('<Tab><Tab>')
        write(screen, 15, 18, i)
        time.sleep(0.5)
        screen.SendKeys('<Enter>')
        time.sleep(0.5)
   
        x1, y1, x2, y2 = 1, 1, 1, 80
        rc = screen.Area(x1,y1,x2,y2); rc = str(rc)
        errorString = "Element not Found"
        if errorString not in rc:
            count = 0
            print("Agora eu estou na tela Element Selection List")
            for x1 in range(1, 22):
                rc = screen.Area(x1,y1,x2,y2); rc = str(rc)
                x2 += 1 # follows x1
                type = "JCLP"
                stage1 = " M "
                stage2 = " P "
                if (type in rc) and (stage1 in rc or stage2 in rc):
                    write(screen, x1, 2, "s")
                    count += 1 # gets one date if smaller than 2
            screen.SendKeys('<Enter>')
            time.sleep(0.5)
            getElementData(count)
        else:
            print("Deu problema")
            print(0, file=open(path + dateFileHMP, "a"))
            print(0, file=open(path + dateFilePRD, "a"))
            getElementDate()

def getElementData(count):
    stage1 = "Stage ID: M"
    stage2 = "Stage ID: P"
    i = 1
    for i in range(count):
        x1, y1, x2, y2 = 1, 1, 1, 80
        qtt = screen.Area(x1,y1,x2,y2); qtt = str(qtt)
        print(qtt, file=open(path + qttFile, "w"))
        with open(path + qttFile, "r") as rfile:
            readfile = rfile.read()
            readfile = readfile[-7:]
            change = re.sub('[^0-9\n]', '', readfile)
        with open(path + qttFile, 'w') as wfile:
            wfile.write(change.replace('\n', ''))
        qtt = open(path + qttFile, 'r')
        num = qtt.read()
        num = int(num)

        x1, y1, x2, y2 = 1, 1, 10, 80
        rc = screen.Area(x1,y1,x2,y2); rc = str(rc)
        if stage1 in rc:
            print("Stage ID: M")
            if num <= rows2:
                if num % rows2 == 0:
                    x1 = margin2 + rows2
                else:
                    x1 = (num % rows2) + margin2
                x2 = x1
                y1, y2 = 31, 43
                dateHMP = screen.Area(x1,y1,x2,y2)
                dateHMP = str(dateHMP)
                newTimestampHMP = dictionary(dateHMP)
                print(newTimestampHMP, file=open(path + dateFileHMP, "a"))
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            else:
                realNum = (num - rows2) / rows2
                pagination = round(realNum)
                if pagination - realNum < 0: # SE RESULTADO FOR 0, num É MÚLTIPLO DE rows.
                    pagination += 1
                i = 1
                while i <= pagination:
                    screen.SendKeys('<Pf8>')
                    time.sleep(0.25)
                    if i == pagination:
                        if num % rows2 == 0:
                            x1 = margin2 + rows2
                        else:
                            x1 = (num % rows2) + margin2
                        x2 = x1
                        y1, y2 = 31, 43
                        dateHMP = screen.Area(x1,y1,x2,y2)
                        dateHMP = str(dateHMP)
                        newTimestampHMP = dictionary(dateHMP)
                        print(newTimestampHMP, file=open(path + dateFileHMP, "a"))
                        break
                    i += 1
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            if count < 2:
                print(0, file=open(path + dateFilePRD, "a"))
        elif stage2 in rc:
            print("Stage ID: P")
            if num <= rows2:
                if num % rows2 == 0:
                    x1 = margin2 + rows2
                else:
                    x1 = (num % rows2) + margin2
                x2 = x1
                y1, y2 = 31, 43
                datePRD = screen.Area(x1,y1,x2,y2)
                datePRD = str(datePRD)
                newTimestampPRD = dictionary(datePRD)
                print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            else:
                realNum = (num - rows2) / rows2
                pagination = round(realNum)
                if pagination - realNum < 0: # SE RESULTADO FOR 0, num É MÚLTIPLO DE rows.
                    pagination += 1
                i = 1
                while i <= pagination:
                    screen.SendKeys('<Pf8>')
                    time.sleep(0.25)
                    if i == pagination:
                        if num % rows2 == 0:
                            x1 = margin2 + rows2
                        else:
                            x1 = (num % rows2) + margin2
                        x2 = x1
                        y1, y2 = 31, 43
                        datePRD = screen.Area(x1,y1,x2,y2)
                        datePRD = str(datePRD)
                        newTimestampPRD = dictionary(datePRD)
                        print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                        break
                    i += 1
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            if count < 2:
                print(0, file=open(path + dateFileHMP, "a"))
        else:
            print("Não entrou no estágio correto")
            exit(1)
        i += 1
    screen.SendKeys('<Pf3>')
    time.sleep(0.5)
    createSheet()

def dictionary(date):
    dictionary = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12", " ": "", ":": ""}
    for key in dictionary.keys():
        date = date.replace(key, dictionary[key])
    dd = date[:2]
    MM = date[2:4]
    yy = date[4:6]
    hh = date[6:8]
    mm = date[8:10]
    newTimestamp = yy + MM + dd + hh + mm
    return newTimestamp

def createSheet():
    with (open(path + newPkgFile, 'r') as pkg,
         open(path + dateCastFile, 'r') as cas,
         open(path + elemFile, 'r') as ele,
         open(path + dateFileHMP, 'r') as dth,
         open(path + dateFilePRD, 'r') as dtp
         ):
            list1 = pkg.read().splitlines()
            list2 = cas.read().splitlines(); list2 = [int(numeric_string) for numeric_string in list2]
            list3 = ele.read().splitlines()
            list4 = dth.read().splitlines(); list4 = [int(numeric_string) for numeric_string in list4]
            list5 = dtp.read().splitlines(); list5 = [int(numeric_string) for numeric_string in list5]
    
    matrix = []
    for i in range(len(list1)):
        matrix.append([list1[i], list2[i], list3[i], list4[i], list5[i]])
    
    removeZeroMatrix = [x for x in matrix if 0 not in x]
    sortedMatrix = sorted(removeZeroMatrix, key=lambda x: x[2], reverse=True)
    
    finalMatrix = []
    verified = set()
    for list in sortedMatrix:
        element = list[2]
        dateH = list[3]
        dateP = list[4]
        if element not in verified and dateH <= dateP:
            finalMatrix.append(list)
            verified.add(element)
            firstDateP = list[4]
        elif element in verified and dateH <= dateP:
            secondDateP = list[4]
            if secondDateP >= firstDateP:
                finalMatrix.pop()
                finalMatrix.append(list)
    
    # Creates a matrix for deny
    deniePackageMatrix = []
    for denie in matrix:
        if denie not in finalMatrix:
            deniePackageMatrix.append(denie)
    
    # Creates excel spreadsheet
    columns = ['Pacote', 'Data Cast', 'Elemento', 'Data HMP', 'Data PRD']
    with (open(path + sheet, 'w', newline='') as csvFileExec,
         open(path + sheet2, 'w', newline='') as csvFileDenie
         ):
            writer = csv.writer(csvFileExec, delimiter=';')
            writer.writerow(columns)
            for row in finalMatrix:
                writer.writerow(row)
    
            writer = csv.writer(csvFileDenie, delimiter=';')
            writer.writerow(columns)
            for row in deniePackageMatrix:
                writer.writerow(row)

cleanFiles()
navigate()
storeTotalPackage()
Ex.terminate()
