import win32com.client
import subprocess
import time
import re
import csv

path = "C:/Users/Tadeu/Documents/"
qtdFile = "Quantidade_ATU"
pkgFile = "Pacotes_ATU" #DESTRUIR ARQUIVOS AO FINAL DO SCRIPT
newPkgFile = "newPkgFile"
dataCastFile = "dataCastFile"
dateFileHMP = "DateFileHMP"
dateFilePRD = "DateFilePRD"
elemFile = "elemFile"
planilha = "executePackage.csv"
planilha2 = "deniePackage.csv"
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

def limpaArquivos():
    open(path + dataCastFile, 'w').close()
    open(path + elemFile, 'w').close()
    open(path + dateFileHMP, 'w').close()
    open(path + dateFilePRD, 'w').close()
    open(path + qtdFile, 'w').close()
    open(path + pkgFile, 'w').close()
    open(path + newPkgFile, 'w').close()

def write(screen,row,col,text):
    screen.row = row
    screen.col = col
    screen.SendKeys(text)

def navegaPacote():
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

def gravaQuantidadePacote():
    x1, y1, x2, y2 = 1, 1, 1, 80
    qtd = screen.Area(x1,y1,x2,y2)
    qtd = str(qtd)
    string = "NO PACKAGES SELECTED"
    if string in qtd:
        print("Nao ha pacotes para executar")
        exit(0)
    else:
        print("Ha pacotes para executar")
        numPacote = qtd.split(" of ")[1]
        numPacote = int(numPacote)
    gravaNomePacote(numPacote)

def gravaNomePacote(num):
    x1, y1, x2, y2 = 6, 4, 20, 19
    if num <= rows:
        x2 = (num % rows) + margin
        pkgs = screen.Area(x1,y1,x2,y2)
        print(pkgs, file=open(path + pkgFile, "w"))
   
        with open(path + pkgFile, 'r') as f:
            contents = f.read().replace(' ', '\n')

        with open(path + newPkgFile, 'w') as f:
            f.write(contents.replace('\n\n', '\n'))
    else:
        pkgs = screen.Area(x1,y1,x2,y2)
        print(pkgs, file=open(path + pkgFile, "w"))
        numReal = (num - rows) / rows
        paginacao = round(numReal)
       
        if paginacao - numReal < 0: # SE RESULTADO FOR 0, num É MÚLTIPLO DE rows.
            paginacao += 1
        i = 1
        while i <= paginacao:
            screen.SendKeys('<Pf8>')
            time.sleep(0.25)
            if i == paginacao:
                if num % rows != 0:
                    x2 = (num % rows) + margin
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
    arrayPacotes = file.read().splitlines()
    file.close()
    time.sleep(0.5)
    screen.SendKeys('<Pf3>')
    time.sleep(0.5)
    gravaElemento(arrayPacotes)

def gravaElemento(arrayPacotes):
    for i in arrayPacotes:
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
            dataCast = screen.Area(x1,y1,x2,y2)
            dataCast = str(dataCast)
            newTimestampCast = Dicionario(dataCast)
            print(newTimestampCast, file=open(path + dataCastFile, "a"))
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
    arrayElementos = file.read().splitlines()
    entraComElemento(arrayElementos)

def entraComElemento(arrayElementos):
    for i in arrayElementos:
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
                x2 += 1 # para seguir o x1
                type = "JCLP"
                stage1 = " M "
                stage2 = " P "
                if (type in rc) and (stage1 in rc or stage2 in rc):
                    write(screen, x1, 2, "s")
                    count += 1 # se tiver menos de 2, pega somente uma data
            screen.SendKeys('<Enter>')
            time.sleep(0.5)
            getElementData(count)
        else:
            print("Deu problema")
            print(0, file=open(path + dateFileHMP, "a"))
            print(0, file=open(path + dateFilePRD, "a"))
            entraComElemento()

def getElementData(count):
    stage1 = "Stage ID: M"
    stage2 = "Stage ID: P"
    i = 1
    for i in range(count):
        x1, y1, x2, y2 = 1, 1, 1, 80
        qtd = screen.Area(x1,y1,x2,y2); qtd = str(qtd)
        print(qtd, file=open(path + qtdFile, "w"))
        with open(path + qtdFile, "r") as rfile:
            readfile = rfile.read()
            readfile = readfile[-7:]
            troca = re.sub('[^0-9\n]', '', readfile)
        with open(path + qtdFile, 'w') as wfile:
            wfile.write(troca.replace('\n', ''))
        qtd = open(path + qtdFile, 'r')
        num = qtd.read()
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
                newTimestampHMP = Dicionario(dateHMP)
                print(newTimestampHMP, file=open(path + dateFileHMP, "a"))
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            else:
                numReal = (num - rows2) / rows2
                paginacao = round(numReal)
                if paginacao - numReal < 0: # SE RESULTADO FOR 0, num É MÚLTIPLO DE rows.
                    paginacao += 1
                i = 1
                while i <= paginacao:
                    screen.SendKeys('<Pf8>')
                    time.sleep(0.25)
                    if i == paginacao:
                        if num % rows2 == 0:
                            x1 = margin2 + rows2
                        else:
                            x1 = (num % rows2) + margin2
                        x2 = x1
                        y1, y2 = 31, 43
                        dateHMP = screen.Area(x1,y1,x2,y2)
                        dateHMP = str(dateHMP)
                        newTimestampHMP = Dicionario(dateHMP)
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
                newTimestampPRD = Dicionario(datePRD)
                print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                screen.SendKeys('<Pf3>')
                time.sleep(0.5)
            else:
                numReal = (num - rows2) / rows2
                paginacao = round(numReal)
                if paginacao - numReal < 0: # SE RESULTADO FOR 0, num É MÚLTIPLO DE rows.
                    paginacao += 1
                i = 1
                while i <= paginacao:
                    screen.SendKeys('<Pf8>')
                    time.sleep(0.25)
                    if i == paginacao:
                        if num % rows2 == 0:
                            x1 = margin2 + rows2
                        else:
                            x1 = (num % rows2) + margin2
                        x2 = x1
                        y1, y2 = 31, 43
                        datePRD = screen.Area(x1,y1,x2,y2)
                        datePRD = str(datePRD)
                        newTimestampPRD = Dicionario(datePRD)
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
    criaPlanilha()

def Dicionario(data):
    dictionary = {"JAN": "01", "FEB": "02", "MAR": "03", "APR": "04", "MAY": "05", "JUN": "06", "JUL": "07", "AUG": "08", "SEP": "09", "OCT": "10", "NOV": "11", "DEC": "12", " ": "", ":": ""}
    for key in dictionary.keys():
        data = data.replace(key, dictionary[key])
    dia = data[:2]
    mes = data[2:4]
    ano = data[4:6]
    hora = data[6:8]
    minu = data[8:10]
    newTimestamp = ano + mes + dia + hora + minu
    return newTimestamp

def criaPlanilha():
    with (open(path + newPkgFile, 'r') as pkg,
         open(path + dataCastFile, 'r') as cas,
         open(path + elemFile, 'r') as ele,
         open(path + dateFileHMP, 'r') as dth,
         open(path + dateFilePRD, 'r') as dtp
         ):
            array1 = pkg.read().splitlines()
            array2 = cas.read().splitlines(); array2 = [int(numeric_string) for numeric_string in array2]
            array3 = ele.read().splitlines()
            array4 = dth.read().splitlines(); array4 = [int(numeric_string) for numeric_string in array4]
            array5 = dtp.read().splitlines(); array5 = [int(numeric_string) for numeric_string in array5]
    
    matrix = []
    for i in range(len(array1)):
        matrix.append([array1[i], array2[i], array3[i], array4[i], array5[i]])
    
    #List comprehension
    removeZeroMatrix = [x for x in matrix if 0 not in x]
    sortedMatrix = sorted(removeZeroMatrix, key=lambda x: x[2], reverse=True) #reverse = True (Sorts in Descending order)
    
    finalMatrix = []
    valorVerificado = set() # SET é usado para armazenar vários itens em uma única variável
    for list in sortedMatrix:
        elemento = list[2]
        dataH = list[3]
        dataP = list[4]
        if elemento not in valorVerificado and dataH <= dataP:
            finalMatrix.append(list)
            valorVerificado.add(elemento)
            primDataP = list[4]
        elif elemento in valorVerificado and dataH <= dataP:
            seguDataP = list[4]
            if seguDataP >= primDataP:
                finalMatrix.pop()
                finalMatrix.append(list)
    
    #Cria matrix para tomar denied
    deniePackageMatrix = []
    for denie in matrix:
        if denie not in finalMatrix:
            deniePackageMatrix.append(denie)
    
    #Cria planilhas no Excel
    colunas = ['Pacote', 'Data Cast', 'Elemento', 'Data HMP', 'Data PRD']
    with (open(path + planilha, 'w', newline='') as csvFileExec,
         open(path + planilha2, 'w', newline='') as csvFileDenie
         ):
            writer = csv.writer(csvFileExec, delimiter=';')
            writer.writerow(colunas)
            for row in finalMatrix:
                writer.writerow(row)
    
            writer = csv.writer(csvFileDenie, delimiter=';')
            writer.writerow(colunas)
            for row in deniePackageMatrix:
                writer.writerow(row)

limpaArquivos()
navegaPacote()
gravaQuantidadePacote()
Ex.terminate()
