import time
import csv
import re
import win32com.client
import subprocess
from MyWindow3 import user, password, partition
from variaveis import *

path = "C:\\"
qtdFile = "Quantidade_ATU.txt"
pkgFile = "Pacotes_ATU.txt"
newPkgFile = "newPkgFile.txt"
dataCastFile = "DataCastFile.txt"
dateFileHMP = "DateFileHMP.txt"
dateFilePRD = "DateFilePRD.txt"
elemFile = "ElemFile.txt"
arquivo = "arquivo.txt"
execucao = "execucao.txt"
rows = 16 # 15
rows2 = 10 # 9
margin = 6 # 5
margin2 = 12 # 11

def start():
    return subprocess.Popen("C:\\Program Files (x86)\\Attachmate\\E!E2K\\Sessions\\RedeCaixa2.edp", shell=True)

def stop(proc):
    proc.terminate()

def terminalConector():
    time.sleep(2) # 10
    system = win32com.client.Dispatch('EXTRA.System')
    session = system.ActiveSession
    global screen 
    screen = session.Screen
    esperaPorString(tela1, lin1, col1)

def limpaArquivos():
    open(path + dataCastFile, 'w').close()
    open(path + elemFile, 'w').close()
    open(path + dateFileHMP, 'w').close()
    open(path + dateFilePRD, 'w').close()
    open(path + arquivo, 'w').close()
    open(path + execucao, 'w').close()

def write(screen,row,col,text):
    screen.row = row
    screen.col = col
    screen.SendKeys(text)

def navegaPacote():
    write(screen, 17, 38, partition.get())
    screen.SendKeys('<Enter>')
    esperaPorString(tela2, lin2, col2)
    write(screen, 2, 1, user.get())
    screen.SendKeys('<Enter>')
    esperaPorString(tela3, lin3, col3)
    write(screen, 8, 20, password.get())
    screen.SendKeys('<Enter>')
    esperaPorString(tela4, lin4, col4)
    esperaPorString(asteriscos, linAsteriscos, colAsteriscos)
    screen.SendKeys('<Pf3>')
    esperaPorString(tela5, lin5, col5)
    write(screen, 3, 16, "tso qedit")
    screen.SendKeys('<Enter><Enter>')
    esperaPorString(tela6, lin6, col6)
    write(screen, 2, 15, "p")   # acrescentar "N" na posição 14/67 da tela CA Endevor SCM Quick Edit 18.1.00
    screen.SendKeys('<Enter>')
    esperaPorString(tela7, lin7, col7)
    write(screen, 2, 15, "1")
    screen.SendKeys('<Tab>')
    write(screen, 14, 26, "atuhmp*")
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
    screen.SendKeys('<Tab><Tab>')
    write(screen, 22, 22, "n")
    screen.SendKeys('<Tab>')
    write(screen, 22, 45, "n")
    screen.SendKeys('<Enter>')
    time.sleep(1) # Para dar tempo de mostrar se há ou não pacotes REVER

def gravaNomePacote():
    x1, y1, x2, y2 = 7, 4, 7, 18
    while True:
        if x1 > 22:
            screen.SendKeys('<Pf8>')
            time.sleep(3)
            x1 = 7
            x2 = x1
        bottomListString = "---"
        linha = screen.Area(x1,y1,x2,y2); linha = str(linha)
        if bottomListString not in linha:
            print(linha, file=open(path + pkgFile, "a"))
            x1 += 1
            x2 = x1
        else:
            print("Chegou ao final da lista")
            screen.SendKeys('<Pf3>')
            time.sleep(0.5)
            with open(path + pkgFile, "r") as file:
                arrayPacotes = file.read().splitlines()
            print(arrayPacotes)
            break
    gravaElemento(arrayPacotes)

def gravaElemento(arrayPacotes):
    for i in arrayPacotes:
        write(screen, 2, 15, "1")
        screen.SendKeys('<Tab>')
        screen.SendKeys(i)
        screen.SendKeys('<Enter>')
        esperaPorString(tela9, lin9, col9)

        x1, y1, x2, y2 = 19, 1, 19, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        string = "CAST"
        if string not in rc: 
            print("Erro, linha errada para consulta na tela PACKAGE DISPLAY")
            exit(1)
        else:
            x1, y1, x2, y2 = 19, 25, 19, 37
            dataCast = str(screen.Area(x1,y1,x2,y2))
            newTimestampCast = Dicionario(dataCast)
            print(newTimestampCast, file=open(path + dataCastFile, "a"))
            write(screen, 2, 15, "s")
            screen.SendKeys('<Enter>')
            esperaPorString(tela10, lin10, col10)

        x1, y1, x2, y2 = 5, 1, 5, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        string = "TRANSFER ELEMENT"
        if string not in rc:
            print("Erro, linha errada para consulta na tela DISPLAY - PACKAGE ID: ATUHMPXXXXXXXX")
            exit(1)
        else:
            pattern = rc[20:].replace("'", "")
            pattern = pattern.replace(" ", "")
            print(pattern, file=open(path + elemFile, "a"))
            screen.SendKeys('<Pf3><Pf3>')
            esperaPorString(tela7, lin7, col7)
    screen.SendKeys('<Pf3>')
    esperaPorString(tela6, lin6, col6)
    file = open(path + elemFile)
    arrayElementos = file.read().splitlines()
    entraComElemento(arrayElementos)

def entraComElemento(arrayElementos):
    for i in arrayElementos:
        screen.SendKeys('<Tab><Tab>')
        write(screen, 12, 18, "CEFHMP")
        screen.SendKeys('<Tab><Tab>')
        write(screen, 13, 18, "PRD")
        screen.SendKeys('<Tab>')
        write(screen, 13, 67, "Y")
        screen.SendKeys('<Tab>')
        write(screen, 14, 18, i[:3])
        screen.SendKeys('<Tab><Tab>')
        screen.SendKeys('<Delete>')
        write(screen, 15, 18, i)
        screen.SendKeys('<Enter>')
        time.sleep(1) # esperaPorString não funciona aqui por causa do loop: esperaPorString(tela11, lin11, col11)
    
        x1, y1, x2, y2 = 1, 1, 1, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        arrayString = ["Element not Found", "Subsystem not Defined"]
        if any(x in rc for x in arrayString):
            print(f"Deu problema no elemento {i}")
            print(0, file=open(path + dateFileHMP, "a"))
            print(0, file=open(path + dateFilePRD, "a"))
            continue
        else:
            count = 0
            esperaPorString(tela11, lin11, col11)
            for x1 in range(1, 22):
                rc = str(screen.Area(x1,y1,x2,y2))
                x2 += 1 # para seguir o x1
                type = "JCLP"
                stage1 = " M "
                stage2 = " P "
                if (type in rc) and (stage1 in rc or stage2 in rc):
                    write(screen, x1, 2, "s")
                    count += 1 # se tiver menos de 2, pega somente uma data
            screen.SendKeys('<Enter>')
            getElementData(count)

def getElementData(count):
    esperaPorString(tela12, lin12, col12)
    stage1 = "Stage ID: M"
    stage2 = "Stage ID: P"
    i = 1
    for i in range(count):
        x1, y1, x2, y2 = 1, 1, 1, 80
        qtd = str(screen.Area(x1,y1,x2,y2))
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
        rc = str(screen.Area(x1,y1,x2,y2))
        if stage1 in rc:
            if num <= rows2: # CÓDIGO SE REPETE AQUI
                if num % rows2 == 0:
                    x1 = margin2 + rows2
                else:
                    x1 = (num % rows2) + margin2
                x2 = x1
                y1, y2 = 31, 43 # TERMINA NESSE PONTO
                dateHMP = str(screen.Area(x1,y1,x2,y2))
                newTimestampHMP = Dicionario(dateHMP)
                print(newTimestampHMP, file=open(path + dateFileHMP, "a"))
                screen.SendKeys('<Pf3>')
                esperaPorString(stagePRD, linStagePRD, colStagePRD)
            else:
                numReal = (num - rows2) / rows2 # CÓDIGO SE REPETE AQUI
                paginacao = round(numReal)
                if paginacao - numReal < 0:
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
                        y1, y2 = 31, 43 # TERMINA NESSE PONTO
                        dateHMP = str(screen.Area(x1,y1,x2,y2))
                        newTimestampHMP = Dicionario(dateHMP)
                        print(newTimestampHMP, file=open(path + dateFileHMP, "a"))
                        break
                    i += 1
                screen.SendKeys('<Pf3>')
                esperaPorString(stagePRD, linStagePRD, colStagePRD)
            if count < 2:
                print(0, file=open(path + dateFilePRD, "a"))
        elif stage2 in rc:
            if num <= rows2: # CÓDIGO SE REPETE AQUI
                if num % rows2 == 0:
                    x1 = margin2 + rows2
                else:
                    x1 = (num % rows2) + margin2
                x2 = x1
                y1, y2 = 31, 43 # TERMINA NESSE PONTO
                datePRD = str(screen.Area(x1,y1,x2,y2))
                newTimestampPRD = Dicionario(datePRD)
                if num == 1:
                    print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                else:
                    for i in range(num,1,-1):
                        x1 -= 1
                        x2 = x1
                        newDatePRD = str(screen.Area(x1,y1,x2,y2))
                        secondTimestampPRD = Dicionario(newDatePRD)
                        if secondTimestampPRD[:6] != newTimestampPRD[:6]:
                            print(secondTimestampPRD, file=open(path + dateFilePRD, "a"))
                            break
                        else:
                            continue
                    else:
                        print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                screen.SendKeys('<Pf3>')
                esperaPorString(tela11, lin11, col11)
            else:
                numReal = (num - rows2) / rows2 # CÓDIGO SE REPETE AQUI
                paginacao = round(numReal)
                if paginacao - numReal < 0:
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
                        y1, y2 = 31, 43 # TERMINA NESSE PONTO
                        datePRD = str(screen.Area(x1,y1,x2,y2))
                        newTimestampPRD = Dicionario(datePRD)
                        for i in range(num,1,-1):
                            x1 -= 1
                            x2 = x1
                            newDatePRD = str(screen.Area(x1,y1,x2,y2))
                            secondTimestampPRD = Dicionario(newDatePRD)
                            if secondTimestampPRD[:6] != newTimestampPRD[:6] and x1 > margin2:
                                print(secondTimestampPRD, file=open(path + dateFilePRD, "a")) # anterior é diferente, grava
                                break
                            elif (secondTimestampPRD[:6] == newTimestampPRD[:6] and x1 == margin2) or x1 == (margin2 - 1): # rever
                                time.sleep(0.5)
                                screen.SendKeys('<Pf7>')
                                time.sleep(0.5)
                                x1 = 23 # colocar esse valor como variável, pode mudar
                                continue
                            else:
                                continue
                        else:
                            print(newTimestampPRD, file=open(path + dateFilePRD, "a"))
                    i += 1
                screen.SendKeys('<Pf3>')
                esperaPorString(tela11, lin11, col11)
            if count < 2:
                print(0, file=open(path + dateFileHMP, "a"))
        else:
            print("Não entrou no estágio correto")
            exit(1)
        i += 1
    screen.SendKeys('<Pf3>')
    esperaPorString(tela6, lin6, col6)

def aproveAndExec(execPackageMatrix):
    write(screen, 2, 15, "p")
    screen.SendKeys('<Enter>')
    esperaPorString(tela7, lin7, col7)
    for list in execPackageMatrix:
        write(screen, 2, 15, "4")
        screen.SendKeys('<Tab>')
        package = list[0]
        write(screen, 14, 26, package)
        screen.SendKeys('<Enter>')
        esperaPorString(tela13, lin13, col13)
        write(screen, 2, 16, "a")
        screen.SendKeys('<Enter>') # após esse Enter, esperar pela tela7, colocar também a string APPROVAL PERFORMED?
        esperaPorString(tela7, lin7, col7)
        write(screen, 2, 15, "5")
        screen.SendKeys('<Enter>')
        esperaPorString(tela14, lin14, col14)
        write(screen, 2, 16, "s")
        screen.SendKeys('<Enter>')
        esperaPorString(tela15, lin15, col15)
        write(screen, 14, 7, "//END@" + package[-4:] + " JOB (END,END,9999,9999),'REROPRJ',")
        write(screen, 15, 7, "// CLASS=K,MSGCLASS=G,TIME=1440")
        write(screen, 16, 7, "/*JOBPARM L=999999,C=999999")
        write(screen, 2, 14, "s")
        screen.SendKeys('<Enter><Enter>') # volta pra tela "Package Foreground Options Menu"
        esperaPorString(tela7, lin7, col7)
        time.sleep(3)   # colocar espera pelas strings 1, n, y, y?
        write(screen, 20, 22, "n")
        esperaPorString(insert1, linInsert1, colInsert1)
        write(screen, 20, 45, "y")
        esperaPorString(insert2, linInsert2, colInsert2)
        write(screen, 21, 45, "y")
        esperaPorString(insert3, linInsert3, colInsert3)
        write(screen, 2, 15, "1")
        esperaPorString(insert4, linInsert4, colInsert4)
        screen.SendKeys('<Enter>')
        esperaPorString(tela9, lin9, col9)
        x1, y1, x2, y2 = 9, 1, 9, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        stringAproved = "EXECUTED"
        stringFailed = "EXEC-FAILED"
        if stringAproved in rc:
            print(f"Executou o pacote {package}")
            print(package, file=open(path + arquivo, "a"))
            print(stringAproved, file=open(path + execucao, "a"))
            screen.SendKeys('<Pf3>')
            esperaPorString(tela7, lin7, col7)
        elif stringFailed in rc:
            print(f"Execucao falhou para o pacote {package}")
            print(stringFailed, file=open(path + execucao, "a"))
            screen.SendKeys('<Pf3>')
            esperaPorString(tela7, lin7, col7)
        else:
            print("Erro na execução do script na tela PACKAGE DISPLAY, strings não apareceram")
            exit(1)
    distPackage()

def distPackage():
    global array6, array7, arrayDist # arrayDist não precisa ser global
    array6, array7 = [], []
    with open(path + execucao, 'r') as Exec:
        array6 = Exec.read().splitlines()
    with open(path + arquivo, 'r') as Dist:
        arrayDist = Dist.read().splitlines()

    condicao = False
    for package in arrayDist:
        write(screen, 2, 15, "6")
        screen.SendKeys('<Tab>')
        write(screen, 14, 26, package)
        screen.SendKeys('<Enter>')
        esperaPorString(tela16, lin16, col16)
        write(screen, 2, 15, "1")
        screen.SendKeys('<Enter>')
        time.sleep(3)   # colocado para substituir a tela17
        stringPart = "DRJH1"
        x1, y1, x2, y2 = 12, 1, 12, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        if stringPart in rc and condicao == False:
            esperaPorString(tela17, lin17, col17) # deu ruim aqui, na segunda rodada ele vai pra tela 18, colocar no if
            write(screen, 12, 2, "s")
            screen.SendKeys('<Enter>')
            esperaPorString(tela18, lin18, col18)   # colocar uma condição caso dê erro na linha de pesquisa
        elif stringPart not in rc:
            print(f"A string {stringPart} está em outra posição")
            exit(1)
        write(screen, 2, 15, "sh")
        screen.SendKeys('<Enter>')
        esperaPorString(tela17, lin17, col17)
        if condicao == False:
            screen.SendKeys('<Pf3>')
            condicao = True
        esperaPorString(tela16, lin16, col16)
        write(screen, 2, 15, "3")
        write(screen, 18, 8, "//ENDS" + package[-4:] + " JOB (END,END,9999,9999),'REROPRJ',")
        write(screen, 19, 8, "// CLASS=Q,MSGCLASS=G,TIME=1440")
        write(screen, 20, 8, "/*JOBPARM L=999999,C=999999")
        screen.SendKeys('<Enter>')
        esperaPorString(shipmentMsg, linShipmentMsg, colShipmentMsg)
        time.sleep(3) # tirar depois que acertar na função esperaporstring
        write(screen, 2, 15, "5")
        screen.SendKeys('<Enter>')
        esperaPorString(tela19, lin19, col19)
        x1, y1, x2, y2 = 5, 1, 5, 80
        rc = str(screen.Area(x1,y1,x2,y2))
        string = "SUBMITTED"
        if string in rc:
            x1 += 1; x2 = x1
            rc = str(screen.Area(x1,y1,x2,y2))
            stageString = "RC=00"
            if stageString in rc:
                print(f"Deu certo a distribuição do pacote {package}")
                flag = 0
            else:
                print(f"Deu algum problema na distribuição do pacote {package}")
                flag = 1
        else:
            print("Linha errada na tela PACKAGE SHIPMENT STATUS")
        screen.SendKeys('<Pf3><Pf3>')
        esperaPorString(tela7, lin7, col7)
    
        for status in array6:
            if status == "EXEC-FAILED":
                array7.append("-")
            elif status == "EXECUTED" and flag == 0:
                array7.append("OK")
            else:
                array7.append("NOT OK")

def criaMatriz():
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
    sortedMatrix = sorted(matrix, key=lambda x: x[2], reverse=True)

    #Cria matrizes para execução, separação e negação e Final
    execPackageMatrix = []
    separaPackageMatrix = []
    deniePackageMatrix = []

    valorVerificado = set()
    for list in sortedMatrix:
        elemento = list[2]
        dataH = list[3]
        dataP = list[4]
        dataC = list[1]
        if elemento not in valorVerificado and ((dataH == dataP and dataP != 0) or (dataH == 0 and dataP != 0)):
            execPackageMatrix.append(list)
            primDataC = dataC
        elif elemento in valorVerificado and ((dataC >= primDataC and dataH == dataP and dataP != 0) or (dataH == 0 and dataP != 0)):
            deniePackageMatrix.append(execPackageMatrix.pop())
            execPackageMatrix.append(list)
        elif elemento in valorVerificado and dataH == dataP and dataP != 0 and dataC < primDataC:
            deniePackageMatrix.append(list)
        elif elemento not in valorVerificado and ((dataH == 0 and dataP == 0) or (dataH != dataP)):
            primDataC = dataC # verificar essa lógica
            separaPackageMatrix.append(list)
        elif elemento in valorVerificado and ((dataH == 0 and dataP == 0) or (dataH != dataP)):
            deniePackageMatrix.append(list)
        valorVerificado.add(elemento)

    #Junto com distPackage, gerará 2 arrays globais para compor a matriz final
    aproveAndExec(execPackageMatrix)
    lastMatrix = []
    for i in range(len(array1)):
        lastMatrix.append([array1[i], array2[i], array3[i], array4[i], array5[i], array6[i], array7[i]])

    #Cria planilhas no Excel
    planilha = "executePackage.csv"
    planilha2 = "deniePackage.csv"
    planilha3 = "separaPackage.csv"
    colunas = ['Pacote', 'Data Cast', 'Elemento', 'Data HMP', 'Data PRD', 'Execucao', 'Shipment']
    with (open(path + planilha, 'w', newline='') as csvFileExec,
         open(path + planilha2, 'w', newline='') as csvFileDenie,
         open(path + planilha3, 'w', newline='') as csvFileSepara
         ):
            writer = csv.writer(csvFileExec, delimiter=';')
            writer.writerow(colunas)
            for row in lastMatrix:
                writer.writerow(row)
            writer = csv.writer(csvFileDenie, delimiter=';')
            writer.writerow(colunas)
            for row in deniePackageMatrix:
                writer.writerow(row)
            writer = csv.writer(csvFileSepara, delimiter=';')
            writer.writerow(colunas)
            for row in separaPackageMatrix:
                writer.writerow(row)

    print('\n')
    for list in execPackageMatrix:
        print(list)

    print('\n')
    for list in lastMatrix:
        print(list)
        
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

def esperaPorString(tela, lin, col):
    #time.sleep(1)
    var = str(screen.WaitForString(tela, lin, col, "", "", 60000))
    if var == "False":
        print(f"ERRO! Verifique posição correta na tela {tela}, linha {lin}, coluna {col}")
        exit(1)
    else:
        time.sleep(1)
        pass

def main():
    proc = start()
    if proc != None:
        terminalConector()
        limpaArquivos()
        navegaPacote()
        gravaNomePacote()
        criaMatriz()
    else:
        stop(proc)
        exit(1)
    stop(proc)
    exit(0)

if __name__ == "__main__":
    main()
