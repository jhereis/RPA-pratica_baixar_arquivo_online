import xlrd
import re
import pyautogui as py
import time
import sys
import ctypes

book = xlrd.open_workbook("C:\\Users\\jeniffer\\Desktop\\ExceldeCadastro.xlsx")
print ("Número de abas: ", book.nsheets)
print ("Nomes das Planilhas:", book.sheet_names())


sh = book.sheet_by_index(0)
print(sh.name, sh.nrows, sh.ncols)
print("Valor da célula CNPJ é ", sh.cell_value(rowx=0, colx=1))
cnpj=sh.cell_value(rowx=0, colx=2)

print("-- ")
py.hotkey("win", "r")
time.sleep(2)
#py.write("“C:\Program Files\Internet Explorer\iexplore.exe” -extoff""")
py.write("chrome.exe")
py.press("enter")

time.sleep(5)
py.hotkey('win','up')
time.sleep(2)

#py.press("tab")
#py.press("tab")

py.write("exemplo de site")
py.press("enter")

time.sleep(2)

try:
    certifica = py.locateCenterOnScreen('C:\\Jheniffer\\exemplo de foto para acessar no site.png')
    py.click(certificado)
except Exception as error:
    print('Erro: Imagem do link CERTIFICADO DIGITAL não localizada: ' + repr(error))
    var_cont = 0
    while var_cont < 1:
        try:
            if (py.locateCenterOnScreen("C:\\Jheniffer\\certificado.png") is None) == False:
                # py.moveTo(10,10)
                certifica = py.locateCenterOnScreen("C:\\Jheniffer\\certificado.png")
                print("Encontrou certificado")
                py.click(certifica)
                var_cont= 1

        except:
            print("Aguardando")

#py.write("i")
#py.press("tab")
#py.press("tab")
time.sleep(5)

#Autenticacao no site

try:
    zamp = py.locateCenterOnScreen('C:\\Jheniffer\\zamp.png')
    py.click(zamp)
except Exception as error:
    print('Erro: Imagem do link zamp não localizada: ' + repr(error))
    var_cont = 0
    while var_cont < 1:
        try:
            if (py.locateCenterOnScreen("C:\\Jheniffer\\zamp.png") is None) == False:
                # py.moveTo(10,10)
                zamp = py.locateCenterOnScreen("C:\\Jheniffer\\zamp.png")
                print("Encontrou certificado")
                py.click(zamp)
                var_cont= 1

        except:
            print("Aguardando")



# SE O RPA ENCONTRAR FOTO DA IMAGEM DO CERTIFICADO, ELE DA UM OK

try:
    ok_certificado = py.locateCenterOnScreen('C:\\Jheniffer\\ok_certificado.png')
    py.click(ok_certificado)
except Exception as error:
    var_cont = 0
    while var_cont < 1:
        try:
            if (py.locateCenterOnScreen("C:\\Jheniffer\\ok_certificado.png") is None) == False:
                # py.moveTo(10,10)
                ok_certificado = py.locateCenterOnScreen("C:\\Jheniffer\\ok_certificado.png")
                print("Encontrou ok_certificado")
                py.click(ok_certificado)
                var_cont = 1

        except:
            print("Aguardando")


time.sleep(2)

try:
    auto = py.locateCenterOnScreen('C:\\Jheniffer\\auto.png')
    py.click(auto)
except Exception as error:
    var_cont = 0
    while var_cont < 1:
        try:
            if (py.locateCenterOnScreen("C:\\Jheniffer\\auto.png") is None) == False:
                # py.moveTo(10,10)
                auto = py.locateCenterOnScreen("C:\\Jheniffer\\auto.png")
                print("Encontrou auto")
                py.click(auto)
                var_cont = 1

        except:
            print("Aguardando")


#botao pesquisa

time.sleep(3)

#
try:
    pesquisa = py.locateCenterOnScreen('C:\\Jheniffer\\pesquisa.png')
    py.click(pesquisa)
except Exception as error:
    var_cont = 0
    while var_cont < 1:
        try:
            if (py.locateCenterOnScreen("C:\\Jheniffer\\pesquisa.png") is None) == False:
                # py.moveTo(10,10)
                auto = py.locateCenterOnScreen("C:\\Jheniffer\\pesquisa.png")
                print("Encontrou pesquisa")
                py.click(pesquisa)
                var_cont = 1

        except:
            print("Aguardando")


py.press("tab")
py.press("tab")
#-----------------------------------------------------------------------------


time.sleep(5)

time.sleep(1000)
#-----------------------------------------------------------------------------
#

print("-- Iniciando consulta de CNPJ")
for rx in range(sh.nrows):
    print(sh.cell_value(rowx=rx, colx=1))
    cnpj = sh.cell_value(rowx=rx, colx=1)
  #.replace(".", "").replace("/", "").replace("-", "")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.press("tab")
    py.write(cnpj)
    py.press("tab")
    py.press("enter")
    var_nome = 'VERIFICAR_DEBITO_'
    time.sleep(3)
    var_cont = 0
    while var_cont < 5:

        try:
            if (py.locateCenterOnScreen("C:\\Conta_corrente\\PASTA DO ROBO\\nao_pendencia.png") is None) == False:
                # py.moveTo(10,10)
                var_nome = 'SEM_DEBITO_'
                var_cont = 5
                print("achou zero pendencia")

        except:
            print("Aguardando nao pendencia")
            var_cont = var_cont + 1
        time.sleep(2)


    #----------------------------------


    '''
    cnpj = cnpj.replace('.', '')
    cnpj = cnpj.replace('/', '')
    cnpj = str(cnpj)
    
    time.sleep(5)
    py.hotkey("Ctrl", "p")
    time.sleep(2)
    time.sleep(2)
    py.hotkey("enter")
    time.sleep(3)
    arquivo2 = (cnpj_num + '_malha_fina')
    py.write(arquivo2)
    time.sleep(2)
    py.hotkey("Ctrl", "l")
    py.write(r'C:\Conta_corrente\ARQUIVOS_RJ')
    time.sleep(2)
    py.hotkey("Alt", "s")
    '''
    foto = py.screenshot()

    cnpj = cnpj.replace('.', '')
    cnpj = cnpj.replace('/', '')
    cnpj = str(cnpj + '.pdf')
    salva_ok = ("C:\\Conta_corrente\\PASTA DO ROBO\\ " + var_nome + cnpj)
    print(salva_ok)
    foto.save(salva_ok)

    time.sleep(3)
