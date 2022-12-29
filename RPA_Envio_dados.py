import pywhatkit
import keyboard
import time
from datetime import datetime
import win32com.client as win32
from win32com.client import Dispatch


arquivo = "C:\\RPA\\RPA _Pai\\Documento fim de ano.xlsx"


try:
    xls = Dispatch('Excel.Application')
    xls.Visible = True
    xls.DisplayAlerts = False
    
    xls.Workbooks.Open(arquivo)  

    xls.Worksheets('Planilha1')
        
    x = 3
    cell = xls.Range("A" + str(x)).Value

    while cell != None: 

        aluno = xls.Range("A" +str(x)).Value
        valor = xls.Range("E" + str(x)).Value
        numero = xls.Range("F"+ str(x)).Value

        mensagem = "Olá responsável pelo/a: " + aluno + ", gostaria das seguintes informações para o próximo ano letivo 2023:" + "\n" + "Ressaltando o valor de R$" + str(valor) + ", para 2023." + "\n" + "Desejo continuar na van" + "\n" + "Não desejo continuar na van" + "\n" + "Período letivo do/a: " + aluno + " caso deseje continuar." + "\n" + "Manhã" + "\n" + "Tarde" + "\n" + "Integral" + "\n"
    
        pywhatkit.sendwhatmsg(numero, mensagem, datetime.now().hour, datetime.now().minute + 2)
        time.sleep(2)
        keyboard.press_and_release('ctrl + w')
    
        x += 1 
        cell = xls.Range("A" + str(x)).Value

    xls.ActiveWorkbook.Save()
    xls.Application.Quit()


except Exception as ex:
    print(ex.args[0])
