import pywhatkit
import keyboard
import time
from datetime import datetime
import win32com.client as win32
from win32com.client import Dispatch
from tkinter import *
from tkinter import filedialog

def ler_arquivo():
    arquivo = filedialog.askopenfilename(initialdir="C:\\RPA\\RPA_Envio_Whats-main", title="Selecione um arquivo", filetypes=(("Planilha do Microsoft Excel", "*.xlsx"),("Planilha do Microsoft Excel","*.xls")))
    
    print(arquivo)
    return arquivo

def envio_mensagem(arquivo):
    try:
        xls = Dispatch('Excel.Application')
        xls.Visible = False
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

janela = Tk()

janela.title("Envio de mensagens para os pais")
janela.geometry("450x130")

lbl_mensagem_excel = Label(janela, text="Selecione o arquivo desejado")
lbl_mensagem_excel.grid(column=0, row=0, padx=10, pady=10)

btn_excel = Button(janela, text="Selecionar", command=ler_arquivo)
btn_excel.grid(column=1, row=0, padx=10, pady=10)

lbl_mensagem = Label(janela, text="Clique no botão abaixo para executar a tarefa")
lbl_mensagem.grid(column=0, row=1, padx=10, pady=10)

btn_executar = Button(janela, text="Executar", command= lambda: envio_mensagem(btn_excel))
btn_executar.grid(column=1, row=1, padx=10, pady=10)

janela.mainloop()