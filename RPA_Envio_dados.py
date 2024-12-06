import pywhatkit
import keyboard
import time
from datetime import datetime, timedelta
import win32com.client
from tkinter import *
from tkinter import filedialog, messagebox
import os
import traceback

arquivo_selecionado = ""

def ler_arquivo():
    global arquivo_selecionado
    arquivo_selecionado = filedialog.askopenfilename(
        initialdir=os.path.dirname(os.path.abspath(__file__)),
        title="Selecione um arquivo",
        filetypes=[("Planilha Excel", "*.xlsx"), ("Planilha Excel", "*.xls")]
    )
    if arquivo_selecionado:
        print(f"Arquivo selecionado: {arquivo_selecionado}")
        messagebox.showinfo("Arquivo Selecionado", "Arquivo carregado com sucesso!")
    else:
        print("Nenhum arquivo selecionado")

def envio_mensagem(arquivo):
    try:
        if not arquivo:
            messagebox.showwarning("Aviso", "Selecione um arquivo primeiro!")
            return

        xls = win32com.client.Dispatch("Excel.Application")
        xls.Visible = False
        xls.DisplayAlerts = False

        workbook = xls.Workbooks.Open(arquivo)
        sheet = workbook.Worksheets("Planilha1")

        x = 3
        cell = sheet.Range("A" + str(x)).Value

        while cell is not None:
            aluno = sheet.Range("A" + str(x)).Value
            valor = sheet.Range("E" + str(x)).Value
            numero = sheet.Range("F" + str(x)).Value

            mensagem = f"""Olá responsável pelo/a: {aluno}, gostaria das seguintes informações para o próximo ano letivo 2023:
            Ressaltando o valor de R$ {valor}, para 2025.
            Desejo continuar na van
            Não desejo continuar na van
            Período letivo do/a: {aluno} caso deseje continuar.
            Manhã
            Tarde
            Integral
            """

            horario = datetime.now()
            pywhatkit.sendwhatmsg(numero, mensagem, horario.hour, horario.minute + 1)
            time.sleep(2)
            keyboard.press_and_release("ctrl + w")

            x += 1
            cell = sheet.Range("A" + str(x)).Value

        workbook.Save()
        xls.Quit()
        messagebox.showinfo("Sucesso", "Mensagens enviadas com sucesso!")

    except Exception as e:
        print(traceback.format_exc())
        messagebox.showerror("Erro", f"Ocorreu um erro: {e}")

janela = Tk()
janela.title("Envio de mensagens para os pais")
janela.geometry("450x150")

lbl_mensagem_excel = Label(janela, text="Selecione o arquivo desejado")
lbl_mensagem_excel.grid(column=0, row=0, padx=10, pady=10)

btn_excel = Button(janela, text="Selecionar", command=ler_arquivo)
btn_excel.grid(column=1, row=0, padx=10, pady=10)

lbl_mensagem = Label(janela, text="Clique no botão abaixo para executar a tarefa")
lbl_mensagem.grid(column=0, row=1, padx=10, pady=10)

btn_executar = Button(janela, text="Executar", command=lambda: envio_mensagem(arquivo_selecionado))
btn_executar.grid(column=1, row=1, padx=10, pady=10)

janela.mainloop()
