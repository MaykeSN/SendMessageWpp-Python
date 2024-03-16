import openpyxl
from urllib.parse import quote
import webbrowser
from time import sleep
import pyautogui
import tkinter as tk
from tkinter import filedialog

def enviar_mensagens():
    arquivo_contatos = filedialog.askopenfilename(title="Selecione o arquivo de contatos", filetypes=[("Arquivos do Excel", "*.xlsx")])
    if arquivo_contatos:
        workbook = openpyxl.load_workbook(arquivo_contatos)
        contatos_almoco = workbook.active

        webbrowser.open('https://web.whatsapp.com/')
        sleep(25)

        for linha in contatos_almoco.iter_rows(min_row=1):
            nome = linha[0].value
            telefone = linha[1].value
            mensagem = f'Eae bobo bora almoçar? Hora do ping pong hihihi!'
            try:
                link_mensagem_zap = f'https://web.whatsapp.com/send?phone={telefone}&text={quote(mensagem)}'
                webbrowser.open(link_mensagem_zap)
                sleep(16)
                seta = pyautogui.locateCenterOnScreen('enviarwpp.png')
                sleep(6)
                pyautogui.click(seta[0], seta[1])
                sleep(6)
                pyautogui.hotkey('ctrl', 'w')
                sleep(6)
            except Exception as e:
                print(f'Não foi possivel enviar mensagem para {nome}: {str(e)}')
                pyautogui.hotkey('ctrl', 'w')
                with open('erros.csv', 'a', newline='', encoding='utf8') as arquivo:
                    arquivo.write(f'Não foi possivel enviar mensagem para {nome}')

# Configurando a interface gráfica
root = tk.Tk()
root.title("Envio de Mensagens via WhatsApp")
root.geometry("400x70")  # Definindo o tamanho da janela

btn_iniciar = tk.Button(root, text="Chamar os meia boca pra almoçar", command=enviar_mensagens, width=40, height=2)  # Definindo o tamanho do botão
btn_iniciar.pack(pady=20)

root.mainloop()
