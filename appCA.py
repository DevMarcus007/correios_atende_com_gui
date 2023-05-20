import tkinter as tk
import os
import sys
from datetime import datetime
from time import sleep
import pandas as pd
from tkcalendar import Calendar
from PIL import Image, ImageTk
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from tkinter.scrolledtext import ScrolledText
from config import *


def show_calendar():
    def select_date():
        date_label["text"] = calendar.selection_get().strftime('%d/%m/%Y')
        calendar_window.destroy()

    calendar_window = tk.Toplevel(window)
    calendar_window.title("Calendário")

    calendar = Calendar(calendar_window, selectmode="day")
    calendar.pack()

    date_button = tk.Button(calendar_window, text="Selecionar", command=select_date)
    date_button.pack()

def erro(texto):
    # Cria uma nova janela popup
    popup = tk.Toplevel(window)
    popup.title("Atenção")

    mensagem = tk.Label(popup, text=texto)
    mensagem.pack(padx=30, pady=30)

    # Botão para fechar a janela popup
    fechar = tk.Button(popup, text="OK", command=popup.destroy)
    fechar.pack(pady=10)

    # Define a posição central da janela popup
    popup.geometry("+{}+{}".format(int(window.winfo_screenwidth()/2 - popup.winfo_reqwidth()/2), int(window.winfo_screenheight()/2 - popup.winfo_reqheight()/2)))

    

def search_atendimentos(usuario, password, date, pagina, mcu):
    def custom_print(*args, **kwargs):
        text = ' '.join(map(str, args))
        text_box.insert(tk.END, text + '\n')
        text_box.see(tk.END)
        window.update()

    sys.stdout.write = custom_print

    df_postados_correios_atende = pd.DataFrame(columns=['Código Objeto', 'CEP Destinatário', 'Dimensões (cm)',
                                                        'Peso (g)', 'Serviço', 'Valor'])
    dia = date[0:2]
    mes = date[2:4]
    ano = date[4:]

    
    navegador = webdriver.Chrome(executable_path= "./chromedriver.exe")

    #Login no sistema------------------------------------------------------------------------------
    navegador.get(pagina[0])
    sleep(2)
    navegador.find_element(By.XPATH,'//*[@id="mcuAgencia2"]').send_keys(mcu)
    tentativas = 5
    while tentativas != 0:
        try:
            navegador.find_element(By.XPATH,'//*[@id="username"]').send_keys(usuario)
            navegador.find_element(By.XPATH,'//*[@id="password"]').send_keys(password)
            navegador.find_element(By.XPATH,'//*[@id="fm1"]/div[2]/button').click()
            break
        except:
            sleep(1)
            tentativas-=1
            pass
    
    navegador.get(pagina[1])
    sleep(1.5)

    #Pesquisa dos atendimentos e criação da lista de atendimentos------------------------------------
   
    navegador.find_element(By.XPATH,'//*[@id="modal-pesquisa"]/i').click()
    sleep(1.5)
    navegador.find_element(By.XPATH,'//*[@id="dataAtendimentoInicial"]').send_keys(f'{dia}/{mes}/{ano}')
    sleep(1.5)
    navegador.find_element(By.XPATH,'//*[@id="dataAtendimentoFinal"]').send_keys(f'{dia}/{mes}/{ano}')
    sleep(1.5)
    navegador.find_element(By.XPATH,'//*[@id="idCorreios"]').clear()
    sleep(1.5)
    navegador.find_element(By.XPATH,'//*[@id="pesquisarAtendimento"]').click()
    
    tentativas = 5
    while tentativas != 0:
        try:
            select = Select(navegador.find_element(By.XPATH, '//*[@id="atendimentosDia_length"]/label/select'))
            select.select_by_value('100')
            break
        except:
            sleep(2)
            tentativas-=1
            pass


    atendimentos = navegador.find_elements(By.ID ,'atendimentosDia')
    table_html = atendimentos[0].get_attribute('outerHTML')
    df = pd.read_html(str(table_html))
    df = df[0]
    lista_atendimentos = df['ATENDIMENTO']
    lista_atendimentos = df['ATENDIMENTO'].to_list()
    lista_atendimentos.pop(-1)
    sleep(1)
    lista_pendentes = []

    for item in lista_atendimentos:
        success = False
        while success == False:
            try:
                navegador.find_element(By.XPATH,'//*[@id="atendimentosDia_filter"]/label/input').clear()       
                navegador.find_element(By.XPATH, '//*[@id="atendimentosDia_filter"]/label/input').send_keys(item)

                navegador.find_element(By.XPATH, ' //*[@id="atendimentosDia"]/tbody/tr/td[2]/a').click()
                sleep(1.5)        
                atendimento = navegador.find_element(By.XPATH,rf'//*[@id="atendimentosDia"]/tbody/tr[{(2)}]/td/div/table')
                sleep(1.5)

                table_html2 = atendimento.get_attribute('outerHTML')
                lista_objetos = pd.read_html(str(table_html2))
                df_temporario = lista_objetos[0]
                df_postados_correios_atende = pd.concat([df_postados_correios_atende,df_temporario])
                sleep(1.5)
                success = True
                print(f'{item} - Sucesso')
            except Exception as e:
                print(f'{item} - Falha. Tentando novamente...')
                continue                    


    navegador.quit()
    print(f'Total de Atendimentos: {len(lista_atendimentos)}')
    print(f'Total de Objetos: {len(df_postados_correios_atende)}')
    return df_postados_correios_atende



def on_button_click():
    date = date_label["text"]
    usuario = entry_user.get()
    
    password = entry_password.get()

    date_obj = datetime.strptime(date, '%d/%m/%Y')
    meses = {
        "January": "JANEIRO",
        "February": "FEVEREIRO",
        "March": "MARÇO",
        "April": "ABRIL",
        "May": "MAIO",
        "June": "JUNHO",
        "July": "JULHO",
        "August": "AGOSTO",
        "September": "SETEMBRO",
        "October": "OUTUBRO",
        "November": "NOVEMBRO",
        "December": "DEZEMBRO"
    }
    mespasta_en = date_obj.strftime('%B')
    mespasta = meses[mespasta_en] + '_' + date_obj.strftime('%Y')

    if len(usuario) != 8:
        erro("Verifique a matrícula digitada")
    else:
        df_postados_correios_atende = search_atendimentos(usuario, password, date, pagina, mcu)

        df_postados_correios_atende['Valor'] = df_postados_correios_atende['Valor'] / 100

        #Verifica a existência das pastas de destino do arquivo gerado
        verifica_pasta(mespasta)

        file_path = pasta_destino(mespasta, date)

    
        df_postados_correios_atende.to_excel(file_path, sheet_name='Relatorio de Postagem', index=False)
        erro(f"Aplicação Finalizada\n - {len(df_postados_correios_atende)} objetos postados no Correios Atende")



window = tk.Tk()
window.title('Correios Atende')

image = Image.open("./logo.webp").convert("RGBA")
image = image.resize((300, 100), Image.ANTIALIAS)
logo = ImageTk.PhotoImage(image)
window.iconbitmap("./icon.ico")
logo_label = tk.Label(image=logo, highlightthickness=0)
logo_label.grid(row=0, column=1, columnspan=5)

title_label = tk.Label(justify='right')
title_label["text"] = "Extração dos dados de Postagem Correios Atende"
title_label["font"] = ("Arial", 15)
title_label.grid(row=4, column=1, columnspan=5)

date_label = tk.Label(justify='left')
date_label["text"] = datetime.now().strftime('%d/%m/%Y')
date_label.grid(row=5, column=1, sticky="w", padx=10, pady=10)

calendar_button = tk.Button(justify='right')
calendar_button["text"] = "Alterar Data"
calendar_button["command"] = show_calendar
calendar_button.grid(row=5, column=2, sticky="w", padx=10, pady=10)

label_user = tk.Label(window, text='Matrícula - Correios Atende:')
label_user.grid(row=6, column=1, sticky="w", padx=10, pady=10)
entry_user = tk.Entry(window)
entry_user.grid(row=6, column=2, columnspan=6, sticky="w", padx=10, pady=10)

label_password = tk.Label(window, text='Senha - Correios Atende:')
label_password.grid(row=7, column=1, sticky="w", padx=10, pady=10)
entry_password = tk.Entry(window, show='*')
entry_password.grid(row=7, column=2, columnspan=6, sticky="w", padx=10, pady=10)

submit_button = tk.Button(window)
submit_button["text"] = "Verificar Postagens"
submit_button["command"] = on_button_click
submit_button.grid(row=8, column=1, columnspan=4, pady=10)


text_box = ScrolledText(window)
text_box.grid(row=9, column=1, columnspan=5, padx=10, pady=10)


window.mainloop()
