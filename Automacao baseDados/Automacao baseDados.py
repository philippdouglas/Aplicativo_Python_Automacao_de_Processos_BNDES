import json
import locale
import tkinter as tk
from tkinter import PhotoImage, scrolledtext
from dotenv import load_dotenv
from tkinter import ttk
from tkinter import messagebox
import threading
import webbrowser
import shutil
from bs4 import BeautifulSoup
from colorama import Fore, Style
import pyautogui
import requests
from parse import *
import pandas as pd
import os
import datetime
from datetime import datetime
from tqdm import tqdm
import time
import keyboard
import tqdm
import io
import subprocess
import sys
import pytz
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
import urllib3
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from pathlib import Path
import glob 
import re
import sys


#==========Variáveis globais

# Variável global para armazenar o userid
saved_userid = ""
saved_pwd = ""

#==========Classe Dica de ferramenta

# Classe para exibir dica de ferramentas nos botões
class ToolTip:
    
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tooltip = None
        self.widget.bind("<Enter>", self.show_tooltip)
        self.widget.bind("<Leave>", self.hide_tooltip)

    def show_tooltip(self, event):
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25

        self.tooltip = tk.Toplevel(self.widget)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")

        label = ttk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1)
        label.pack()

    def hide_tooltip(self, event):
        if self.tooltip:
            self.tooltip.destroy()
            self.tooltip = None
'''
#==========Classe Login

# Classe para login no proxy da rede
class LoginApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Automação baseDados')
        self.iconbitmap("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/assets/Ico.ico")
        self.resizable(width=False, height=False)
        self.logo_image = tk.PhotoImage(file='C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/assets/Logo CO.png')
        self.logo_label = tk.Label(self, image=self.logo_image)

        # Obtém as dimensões da área de trabalho
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()

        # Define as dimensões e a posição da janela centralizada
        login_width = 400
        login_height = 150
        login_x = (screen_width // 2) - (login_width // 2)
        login_y = (screen_height // 2) - (login_height // 2) - 200
        self.geometry(f"{login_width}x{login_height}+{login_x}+{login_y}")

        # Intercepta o evento de fechamento da janela
        self.protocol("WM_DELETE_WINDOW", self.on_closing)

        login_label = tk.Label(self, text='Login')
        self.login_entry = tk.Entry(self)
        self.tooltip = ToolTip(self.login_entry, 'Insira seu usuário e senha de rede')
        senha_label = tk.Label(self, text='Senha')
        self.senha_entry = tk.Entry(self, show='*')
        logar_button = tk.Button(self, text='   Logar   ', command=self.logar)
        self.bind('<Return>', self.press_enter)

        self.logo_label.grid(row=0, column=0, padx=10, pady=20, rowspan=2)
        login_label.grid(row=0, column=1, padx=10, pady=20)
        self.login_entry.grid(row=0, column=2, padx=10, pady=15)
        senha_label.grid(row=1, column=1, padx=10, pady=5)
        self.senha_entry.grid(row=1, column=2, padx=10, pady=5)
        logar_button.grid(row=2, column=1, columnspan=2, padx=30, pady=5)

    def on_closing(self):
        if messagebox.askokcancel("Fechar", "Deseja fechar o aplicativo?"):
            self.destroy()
            sys.exit()

    def press_enter(self, event):
        self.logar()

    def conf_proxy(self, login, pwd):
        global userid  # Acessa a variável global
        userid = login
        pwd = pwd.replace('#', '%23').replace('@', '%40')
        os.environ['http_proxy'] = f"http://{userid}:{pwd}@proxy.inf.bndes.net:8080"
        os.environ['https_proxy'] = f"http://{userid}:{pwd}@proxy.inf.bndes.net:8080"

    def logar(self):
        global saved_userid, saved_pwd  # Acessa as variáveis globais
        login = self.login_entry.get()
        senha = self.senha_entry.get()
        self.conf_proxy(login, senha)
        saved_userid = login
        saved_pwd = senha.replace('#', '%23').replace('@', '%40')

        self.close_window()

    def close_window(self):
        self.destroy()
        self.quit()

if __name__ == '__main__':
    app = LoginApp()
    app.mainloop()
'''
#============================================================================================================================================================================================================================

#==========Menu ABN

# Variável de controle para verificar se a função está em andamento          
buscarABN = False
  
def tab1():
    
    # Verifica o número de linhas onde a coluna Setor_Tema está vazia e verifica a data da ultima linha de Setor_Tema Preenchida                       
    def statusABN():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file_path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx'
        sheet_name = 'ABN_producao'
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        empty_lines = df[df['Setor_Tema'].isnull() & (df['Ano_publicacao2'] >= 2023)].shape[0]
        last_class_date = df[df['Setor_Tema'].notnull()]['Data'].max().strftime('%d/%m/%Y')
        text = f" {current_time}- Status classificação ABN.xlsx\n Postagens para classificar: {empty_lines} || Última classificação: {last_class_date} "
        label_text.configure(text=text)
        label_text.configure(relief="ridge") 
        print_to_text(f"{current_time} \u2714Status ABN.xlsx, Atualizado...")
    
    # Busca Título,Link,Data e Hora das postagens da página Agencia de noticia do bndes   
    def buscarABN():
        
        # Desabilitar os botões em quanto a função buscarABN estiver sendo executada
        button_buscar.configure(state="disabled")
        button_classificar.configure(state="disabled")
        button_atualizar.configure(state="disabled")
        restaurar_button.configure(state="disabled")
        
        global buscarABN
        # Verificar se a função já está em andamento
        if buscarABN:
            return
        # Definir a variável de controle como global
        buscarABN = True
        '''
        # Executa a autenticação de usuário de rede       
        global saved_userid
        global saved_pwd 
        saved_pwd = saved_pwd.replace('#', '%23').replace('@', '%40')
        http_proxy = f"http://{saved_userid}:{saved_pwd}@proxy.inf.bndes.net:8080"
        os.environ['http_proxy'] = http_proxy
        os.environ['https_proxy'] = http_proxy
        '''       
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        # Faz o Backup do arquivo base antes de fazer as buscas e a atualização
        try:
            from datetime import datetime
            data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx'
            destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup ABN/'
            nome_arquivo = f'ABN_{data_hora_atual}.xlsx'
            destino_completo = destino + nome_arquivo
            shutil.copy(origem, destino_completo)
            print_to_text(f"{data_hora_atual} \u2714Backup do arquivo ABN.xlsx concluído...")
        except Exception as e:
            print_to_text(f"{data_hora_atual} !Ocorreu um erro ao fazer o backup do arquivo ABN.xlsx: {str(e)}")
            
        # Busca               
        try:
        
            titulos = []
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36 Edge/B08C3901'
            }
            print_to_text(f"{current_time} ->buscarABN Título...")
            contador = 1
            for i in range(1, total_iterations + 1):
                url = f"https://agenciadenoticias.bndes.gov.br/system/modules/br.gov.bndes.agc/ajax_elements/ultimas-noticias-ajax.jsp?page={i}"
                response = requests.get(url, headers=headers)
                html = response.text
                soup = BeautifulSoup(html, "html.parser")
                text = soup.select("h4")
                               
                if len(text) == 0:
                    break
                             
                for t in text:
                    titulos.append(t.text.strip())
                    print_to_text2(f"{contador}. {t.text.strip()}")
                    contador += 1
                df_titulo = pd.DataFrame({"Título": titulos})
               
            #Link
            links = []
            print_to_text(f"{current_time} ->buscarABN Link...")
            contador = 1    
            for i in range(1, total_iterations + 1):
                url = f"https://agenciadenoticias.bndes.gov.br/system/modules/br.gov.bndes.agc/ajax_elements/ultimas-noticias-ajax.jsp?page={i}"
                response = requests.get(url)
                html = response.text
                soup = BeautifulSoup(html, "html.parser")
                text = soup.select("a")
                    
                if len(text) == 0:
                    break

                for t in text:
                    links.append("https://agenciadenoticias.bndes.gov.br" + t['href'])
                    print_to_text2(f"{contador}. https://agenciadenoticias.bndes.gov.br" + t['href'])
                    contador += 1
                df_link = pd.DataFrame({"Link": links})
                               
            #Data
            data = []
            print_to_text(f"{current_time} ->buscarABN Data...")
            contador = 1 
            for i in range(1, total_iterations + 1):
                url = f"https://agenciadenoticias.bndes.gov.br/system/modules/br.gov.bndes.agc/ajax_elements/ultimas-noticias-ajax.jsp?page={i}"
                response = requests.get(url)
                html = response.text
                soup = BeautifulSoup(html, "html.parser")
                text = soup.select(".info-data strong")
                
                if len(text) == 0:
                    break

                for t in text:
                    data.append(t.text.strip())
                    print_to_text2(f"{contador}. "+t.text.strip())
                    contador += 1
                df_data = pd.DataFrame({"Data": data})

            #Hora
            hora = []
            print_to_text(f"{current_time} ->buscarABN Hora...")
            contador = 1 
            for i in range(1, total_iterations + 1):
                url = f"https://agenciadenoticias.bndes.gov.br/system/modules/br.gov.bndes.agc/ajax_elements/ultimas-noticias-ajax.jsp?page={i}"
                response = requests.get(url)
                html = response.text

                soup = BeautifulSoup(html, "html.parser")

                spans = soup.select("span.info-data")
                if len(spans) == 0:
                    break

                for span in spans:
                    content = span.contents
                    if len(content) > 0:
                        hora.append(content[0].strip())
                        print_to_text2(f"{contador}. "+content[0].strip())
                        contador += 1
            df_hora = pd.DataFrame({"Hora": hora})
            
            def verificar_igualdade(df_titulo, df_link, df_data, df_hora):
                num_titulos = len(df_titulo)
                num_links = len(df_link)
                num_datas = len(df_data)
                num_horas = len(df_hora)

                if num_titulos == num_links == num_datas == num_horas:
                    print_to_text(f"{current_time} \u2714Número de títulos, links, datas e horas são iguais.")
                else:
                    print_to_text(f"{current_time} !Erro: Número de títulos, links, datas e horas são diferentes.")
                print_to_text(f"{current_time} ->Número de Títulos: {num_titulos}")
                print_to_text(f"{current_time} ->Número de Links: {num_links}")
                print_to_text(f"{current_time} ->Número de Datas: {num_datas}")
                print_to_text(f"{current_time} ->Número de Horas: {num_horas}")

            verificar_igualdade(df_titulo, df_link, df_data, df_hora)

            #Colunas de Classificação
            df_tipo_Releases = pd.DataFrame({"Tipo Releases": []})
            df_setor_tema = pd.DataFrame({"Setor_Tema": []})
            df_iniciativa_produto = pd.DataFrame({"Iniciativa_Produto": []})
            df_estrategia = pd.DataFrame({"Estrategia": []})
            df_classificacao = pd.concat([df_tipo_Releases, df_setor_tema, df_iniciativa_produto, df_estrategia], axis=1)

            #Filtros Data    
            df_data['Data'] = pd.to_datetime(df_data['Data'], format="%d/%m/%Y")
            df_dataFiltros = df_data
            df_dataFiltros['Dia_publicacao'] = df_data['Data'].dt.day
            df_dataFiltros['Mes_publicacao2'] = df_data['Data'].dt.month
            df_dataFiltros['Mes_publicacao2'] = df_data['Data'].dt.strftime('%b')
            df_dataFiltros['Mes_publicacao2'] = df_data['Mes_publicacao2']
            df_dataFiltros = df_dataFiltros.replace("Dec", "12")
            df_dataFiltros = df_dataFiltros.replace("Nov", "11")
            df_dataFiltros = df_dataFiltros.replace("Oct", "10")
            df_dataFiltros = df_dataFiltros.replace("Sep", "9")
            df_dataFiltros = df_dataFiltros.replace("Aug", "8")
            df_dataFiltros = df_dataFiltros.replace("Jul", "7")
            df_dataFiltros = df_dataFiltros.replace("Jun", "6")
            df_dataFiltros = df_dataFiltros.replace("May", "5")
            df_dataFiltros = df_dataFiltros.replace("Apr", "4")
            df_dataFiltros = df_dataFiltros.replace("Mar", "3")
            df_dataFiltros = df_dataFiltros.replace("Feb", "2")
            df_dataFiltros = df_dataFiltros.replace("Jan", "1")
            df_dataFiltros['Ano_publicacao2'] = df_data['Data'].dt.year
            semana = {
                    0: 'Seg',
                    1: 'Ter',
                    2: 'Qua',
                    3: 'Qui',
                    4: 'Sex',
                    5: 'Sáb',
                    6: 'Dom'
                    }
            df_dataFiltros['Dia_da_semana_publicacao'] = df_data['Data'].dt.weekday.map(semana)
            df_dataFiltros = df_dataFiltros.drop('Data', axis=1)
            df_data = df_data.drop('Dia_publicacao', axis=1)
            df_data = df_data.drop('Mes_publicacao2', axis=1)

            #Concat dataframes
            df_feed = pd.concat([df_link, df_titulo, df_data, df_hora, df_classificacao, df_dataFiltros], axis=1)

            #Filtragem 2023
            df_feed = df_feed[df_feed['Data'] >= '01-01-2023']
                
            #Deleta postagens que não possuem textos
            df_feed = df_feed[df_feed['Título'].str.strip() != ""]
            df_feed.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/ABN(feed).xlsx", index=False)
            print_to_text(f"{current_time} \u2714Buscas realizadas com sucesso...")
            
            #Merge

            try:
                print_to_text(f"{current_time} ->Atualizando arquivo ABN.xlsx...")
                df_base = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx', sheet_name='ABN_producao')
                df_base_resultados = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx', sheet_name='ABN_resultados')

                df_feed = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/ABN(feed).xlsx')
                df_feed = df_feed[~df_feed["Título"].isin(df_base["Título"])]
                df_base = pd.concat([df_base, df_feed], ignore_index=True)
                df_base = df_base.sort_values(by='Data', ascending=False)

                writer = pd.ExcelWriter('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx', engine='xlsxwriter')
                df_base.to_excel(writer, sheet_name='ABN_producao', index=False)  # Adiciona o parâmetro index=False
                df_base_resultados.to_excel(writer, sheet_name='ABN_resultados', index=False)  # Adiciona o parâmetro index=False
                writer.close()
                print_to_text(f"{current_time} \u2714Arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx Atualizado...")

            except Exception as e:
                print_to_text("Ocorreu um erro durante a junção do arquivo ABN.xlsx com o arquivo ABN(Feed):" + str(e))
                                     
        except Exception as e:
            print_to_text(f"{current_time} !Erro o arquivo ..02_Tratadas/ABN.xlsx está aberto por você ou por outro usuário,feche o arquivo para continuar, ou" + str(e))
        
                 
        # Atualizar a variável de controle
        buscarABN = False
        # Reabilitar o botão após a conclusão da função
        button_buscar.configure(state="normal")
        button_classificar.configure(state="normal")
        button_atualizar.configure(state="normal")
        restaurar_button.configure(state="normal")

    # Funçao para executar 1 função por vez
    def start_buscar_thread():
        # Criar uma nova thread para executar a função buscarABN
        buscar_thread = threading.Thread(target=buscarABN)
        buscar_thread.start()
                                    
    # Função para abrir o arquivo base e formatar a tabela em planilha
    def classificarABN():

        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            
        try:
            print_to_text(f"{current_time} ->Abrindo o arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx...\n{current_time} ->Aguarde a formatação da planilha em Tabela...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx'
            os.startfile(path)
            time.sleep(15)
            # Alternar para a planilha "ABN_producao"
            pyautogui.hotkey('ctrl', 'pgup')
            keyboard.press_and_release('ctrl+alt+t')
            time.sleep(2)
            keyboard.press_and_release('enter')
            time.sleep(1)
            keyboard.press_and_release('ctrl+b')
            print_to_text(f"{current_time} \u2714Formatação concluída...")
        except Exception as e:
            print_to_text(f"{current_time} !Falha [Classificar] Ocorreu o erro na função classificarABN(): {e}")

    # Função para fazer uma copia do arquivo da pasta 02_Tratadas para as pastas 03_Sistema e 04_Publico               
    def atualizaçãoPastas_ABN():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            print_to_text(f"{current_time} \u2714Atualizando pastas...")
            src_file = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/ABN.xlsx'
            dst_folder = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'
            dst_folder2 = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
            shutil.copy(src_file, dst_folder)
            shutil.copy(src_file, dst_folder2)
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'Atualizada...")
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'Atualizada...")
        except Exception as e:
            print_to_text(f"{current_time}- !Falha [Atualizar pastas] Ocorreu um erro ao copiar o arquivo para as pastas 03_Sistema e 04_Público: {e}")

    # Função para abrir o Manual de Automações.docx
    def open_manual():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
        try:
            print_to_text(f"{current_time} ->Abrindo C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx'
            os.startfile(path)
        except Exception as e:
            print_to_text(f"{current_time} !Erro, ocorreu um erro ao abrir o arquivo Manual automação.docx : {e}")
        
    # Função para restaurar o arquivo da pasta 02_Tratadas com o último backup da pasta 05_Backups
    def restaurar():
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Diretórios de origem e destino
        backup_dir = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup ABN/'
        dest_dirs = [
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
        ]

        # Obter o arquivo mais recente na pasta de backup
        backup_path = Path(backup_dir)
        files = backup_path.iterdir()
        latest_file = max(files, key=lambda f: f.stat().st_ctime) if files else None

        if latest_file:
            source_file = str(latest_file)

            # Copiar o arquivo para os diretórios de destino
            for dest_dir in dest_dirs:
                dest_file = Path(dest_dir) / 'ABN.xlsx'
                shutil.copy(source_file, str(dest_file))

            print_to_text(f"{current_time} \u2714 ABN.xlsx restaurado")
        else:
            print_to_text(f"{current_time} Nenhum arquivo de backup encontrado")
          
#============================================================================================================================================================================================================================

    #==========Interface menu ABN
          
    tab1_frame = ttk.Frame(tab_control)
    tab_control.add(tab1_frame, text='ABN         ')     
    tab1_frame.configure(height=100)

    def open_url(event):
        webbrowser.open("https://agenciadenoticias.bndes.gov.br")

    label = ttk.Label(tab1_frame, text='Atualização base de dados ABN, informações obtidas através do site: https://agenciadenoticias.bndes.gov.br\n\n', cursor='hand2')
    label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    label.bind("<Button-1>", open_url)

    button_frame = ttk.Frame(tab1_frame)
    button_frame.grid(row=2, column=0, padx=20, sticky="w")

    #==========Botão Status ABN
    
    button_status = tk.Button(button_frame, text="Status", command=statusABN, width=12)
    button_status.grid(row=0, column=0, padx=5, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab1_frame, text="Verifica o número de postagens sem classificação e a data da última classificação", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_status.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_status.bind("<Leave>", hide_status_tooltip)
  
  
    #==========Botão Buscar ABN
    
    button_buscar = tk.Button(button_frame, text="Buscar", command=buscarABN, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)
 
    # Botão Máscara Buscar ABN para não haver duplicação de command
    button_buscar = tk.Button(button_frame, text="Buscar", command=start_buscar_thread, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)

    # Função para exibir a legenda do botão buscar
    def show_buscar_tooltip(event):
        global tooltip_buscar
        tooltip_buscar = tk.Label(tab1_frame, text="Busca as 40 notícias mais recentes da Agência de Notícias do BNDES e atualiza a base ABN.xlsx", background="white", relief="solid", borderwidth=1)
        tooltip_buscar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão buscar
    button_buscar.bind("<Enter>", show_buscar_tooltip)

    # Função para ocultar a legenda do botão buscar
    def hide_buscar_tooltip(event):
        global tooltip_buscar
        if tooltip_buscar is not None:
            tooltip_buscar.destroy()
            tooltip_buscar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão buscar
    button_buscar.bind("<Leave>", hide_buscar_tooltip)


    #==========Botão Classificar ABN
    
    button_classificar = tk.Button(button_frame, text="Classificar", command=classificarABN, width=12)
    button_classificar.grid(row=0, column=2, padx=5, pady=5)

    # Função para exibir a legenda do botão classificar
    def show_classificar_tooltip(event):
        global tooltip_classificar
        tooltip_classificar = tk.Label(tab1_frame, text="Executa o Excel abrindo o arquivo ABN.xlsx da pasta 02_Tratadas e faz a formatação para planilha", background="white", relief="solid", borderwidth=1)
        tooltip_classificar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão classificar
    button_classificar.bind("<Enter>", show_classificar_tooltip)

    # Função para ocultar a legenda do botão classificar
    def hide_classificar_tooltip(event):
        global tooltip_classificar
        if tooltip_classificar is not None:
            tooltip_classificar.destroy()
            tooltip_classificar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão classificar
    button_classificar.bind("<Leave>", hide_classificar_tooltip)


    #==========Botão Atualizar Pastas
    
    button_atualizar = tk.Button(button_frame, text="Atualizar Pastas", command=atualizaçãoPastas_ABN, width=12)
    button_atualizar.grid(row=0, column=3, padx=5, pady=5)
    
    def show_atualizar_tooltip(event):
        global tooltip_atualizar
        tooltip_atualizar = tk.Label(tab1_frame, text="Faz uma cópia do arquivo ABN.xlsx para as pastas 03_Sistema e 04_Público", background="white", relief="solid", borderwidth=1)
        tooltip_atualizar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão atualizar
    button_atualizar.bind("<Enter>", show_atualizar_tooltip)

    # Função para ocultar a legenda do botão atualizar
    def hide_atualizar_tooltip(event):
        global tooltip_atualizar
        if tooltip_atualizar is not None:
            tooltip_atualizar.destroy()
            tooltip_atualizar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão atualizar
    button_atualizar.bind("<Leave>", hide_atualizar_tooltip)

    #==========Botão Manual
    
    button_manual = tk.Button(tab1_frame, text="Manual", command=open_manual, width=12)
    button_manual.grid(row=4, column=0, padx=130, pady=5, sticky="w")
    
    def show_manual_tooltip(event):
        global tooltip_manual
        tooltip_manual = tk.Label(tab1_frame, text="Abre o Manual de Automações.docx", background="white", relief="solid", borderwidth=1)
        tooltip_manual.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão manual
    button_manual.bind("<Enter>", show_manual_tooltip)

    # Função para ocultar a legenda do botão manual
    def hide_manual_tooltip(event):
        global tooltip_manual
        if tooltip_manual is not None:
            tooltip_manual.destroy()
            tooltip_manual = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão manual
    button_manual.bind("<Leave>", hide_manual_tooltip)

    #==========Label==========
    
    label_text = ttk.Label(button_frame, text="")
    label_text.grid(row=0, column=5, columnspan=4, padx=20, pady=5, sticky="e")

    
    #==========Frames==========
    
    button_frame.columnconfigure(4, weight=1)  # Configurar coluna para expansão horizontal

    text_frame = ttk.Frame(tab1_frame)
    text_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")


    #==========Botão Restaurar ABN
    
    # Cria um popup para executar ou não a funçao restaurar
    def abrir_popup():
        resultado = messagebox.askquestion("Confirmação", "Tem certeza de que deseja restaurar o arquivo ABN.xlsx \ncom o backup mais recente da pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup ABN' ?")
        if resultado == 'yes':
            restaurar()
        else:
            messagebox.showinfo("Informação", "Restauração cancelada")

    restaurar_button = tk.Button(tab1_frame, text="Restaurar Base", command=abrir_popup, width=12)
    restaurar_button.grid(row=4, column=0, padx=25, pady=5, sticky="w")

    # Função para exibir a legenda do botão restaurar
    def show_restaurar_tooltip(event):
        global tooltip_restaurar
        tooltip_restaurar = tk.Label(tab1_frame, text="Restaura o arquivo ABN.xlsx das pastas 02_Tratadas, 03_Sistema e 04_Público com o backup mais recente da pasta 05_BackUp/@Backup ABN", background="white", relief="solid", borderwidth=1)
        tooltip_restaurar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão restaurar
    restaurar_button.bind("<Enter>", show_restaurar_tooltip)

    # Função para ocultar a legenda do botão restaurar
    def hide_restaurar_tooltip(event):
        global tooltip_restaurar
        if tooltip_restaurar is not None:
            tooltip_restaurar.destroy()
            tooltip_restaurar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão restaurar
    restaurar_button.bind("<Leave>", hide_restaurar_tooltip)

#============================================================================================================================================================================================================================

    #==========Área de exibição de prints
    
    # Criar o PanedWindow
    paned_window = ttk.PanedWindow(text_frame, orient='vertical')
    paned_window.pack(fill="both", expand=True, pady=(15, 0))

    # Área 1
    text_widget1 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget1, weight=1)  # Definir peso igual para a área 1

    # Área 2
    text_widget2 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget2, weight=1)
    
    # Configurar as barras de rolagem
    text_scrollbar1 = ttk.Scrollbar(text_widget1, command=text_widget1.yview)
    text_widget1.config(yscrollcommand=text_scrollbar1.set)
    text_scrollbar1.pack(side="right", fill="y")

    text_scrollbar2 = ttk.Scrollbar(text_widget2, command=text_widget2.yview)
    text_widget2.config(yscrollcommand=text_scrollbar2.set)
    text_scrollbar2.pack(side="right", fill="y")

    total_iterations = 10 # Número do range das buscas 10 iterações equivalem a 40 buscas.

    # Criar a barra de progresso com o estilo personalizado
    progress_bar = ttk.Progressbar(left_frame, mode='indeterminate', maximum=total_iterations)
    progress_bar.grid(row=8, column=0, padx=10, pady=5, sticky="ew")

    # Definir o tamanho da barra de progresso (comprimento em pixels)
    progress_bar.config(length=10)  # Substitua 300 pelo valor desejado

    def print_to_text(text):
        text_widget1.insert("end", text + "\n")
        text_widget1.see("end")
        text_widget1.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget1
    
    def print_to_text2(text):
        text_widget2.insert("end", text + "\n")
        text_widget2.see("end")
        text_widget2.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget2

    # Configurar o redimensionamento do tab1_frame
    tab1_frame.columnconfigure(0, weight=1)  # Expandir coluna 0
    tab1_frame.rowconfigure(3, weight=1)  # Expandir linha 3

#============================================================================================================================================================================================================================

#==========Menu Blog

# Variável de controle para verificar se a função está em andamento
buscarBlog = False 
 
def tab2():
  
    def statusBlog():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file_path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx'
        sheet_name = 'Blog_producao'
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        empty_lines = df[df['Setor_Tema'].isnull() & (df['Ano_publicacao2'] >= 2023)].shape[0]
        last_class_date = df[df['Setor_Tema'].notnull()]['Data'].max().strftime('%d/%m/%Y')
        text = f" {current_time}- Status classificação Blog.xlsx\n Postagens para classificar: {empty_lines} || Última classificação: {last_class_date} "
        label_text.configure(text=text)
        label_text.configure(relief="ridge") 
        print_to_text(f"{current_time} \u2714Status Blog.xlsx, Atualizado...")

#============================================================================================================================================================================================================================       
    def buscarBlog():
        
        # Desabilitar o botão
        button_buscar.configure(state="disabled")
        button_classificar.configure(state="disabled")
        button_atualizar.configure(state="disabled")
        restaurar_button.configure(state="disabled")
        
        global buscarBlog
        # Verificar se a função já está em andamento
        if buscarBlog:
            return
        # Definir a variável de controle como global
        buscarBlog = True
              
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        #Backup
        try:
            from datetime import datetime
            data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx'
            destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Blog/'
            nome_arquivo = f'Blog_{data_hora_atual}.xlsx'
            destino_completo = destino + nome_arquivo
            shutil.copy(origem, destino_completo)
            print_to_text(f"{data_hora_atual} \u2714Backup do arquivo Blog.xlsx concluído...")
        except Exception as e:
            print_to_text(f"{data_hora_atual} !Ocorreu um erro ao fazer o backup do arquivo Blog.xlsx: {str(e)}")
            
        #scrap               
        try:
            
            #### inicio busca Blog
            url = "https://agenciadenoticias.bndes.gov.br/blogdodesenvolvimento/index.html?reloaded&page=1"
            response = requests.get(url.format(1))
            html = response.text
            soup = BeautifulSoup(html, "html.parser")
            pagination_div = soup.find('div', {'class': 'pag'})
            pag_total_span = pagination_div.find('span', {'class': 'pag-total'})
            pag_total = int(pag_total_span.text.replace('de', '').replace('páginas', '').strip())
            max_pages = pag_total
            url = 'https://agenciadenoticias.bndes.gov.br/blogdodesenvolvimento/index.html?reloaded&page={}'
            page_num = 1
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36 Edge/B08C3901'
            }
            df_blog = pd.DataFrame(columns=['Link', 'Título', 'Data', 'Hora'])

            print_to_text(f"{current_time} \u2714Buscando Link, Título, Data, Hora...")
            contador = 1
            for page_num in range(1, 8):  # Range de páginas alcançe 58 postagens mais recentes
                current_url = url.format(page_num)
                response = requests.get(current_url, headers=headers)

                if response.status_code != 200:
                    break

                
                soup = BeautifulSoup(response.content, 'html.parser')
                blog_main = soup.find('section', {'class': 'blog-main'})
                h2_tags = blog_main.find_all("h2") 
                data_tags = blog_main.select('div.post-data strong')
                leia_mais_tags = soup.find_all('p', class_='leia-mais')
                hora = soup.find_all('div', {'class': 'post-data'})
                
                for h2, data, tag, time in zip(h2_tags, data_tags, leia_mais_tags, hora):
                    href = 'https://agenciadenoticias.bndes.gov.br' + tag.a['href']
                    df_temp = pd.DataFrame({'Link': href, 'Título': h2.text, 'Data': data.text, 'Hora': time.text}, index=[0])
                    df_blog = pd.concat([df_blog, df_temp], ignore_index=True)

                    print_to_text2(f"{contador}. Link: {href}")
                    print_to_text2(f"{contador}. Título: {h2.text}")
                    print_to_text2(f"{contador}. Data: {data.text}")
                    print_to_text2(f"{contador}. Hora: {time.text}")
                    contador += 1
              
            df_blog['Data'] = pd.to_datetime(df_blog['Data'])    
            df_blog['Hora'] = df_blog['Hora'].str.replace('\n', '').str.replace('\t', '')
            df_blog['Hora'] = df_blog['Hora'].str.split('|').str[0]
            df_blog['Hora'] = pd.to_datetime(df_blog['Hora'])
            df_blog['Hora'] = pd.to_datetime(df_blog['Hora']).dt.strftime('%H:%M')
            df_blog = df_blog.sort_values(by='Data', ascending=False)
            
            ########## Filtros Data ##########
            df_blog['Data'] = pd.to_datetime(df_blog['Data'])
            df_dataFiltros = df_blog
            df_dataFiltros['Dia_publicacao'] = df_blog['Data'].dt.day
            df_dataFiltros['Mes_publicacao2'] = df_blog['Data'].dt.month
            df_dataFiltros['Mes_publicacao2'] = df_blog['Data'].dt.strftime('%b')
            df_dataFiltros['Mes_publicacao2'] = df_blog['Mes_publicacao2']
            df_dataFiltros = df_dataFiltros.replace("Dec", "dez")
            df_dataFiltros = df_dataFiltros.replace("Nov", "nov")
            df_dataFiltros = df_dataFiltros.replace("Oct", "out")
            df_dataFiltros = df_dataFiltros.replace("Sep", "set")
            df_dataFiltros = df_dataFiltros.replace("Aug", "ago")
            df_dataFiltros = df_dataFiltros.replace("Jul", "jul")
            df_dataFiltros = df_dataFiltros.replace("Jun", "jun")
            df_dataFiltros = df_dataFiltros.replace("May", "mai")
            df_dataFiltros = df_dataFiltros.replace("Apr", "abr")
            df_dataFiltros = df_dataFiltros.replace("Mar", "mar")
            df_dataFiltros = df_dataFiltros.replace("Feb", "fev")
            df_dataFiltros = df_dataFiltros.replace("Jan", "jan")
            df_dataFiltros['Ano_publicacao2'] = df_blog['Data'].dt.year
            semana = {0: 'Seg',1: 'Ter',2: 'Qua',3: 'Qui',4: 'Sex',5: 'Sáb',6: 'Dom'}
            df_dataFiltros['Dia_da_semana_publicacao'] = df_blog['Data'].dt.weekday.map(semana)
            df_dataFiltros = df_dataFiltros.drop(['Data'], axis=1)
            df_dataFiltros = df_dataFiltros.drop(['Link'], axis=1)
            df_dataFiltros = df_dataFiltros.drop(['Título'], axis=1)
            df_dataFiltros = df_dataFiltros.drop(['Hora'], axis=1)
            
            ########## Criação de colunas de classificação para merge com arquivo base Bases/02_Tratadas/Blog.xlsx ##########
            colunas = ['Tipo Blog']
            df_tipo = pd.DataFrame(columns=colunas)
            df_tipo['Tipo Blog'] = df_tipo['Tipo Blog'].astype(object)

            colunas = ['Setor_Tema']
            df_setor = pd.DataFrame(columns=colunas)
            df_setor['Setor_Tema'] = df_setor['Setor_Tema'].astype(object)

            colunas = ['Iniciativa_Produto']
            df_iniciativa = pd.DataFrame(columns=colunas)
            df_iniciativa['Iniciativa_Produto'] = df_iniciativa['Iniciativa_Produto'].astype(object)

            colunas = ['Estratégia']
            df_estrategia = pd.DataFrame(columns=colunas)
            df_estrategia['Estratégia'] = df_estrategia['Estratégia'].astype(object)

            #Organização
            df_blog = df_blog.drop(['Dia_publicacao'], axis=1)
            df_blog = df_blog.drop(['Mes_publicacao2'], axis=1)

            #Merge
            df_blog = pd.concat([df_blog, df_tipo, df_setor, df_iniciativa, df_estrategia, df_dataFiltros ], axis=1)

            #Filtragem
            df_blog = df_blog[df_blog['Data'] >= '2023-06-01']
            df_blog.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Blog(Feed).xlsx", index=False)
            
            ########## Merge com a base ##########
            
            try:
                print_to_text(f"{current_time} \u2714Atualizando arquivo...")    
                df_base = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx", sheet_name='Blog_producao')
                df_base_desempenho = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx', sheet_name='Blog_desempenho')
                df_feed = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Blog(Feed).xlsx")

                #Filtrar apenas as linhas do dataframe "df_feed" que não estão duplicadas no dataframe "df_base"
                df_feed = df_feed[~df_feed["Título"].isin(df_base["Título"])]

                #Concatenar os dataframes "df_base" e "df_feed"
                df_base = pd.concat([df_base, df_feed], ignore_index=True)
                df_base = df_base.sort_values(by='Data', ascending=False)
         
                writer = pd.ExcelWriter('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx', engine='xlsxwriter')
                df_base.to_excel(writer, sheet_name='Blog_producao', index=False)  # Adiciona o parâmetro index=False
                df_base_desempenho.to_excel(writer, sheet_name='Blog_desempenho', index=False)  # Adiciona o parâmetro index=False
                writer.close()
                print_to_text(f"{current_time} \u2714Arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx Atualizado...")

            except Exception as e:
                print_to_text("Ocorreu um erro durante a junção do arquivo Blog.xlsx com o arquivo Blog(Feed):" + str(e))
                                     
        except Exception as e:
            print_to_text(f"{current_time} !Erro o arquivo ..02_Tratadas/Blog.xlsx está aberto por você ou por outro usuário,feche o arquivo para continuar, ou" + str(e))
                 
        # Atualizar a variável de controle
        buscarBlog = False
        # Reabilitar o botão após a conclusão da função
        button_buscar.configure(state="normal")
        button_classificar.configure(state="normal")
        button_atualizar.configure(state="normal")
        restaurar_button.configure(state="normal")

#============================================================================================================================================================================================================================
    def start_buscar_thread():
        # Criar uma nova thread para executar a função buscarBlog
        buscar_thread = threading.Thread(target=buscarBlog)
        buscar_thread.start()

#============================================================================================================================================================================================================================                                      
    def classificarBlog():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            print_to_text(f"{current_time} ->Abrindo o arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx...\n{current_time} ->Aguarde a formatação da planilha em Tabela...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx'
            os.startfile(path)
            time.sleep(15)
            # Alternar para a planilha "Blog_producao"
            pyautogui.hotkey('ctrl', 'pgup')
            keyboard.press_and_release('ctrl+alt+t')
            time.sleep(2)
            keyboard.press_and_release('enter')
            time.sleep(1)
            keyboard.press_and_release('ctrl+b')
            print_to_text(f"{current_time} \u2714Formatação concluída...")
        except Exception as e:
            print_to_text(f"{current_time} !Falha [Classificar] Ocorreu o erro na função classificarBlog(): {e}")

#============================================================================================================================================================================================================================                
    def atualizaçãoPastas_Blog():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            print_to_text(f"{current_time} \u2714Atualizando pastas...")
            src_file = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Blog.xlsx'
            dst_folder = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'
            dst_folder2 = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
            shutil.copy(src_file, dst_folder)
            shutil.copy(src_file, dst_folder2)
            #print_to_text(f"{current_time} ->Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Blog/'Atualizada...")
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'Atualizada...")
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'Atualizada...")
        except Exception as e:
            print_to_text(f"{current_time}- !Falha [Atualizar pastas] Ocorreu um erro ao copiar o arquivo para as pastas 03_Sistema e 04_Público: {e}")

#============================================================================================================================================================================================================================
    # Função para abrir o Manual de Automações.docx
    def open_manual():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
        try:
            print_to_text(f"{current_time} ->Abrindo C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx'
            os.startfile(path)
        except Exception as e:
            print_to_text(f"{current_time} !Erro, ocorreu um erro ao abrir o arquivo Manual automação.docx : {e}")
#============================================================================================================================================================================================================================
    def restaurar():
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Diretórios de origem e destino
        backup_dir = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Blog/'
        dest_dirs = [
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
        ]

        # Obter o arquivo mais recente na pasta de backup
        backup_path = Path(backup_dir)
        files = backup_path.iterdir()
        latest_file = max(files, key=lambda f: f.stat().st_ctime) if files else None

        if latest_file:
            source_file = str(latest_file)

            # Copiar o arquivo para os diretórios de destino
            for dest_dir in dest_dirs:
                dest_file = Path(dest_dir) / 'Blog.xlsx'
                shutil.copy(source_file, str(dest_file))

            print_to_text(f"{current_time} \u2714 Blog.xlsx restaurado")
        else:
            print_to_text(f"{current_time} Nenhum arquivo de backup encontrado")

#============================================================================================================================================================================================================================

    #==========Interface menu Blog
          
    tab2_frame = ttk.Frame(tab_control)
    tab_control.add(tab2_frame, text='Blog         ')     
    tab2_frame.configure(height=100)

    def open_url(event):
        webbrowser.open("https://agenciadenoticias.bndes.gov.br/blogdodesenvolvimento")

    label = ttk.Label(tab2_frame, text='Atualização base de dados Blog, informações obtidas através do site: https://agenciadenoticias.bndes.gov.br/blogdodesenvolvimento\n\n', cursor='hand2')
    label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    label.bind("<Button-1>", open_url)

    button_frame = ttk.Frame(tab2_frame)
    button_frame.grid(row=2, column=0, padx=20, sticky="w")
  
#============================================================================================================================================================================================================================ 
  
    #==========Botão Status Blog
    
    button_status = tk.Button(button_frame, text="Status", command=statusBlog, width=12)
    button_status.grid(row=0, column=0, padx=5, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab2_frame, text="Verifica o número de postagens sem classificação e a data da última classificação", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_status.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_status.bind("<Leave>", hide_status_tooltip)
  
#============================================================================================================================================================================================================================ 
  
    #==========Botão Buscar Blog
    
    button_buscar = tk.Button(button_frame, text="Buscar", command=buscarBlog, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)
 
    # Botão Máscara Buscar Blog para não haver duplicação de command
    button_buscar = tk.Button(button_frame, text="Buscar", command=start_buscar_thread, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)

    # Função para exibir a legenda do botão buscar
    def show_buscar_tooltip(event):
        global tooltip_buscar
        tooltip_buscar = tk.Label(tab2_frame, text="Busca as notícias Blog do Desenvolvimento e atualiza a base Blog.xlsx", background="white", relief="solid", borderwidth=1)
        tooltip_buscar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão buscar
    button_buscar.bind("<Enter>", show_buscar_tooltip)

    # Função para ocultar a legenda do botão buscar
    def hide_buscar_tooltip(event):
        global tooltip_buscar
        if tooltip_buscar is not None:
            tooltip_buscar.destroy()
            tooltip_buscar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão buscar
    button_buscar.bind("<Leave>", hide_buscar_tooltip)


#============================================================================================================================================================================================================================

    #==========Botão Classificar Blog
    
    button_classificar = tk.Button(button_frame, text="Classificar", command=classificarBlog, width=12)
    button_classificar.grid(row=0, column=2, padx=5, pady=5)

    # Função para exibir a legenda do botão classificar
    def show_classificar_tooltip(event):
        global tooltip_classificar
        tooltip_classificar = tk.Label(tab2_frame, text="Executa o Excel abrindo o arquivo Blog.xlsx da pasta 02_Tratadas e faz a formatação para planilha", background="white", relief="solid", borderwidth=1)
        tooltip_classificar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão classificar
    button_classificar.bind("<Enter>", show_classificar_tooltip)

    # Função para ocultar a legenda do botão classificar
    def hide_classificar_tooltip(event):
        global tooltip_classificar
        if tooltip_classificar is not None:
            tooltip_classificar.destroy()
            tooltip_classificar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão classificar
    button_classificar.bind("<Leave>", hide_classificar_tooltip)

#============================================================================================================================================================================================================================

    #==========Botão Classificar Blog
    
    button_atualizar = tk.Button(button_frame, text="Atualizar Pastas", command=atualizaçãoPastas_Blog, width=12)
    button_atualizar.grid(row=0, column=3, padx=5, pady=5)
    
    def show_atualizar_tooltip(event):
        global tooltip_atualizar
        tooltip_atualizar = tk.Label(tab2_frame, text="Faz uma cópia do arquivo Blog.xlsx para as pastas 03_Sistema e 04_Público", background="white", relief="solid", borderwidth=1)
        tooltip_atualizar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão atualizar
    button_atualizar.bind("<Enter>", show_atualizar_tooltip)

    # Função para ocultar a legenda do botão atualizar
    def hide_atualizar_tooltip(event):
        global tooltip_atualizar
        if tooltip_atualizar is not None:
            tooltip_atualizar.destroy()
            tooltip_atualizar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão atualizar
    button_atualizar.bind("<Leave>", hide_atualizar_tooltip)

#============================================================================================================================================================================================================================    
    
    #==========Botão Manual
    
    button_manual = tk.Button(tab2_frame, text="Manual", command=open_manual, width=12)
    button_manual.grid(row=4, column=0, padx=130, pady=5, sticky="w")
    
    def show_manual_tooltip(event):
        global tooltip_manual
        tooltip_manual = tk.Label(tab2_frame, text="Abre o Manual de Automações", background="white", relief="solid", borderwidth=1)
        tooltip_manual.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão manual
    button_manual.bind("<Enter>", show_manual_tooltip)

    # Função para ocultar a legenda do botão manual
    def hide_manual_tooltip(event):
        global tooltip_manual
        if tooltip_manual is not None:
            tooltip_manual.destroy()
            tooltip_manual = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão manual
    button_manual.bind("<Leave>", hide_manual_tooltip)
    
#============================================================================================================================================================================================================================

    #Label
    
    label_text = ttk.Label(button_frame, text="")
    label_text.grid(row=0, column=5, columnspan=4, padx=20, pady=5, sticky="e")

#============================================================================================================================================================================================================================
    #Frames
    
    button_frame.columnconfigure(4, weight=1)  # Configurar coluna para expansão horizontal

    text_frame = ttk.Frame(tab2_frame)
    text_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

#============================================================================================================================================================================================================================

    #==========Botão Restaurar
    
    def abrir_popup():
        resultado = messagebox.askquestion("Confirmação", "Tem certeza de que deseja restaurar o arquivo Blog.xlsx \ncom o backup mais recente da pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Blog' ?")
        if resultado == 'yes':
            restaurar()
        else:
            messagebox.showinfo("Informação", "Restauração cancelada")

    restaurar_button = tk.Button(tab2_frame, text="Restaurar Base", command=abrir_popup, width=12)
    restaurar_button.grid(row=4, column=0, padx=25, pady=5, sticky="w")

    # Função para exibir a legenda do botão restaurar
    def show_restaurar_tooltip(event):
        global tooltip_restaurar
        tooltip_restaurar = tk.Label(tab2_frame, text="Restaura o arquivo Blog.xlsx das pastas 02_Tratadas, 03_Sistema e 04_Público com o backup mais recente da pasta 05_BackUp/@Backup Blog", background="white", relief="solid", borderwidth=1)
        tooltip_restaurar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão restaurar
    restaurar_button.bind("<Enter>", show_restaurar_tooltip)

    # Função para ocultar a legenda do botão restaurar
    def hide_restaurar_tooltip(event):
        global tooltip_restaurar
        if tooltip_restaurar is not None:
            tooltip_restaurar.destroy()
            tooltip_restaurar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão restaurar
    restaurar_button.bind("<Leave>", hide_restaurar_tooltip)

##############################################################################################################################################################################################################################

    # Área de exibição de prints
    
    # Criar o PanedWindow
    paned_window = ttk.PanedWindow(text_frame, orient='vertical')
    paned_window.pack(fill="both", expand=True, pady=(15, 0))

    # Área 1
    text_widget1 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget1, weight=1)  # Definir peso igual para a área 1

    # Área 2
    text_widget2 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget2, weight=1)
    
    # Configurar as barras de rolagem
    text_scrollbar1 = ttk.Scrollbar(text_widget1, command=text_widget1.yview)
    text_widget1.config(yscrollcommand=text_scrollbar1.set)
    text_scrollbar1.pack(side="right", fill="y")

    text_scrollbar2 = ttk.Scrollbar(text_widget2, command=text_widget2.yview)
    text_widget2.config(yscrollcommand=text_scrollbar2.set)
    text_scrollbar2.pack(side="right", fill="y")

    total_iterations = 10 # Número do range das buscas 10 iterações equivalem a 40 buscas.

    # Criar a barra de progresso com o estilo personalizado
    progress_bar = ttk.Progressbar(left_frame, mode='indeterminate', maximum=total_iterations)
    progress_bar.grid(row=8, column=0, padx=10, pady=5, sticky="ew")

    # Definir o tamanho da barra de progresso (comprimento em pixels)
    progress_bar.config(length=10)  # Substitua 300 pelo valor desejado

    def print_to_text(text):
        text_widget1.insert("end", text + "\n")
        text_widget1.see("end")
        text_widget1.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget1
    
    def print_to_text2(text):
        text_widget2.insert("end", text + "\n")
        text_widget2.see("end")
        text_widget2.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget2

#============================================================================================================================================================================================================================

    # Configurar o redimensionamento do tab2_frame
    tab2_frame.columnconfigure(0, weight=1)  # Expandir coluna 0
    tab2_frame.rowconfigure(3, weight=1)  # Expandir linha 3

    #Blog - last line  

#============================================================================================================================================================================================================================

#==========Menu Release

# Variável de controle para verificar se a função está em andamento
buscarReleases = False  
def tab3():

    def statusReleases():

        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        file_path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx'
        sheet_name = 'Releases'
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        empty_lines = df[df['Setor_Tema'].isnull() & (df['Ano'] >= 2023)].shape[0]
        last_class_date = df[df['Setor_Tema'].notnull()]['Data_publicacao'].max().strftime('%d/%m/%Y')
        text = f" {current_time}- Status classificação Releases.xlsx\n Postagens para classificar: {empty_lines} || Última classificação: {last_class_date} "
        label_text.configure(text=text)
        label_text.configure(relief="ridge") 
        print_to_text(f"{current_time} \u2714Status Releases.xlsx, Atualizado...")

#============================================================================================================================================================================================================================       
    def buscarReleases():
        
        # Desabilitar o botão
        button_buscar.configure(state="disabled")
        button_classificar.configure(state="disabled")
        button_atualizar.configure(state="disabled")
        restaurar_button.configure(state="disabled")
        
        global buscarReleases
        # Verificar se a função já está em andamento
        if buscarReleases:
            return
        # Definir a variável de controle como global
        buscarReleases = True
                             
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        #Backup
        try:
            from datetime import datetime
            data_hora_atual = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
            origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx'
            destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Releases/'
            nome_arquivo = f'Releases_{data_hora_atual}.xlsx'
            destino_completo = destino + nome_arquivo
            shutil.copy(origem, destino_completo)
            print_to_text(f"{data_hora_atual} \u2714Backup do arquivo Releases.xlsx concluído...")
        except Exception as e:
            print_to_text(f"{data_hora_atual} !Ocorreu um erro ao fazer o backup do arquivo Releases.xlsx: {str(e)}")
            
        #scrap               
          
        #### inicio busca Releases
        try:
            
            api = requests.get('https://www.bndes.gov.br/WCMUtil/api/noticia/')
            api = json.loads(api.content)
           
            url_list = []
            titulo_list = []
            data_list = []

            for item in api['data']:
                url_list.append(item['url'])
                titulo_list.append(item['titulo'])
                data_list.append(item['data'])

            api = pd.DataFrame(api['data'])
            api = api.drop('tagImagem', axis=1)
            api = api.drop('resumo', axis=1)
            api = api.rename(columns={'titulo': 'Título', 'url': 'Link', 'data': 'Data'})
            api = api.loc[:, ['Título', 'Link', 'Data']]
            
            #Output
            from datetime import datetime

            # Combina as três listas em uma lista de tuplas
            combined_list = list(zip(data_list, url_list, titulo_list))

            # Ordena a lista combinada com base na data
            sorted_list = sorted(combined_list, key=lambda x: int(x[0]))

            print_to_text2("Data, URL, e Título List:")
            for indice, (data, url, titulo) in enumerate(sorted_list, start=1):
                timestamps_s = int(data) // 1000
                date = datetime.fromtimestamp(timestamps_s)
                
                # Verifica se o ano da data é igual ou posterior a 2021
                if date.year >= 2021:
                    date_string = date.strftime('%Y-%m-%d %H:%M:%S')
                    url = url.replace('bndes_institucional/home/imprensa/noticias/conteudo/', 'https://www.bndes.gov.br/wps/portal/site/home/imprensa/noticias/conteudo/')
                    print_to_text2(f"{indice}. Data: {date_string}\n   URL: {url}\n   Título: {titulo}")

              
            #Complemento da url da API
            api['Link'] = api['Link'].replace('bndes_institucional/home/imprensa/noticias/conteudo/', 'https://www.bndes.gov.br/wps/portal/site/home/imprensa/noticias/conteudo/', regex=True)
            
            # Data_publicacao

            data = pd.DataFrame(api['Data'])

            # extrai coluna 'Data_publicacao' do DataFrame
            timestamps_ms = data['Data']

            # converte strings para números inteiros e em seguida, para segundos
            timestamps_s = [int(ts) // 1000 for ts in timestamps_ms]

            # cria lista de objetos datetime a partir dos timestamps em segundos
            dates = [datetime.fromtimestamp(ts) for ts in timestamps_s]

            # formata as datas como strings
            date_strings = [dt.strftime('%Y-%m-%d %H:%M:%S') for dt in dates]

            # substitui coluna original pelos valores formatados
            data['Data'] = date_strings
                       
            # converte para datatime para compatibilidade com .xlsx
            colunas = ['Data_publicacao']
            df_data = pd.DataFrame(columns=colunas)
            df_data['Data_publicacao'] = data['Data']

            # converter a coluna "data" em objeto de data
            df_data['Data_publicacao'] = pd.to_datetime(df_data['Data_publicacao'])

            # adicionar colunas para dia, mes, ano e dia da semana

            df_data['Dia'] = df_data['Data_publicacao'].dt.day

            df_data['Mês'] = df_data['Data_publicacao'].dt.month
            df_data['Mês'] = df_data['Data_publicacao'].dt.strftime('%b')
            df_data['Mês'] = df_data['Mês']
            df_data = df_data.replace("Dec", "Dez")
            df_data = df_data.replace("Oct", "Out")
            df_data = df_data.replace("Sep", "Set")
            df_data = df_data.replace("Aug", "Ago")
            df_data = df_data.replace("May", "Mai")
            df_data = df_data.replace("Apr", "Abr")
            df_data = df_data.replace("Feb", "Fev")

            df_data['Ano'] = df_data['Data_publicacao'].dt.year

            dias_da_semana = {
                0: 'Seg',
                1: 'Ter',
                2: 'Qua',
                3: 'Qui',
                4: 'Sex',
                5: 'Sáb',
                6: 'Dom'
            }
            df_data['Dia_da_semana'] = df_data['Data_publicacao'].dt.weekday.map(dias_da_semana)

            # tirando coluna Data_publicacao do df_data
            df_data = df_data.loc[:, ['Dia', 'Mês', 'Ano', 'Dia_da_semana']]

            # renomeando coluna do data
            data = data.rename(columns={'Data': 'Data_publicacao'})

            #Montagem Feed.xlsx

            #Criação de colunas vazias para merge com arquivo base Bases/02_Tratadas/Releases.xlsx

            colunas = ['Repercussao']
            df_repercussao = pd.DataFrame(columns=colunas)
            df_repercussao['Repercussao'] = df_repercussao['Repercussao'].astype(object)

            colunas = ['Estrategia']
            df_estrategia = pd.DataFrame(columns=colunas)
            df_estrategia['Estrategia'] = df_estrategia['Estrategia'].astype(object)

            colunas = ['Histograma']
            df_histograma = pd.DataFrame(columns=colunas)
            df_histograma['Histograma'] = df_histograma['Histograma'].astype(object)

            colunas = ['Tipo']
            df_tipo = pd.DataFrame(columns=colunas)
            df_tipo['Tipo'] = df_tipo['Tipo'].astype(object)

            colunas = ['Setor_Tema']
            df_setor = pd.DataFrame(columns=colunas)
            df_setor['Setor_Tema'] = df_setor['Setor_Tema'].astype(object)

            colunas = ['Iniciativa_Produto']
            df_iniciativa = pd.DataFrame(columns=colunas)
            df_iniciativa['Iniciativa_Produto'] = df_iniciativa['Iniciativa_Produto'].astype(object)

            colunas = ['Subgrupo (Fabiano)']
            df_subgrupo = pd.DataFrame(columns=colunas)
            df_subgrupo['Subgrupo (Fabiano)'] = df_subgrupo['Subgrupo (Fabiano)'].astype(object)

            colunas = ['Dia_da_semana']
            df_dia_da_semana = pd.DataFrame(columns=colunas)
            df_dia_da_semana['Dia_da_semana'] = df_dia_da_semana['Dia_da_semana'].astype(object)

            colunas = ['CONTEUDO']
            df_conteudo = pd.DataFrame(columns=colunas)
            df_conteudo['CONTEUDO'] = df_conteudo['CONTEUDO'].astype(object)

            colunas = ['Grupo Conteúdo']
            df_grupo_conteúdo = pd.DataFrame(columns=colunas)
            df_grupo_conteúdo['Grupo Conteúdo'] = df_grupo_conteúdo['Grupo Conteúdo'].astype(object)

            #Merge
            df_feed = pd.concat([api, data, df_repercussao, df_histograma, df_tipo, df_setor, df_iniciativa, df_estrategia, df_subgrupo, df_data, df_conteudo, df_grupo_conteúdo], axis=1)

            #Filtragem 2023
            df_feed = df_feed[df_feed['Data_publicacao'] >= '2023-01-01']
            
            #Ordenação
            df_feed = df_feed.sort_values(by='Data_publicacao', ascending=False)

            # Substituir os espaços duplos na coluna 'Título'
            df_feed['Título'] = df_feed['Título'].str.replace('  ', ' ')

            #Salvando em .xlsx
            df_feed.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Releases(Feed).xlsx", sheet_name="Releases", index=False)
            
            #Formatações
            df_base = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx", sheet_name='Releases')
            df_feed = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Releases(Feed).xlsx")

            # Converte a coluna "Data_publicacao" para datetime
            df_feed['Data_publicacao'] = pd.to_datetime(df_feed['Data_publicacao'])

            # Adequação de formatos df_base
            df_base['Repercussao'] = df_base['Repercussao'].astype(str)
            df_base['Estrategia'] = df_base['Estrategia'].astype(str)
            df_base['Histograma'] = df_base['Histograma'].astype(str)
            df_base['Tipo'] = df_base['Tipo'].astype(str)
            df_base['Setor_Tema'] = df_base['Setor_Tema'].astype(str)
            df_base['Iniciativa_Produto'] = df_base['Iniciativa_Produto'].astype(str)
            df_base['Subgrupo (Fabiano)'] = df_base['Subgrupo (Fabiano)'].astype(str)
            df_base['Dia'] = df_base['Dia'].astype(int)
            df_base['Mês'] = df_base['Mês'].astype(str)
            df_base['Ano'] = df_base['Ano'].astype(int)
            df_base['Dia_da_semana'] = df_base['Dia_da_semana'].astype(str)
            df_base['CONTEUDO'] = df_base['CONTEUDO'].astype(str)
            df_base['Grupo Conteúdo'] = df_base['Grupo Conteúdo'].astype(str)

            # Adequação de formatos df_feed
            df_feed['Repercussao'] = df_feed['Repercussao'].astype(str)
            df_feed['Estrategia'] = df_feed['Estrategia'].astype(str)
            df_feed['Histograma'] = df_feed['Histograma'].astype(str)
            df_feed['Tipo'] = df_feed['Tipo'].astype(str)
            df_feed['Setor_Tema'] = df_feed['Setor_Tema'].astype(str)
            df_feed['Iniciativa_Produto'] = df_feed['Iniciativa_Produto'].astype(str)
            df_feed['Subgrupo (Fabiano)'] = df_feed['Subgrupo (Fabiano)'].astype(str)
            df_feed['Dia'] = df_feed['Dia'].astype(int)
            df_feed['Mês'] = df_feed['Mês'].astype(str)
            df_feed['Ano'] = df_feed['Ano'].astype(int)
            df_feed['Dia_da_semana'] = df_feed['Dia_da_semana'].astype(str)
            df_feed['CONTEUDO'] = df_feed['CONTEUDO'].astype(str)
            df_feed['Grupo Conteúdo'] = df_feed['Grupo Conteúdo'].astype(str)
            
            df_base.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx", sheet_name="Releases", index=False)
            df_feed.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Releases(Feed).xlsx", sheet_name="Releases", index=False)
            
            ########## Merge com a base ##########
            
            try:
                print_to_text(f"{current_time} \u2714Atualizando arquivo...")
                df_base = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx", sheet_name='Releases')
                df_feed = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Releases(Feed).xlsx")
                
                # Converter as colunas de data para datetime e remover a informação de timezone
                df_base['Data_publicacao'] = pd.to_datetime(df_base['Data_publicacao']).dt.tz_localize(None)
                df_feed['Data_publicacao'] = pd.to_datetime(df_feed['Data_publicacao']).dt.tz_localize(None)

                # Ordenação por mais recente
                df_base = df_base.sort_values(by='Data_publicacao', ascending=True)
                df_feed = df_feed.sort_values(by='Data_publicacao', ascending=False)

                df_base.set_index("Título", inplace=True)
                df_feed.set_index("Título", inplace=True)

                df_base.update(df_feed)
                novas_linhas = df_feed[~df_feed.index.isin(df_base.index)]
                df_base.update(df_feed)
                df_base = pd.concat([df_base, novas_linhas])
                df_base.reset_index(inplace=True)
                df_base = df_base.drop('Data', axis=1)
                df_base = df_base.sort_values(by='Data_publicacao', ascending=False)

                writer = pd.ExcelWriter('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx', engine='xlsxwriter')
                df_base.to_excel(writer, sheet_name='Releases', index=False)
                writer.close()
                print_to_text(f"{current_time} \u2714Arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx Atualizado...")

            except Exception as e:
                print_to_text("Ocorreu um erro durante a junção do arquivo Releases.xlsx com o arquivo Releases(Feed):" + str(e))
                                     
        except Exception as e:
            print_to_text(f"{current_time} !Erro o arquivo ..02_Tratadas/Releases.xlsx está aberto por você ou por outro usuário,feche o arquivo para continuar, ou" + str(e))
                
             
        # Atualizar a variável de controle
        buscarReleases = False
        # Reabilitar o botão após a conclusão da função
        button_buscar.configure(state="normal")
        button_classificar.configure(state="normal")
        button_atualizar.configure(state="normal")
        restaurar_button.configure(state="normal")

#============================================================================================================================================================================================================================
    def start_buscar_thread():
        # Criar uma nova thread para executar a função buscarReleases
        buscar_thread = threading.Thread(target=buscarReleases)
        buscar_thread.start()

#============================================================================================================================================================================================================================                                      
    def classificarReleases():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            print_to_text(f"{current_time} ->Abrindo o arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx...\n{current_time} ->Aguarde a formatação da planilha em Tabela...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx'
            os.startfile(path)
            time.sleep(15)
            keyboard.press_and_release('ctrl+alt+t')
            time.sleep(2)
            keyboard.press_and_release('enter')
            time.sleep(1)
            keyboard.press_and_release('ctrl+b')
            print_to_text(f"{current_time} \u2714Formatação concluída...")
        except Exception as e:
            print_to_text(f"{current_time} !Falha [Classificar] Ocorreu o erro na função classificarReleases(): {e}")

#============================================================================================================================================================================================================================                
    def atualizaçãoPastas_Releases():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            print_to_text(f"{current_time} \u2714Atualizando pastas...")
            src_file = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/Releases.xlsx'
            dst_folder = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'
            dst_folder2 = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
            shutil.copy(src_file, dst_folder)
            shutil.copy(src_file, dst_folder2)
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'Atualizada...")
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'Atualizada...")
        except Exception as e:
            print_to_text(f"{current_time}- !Falha [Atualizar pastas] Ocorreu um erro ao copiar o arquivo para as pastas 03_Sistema e 04_Público: {e}")

#============================================================================================================================================================================================================================
    def restaurar():
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

        # Diretórios de origem e destino
        backup_dir = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Releases/'
        dest_dirs = [
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas/',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/',
            'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
        ]

        # Obter o arquivo mais recente na pasta de backup
        backup_path = Path(backup_dir)
        files = backup_path.iterdir()
        latest_file = max(files, key=lambda f: f.stat().st_ctime) if files else None

        if latest_file:
            source_file = str(latest_file)

            # Copiar o arquivo para os diretórios de destino
            for dest_dir in dest_dirs:
                dest_file = Path(dest_dir) / 'Releases.xlsx'
                shutil.copy(source_file, str(dest_file))

            print_to_text(f"{current_time} \u2714 Releases.xlsx restaurado")
        else:
            print_to_text(f"{current_time} Nenhum arquivo de backup encontrado")

#============================================================================================================================================================================================================================

    # Função para abrir o Manual de Automações.docx
    def open_manual():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                
        try:
            print_to_text(f"{current_time} ->Abrindo C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx...")  
            path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Manual automação.docx'
            os.startfile(path)
        except Exception as e:
            print_to_text(f"{current_time} !Erro, ocorreu um erro ao abrir o arquivo Manual automação.docx : {e}")
            
#============================================================================================================================================================================================================================

    #==========Interface menu Releases
          
    tab3_frame = ttk.Frame(tab_control)
    tab_control.add(tab3_frame, text='Releases         ')     
    tab3_frame.configure(height=100)

    def open_url(event):
        webbrowser.open("https://www.bndes.gov.br/wps/portal/site/home/imprensa/noticias")

    label = ttk.Label(tab3_frame, text='Atualização base de dados Releases, informações obtidas através do site: https://www.bndes.gov.br/wps/portal/site/home/imprensa/noticias\n\n', cursor='hand2')
    label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
    label.bind("<Button-1>", open_url)

    button_frame = ttk.Frame(tab3_frame)
    button_frame.grid(row=2, column=0, padx=20, sticky="w")
    
  
#============================================================================================================================================================================================================================ 
  
    #==========Botão Status Releases
    
    button_status = tk.Button(button_frame, text="Status", command=statusReleases, width=12)
    button_status.grid(row=0, column=0, padx=5, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab3_frame, text="Verifica o número de postagens sem classificação e a data da última classificação", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_status.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_status.bind("<Leave>", hide_status_tooltip)
  
#============================================================================================================================================================================================================================ 
  
    #==========Botão Buscar Releases
    
    button_buscar = tk.Button(button_frame, text="Buscar", command=buscarReleases, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)
 
    # Botão Máscara Buscar Releases para não haver duplicação de command
    button_buscar = tk.Button(button_frame, text="Buscar", command=start_buscar_thread, width=12)
    button_buscar.grid(row=0, column=1, padx=5, pady=5)

    # Função para exibir a legenda do botão buscar
    def show_buscar_tooltip(event):
        global tooltip_buscar
        tooltip_buscar = tk.Label(tab3_frame, text="Busca as notícias da imprensa e atualiza a base Releases.xlsx", background="white", relief="solid", borderwidth=1)
        tooltip_buscar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão buscar
    button_buscar.bind("<Enter>", show_buscar_tooltip)

    # Função para ocultar a legenda do botão buscar
    def hide_buscar_tooltip(event):
        global tooltip_buscar
        if tooltip_buscar is not None:
            tooltip_buscar.destroy()
            tooltip_buscar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão buscar
    button_buscar.bind("<Leave>", hide_buscar_tooltip)


#============================================================================================================================================================================================================================

    #==========Botão Classificar Releases
    
    button_classificar = tk.Button(button_frame, text="Classificar", command=classificarReleases, width=12)
    button_classificar.grid(row=0, column=2, padx=5, pady=5)

    # Função para exibir a legenda do botão classificar
    def show_classificar_tooltip(event):
        global tooltip_classificar
        tooltip_classificar = tk.Label(tab3_frame, text="Executa o Excel abrindo o arquivo Releases.xlsx da pasta 02_Tratadas e faz a formatação para planilha", background="white", relief="solid", borderwidth=1)
        tooltip_classificar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão classificar
    button_classificar.bind("<Enter>", show_classificar_tooltip)

    # Função para ocultar a legenda do botão classificar
    def hide_classificar_tooltip(event):
        global tooltip_classificar
        if tooltip_classificar is not None:
            tooltip_classificar.destroy()
            tooltip_classificar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão classificar
    button_classificar.bind("<Leave>", hide_classificar_tooltip)

#============================================================================================================================================================================================================================

    #==========Botão Atualizar Pastas
    
    button_atualizar = tk.Button(button_frame, text="Atualizar Pastas", command=atualizaçãoPastas_Releases, width=12)
    button_atualizar.grid(row=0, column=3, padx=5, pady=5)
    
    def show_atualizar_tooltip(event):
        global tooltip_atualizar
        tooltip_atualizar = tk.Label(tab3_frame, text="Faz uma cópia do arquivo Releases.xlsx para as pastas 03_Sistema e 04_Público", background="white", relief="solid", borderwidth=1)
        tooltip_atualizar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão atualizar
    button_atualizar.bind("<Enter>", show_atualizar_tooltip)

    # Função para ocultar a legenda do botão atualizar
    def hide_atualizar_tooltip(event):
        global tooltip_atualizar
        if tooltip_atualizar is not None:
            tooltip_atualizar.destroy()
            tooltip_atualizar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão atualizar
    button_atualizar.bind("<Leave>", hide_atualizar_tooltip)
 
 #============================================================================================================================================================================================================================    
    
    #==========Botão Manual
    
    button_manual = tk.Button(tab3_frame, text="Manual", command=open_manual, width=12)
    button_manual.grid(row=4, column=0, padx=130, pady=5, sticky="w")
    
    def show_manual_tooltip(event):
        global tooltip_manual
        tooltip_manual = tk.Label(tab3_frame, text="Abre o Manual de Automações", background="white", relief="solid", borderwidth=1)
        tooltip_manual.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão manual
    button_manual.bind("<Enter>", show_manual_tooltip)

    # Função para ocultar a legenda do botão manual
    def hide_manual_tooltip(event):
        global tooltip_manual
        if tooltip_manual is not None:
            tooltip_manual.destroy()
            tooltip_manual = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão manual
    button_manual.bind("<Leave>", hide_manual_tooltip)  
     
#============================================================================================================================================================================================================================

    #==========Label
    
    label_text = ttk.Label(button_frame, text="")
    label_text.grid(row=0, column=5, columnspan=4, padx=20, pady=5, sticky="e")

#============================================================================================================================================================================================================================
    #==========Frames
    
    button_frame.columnconfigure(4, weight=1)  # Configurar coluna para expansão horizontal

    text_frame = ttk.Frame(tab3_frame)
    text_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

#============================================================================================================================================================================================================================

    #==========Botão Restaurar
    
    def abrir_popup():
        resultado = messagebox.askquestion("Confirmação", "Tem certeza de que deseja restaurar o arquivo Releases.xlsx \ncom o backup mais recente da pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Releases' ?")
        if resultado == 'yes':
            restaurar()
        else:
            messagebox.showinfo("Informação", "Restauração cancelada")

    restaurar_button = tk.Button(tab3_frame, text="Restaurar Base", command=abrir_popup, width=12)
    restaurar_button.grid(row=4, column=0, padx=25, pady=5, sticky="w")

    # Função para exibir a legenda do botão restaurar
    def show_restaurar_tooltip(event):
        global tooltip_restaurar
        tooltip_restaurar = tk.Label(tab3_frame, text="Restaura o arquivo Releases.xlsx das pastas 02_Tratadas, 03_Sistema e 04_Público com o backup mais recente da pasta 05_BackUp/@Backup Releases", background="white", relief="solid", borderwidth=1)
        tooltip_restaurar.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão restaurar
    restaurar_button.bind("<Enter>", show_restaurar_tooltip)

    # Função para ocultar a legenda do botão restaurar
    def hide_restaurar_tooltip(event):
        global tooltip_restaurar
        if tooltip_restaurar is not None:
            tooltip_restaurar.destroy()
            tooltip_restaurar = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão restaurar
    restaurar_button.bind("<Leave>", hide_restaurar_tooltip)

##############################################################################################################################################################################################################################

    #==========Área de exibição de prints
    
    # Criar o PanedWindow
    paned_window = ttk.PanedWindow(text_frame, orient='vertical')
    paned_window.pack(fill="both", expand=True, pady=(15, 0))

    # Área 1
    text_widget1 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget1, weight=1)  # Definir peso igual para a área 1

    # Área 2
    text_widget2 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget2, weight=1)
    
    # Configurar as barras de rolagem
    text_scrollbar1 = ttk.Scrollbar(text_widget1, command=text_widget1.yview)
    text_widget1.config(yscrollcommand=text_scrollbar1.set)
    text_scrollbar1.pack(side="right", fill="y")

    text_scrollbar2 = ttk.Scrollbar(text_widget2, command=text_widget2.yview)
    text_widget2.config(yscrollcommand=text_scrollbar2.set)
    text_scrollbar2.pack(side="right", fill="y")

    total_iterations = 10 # Número do range das buscas 10 iterações equivalem a 40 buscas.

    # Criar a barra de progresso com o estilo personalizado
    progress_bar = ttk.Progressbar(left_frame, mode='indeterminate', maximum=total_iterations)
    progress_bar.grid(row=8, column=0, padx=10, pady=5, sticky="ew")

    # Definir o tamanho da barra de progresso (comprimento em pixels)
    progress_bar.config(length=10)  # Substitua 300 pelo valor desejado

    def print_to_text(text):
        text_widget1.insert("end", text + "\n")
        text_widget1.see("end")
        text_widget1.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget1
    
    def print_to_text2(text):
        text_widget2.insert("end", text + "\n")
        text_widget2.see("end")
        text_widget2.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget2

#============================================================================================================================================================================================================================

    # Configurar o redimensionamento do tab3_frame
    tab3_frame.columnconfigure(0, weight=1)  # Expandir coluna 0
    tab3_frame.rowconfigure(3, weight=1)  # Expandir linha 3

    #Releases - last line  

#============================================================================================================================================================================================================================

#==========Menu Seguidores

def tab5():
    
    import datetime
    current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")

#============================================================================================================================================================================================================================

    def scrape_Facebook():
        
        options = webdriver.ChromeOptions()
        #options.add_argument("--headless")
        options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
        options.add_argument("--log-level=3")
        driver = webdriver.Chrome(options=options)
        
        def login():
            
            driver.get("https://www.facebook.com/")
            email_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="email"]')))
            email_field.send_keys("e-mail")
            time.sleep(2)
            password_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="pass"]')))
            password_field.send_keys("senha")
            time.sleep(2)
            pyautogui.press('tab', presses=2)
            pyautogui.press('enter')
            time.sleep(5)
        
        login()
        
        def scrape():

            lista_urls_Facebook = ['https://www.facebook.com/bancodobrasil'
                                ,'https://www.facebook.com/bancodonordeste'
                                ,'https://www.facebook.com/bndes.imprensa'
                                ,'https://www.facebook.com/bradesco'
                                ,'https://www.facebook.com/BTGPactual'
                                ,'https://www.facebook.com/caixa'
                                ,'https://www.facebook.com/CitiBrasil'
                                ,'https://www.facebook.com/jpmorganchase'
                                ,'https://www.facebook.com/morganstanley'
                                ,'https://www.facebook.com/nubank'
                                ,'https://www.facebook.com/santanderbrasil']

        
            Facebook = 'Facebook'

            df1 = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df1.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook1.xlsx', index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores Facebook...")
        

            for url in lista_urls_Facebook:
                driver.get(url)
          
                while True:
                    try:
                        soup = BeautifulSoup(driver.page_source, 'html.parser')
                        a_tags = soup.find_all('a', {'class': 'x1i10hfl xjbqb8w x1ejq31n xd10rxx x1sy0etr x17r0tee x972fbf xcfux6l x1qhh985 xm0m39n x9f619 x1ypdohk xt0psk2 xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x16tdsg8 xggy1nq x1a2a7pz xt0b8zv x1hl2dhg xi81zsa x1s688f'})
                        
                        if len(a_tags) < 2:
                            raise ValueError('second <a> element not found')
                        
                        seguidoresFacebook = a_tags[1]
                        
                        print_to_text2(f'Link: {url}')
                        print_to_text2(f'Seguidores: {seguidoresFacebook.text}\n')
                        break
                    except (TypeError, KeyError, ValueError):
                        continue

                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m")

                df1 = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook1.xlsx')
                df1_new = pd.DataFrame({'Grupo': url
                                                .replace('https://www.facebook.com/', '')
                                                .replace('bancodobrasil', 'Bancos Comerciais')
                                                .replace('bancodonordeste', 'Bancos Desenvolvimento')
                                                .replace('bndes.imprensa', 'Bancos Desenvolvimento')
                                                .replace('bradesco', 'Bancos Comerciais')
                                                .replace('BTGPactual', 'Banco Boutique')
                                                .replace('caixa', 'Bancos Comerciais')
                                                .replace('CitiBrasil', 'Banco Boutique')
                                                .replace('jpmorganchase', 'Banco Boutique')
                                                .replace('morganstanley', 'Banco Boutique')
                                                .replace('nubank', 'Bancos Comerciais')
                                                .replace('santanderbrasil', 'Bancos Comerciais')
                                                , 'Instituicao': url
                                                .replace('https://www.facebook.com/', '')
                                                .replace('bancodobrasil', 'BB')
                                                .replace('bancodonordeste', 'BNB')
                                                .replace('bndes.imprensa', 'Bancos Desenvolvimento')
                                                .replace('bradesco', 'BRADESCO')
                                                .replace('BTGPactual', 'BTG')
                                                .replace('caixa', 'Caixa')
                                                .replace('CitiBrasil', 'City Bank')
                                                .replace('jpmorganchase', 'JPMorgan')
                                                .replace('morganstanley', 'Morgan Stanley')
                                                .replace('nubank', 'NUBANK')
                                                .replace('santanderbrasil', 'Santander')
                                                , 'Plataforma': Facebook
                                                , 'data': data_atual
                                                , 'Seguidores': seguidoresFacebook
                                                , 'Ano': ano
                                                , 'Mês': mes
                                                , 'Link': url}, index=[0])
                df1 = pd.concat([df1, df1_new])
                # Filtragens
                df1['Seguidores'] = df1['Seguidores'].apply(lambda x: x.replace('Seguidores:', ''))
                df1['Seguidores'] = df1['Seguidores'].apply(lambda x: x.replace('mil', '000')) 
                df1['Seguidores'] = df1['Seguidores'].apply(lambda x: x.replace(' ', ''))     
                df1.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook1.xlsx', index=False)

           
            lista_urls_Facebook2 = ['https://www.facebook.com/AFDOfficiel'
                                ,'https://www.facebook.com/bancocentraldobrasil/'
                                ,'https://www.facebook.com/imf'
                                ,'https://www.facebook.com/itau'
                                ,'https://www.facebook.com/minfazenda'
                                ,'https://www.facebook.com/petrobras'
                                ,'https://www.facebook.com/sebrae'
                                ,'https://www.facebook.com/worldbank']
            Facebook = 'Facebook'

            df2 = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df2.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook2.xlsx', index=False)

            for url in lista_urls_Facebook2:
                driver.get(url)

                while True:
                    try:
                        soup = BeautifulSoup(driver.page_source, 'html.parser')
                        seguidoresFacebook2 = soup.find('a', {'class': 'x1i10hfl xjbqb8w x1ejq31n xd10rxx x1sy0etr x17r0tee x972fbf xcfux6l x1qhh985 xm0m39n x9f619 x1ypdohk xt0psk2 xe8uvvx xdj266r x11i5rnm xat24cr x1mh8g0r xexx8yu x4uap5 x18d9i69 xkhd6sd x16tdsg8 xggy1nq x1a2a7pz xt0b8zv x1hl2dhg xi81zsa x1s688f'})
                        if seguidoresFacebook2 is None:
                            raise ValueError('element not found')
                        print_to_text2(f'Link: {url}')
                        print_to_text2(f'Seguidores: {seguidoresFacebook2}\n')
                        break
                    except (TypeError, KeyError, ValueError):
                        continue               

                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m")

                df2 = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook2.xlsx')
                df2_new = pd.DataFrame({'Grupo': url
                                                .replace('https://www.facebook.com/', '')
                                                .replace('AFDOfficiel', 'Bancos Desenvolvimento')
                                                .replace('bancocentraldobrasil/', 'Outros')                                        
                                                .replace('imf', 'Bancos Desenvolvimento')
                                                .replace('itau', 'Bancos Comerciais')                                           
                                                .replace('minfazenda', 'Outros')
                                                .replace('petrobras', 'Outros')
                                                .replace('sebrae', 'Outros')
                                                .replace('worldbank', 'Bancos Desenvolvimento')
                                                , 'Instituicao': url
                                                .replace('https://www.facebook.com/', '')
                                                .replace('AFDOfficiel', 'ADF')                                          
                                                .replace('bancocentraldobrasil/', 'BACEN')
                                                .replace('imf', 'IMF')                              
                                                .replace('itau', 'ITAU')
                                                .replace('minfazenda', 'Ministério Fazenda')                                           
                                                .replace('petrobras', 'Petrobras')
                                                .replace('sebrae', 'SEBRAE')
                                                .replace('worldbank', 'WorldBank')
                                                , 'Plataforma': Facebook
                                                , 'data': data_atual
                                                , 'Seguidores': seguidoresFacebook2
                                                , 'Ano': ano
                                                , 'Mês': mes
                                                , 'Link': url}, index=[0])
                df2 = pd.concat([df2, df2_new])
                df2['Seguidores'] = df2['Seguidores'].apply(lambda x: x.replace('curtidas', ''))
                df2['Seguidores'] = df2['Seguidores'].apply(lambda x: x.replace('Seguidores:', ''))
                df2['Seguidores'] = df2['Seguidores'].apply(lambda x: x.replace('mil', '000'))                           
                df2.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook2.xlsx', index=False)
                df2['Seguidores'] = df2['Seguidores'].apply(lambda x: x.replace(' ', ''))
        
            #==========Merge lista 1 e 2
            df1 = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook1.xlsx')
            df2 = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook2.xlsx')
            df_facebook = pd.concat([df1, df2], ignore_index=True)
            df_facebook = df_facebook.sort_values(by=['Instituicao'])
            df_facebook.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook.xlsx', index=False)
            os.remove('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook1.xlsx')
            os.remove('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Facebook2.xlsx')
            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Facebook')

            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)
                data_e_hora_atuais = datetime.datetime.now()
                data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
                arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Facebook/Seguidores_Facebook.xlsx'
                arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Facebook/Seguidores_Facebook_' + data_e_hora_formatada + '.xlsx'
                os.rename(arquivo_origem, arquivo_destino)
                local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
                
                for arquivo in os.listdir(local):
                    os.remove(os.path.join(local, arquivo))    
            print_to_text(f"{current_time} \u2714Busca seguidores Facebook concluída...")
            
        scrape()

#============================================================================================================================================================================================================================
    
    def scrape_Flickr():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        lista_urls_Flickr = ['https://www.flickr.com/photos/145226591@N08/','https://www.flickr.com/photos/imfphoto','https://www.flickr.com/photos/itaucultural/'
                            ,'https://www.flickr.com/photos/ministeriodaeconomia/','https://www.flickr.com/photos/worldbank']
        Flickr = 'Flickr'

        df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês','Link'])
        df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Flickr.xlsx', index=False)

        print_to_text(f"{current_time} ->Iniciando busca Seguidores Flickr...")
        for url in lista_urls_Flickr:
            headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
            req = requests.get(url, headers=headers)
            soup = BeautifulSoup(req.content, 'html.parser')
            seguidoresFlickr = soup.find('p', {'class': 'followers truncate no-shrink'}).text[0:12]

            import datetime
            agora = datetime.datetime.now()
            data_atual = agora.strftime("%d/%m/%Y")
            ano = agora.strftime("%Y")
            mes = agora.strftime("%m")

            df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Flickr.xlsx')
            df = pd.concat([df,pd.DataFrame({'Grupo':url
                                            .replace('https://www.flickr.com/photos/145226591@N08/','BNDES')
                                            .replace('https://www.flickr.com/photos/imfphoto','Bancos Desenvolvimento')
                                            .replace('https://www.flickr.com/photos/itaucultural/','Bancos Comerciais')
                                            .replace('https://www.flickr.com/photos/ministeriodaeconomia/','Outros')     
                                            .replace('https://www.flickr.com/photos/worldbank','Bancos Desenvolvimento')                                        
                                            ,'Plataforma':Flickr
                                            ,'Instituicao':url.replace('https://www.flickr.com/photos/145226591@N08/','BNDES').replace('https://www.flickr.com/photos/imfphoto','IMF')
                                            .replace('https://www.flickr.com/photos/itaucultural/','ITAU').replace('https://www.flickr.com/photos/ministeriodaeconomia/','Ministério Fazenda')
                                            .replace('https://www.flickr.com/photos/worldbank','WorldBank')
                                            ,'data': data_atual
                                            ,'Seguidores':seguidoresFlickr.replace('Followers•','').replace('Follower','').replace(' Followe','')
                                            ,'Ano':ano
                                            ,'Mês':mes
                                            ,'Link':url}, index=[0])], ignore_index=True)
            print_to_text2(str(url))
            print_to_text2(str(seguidoresFlickr.replace('Followers•','').replace('Follower','').replace(' Followe','')))

            def convert_count(seguidoresFlickr):
                
                if isinstance(seguidoresFlickr, str):
                    if 'K' in seguidoresFlickr:
                        seguidoresFlickr = seguidoresFlickr.replace('K', '')
                        seguidoresFlickr = float(seguidoresFlickr) * 1000
                    elif 'mi' in seguidoresFlickr:
                        seguidoresFlickr = seguidoresFlickr.replace('M', '')
                        seguidoresFlickr = float(seguidoresFlickr) * 1000000
                return seguidoresFlickr

            df['Seguidores'] = df['Seguidores'].apply(convert_count)
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Flickr.xlsx', index=False)

        origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
        destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Flickr')
        for arquivo in origem.iterdir():
            shutil.copy(arquivo, destino)

        data_e_hora_atuais = datetime.datetime.now()
        data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
        arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Flickr/Seguidores_Flickr.xlsx'
        arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Flickr/Seguidores_Flickr_' + data_e_hora_formatada + '.xlsx'
        os.rename(arquivo_origem, arquivo_destino)

        local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
        for arquivo in os.listdir(local):
            os.remove(os.path.join(local, arquivo))

        print_to_text(f"{current_time} \u2714Busca seguidores Flickr concluída...")

#============================================================================================================================================================================================================================

    def scrape_instagram():
            
            options = webdriver.ChromeOptions()
            options.add_argument("--headless")
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
            options.add_argument("--log-level=3")
            driver = webdriver.Chrome(options=options)

            def login():

                driver.get("https://www.instagram.com/")
                email_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="loginForm"]/div/div[1]/div/label/input')))
                email_field.send_keys("e-mail")
                password_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="loginForm"]/div/div[2]/div/label/input')))
                password_field.send_keys("senha")
                submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="loginForm"]/div/div[3]/button')))
                submit_button.click()
                time.sleep(5)
             
            login()

            def scrape():
                lista_urls_Instagram = ['https://www.instagram.com/AFD_France/'
                                        ,'https://www.instagram.com/bancocentraldobrasil/'
                                        ,'https://www.instagram.com/bancodobrasil/'
                                        ,'https://www.instagram.com/bancodonordeste/'
                                        ,'https://www.instagram.com/bndesgovbr/'
                                        ,'https://www.instagram.com/bradesco/'
                                        ,'https://www.instagram.com/btgpactual/'
                                        ,'https://www.instagram.com/caixa/'
                                        ,'https://www.instagram.com/citibrasil/?hl=pt-br'
                                        ,'https://www.instagram.com/the_imf/?fbclid=IwAR3AWRfpk0AQ0dm-b2jDaRimHUoxAsWcvb1jTmgNLuWZOOwpCGUVeRPfSjg'
                                        ,'https://www.instagram.com/itau/'
                                        ,'https://www.instagram.com/jpmorgan/'
                                        ,'https://www.instagram.com/kfw.stories/'
                                        ,'https://www.instagram.com/min.fazenda/'
                                        ,'https://www.instagram.com/morgan.stanley/'
                                        ,'https://www.instagram.com/nubank/?hl=pt-br'
                                        ,'https://www.instagram.com/petrobras/'
                                        ,'https://www.instagram.com/santanderbrasil/'
                                        ,'https://www.instagram.com/sebrae/'
                                        ,'https://www.instagram.com/worldbank/']                                 
                                        
                Instagram = 'Instagram'

                df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Instagram.xlsx', index=False)

                print_to_text(f"{current_time} ->Iniciando busca Seguidores Instagram...")
                
                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y %H:%M:%S")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m") 

                for url in lista_urls_Instagram:
                    driver.get(url)
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    try:
                        seguidoresInstagram = soup.find('meta', attrs={'name': 'description'})['content']
                        seguidoresInstagram = seguidoresInstagram.split()[0]
                    except:
                        seguidoresInstagram = 0
                    print_to_text2(f"{current_time} {url}")
                    print_to_text2(f"{seguidoresInstagram}")
                                      
                    df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Instagram.xlsx')
                    df = pd.concat([df,pd.DataFrame({'Grupo':url
                                                .replace('https://www.instagram.com/', '')                       
                                                .replace('AFD_France/','Bancos Desenvolvimento')
                                                .replace('bancocentraldobrasil/','Outros')
                                                .replace('bancodobrasil/','Bancos Comerciais')
                                                .replace('bancodonordeste/','Bancos Desenvolvimento')
                                                .replace('bndesgovbr/','BNDES')
                                                .replace('bradesco/','Bancos Comerciais')
                                                .replace('btgpactual/','Banco Boutique')
                                                .replace('caixa/','Bancos Comerciais')
                                                .replace('citibrasil/?hl=pt-br','Banco Boutique')
                                                .replace('the_imf/?fbclid=IwAR3AWRfpk0AQ0dm-b2jDaRimHUoxAsWcvb1jTmgNLuWZOOwpCGUVeRPfSjg','Bancos Desenvolvimento')
                                                .replace('itau/','Bancos Comerciais')
                                                .replace('jpmorgan/','Banco Boutique')
                                                .replace('kfw.stories/','Bancos Desenvolvimento')
                                                .replace('min.fazenda/','Outros')
                                                .replace('morgan.stanley/','Banco Boutique')
                                                .replace('nubank/?hl=pt-br','Bancos Comerciais')
                                                .replace('petrobras/','Outros')                                          
                                                .replace('santanderbrasil/','Bancos Comerciais')                                          
                                                .replace('sebrae/','Outros')
                                                .replace('worldbank/','Bancos Desenvolvimento')
                                        ,'Plataforma':Instagram
                                        ,'Instituicao':url
                                                .replace('https://www.instagram.com/', '')                       
                                                .replace('AFD_France/','AFD')
                                                .replace('bancocentraldobrasil/','BACEN')
                                                .replace('bancodobrasil/','BB')
                                                .replace('bancodonordeste/','BNB')
                                                .replace('bndesgovbr/','BNDES')
                                                .replace('bradesco/','BRADESCO')
                                                .replace('btgpactual/','BTG')
                                                .replace('caixa/','Caixa')
                                                .replace('citibrasil/?hl=pt-br','City Bank')
                                                .replace('the_imf/?fbclid=IwAR3AWRfpk0AQ0dm-b2jDaRimHUoxAsWcvb1jTmgNLuWZOOwpCGUVeRPfSjg','IMF')
                                                .replace('itau/','ITAU')
                                                .replace('jpmorgan/','JPMorgan')
                                                .replace('kfw.stories/','KfW')
                                                .replace('min.fazenda/','Ministério Fazenda')
                                                .replace('morgan.stanley/','Morgan Stanley')
                                                .replace('nubank/?hl=pt-br','NUBANK')
                                                .replace('petrobras/','Petrobras')                                          
                                                .replace('santanderbrasil/','Santander')                                          
                                                .replace('sebrae/','SEBRAE')
                                                .replace('worldbank/','Worldbank')
                                        ,'data': data_atual
                                        ,'Seguidores':seguidoresInstagram
                                        ,'Ano':ano
                                        ,'Mês':mes
                                        ,'Link':url}, index=[0])], ignore_index=True)

                    df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Instagram.xlsx', index=False)
                    
                origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
                destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Instagram')
                for arquivo in origem.iterdir():
                    shutil.copy(arquivo, destino)

                data_e_hora_atuais = datetime.datetime.now()
                data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
                arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Instagram/Seguidores_Instagram.xlsx'
                arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Instagram/Seguidores_Instagram_' + data_e_hora_formatada + '.xlsx'
                os.rename(arquivo_origem, arquivo_destino)

                local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
                for arquivo in os.listdir(local):
                    os.remove(os.path.join(local, arquivo))
                
                print_to_text(f"{current_time} \u2714Busca seguidores Instagram concluída...")   
                
            scrape()             

#============================================================================================================================================================================================================================

    def scrape_linkedin():

                print_to_text(f"{current_time} ->Iniciando busca Seguidores Linkedin...")
                
                options = webdriver.ChromeOptions()
                options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36")
                driver = webdriver.Chrome(options=options)

                def login():

                    driver.get("https://www.linkedin.com/home")
                    email_field = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="session_key"]')))
                    email_field.send_keys("e-mail")
                    password_field = WebDriverWait(driver, 15).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="session_password"]')))
                    password_field.send_keys("senha")
                    submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="main-content"]/section[1]/div/div/form/div[2]/button')))
                    submit_button.click()
                    #time.sleep(15)
                        
                login()

                def scrape():
                    
                    lista_urls_Linkedin = ['https://www.linkedin.com/company/agence-francaise-de-developpement'
                                            ,'https://www.linkedin.com/company/banco-central-do-brasil/?originalSubdomain=br'
                                            ,'https://www.linkedin.com/company/bancodobrasil/'
                                            ,'https://www.linkedin.com/company/bndes/mycompany/'
                                            ,'https://www.linkedin.com/company/bradesco/'
                                            ,'https://www.linkedin.com/company/btgpactual/'
                                            ,'https://www.linkedin.com/company/citi-brasil/'
                                            ,'https://www.linkedin.com/company/international-monetary-fund/'
                                            ,'https://www.linkedin.com/company/itau/'
                                            ,'https://www.linkedin.com/company/jpmorgan/'
                                            ,'https://www.linkedin.com/company/kfw/?originalSubdomain=de'
                                            ,'https://www.linkedin.com/company/ministeriodafazenda/'
                                            ,'https://www.linkedin.com/company/morgan-stanley/'
                                            ,'https://www.linkedin.com/company/nubank/'
                                            ,'https://www.linkedin.com/company/petrobras/'
                                            ,'https://www.linkedin.com/company/grupo-santander-brasil/'
                                            ,'https://www.linkedin.com/company/sebrae/'
                                            ,'https://www.linkedin.com/company/the-world-bank/']                              
                                            
                    Linkedin = 'Linkedin'

                    df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
                    df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Linkedin.xlsx', index=False)

                    for url in lista_urls_Linkedin:
                        driver.get(url)
                        max_attempts = 3
                        attempts = 0
                        while attempts < max_attempts:
                            
                            try:
                                seguidoresLinkedin = driver.find_element(By.XPATH, '/html/body/div[4]/div[3]/div/div[2]/div/div[2]/main/div[1]').text
                                seguidoresLinkedin = re.search(r'(\d[\d\.]*) seguidores', seguidoresLinkedin).group(1)
                                seguidoresLinkedin = seguidoresLinkedin.replace('.', '') + ' seguidores'
                                print_to_text2(f"{current_time} {url}")
                                print_to_text2(f"{seguidoresLinkedin}")
                                break
                            except:
                                attempts += 1
                                if attempts == max_attempts:
                                    attempts = 0
                                    time.sleep(1)      
                                                                
                        import datetime
                        agora = datetime.datetime.now()
                        data_atual = agora.strftime("%d/%m/%Y")
                        ano = agora.strftime("%Y")
                        mes = agora.strftime("%m") 
                        
                        df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Linkedin.xlsx')
                        df = pd.concat([df,pd.DataFrame({'Grupo':url
                                                    .replace('https://www.linkedin.com/company/', '')                       
                                                    .replace('banco-central-do-brasil/?originalSubdomain=br','Outros')
                                                    .replace('bancodobrasil/','Bancos Comerciais')
                                                    .replace('bndes/mycompany/','BNDES')
                                                    .replace('bradesco/','Bancos Comerciais')
                                                    .replace('btgpactual/','Banco Boutique')
                                                    .replace('citi-brasil/','Banco Boutique')
                                                    .replace('international-monetary-fund/','Bancos Desenvolvimento')
                                                    .replace('itau/','Bancos Comerciais')
                                                    .replace('jpmorgan/','Banco Boutique')
                                                    .replace('kfw/?originalSubdomain=de','Bancos Desenvolvimento')
                                                    .replace('ministeriodafazenda/','Outros')
                                                    .replace('morgan-stanley/','Banco Boutique')
                                                    .replace('nubank/','Bancos Comerciais')
                                                    .replace('petrobras/','Outros')                                          
                                                    .replace('grupo-santander-brasil/','Bancos Comerciais')                                          
                                                    .replace('sebrae/','Outros')
                                                    .replace('the-world-bank/','Bancos Desenvolvimento')
                                            ,'Plataforma':Linkedin
                                            ,'Instituicao':url
                                                    .replace('https://www.linkedin.com/company/', '')                       
                                                    .replace('banco-central-do-brasil/?originalSubdomain=br','BACEN')
                                                    .replace('bancodobrasil/','BB')
                                                    .replace('bndes/mycompany/','BNDES')
                                                    .replace('bradesco/','BRADESCO')
                                                    .replace('btgpactual/','BTG')
                                                    .replace('citi-brasil/','City Bank')
                                                    .replace('international-monetary-fund/','IMF')
                                                    .replace('itau/','ITAU')
                                                    .replace('jpmorgan/','JPMorgan')
                                                    .replace('kfw/?originalSubdomain=de','KfW')
                                                    .replace('ministeriodafazenda/','Ministério Fazenda')
                                                    .replace('morgan-stanley/','Morgan Stanley')
                                                    .replace('nubank/','NUBANK')
                                                    .replace('petrobras/','Petrobras')                                          
                                                    .replace('grupo-santander-brasil/','Santander')                                          
                                                    .replace('sebrae/','SEBRAE')
                                                    .replace('the-world-bank/','Worldbank')
                                            ,'data': data_atual
                                            ,'Seguidores':seguidoresLinkedin
                                            ,'Ano':ano
                                            ,'Mês':mes
                                            ,'Link':url}, index=[0])], ignore_index=True)     
                            
                        df['Seguidores'] = df['Seguidores'].str.extract(r'(\d[\d\.]*) seguidores', expand=False)
                        df['Seguidores'] = df['Seguidores'].str.replace('[^0-9]+', '', regex=True) + ' seguidores'
                        df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Linkedin.xlsx', index=False)

                    df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Linkedin.xlsx')
                    df['Seguidores'] = df['Seguidores'].str.replace(' seguidores', '')
                    df['Grupo'] = df['Grupo'].str.replace('agence-francaise-de-developpement', 'Bancos Desenvolvimento')
                    df['Instituicao'] = df['Instituicao'].str.replace('-', '')
                    df['Instituicao'] = df['Instituicao'].str.replace('agencefrancaisededeveloppement', 'AFD')
                    df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Linkedin.xlsx', index=False)

                
                    origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
                    destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Linkedin')
                    for arquivo in origem.iterdir():
                        shutil.copy(arquivo, destino)

                    data_e_hora_atuais = datetime.datetime.now()
                    data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
                    arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Linkedin/Seguidores_Linkedin.xlsx'
                    arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Linkedin/Seguidores_Linkedin_' + data_e_hora_formatada + '.xlsx'
                    os.rename(arquivo_origem, arquivo_destino)

                    local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
                    for arquivo in os.listdir(local):
                        os.remove(os.path.join(local, arquivo))
                    
                    print_to_text(f"{current_time} \u2714Busca seguidores Linkedin concluída...")   
                    
                scrape()                  

#============================================================================================================================================================================================================================
  
    def scrape_Spotify():
        
            options = webdriver.ChromeOptions()
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/90.0.4430.212 Safari/537.36")
            options.add_argument("--log-level=3")
            driver = webdriver.Chrome(options=options)
            driver.set_window_size(300, 50)
            lista_urls_Spotify = ['https://open.spotify.com/user/bancodobrasil']                                                      
            Spotify = 'Spotify'
            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Spotify.xlsx', index=False)
            print_to_text(f"{current_time} ->Iniciando busca Seguidores Spotify...")
            
            for url in lista_urls_Spotify:
                driver.get(url)
                max_attempts = 10
                attempts = 0
                while attempts < max_attempts:                   
                    try:
                        seguidoresSpotify = driver.find_element(By.CSS_SELECTOR, '#main > div > div.ZQftYELq0aOsg6tPbVbV > div.jEMA2gVoLgPQqAFrPhFw.lPapCDz3v_LipgXwe8gi > div.main-view-container > div.os-host.os-host-foreign.os-theme-spotify.os-host-resize-disabled.os-host-scrollbar-horizontal-hidden.main-view-container__scroll-node.os-host-transition.os-host-overflow.os-host-overflow-y > div.os-padding > div > div > div.main-view-container__scroll-node-child > main > section > div > div.contentSpacing.NXiYChVp4Oydfxd7rT5r.RMDSGDMFrx8eXHpFphqG > div.RP2rRchy4i8TIp1CTmb7 > div > span:nth-child(2) > a').text
                        print_to_text2(f"{current_time} {url}")
                        print_to_text2(f"{seguidoresSpotify}")
                        break
                    except:
                        attempts += 1
                        if attempts == max_attempts:
                            attempts = 0
                            time.sleep(1)
                        driver.refresh()
        
                import datetime
                data = datetime.datetime.now(tz=pytz.utc)
                data = data.strftime('%d/%m/%Y')
                agora = datetime.datetime.now()
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m") 

                df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Spotify.xlsx')            
                novo_registro = pd.DataFrame({'Grupo':url.replace('https://open.spotify.com/user/bancodobrasil','Bancos Comerciais')        
                                    ,'Plataforma':Spotify
                                    ,'Instituicao':url.replace('https://open.spotify.com/user/bancodobrasil','BB')
                                    ,'data':data
                                    ,'Seguidores':seguidoresSpotify
                                    ,'Ano':ano
                                    ,'Mês':mes
                                    ,'Link':url}, index=[0])
                df = pd.concat([df, novo_registro], ignore_index=True)
                df['Seguidores'] = df['Seguidores'].apply(lambda x: x.replace(' seguidores', ''))
                df['Seguidores'] = df['Seguidores'].apply(lambda x: x.replace('.', ''))
                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Spotify.xlsx', index=False)

            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Spotify')
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)
                data_e_hora_atuais = datetime.datetime.now()
                data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
                arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Spotify/Seguidores_Spotify.xlsx'
                arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Spotify/Seguidores_Spotify_' + data_e_hora_formatada + '.xlsx'
                os.rename(arquivo_origem, arquivo_destino)
                local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
                for arquivo in os.listdir(local):
                    os.remove(os.path.join(local, arquivo))
            print_to_text(f"{current_time} \u2714Busca seguidores Spotify concluída...")  

#============================================================================================================================================================================================================================

    def scrape_Telegram():
        
            urls = ['https://t.me/bancocentraloficial', 'https://t.me/btgpactual']
            Telegram = 'Telegram'
            file_path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Telegram.xlsx'
            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel(file_path, index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores Telegram...")
            
            for url in urls:
                try:
                    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
                    response = requests.get(url, headers=headers)
                    soup = BeautifulSoup(response.content, 'html.parser')
                    subscribers_element = soup.find('div', {'class': 'tgme_page_extra'})
                    subscribers = subscribers_element.text.strip()
                    subscribers = subscribers.replace('subscribers', '').replace(' ', '')
                    subscribers = re.sub(r'[^\d]', '', subscribers)
                    subscribers = int(subscribers)
                    subscribers = '{:,.0f}'.format(subscribers).replace(',', '.')
                    print_to_text2(f"{current_time} {url}")
                    print_to_text2(f"{subscribers}")
                except:
                    subscribers = '0 subscribers'

                if os.path.exists(file_path):
                    df = pd.read_excel(file_path)
                else:
                    df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])

                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m")

                novo_registro = pd.DataFrame({
                    'Grupo': url.replace('https://t.me/','').replace('bancocentraloficial', 'Outros').replace('btgpactual', 'Banco Boutique')
                    , 'Instituicao': url.replace('https://t.me/', '').replace('bancocentraloficial', 'BACEN').replace('btgpactual', 'BTG')
                    , 'Plataforma': Telegram
                    , 'data': data_atual
                    , 'Seguidores': subscribers
                    , 'Ano': ano
                    , 'Mês': mes
                    ,'Link':url}, index=[0])
                
                df = pd.concat([df, novo_registro], ignore_index=True)
                df['Seguidores'] = df['Seguidores'].astype(str)
                df['Seguidores'] = df['Seguidores'].apply(lambda x: x.replace('.', ''))
                df['Seguidores'] = df['Seguidores'].astype(int)
                df.to_excel(file_path, index=False)

            # Copiando os arquivos temp para base com data e hora
            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Telegram')
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)
                
            import datetime
            data_e_hora_atuais = datetime.datetime.now()
            data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
            arquivo_destino = f"C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Telegram/Seguidores_Telegram_{data_e_hora_formatada}.xlsx"
            os.rename(file_path, arquivo_destino)

            # Limpando pasta temp
            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))
            
            arquivo = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Telegram/Seguidores_Telegram.xlsx'
            os.remove(arquivo)

            print_to_text(f"{current_time} \u2714Busca seguidores Telegram concluída...")

#============================================================================================================================================================================================================================

    def scrape_tiktok():
        
            lista_urls_TikTok = ['https://www.tiktok.com/@bancodobrasil','https://www.tiktok.com/@bradesco','https://www.tiktok.com/@btgpactual','https://www.tiktok.com/@ministeriodafazenda','https://www.tiktok.com/@nubank']      
            TikTok = 'Tik Tok'

            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_TikTok.xlsx', index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores TikTok...")
                    
            for url in lista_urls_TikTok:
                #chromedriver_path = r"T:/DECOM_DEMKT/Dados/Bases/07_Automacao funcoes/dist/Automacao baseDados/webdriver/chromedriver.exe"
                options = webdriver.ChromeOptions()
                options.add_argument("--headless")
                options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
                options.add_argument("--log-level=3")
                #driver = webdriver.Chrome(executable_path=chromedriver_path, options=options)
                driver = webdriver.Chrome(options=options)
            
                driver.get(url)
                max_attempts = 10
                attempts = 0
                while attempts < max_attempts:                   
                    try:
                        seguidoresTikTok = driver.find_element(By.XPATH, '//*[@id="main-content-others_homepage"]/div/div[1]/h3/div[2]/strong').text
                        print_to_text2(f"{current_time} {url}")
                        print_to_text2(f"{seguidoresTikTok}:")
                        break
                    except:
                        attempts += 1
                        if attempts == max_attempts:
                            attempts = 0
                            time.sleep(1)
                        driver.refresh()
            
                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m")     
                        
                df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_TikTok.xlsx')
                df_new = pd.DataFrame({'Grupo':url
                                    .replace('https://www.tiktok.com/@','')
                                    .replace('obndes','BNDES')
                                    .replace('bancodobrasil','Bancos Comerciais')
                                    .replace('?','').replace('=','').replace('_','').replace(';','').replace('&','').replace('trafficty','').replace('peothersrefererurlampbradescoreferervideoid','').replace('6894298726983126273','')
                                    .replace('bradesco','Bancos Comerciais')
                                    .replace('btgpactual','Banco Boutique')
                                    .replace('ministeriodafazenda','Outros')
                                    .replace('nubank','Bancos Comerciais')
                                    ,'Instituicao':url
                                    .replace('https://www.tiktok.com/@','')
                                    .replace('obndes','BNDES')
                                    .replace('bancodobrasil','BB')
                                    .replace('?','').replace('=','').replace('_','').replace(';','').replace('&','').replace('trafficty','').replace('peothersrefererurlampbradescoreferervideoid','').replace('6894298726983126273','')
                                    .replace('bradesco','BRADESCO')
                                    .replace('btgpactual','BTG')
                                    .replace('ministeriodafazenda','Ministério Fazenda')
                                    .replace('nubank','NUBANK')
                                    ,'Plataforma':TikTok
                                    ,'data':data_atual
                                    ,'Seguidores':seguidoresTikTok.replace('K', 'k').replace('M', 'mi')
                                    ,'Ano':ano
                                    ,'Mês':mes
                                    ,'Link':url}, index=[0])
                df = pd.concat([df, df_new])
                
                def convert_count(seguidoresTikTok):
                    
                    if isinstance(seguidoresTikTok, str):
                        if 'k' in seguidoresTikTok:
                            seguidoresTikTok = seguidoresTikTok.replace('k', '')
                            seguidoresTikTok = float(seguidoresTikTok) * 1000
                        elif 'mi' in seguidoresTikTok:
                            seguidoresTikTok = seguidoresTikTok.replace('mi', '')
                            seguidoresTikTok = float(seguidoresTikTok) * 1000000
                    return seguidoresTikTok

                df['Seguidores'] = df['Seguidores'].apply(convert_count)
                
                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_TikTok.xlsx', index=False)

            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/TikTok')
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)
            
            data_e_hora_atuais = datetime.datetime.now()
            data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
            arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/TikTok/Seguidores_TikTok.xlsx'
            arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/TikTok/Seguidores_TikTok_' + data_e_hora_formatada + '.xlsx'
            os.rename(arquivo_origem, arquivo_destino)

            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))

            print_to_text(f"{current_time} \u2714Busca seguidores TikTok concluída...")

#============================================================================================================================================================================================================================

    def scrape_twitter():
            
            options = webdriver.ChromeOptions()
            options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
            options.add_argument("--log-level=3")
            driver = webdriver.Chrome(options=options)
            
            def login():

                driver.get("https://twitter.com/i/flow/login")
                email_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[5]/label/div/div[2]/div/input')))
                email_field.send_keys("e-mail")
                avançar_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div/div/div/div[6]/div/span/span')))
                avançar_button.click()
                nome_usuario = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div[2]/label/div/div[2]/div/input')))
                nome_usuario.send_keys("usuário")
                avançar_button_usuario = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div/div/div/div/span/span')))
                avançar_button_usuario.click()
                password_field = WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div[1]/div/div/div[3]/div/label/div/div[2]/div[1]/input')))
                password_field.send_keys("senha")
                submit_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable((By.XPATH, '//*[@id="layers"]/div/div/div/div/div/div/div[2]/div[2]/div/div/div[2]/div[2]/div[2]/div/div[1]/div/div/div/div/span/span')))
                submit_button.click()
                time.sleep(5)
             
            login()


            lista_urls_twitter = ['https://twitter.com/AFD_France'
                            ,'https://twitter.com/bancocentralbr'
                            ,'https://twitter.com/bancodobrasil'
                            ,'https://twitter.com/bndes'
                            ,'https://twitter.com/bradesco'
                            ,'https://twitter.com/BTGPactual'
                            ,'https://twitter.com/caixa'
                            ,'https://twitter.com/imfnews'
                            ,'https://twitter.com/itau'
                            ,'https://twitter.com/jpmorgan'
                            ,'https://twitter.com/KfW_FZ_int'
                            ,'https://twitter.com/MinFazenda'
                            ,'https://twitter.com/morganstanley'
                            ,'https://twitter.com/nubank'
                            ,'https://twitter.com/petrobras'
                            ,'https://twitter.com/santander_br'
                            ,'https://twitter.com/sebrae'
                            ,'https://twitter.com/worldbank']                                
                                        
            Twitter = 'Twitter'

            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_twitter.xlsx', index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores Twitter...")
                

            for url in lista_urls_twitter:
                driver.get(url)
                wait = WebDriverWait(driver, 10)
                try:
                    inscritos = wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="react-root"]/div/div/div[2]/main/div/div/div/div/div/div[3]/div/div/div/div/div[5]/div[2]/a/span[1]/span')))
                    seguidoresTwitter = inscritos.text
                    print_to_text2(f"{current_time} {url}")
                    print_to_text2(f"{seguidoresTwitter}")

                except:
                    seguidoresTwitter = 0
                    
                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m") 
                    
                df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_twitter.xlsx')
                df = pd.concat([df,pd.DataFrame({'Grupo':url
                                                        .replace('https://twitter.com/','')         
                                                        .replace('AFD_France','Bancos Desenvolvimento')
                                                        .replace('bancocentralbr','Outros')
                                                        .replace('bancodobrasil','Bancos Comerciais')
                                                        .replace('bndes','BNDES')
                                                        .replace('bradesco','Bancos Comerciais')
                                                        .replace('BTGPactual','Banco Boutique')
                                                        .replace('caixa','Bancos Comerciais')
                                                        .replace('imfnews','Bancos Desenvolvimento')
                                                        .replace('itau','Bancos Comerciais')
                                                        .replace('jpmorgan','Banco Boutique')
                                                        .replace('KfW_FZ_int','Bancos Desenvolvimento')
                                                        .replace('MinFazenda','Outros')                            
                                                        .replace('morganstanley','Banco Boutique')
                                                        .replace('nubank','Bancos Comerciais')
                                                        .replace('petrobras','Outros')
                                                        .replace('santander_br','Bancos Comerciais')
                                                        .replace('sebrae','Outros')
                                                        .replace('worldbank','Bancos Desenvolvimento')
                                                    ,'Plataforma':Twitter
                                                    ,'Instituicao':url
                                                        .replace('https://twitter.com/','') 
                                                        .replace('AFD_France','AFD')
                                                        .replace('bancocentralbr','BACEN')
                                                        .replace('bancodobrasil','BB')
                                                        .replace('bndes','BNDES')
                                                        .replace('bradesco','BRADESCO')
                                                        .replace('BTGPactual','BTG')
                                                        .replace('caixa','Caixa')
                                                        .replace('imfnews','IMF')
                                                        .replace('itau','ITAU')
                                                        .replace('jpmorgan','JPMorgan')
                                                        .replace('KfW_FZ_int','KfW')
                                                        .replace('MinFazenda','Ministério Fazenda')                           
                                                        .replace('morganstanley','Morgan Stanley')
                                                        .replace('nubank','NuBank')
                                                        .replace('petrobras','Petrobras')
                                                        .replace('santander_br','Santander')
                                                        .replace('sebrae','SEBRAE')
                                                        .replace('worldbank','WorldBank')
                                                    ,'data': data_atual
                                                    ,'Seguidores':seguidoresTwitter
                                                    ,'Ano':ano
                                                    ,'Mês':mes
                                                    ,'Link':url}, index=[0])], ignore_index=True)

                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_twitter.xlsx', index=False)
                    
            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/twitter')
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)

            data_e_hora_atuais = datetime.datetime.now()
            data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M_%S')
            arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/twitter/Seguidores_twitter.xlsx'
            arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/twitter/Seguidores_twitter_' + data_e_hora_formatada + '.xlsx'
            os.rename(arquivo_origem, arquivo_destino)

            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))
                
            print_to_text(f"{current_time} \u2714Busca seguidores Twitter concluída...") 

#============================================================================================================================================================================================================================

    def scrape_Xing():
        
            lista_urls_Xing = ['https://www.xing.com/pages/kfw']
            Xing = 'Xing'
            
            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Xing.xlsx', index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores Xing...")
            
            for url in lista_urls_Xing:
                try:
                    headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'}
                    r = requests.get(url, headers=headers)
                    soup = BeautifulSoup(r.content, 'html.parser')
                    seguidoresXing = soup.select_one('#content div p span')
                    seguidoresXing = seguidoresXing.text.strip()
                    print_to_text2(f"{current_time} {url}")
                    print_to_text2(f"{seguidoresXing}")
                except:
                    seguidoresXing = '0'
                
                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m") 
                
                df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Xing.xlsx')
                df = pd.concat([df,pd.DataFrame({'Grupo':url
                                    .replace('https://www.xing.com/pages/kfw', 'Bancos Desenvolvimento')                       
                                    ,'Plataforma':Xing
                                    ,'Instituicao':url
                                    .replace('https://www.xing.com/pages/kfw','KfW')
                                    ,'data': data_atual
                                    ,'Seguidores':seguidoresXing
                                    ,'Ano':ano
                                    ,'Mês':mes
                                    ,'Link':url}, index=[0])], ignore_index=True)
                df['Seguidores'] = df['Seguidores'].apply(lambda x: x.replace('.', ''))
                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Xing.xlsx', index=False)
                
            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Xing')
            
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)

            data_e_hora_atuais = datetime.datetime.now()
            data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
            arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Xing/Seguidores_Xing.xlsx'
            arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Xing/Seguidores_Xing_' + data_e_hora_formatada + '.xlsx'
            os.rename(arquivo_origem, arquivo_destino)

            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))
            
            print_to_text(f"{current_time} \u2714Busca seguidores Xing concluída...")    

#============================================================================================================================================================================================================================

    def scrape_Youtube():

            lista_urls_Youtube = ['https://www.youtube.com/GroupeAFD'
                                ,'https://www.youtube.com/user/BancoCentralBR'
                                ,'https://www.youtube.com/bancodobrasil'
                                ,'https://www.youtube.com/user/imprensaBancoNE'
                                ,'https://www.youtube.com/user/bndesgovbr'
                                ,'https://www.youtube.com/bradesco'
                                ,'https://www.youtube.com/c/btgpactual'
                                ,'https://www.youtube.com/canalcaixa'
                                ,'https://www.youtube.com/user/CitiBrasil'
                                ,'https://www.youtube.com/IMF'
                                ,'https://www.youtube.com/itau'
                                ,'https://www.youtube.com/user/jpmorgan'
                                ,'https://www.youtube.com/kfw'
                                ,'https://www.youtube.com/@MinFazenda'
                                ,'https://www.youtube.com/user/mgstnly'
                                ,'https://www.youtube.com/channel/UCgsDX3hTwiPdtGHJjMFfDxg'
                                ,'https://www.youtube.com/user/canalpetrobras'
                                ,'https://www.youtube.com/channel/UCxvisnfGI7j6SCnpdz0lJpg'
                                ,'https://www.youtube.com/user/tvsebrae'
                                ,'https://www.youtube.com/user/WorldBank']

            Youtube = 'Youtube'

            #Criando arquivo base com cabeçalho
            df = pd.DataFrame(columns=['Grupo', 'Instituicao', 'Plataforma', 'data', 'Seguidores', 'Ano', 'Mês', 'Link'])
            df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Youtube.xlsx', index=False)

            print_to_text(f"{current_time} ->Iniciando busca Seguidores Youtube...")
                        
            for url in lista_urls_Youtube:
                options = webdriver.ChromeOptions()
                options.add_argument('--headless')
                options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.36")
                options.add_argument("--log-level=3")

                try:
                    driver = webdriver.Chrome(options=options)  # Inicialize o driver do Chrome
                    driver.get(url)
                    print_to_text2(f"{current_time} {url}")
                except:
                    if driver:
                        driver.quit()
                    continue

                try:
                    channel_title = driver.find_element(By.XPATH, '//yt-formatted-string[contains(@class, "ytd-channel-name")]').text
                except:
                    channel_title = ''

                try:
                    subscriber_count = driver.find_element(By.XPATH, '//yt-formatted-string[@id="subscriber-count"]').text
                    print_to_text2(f"{subscriber_count}")
                except:
                    subscriber_count = '0 Seguidores'

                WAIT_IN_SECONDS = 3
                last_height = driver.execute_script("return document.documentElement.scrollHeight")

                #Criando coluna com data atual
                import datetime
                agora = datetime.datetime.now()
                data_atual = agora.strftime("%d/%m/%Y")
                ano = agora.strftime("%Y")
                mes = agora.strftime("%m")

                df = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Youtube.xlsx')
                novo_registro = pd.DataFrame({'Grupo':channel_title
                                            .replace('AFD - Agence française de développement','Bancos Desenvolvimento')
                                            .replace('Banco Central do Brasil','Outros')
                                            .replace('Banco do Brasil','Bancos Comerciais')
                                            .replace('Banco do Nordeste','Bancos Desenvolvimento')
                                            .replace('BNDES','BNDES')
                                            .replace('bradesco','Bancos Comerciais')
                                            .replace('BTG Pactual','Banco Boutique')
                                            .replace('CAIXA','Bancos Comerciais')
                                            .replace('Citi Brasil','Banco Boutique')
                                            .replace('IMF','Bancos Desenvolvimento')
                                            .replace('CAIXA','Bancos Comerciais')
                                            .replace('Itaú','Bancos Comerciais')
                                            .replace('jpmorgan','Banco Boutique')
                                            .replace('KfW Bankengruppe','Bancos Desenvolvimento')
                                            .replace('Ministério da Fazenda','Outros')
                                            .replace('Morgan Stanley','Banco Boutique')
                                            .replace('Nubank','Bancos Comerciais')                                          
                                            .replace('Santander Brasil','Bancos Comerciais')
                                            .replace('Petrobras','Outros')                                          
                                            .replace('Sebrae','Outros')
                                            .replace('World Bank','Bancos Desenvolvimento')
                                            ,'Instituicao':channel_title
                                            .replace('AFD - Agence française de développement','AFD')
                                            .replace('Banco Central do Brasil','BACEN')
                                            .replace('Banco do Brasil','BB')
                                            .replace('Banco do Nordeste','BNB')
                                            .replace('','')#BNDES
                                            .replace('bradesco','BRADESCO')
                                            .replace('BTG Pactual','BTG')
                                            .replace('CAIXA','Caixa')
                                            .replace('Citi Brasil','City Bank')
                                            .replace('CAIXA','Caixa')
                                            .replace('Itaú','ITAU')
                                            .replace('jpmorgan','JPMorgan')
                                            .replace('KfW Bankengruppe','KfW')
                                            .replace('','')#Ministério Fazenda
                                            .replace('MorganStanley','Morgan Stanley')
                                            .replace('','')#NUBANK
                                            .replace('','')#Petrobras                                          
                                            .replace('Santander Brasil','Santander')
                                            .replace('Sebrae','SEBRAE')
                                            .replace('World Bank','WorldBank')
                                            ,'Plataforma':Youtube
                                            ,'data':data_atual
                                            ,'Seguidores':subscriber_count.replace('inscritos', '').replace('subscritores', '').replace('de', '').replace('mil', 'k').replace(',', '.').replace(' ', '')
                                            ,'Ano':ano
                                            ,'Mês':mes
                                            ,'Link':url}, index=[0])
                                                        
                df = pd.concat([df, novo_registro], ignore_index=True)
                
                def convert_count(subscriber_count):
                    if isinstance(subscriber_count, str):
                        if 'k' in subscriber_count:
                            subscriber_count = subscriber_count.replace('k', '')
                            subscriber_count = float(subscriber_count) * 1000
                        elif 'mi' in subscriber_count:
                            subscriber_count = subscriber_count.replace('mi', '')
                            subscriber_count = float(subscriber_count) * 1000000
                    return subscriber_count

                df['Seguidores'] = df['Seguidores'].apply(convert_count)

                df.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/Seguidores_Youtube.xlsx', index=False)

            #Copiando os arquivos temp para base com data e hora
            origem = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')
            destino = Path('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Youtube')
            for arquivo in origem.iterdir():
                shutil.copy(arquivo, destino)

            data_e_hora_atuais = datetime.datetime.now()
            data_e_hora_formatada = data_e_hora_atuais.strftime('%d_%m_%Y_%H_%M')
            arquivo_origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Youtube/Seguidores_Youtube.xlsx'
            arquivo_destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/Youtube/Seguidores_Youtube_' + data_e_hora_formatada + '.xlsx'
            os.rename(arquivo_origem, arquivo_destino)

            #Limpando pasta temp
            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))
                
            print_to_text(f"{current_time} \u2714Busca seguidores Youtube concluída...")

#============================================================================================================================================================================================================================
           
    def criarFeed():
        
        def montaBase():
            
            print_to_text(f"{current_time} ->Montando Feed Seguidores...")
        
            i = 1        
            pastas = ['Facebook', 'Flickr', 'Instagram', 'Linkedin', 'Spotify', 'Telegram', 'TikTok', 'twitter', 'Xing', 'Youtube']
            
            for pasta in pastas:
                path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/baseDados/' + pasta + '/'
                date_time = max([f for f in os.listdir(path) if os.path.isfile(os.path.join(path, f))])
                shutil.copyfile(os.path.join(path, date_time), 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/' + date_time) 
                    
            for filename in os.listdir("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/"): 
                dst ="seguidores" + str(i) + ".xlsx"
                src ='C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'+ filename 
                dst ='C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'+ dst 
                os.rename(src, dst) 
                i += 1
                
            arquivos = glob.glob('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/seguidores*.xlsx')
            dados = [pd.read_excel(arquivo) for arquivo in arquivos]
            baseDadosConcat = pd.concat(dados)
            baseDadosConcat.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/baseDadosConcat.xlsx', sheet_name='Seguidores')
            baseDadosConcat = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/baseDadosConcat.xlsx', sheet_name='Seguidores')
            baseDadosConcat.drop(baseDadosConcat.columns[0], axis=1, inplace=True)
            baseDadosConcat.to_excel('C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/baseDadosConcat.xlsx', index=False)
            current_path = os.getcwd()
            filename = "baseDadosConcat.xlsx"
            source_path = os.path.join(current_path, 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp')


            import datetime
            from datetime import datetime
            date_time = datetime.now().strftime('%d_%m_%Y_%H_%M_%S')
            filename_with_date_time = 'Base_Seguidores_' + date_time + '.xlsx'
            destination_path = os.path.join(current_path, 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Bases')

            os.rename(os.path.join(source_path, filename), os.path.join(destination_path, filename_with_date_time))
            local = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/temp/'
            for arquivo in os.listdir(local):
                os.remove(os.path.join(local, arquivo))
            print_to_text(f"{current_time} \u2714Feed Seguidores criada com sucesso...")
                
        montaBase()

#============================================================================================================================================================================================================================

    def abrirArquivo():
        pasta = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Bases'
        arquivos = os.listdir(pasta)
        arquivos_xlsx = [arq for arq in arquivos if arq.endswith(".xlsx")]
        if arquivos_xlsx:
            mais_recente = max(arquivos_xlsx, key=lambda x: os.path.getmtime(os.path.join(pasta, x)))
            caminho_arquivo = os.path.join(pasta, mais_recente)
            if os.path.exists(caminho_arquivo):
                os.startfile(caminho_arquivo)
                time.sleep(5)
                keyboard.press_and_release('ctrl+alt+t')
                time.sleep(1)
                keyboard.press_and_release('enter')
                time.sleep(1)
                keyboard.press_and_release('ctrl+b')
                time.sleep(1)
            else:
                print_to_text(f"O arquivo '{caminho_arquivo}' não existe.")
        else:
            print_to_text("Não há arquivos XLSX na pasta.")

#============================================================================================================================================================================================================================

    def atualizarBase():
        print_to_text(f"{current_time} -> Atualizando Base Seguidores.xlsx...")

        # Copia o arquivo mais recente da pasta Bases
        directory = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Bases'
        files = glob.glob(os.path.join(directory, '*.xlsx'))
        most_recent_file = max(files, key=os.path.getmtime)

        # Cria o dataframe e deleta a coluna Link
        seguidores_feed = pd.read_excel(most_recent_file)
        seguidores_feed = seguidores_feed.drop('Link', axis=1)

        # Salva o dataframe em xlsx
        seguidores_feed.to_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Seguidores(feed).xlsx", index=False)

        # Le os arquivos Base e feed
        seguidores_base = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx', sheet_name=None)
        seguidores_feed = pd.read_excel("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/Seguidores(feed).xlsx")

        # Atualiza a base com o feed
        for sheet_name, df in seguidores_base.items():
            if sheet_name == 'Seguidores':
                seguidores_base[sheet_name] = pd.concat([df, seguidores_feed], ignore_index=True)

        # Salva a base atualizada em xlsx preservando as abas
        with pd.ExcelWriter("C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx", engine="openpyxl") as writer:
            for sheet_name, df in seguidores_base.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        print_to_text(f"{current_time} -> Abrindo o arquivo C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx...\n{current_time} -> Aguarde a formatação da planilha em Tabela...")
        path = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx'
        os.startfile(path)
        time.sleep(5)
        keyboard.press_and_release('ctrl+alt+t')
        time.sleep(2)
        keyboard.press_and_release('enter')
        time.sleep(1)
        keyboard.press_and_release('ctrl+b')
        print_to_text(f"{current_time} \u2714 Formatação concluída...")
        print_to_text(f"{current_time} \u2714 Base Seguidores atualizada com sucesso...")

#============================================================================================================================================================================================================================

    def atualizaçãoPastas_Seguidores():
        
        import datetime
        current_time = datetime.datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        
        try:
            
            src_file = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx'
            dst_folder = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'
            dst_folder2 = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'
            shutil.copy(src_file, dst_folder)
            shutil.copy(src_file, dst_folder2)
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema/'Atualizada...")
            print_to_text(f"{current_time} \u2714Pasta 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público/'Atualizada...")
            
        except Exception as e:
            print_to_text(f"{current_time}- !Falha [Atualizar pastas] Ocorreu um erro ao copiar o arquivo para as pastas 03_Sistema e 04_Público: {e}")

#============================================================================================================================================================================================================================

    def listaLinks():
        listaLinks = pd.read_excel('C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_TratadasSeguidores.xlsx', sheet_name='Links')
        
        for index, row in listaLinks.iterrows():
            instituicao = row['Instituição']
            plataforma = row['Plataforma']
            endereco = row['Endereço']
            print_to_text2(f"Instituição: {instituicao}")
            print_to_text2(f"Plataforma: {plataforma}")
            print_to_text2(f"Endereço: {endereco}")
            print_to_text2("")

#============================================================================================================================================================================================================================

    #==========Interface menu Seguidores
    
    tab5_frame = ttk.Frame(tab_control)
    tab_control.add(tab5_frame, text='Seguidores         ')     
    tab5_frame.configure(height=100)
    
    label = ttk.Label(tab5_frame, text='Atualização do número de seguidores da base Seguidores')
    label.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    button_frame = ttk.Frame(tab5_frame)
    button_frame.grid(row=1, column=0, padx=5, sticky="w")

    button_frame2 = ttk.Frame(tab5_frame)
    button_frame2.grid(row=2, column=0, padx=5, sticky="w")

#============================================================================================================================================================================================================================   
   
    #==========Botões de Busca
    
    button_width = 11 # Largura dos botões
  
    button_Facebook = tk.Button(button_frame, text="Facebook", command=scrape_Facebook, width=button_width)
    button_Facebook.grid(row=0, column=0, padx=1, pady=5)
        
    button_Flickr = tk.Button(button_frame, text="Flickr", command=scrape_Flickr, width=button_width)
    button_Flickr.grid(row=0, column=1, padx=1, pady=5)

    button_Instagram = tk.Button(button_frame, text="Instagram", command=scrape_instagram, width=button_width)
    button_Instagram.grid(row=0, column=2, padx=1, pady=5)

    button_linkedin = tk.Button(button_frame, text="linkedin", command=scrape_linkedin, width=button_width)
    button_linkedin.grid(row=0, column=3, padx=1, pady=5)
        
    button_Spotify = tk.Button(button_frame, text="Spotify", command=scrape_Spotify, width=button_width)
    button_Spotify.grid(row=0, column=4, padx=1, pady=5)
        
    button_Telegram = tk.Button(button_frame, text="Telegram", command=scrape_Telegram, width=button_width)
    button_Telegram.grid(row=0, column=5, padx=1, pady=5)
            
    button_TikTok = tk.Button(button_frame, text="TikTok", command=scrape_tiktok, width=button_width)
    button_TikTok.grid(row=0, column=6, padx=1, pady=5)
        
    button_Twitter = tk.Button(button_frame, text="Twitter", command=scrape_twitter, width=button_width)
    button_Twitter.grid(row=0, column=7, padx=1, pady=5)
            
    button_Xing = tk.Button(button_frame, text="Xing", command=scrape_Xing, width=button_width)
    button_Xing.grid(row=0, column=8, padx=1, pady=5)
            
    button_Youtube = tk.Button(button_frame, text="Youtube", command=scrape_Youtube, width=button_width)
    button_Youtube.grid(row=0, column=9, padx=1, pady=5)
#============================================================================================================================================================================================================================
    
    #==========Botão Criar feed
    
    button_criarFeed = tk.Button(button_frame2, text="Criar Seguidores(feed)", command=criarFeed, width=17)
    button_criarFeed.grid(row=1, column=0, padx=2, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    #==========Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab5_frame, text="Cria a o arquivo Seguidores(feed).xlsx com os arquivos mais recentes da pasta baseDados ", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_criarFeed.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_criarFeed.bind("<Leave>", hide_status_tooltip)
#============================================================================================================================================================================================================================
    
    #==========Botão Abrir Seguidores(feed)
    
    button_criarFeed = tk.Button(button_frame2, text="Abrir Seguidores(feed)", command=abrirArquivo, width=17)
    button_criarFeed.grid(row=1, column=1, padx=2, pady=5)

#============================================================================================================================================================================================================================
    
    #==========Botão Atualizar Seguidores
    
    button_atualizarBase = tk.Button(button_frame2, text="Atualizar Seguidores", command=atualizarBase, width=17)
    button_atualizarBase.grid(row=1, column=2, padx=2, pady=5)
    
    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab5_frame, text="Atualiza a base Seguidores.xlsx com o Feed", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_atualizarBase.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_atualizarBase.bind("<Leave>", hide_status_tooltip)
    
#============================================================================================================================================================================================================================
    
    #==========Botão Atualizar pastas
    
    button_atualizarPastas = tk.Button(button_frame2, text="Atualizar pastas", command=atualizaçãoPastas_Seguidores, width=17)
    button_atualizarPastas.grid(row=1, column=3, padx=2, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab5_frame, text="Faz uma cópia do arquivo Seguidores.xlsx para as pastas 03_Sistema e 04_Público", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_atualizarPastas.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_atualizarPastas.bind("<Leave>", hide_status_tooltip)
               
#============================================================================================================================================================================================================================
    
    #==========Botão Lista Links Seguidores
    
    button_listaLinks = tk.Button(button_frame2, text="Lista Links Seguidores", command=listaLinks, width=17)
    button_listaLinks.grid(row=1, column=4, padx=2, pady=5)

    # Variável para armazenar a legenda
    tooltip = None

    # Função para exibir a legenda
    def show_status_tooltip(event):
        global tooltip
        tooltip = tk.Label(tab5_frame, text="Exibe a lista de Links da base Seguidores", background="white", relief="solid", borderwidth=1)
        tooltip.grid(row=0, column=0, padx=10, pady=10, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_listaLinks.bind("<Enter>", show_status_tooltip)

    # Função para ocultar a legenda
    def hide_status_tooltip(event):
        global tooltip
        if tooltip is not None:
            tooltip.destroy()
            tooltip = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão
    button_listaLinks.bind("<Leave>", hide_status_tooltip)
               
#============================================================================================================================================================================================================================

    #==========Label
    label_text = ttk.Label(button_frame, text="")
    label_text.grid(row=0, column=5, columnspan=4, padx=20, pady=5, sticky="e")

#============================================================================================================================================================================================================================
    
    #==========Frames
    button_frame.columnconfigure(4, weight=1)  # Configurar coluna para expansão horizontal

    text_frame = ttk.Frame(tab5_frame)
    text_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

#============================================================================================================================================================================================================================

    #==========Área de exibição de prints
    
    # Criar o PanedWindow
    paned_window = ttk.PanedWindow(text_frame, orient='vertical')
    paned_window.pack(fill="both", expand=True, pady=(15, 0))

    # Área 1
    text_widget1 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget1, weight=1)  # Definir peso igual para a área 1

    # Área 2
    text_widget2 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget2, weight=1)
    
    # Configurar as barras de rolagem
    text_scrollbar1 = ttk.Scrollbar(text_widget1, command=text_widget1.yview)
    text_widget1.config(yscrollcommand=text_scrollbar1.set)
    text_scrollbar1.pack(side="right", fill="y")

    text_scrollbar2 = ttk.Scrollbar(text_widget2, command=text_widget2.yview)
    text_widget2.config(yscrollcommand=text_scrollbar2.set)
    text_scrollbar2.pack(side="right", fill="y")

    total_iterations = 10 # Número do range das buscas 10 iterações equivalem a 40 buscas.

    # Criar a barra de progresso com o estilo personalizado
    progress_bar = ttk.Progressbar(left_frame, mode='indeterminate', maximum=total_iterations)
    progress_bar.grid(row=8, column=0, padx=10, pady=5, sticky="ew")

    # Definir o tamanho da barra de progresso (comprimento em pixels)
    progress_bar.config(length=10)  # Substitua 300 pelo valor desejado

    def print_to_text(text):
        text_widget1.insert("end", text + "\n")
        text_widget1.see("end")
        text_widget1.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget1
    
    def print_to_text2(text):
        text_widget2.insert("end", text + "\n")
        text_widget2.see("end")
        text_widget2.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget2

    # Configurar o redimensionamento do tab5_frame
    tab5_frame.columnconfigure(0, weight=1)  # Expandir coluna 0
    tab5_frame.rowconfigure(3, weight=1)  # Expandir linha 3

#============================================================================================================================================================================================================================

#==========Menu Backups
     
def tab6():
                                
    def backupTradadas():
        
        origem = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas'
        destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_BackUp/@Backup Tratadas/'

        # Obter a data e hora atual
        import datetime
        current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

        print_to_text(f"{current_time} -> Iniciando Backup da pasta 02_Tratadas.")
        try:
            # Verificar se a pasta de destino existe, caso contrário, criar a pasta
            if not os.path.exists(destino):
                os.makedirs(destino)

            # Listar todos os arquivos na pasta de origem
            arquivos = [arquivo for arquivo in os.listdir(origem) if arquivo.endswith('.xlsx')]

            # Copiar os arquivos para a pasta de destino com a data e hora atual no nome
            for arquivo in arquivos:
                nome_arquivo, extensao = os.path.splitext(arquivo)
                novo_nome = f"{nome_arquivo}_{current_time}.xlsx"
                shutil.copy2(os.path.join(origem, arquivo), os.path.join(destino, novo_nome))
                current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
                print_to_text2(f"{current_time} -> {novo_nome}")

            print_to_text(f"{current_time} \u2714 Backup da pasta 02_Tratadas concluído com sucesso.")
        except Exception as e:
            print_to_text(f"{current_time} ! Erro ao fazer o backup da pasta 02_Tratadas: {e}")

#============================================================================================================================================================================================================================                                      
                                      
    def backupPaineis():
        origem_principal = 'D:/DECOM_DEMKT/Paineis de dados'
        origem_secundaria = 'D:/DECOM_DEMKT/Paineis de dados/2_Homologação'
        destino = 'C:/Automacao_DECOM&DEMKT_BNDES/Bases/05_Backup/@Backup Seguidores'

        # Obter a data atual
        import datetime
        current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")

        print_to_text(f"{current_time} -> Iniciando Backup das pastas Paineis de dados.")
        try:
            # Verificar se a pasta de destino existe, caso contrário, criar a pasta
            if not os.path.exists(destino):
                os.makedirs(destino)

            # Backup da pasta de origem principal
            backup_pasta(origem_principal, destino, current_time)

            # Backup da pasta de origem secundária
            backup_pasta(origem_secundaria, destino, current_time)

            print_to_text(f"{current_time} \u2714 Backup das pastas Paineis de dados concluído com sucesso.")
        except Exception as e:
            print_to_text(f"{current_time} ! Erro ao fazer o backup das pastas Paineis de dados: {e}")


    def backup_pasta(origem, destino, current_time):
        # Listar todos os arquivos na pasta de origem
        arquivos = os.listdir(origem)

        # Copiar os arquivos com a extensão .pbix para a pasta de destino com a data atual no nome
        for arquivo in arquivos:
            if arquivo.endswith(".pbix"):
                nome_arquivo, extensao = os.path.splitext(arquivo)
                novo_nome = f"{nome_arquivo}_{current_time}.pbix"
                shutil.copy2(os.path.join(origem, arquivo), os.path.join(destino, novo_nome))
                print_to_text2(f"{current_time} -> {novo_nome}")

#============================================================================================================================================================================================================================                                      
                                      
    def atualizaçãoSistema():
        origem = "C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas"
        destino = "C:/Automacao_DECOM&DEMKT_BNDES/Bases/03_Sistema"

        import datetime
        current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        
        # Verificar se a pasta de destino existe, caso contrário, criá-la
        if not os.path.exists(destino):
            os.makedirs(destino)

        print_to_text(f"{current_time} -> Iniciando cópia dos arquivos 02_Tratadas para 03_Sistemas.")
        
        # Percorrer todos os arquivos na pasta de origem
        for nome_arquivo in os.listdir(origem):
            # Construir o caminho completo para o arquivo de origem e destino
            caminho_origem = os.path.join(origem, nome_arquivo)
            caminho_destino = os.path.join(destino, nome_arquivo)

            # Verificar se o item na pasta de origem é um arquivo e não um diretório
            if os.path.isfile(caminho_origem):
                # Copiar o arquivo para a pasta de destino
                shutil.copy2(caminho_origem, caminho_destino)

                # Imprimir o nome do arquivo copiado
                current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
                print_to_text2(f"{current_time} -> Arquivo copiado: {nome_arquivo}")

        print_to_text(f"{current_time} \u2714 Cópia dos arquivos 02_Tratadas para 03_Sistemas concluída.")

#============================================================================================================================================================================================================================                                      
                                      
    def atualizaçãoPúblico():
        origem = "C:/Automacao_DECOM&DEMKT_BNDES/Bases/02_Tratadas"
        destino = "C:/Automacao_DECOM&DEMKT_BNDES/Bases/04_Público"

        import datetime
        current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
        
        # Verificar se a pasta de destino existe, caso contrário, criá-la
        if not os.path.exists(destino):
            os.makedirs(destino)

        print_to_text(f"{current_time} -> Iniciando cópia dos arquivos 02_Tratadas para 04_Público.")
        
        # Percorrer todos os arquivos na pasta de origem
        for nome_arquivo in os.listdir(origem):
            # Construir o caminho completo para o arquivo de origem e destino
            caminho_origem = os.path.join(origem, nome_arquivo)
            caminho_destino = os.path.join(destino, nome_arquivo)

            # Verificar se o item na pasta de origem é um arquivo e não um diretório
            if os.path.isfile(caminho_origem):
                # Copiar o arquivo para a pasta de destino
                shutil.copy2(caminho_origem, caminho_destino)

                # Imprimir o nome do arquivo copiado
                current_time = datetime.datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
                print_to_text2(f"{current_time} -> Arquivo copiado: {nome_arquivo}")

        print_to_text(f"{current_time} \u2714 Cópia dos arquivos 02_Tratadas para 04_Público concluída.")  
              
#============================================================================================================================================================================================================================    
    
    #==========Interface menu Backups
          
    tab6_frame = ttk.Frame(tab_control)
    tab_control.add(tab6_frame, text='Backups     ')     
    tab6_frame.configure(height=100)

    label = ttk.Label(tab6_frame, text='Rotinas de Backups\n\n')
    label.grid(row=0, column=0, padx=10, pady=10, sticky="w")
 
    button_frame = ttk.Frame(tab6_frame)
    button_frame.grid(row=2, column=0, padx=20, sticky="w")
  
#============================================================================================================================================================================================================================   

    #==========Botão Backup Tratadas
    
    button_backupTradadas = tk.Button(button_frame, text="Backup 02_Tratadas", command=backupTradadas, width=18)
    button_backupTradadas.grid(row=0, column=0, padx=5, pady=5)
    
    # Função para exibir a legenda do botão
    def show_backupTradadas_tooltip(event):
        global tooltip_backupTradadas
        tooltip_backupTradadas = tk.Label(tab6_frame, text="Realiza o backup de todo conteudo da pasta 02_Tratadas para a pasta 05_BackUp ", background="white", relief="solid", borderwidth=1)
        tooltip_backupTradadas.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_backupTradadas.bind("<Enter>", show_backupTradadas_tooltip)

    # Função para ocultar a legenda do botão
    def hide_backupTradadas_tooltip(event):
        global tooltip_backupTradadas
        if tooltip_backupTradadas is not None:
            tooltip_backupTradadas.destroy()
            tooltip_backupTradadas = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão 
    button_backupTradadas.bind("<Leave>", hide_backupTradadas_tooltip)

#============================================================================================================================================================================================================================

    #==========Botão Backup Paineis de dados
    
    button_backupPaineis = tk.Button(button_frame, text="Backup Paineis", command=backupPaineis, width=18)
    button_backupPaineis.grid(row=0, column=1, padx=5, pady=5)

    # Função para exibir a legenda do botão
    def show_backupPaineis_tooltip(event):
        global tooltip_backupPaineis
        tooltip_backupPaineis = tk.Label(tab6_frame, text="Realiza o backup de todo conteudo da pasta c:\\DECOM_DEMKT\\Dados\\Paineis de dados, para a pasta 05_BackUp ", background="white", relief="solid", borderwidth=1)
        tooltip_backupPaineis.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_backupPaineis.bind("<Enter>", show_backupPaineis_tooltip)

    # Função para ocultar a legenda do botão
    def hide_backupPaineis_tooltip(event):
        global tooltip_backupPaineis
        if tooltip_backupPaineis is not None:
            tooltip_backupPaineis.destroy()
            tooltip_backupPaineis = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão 
    button_backupPaineis.bind("<Leave>", hide_backupPaineis_tooltip)

#============================================================================================================================================================================================================================

    #==========Botão Atualização 03_Sistemas
    
    button_atualizaçãoSistema = tk.Button(button_frame, text="Atualização 03_Sistema", command=atualizaçãoSistema, width=18)
    button_atualizaçãoSistema.grid(row=0, column=2, padx=5, pady=5)

    # Função para exibir a legenda do botão
    def show_atualizaçãoSistema_tooltip(event):
        global tooltip_atualizaçãoSistema
        tooltip_atualizaçãoSistema = tk.Label(tab6_frame, text="Realiza o backup de todo conteudo da pasta 02_Tratadas, para a pasta 03_Sistema ", background="white", relief="solid", borderwidth=1)
        tooltip_atualizaçãoSistema.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_atualizaçãoSistema.bind("<Enter>", show_atualizaçãoSistema_tooltip)

    # Função para ocultar a legenda do botão
    def hide_atualizaçãoSistema_tooltip(event):
        global tooltip_atualizaçãoSistema
        if tooltip_atualizaçãoSistema is not None:
            tooltip_atualizaçãoSistema.destroy()
            tooltip_atualizaçãoSistema = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão 
    button_atualizaçãoSistema.bind("<Leave>", hide_atualizaçãoSistema_tooltip)

#============================================================================================================================================================================================================================

    #==========Botão Atualização 04_Público
    
    button_atualizaçãoPúblico = tk.Button(button_frame, text="Atualização 04_Público", command=atualizaçãoPúblico, width=18)
    button_atualizaçãoPúblico.grid(row=0, column=3, padx=5, pady=5)

    # Função para exibir a legenda do botão
    def show_atualizaçãoPúblico_tooltip(event):
        global tooltip_atualizaçãoPúblico
        tooltip_atualizaçãoPúblico = tk.Label(tab6_frame, text="Realiza o backup de todo conteudo da pasta 02_Tratadas, para a pasta 04_Público ", background="white", relief="solid", borderwidth=1)
        tooltip_atualizaçãoPúblico.grid(row=0, column=0, padx=20, sticky="w")

    # Vincular a exibição da legenda ao evento de passar o mouse sobre o botão
    button_atualizaçãoPúblico.bind("<Enter>", show_atualizaçãoPúblico_tooltip)

    # Função para ocultar a legenda do botão
    def hide_atualizaçãoPúblico_tooltip(event):
        global tooltip_atualizaçãoPúblico
        if tooltip_atualizaçãoPúblico is not None:
            tooltip_atualizaçãoPúblico.destroy()
            tooltip_atualizaçãoPúblico = None

    # Vincular a ocultação da legenda ao evento de mover o mouse fora do botão 
    button_atualizaçãoPúblico.bind("<Leave>", hide_atualizaçãoPúblico_tooltip)
    
#============================================================================================================================================================================================================================ 
    
    #==========Label
    
    label_text = ttk.Label(button_frame, text="")
    label_text.grid(row=0, column=5, columnspan=4, padx=20, pady=5, sticky="e")

#============================================================================================================================================================================================================================
    #==========Frames
    
    button_frame.columnconfigure(4, weight=1)  # Configurar coluna para expansão horizontal

    text_frame = ttk.Frame(tab6_frame)
    text_frame.grid(row=3, column=0, padx=10, pady=5, sticky="nsew")

#============================================================================================================================================================================================================================

    #==========Área de exibição de prints
    
    # Criar o PanedWindow
    paned_window = ttk.PanedWindow(text_frame, orient='vertical')
    paned_window.pack(fill="both", expand=True, pady=(15, 0))

    # Área 1
    text_widget1 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget1, weight=1)  # Definir peso igual para a área 1

    # Área 2
    text_widget2 = tk.Text(paned_window, width=20)
    paned_window.add(text_widget2, weight=1)
    
    # Configurar as barras de rolagem
    text_scrollbar1 = ttk.Scrollbar(text_widget1, command=text_widget1.yview)
    text_widget1.config(yscrollcommand=text_scrollbar1.set)
    text_scrollbar1.pack(side="right", fill="y")

    text_scrollbar2 = ttk.Scrollbar(text_widget2, command=text_widget2.yview)
    text_widget2.config(yscrollcommand=text_scrollbar2.set)
    text_scrollbar2.pack(side="right", fill="y")

    total_iterations = 10 # Número do range das buscas 10 iterações equivalem a 40 buscas.

    # Criar a barra de progresso com o estilo personalizado
    progress_bar = ttk.Progressbar(left_frame, mode='indeterminate', maximum=total_iterations)
    progress_bar.grid(row=8, column=0, padx=10, pady=5, sticky="ew")

    # Definir o tamanho da barra de progresso (comprimento em pixels)
    progress_bar.config(length=10)  # Substitua 300 pelo valor desejado
  
    def print_to_text(text):
        text_widget1.insert("end", text + "\n")
        text_widget1.see("end")
        text_widget1.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget1
    
    def print_to_text2(text):
        text_widget2.insert("end", text + "\n")
        text_widget2.see("end")
        text_widget2.update_idletasks()
        progress_bar.step(1)

    sys.stdout = text_widget2

#============================================================================================================================================================================================================================

    # Configurar o redimensionamento do tab1_frame
    tab6_frame.columnconfigure(0, weight=1)  # Expandir coluna 0
    tab6_frame.rowconfigure(3, weight=1)  # Expandir linha 3

    #Backups - last line  

#============================================================================================================================================================================================================================

#==========Interface Gráfica Principal
  
root = tk.Tk()
root.title(f"DECOM_DEMKT -- Automação Base de Dados -- usuário: {saved_userid}")
root.iconbitmap("C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/assets/ico.ico")
#root.geometry("1200x650")

# Obtém as dimensões da área de trabalho
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Define as dimensões e a posição da janela centralizada
root_width = 1200
root_height = 650
root_x = (screen_width // 2) - (root_width // 2)
root_y = (screen_height // 2) - (root_height // 2)
root.geometry(f"{root_width}x{root_height}+{root_x}+{root_y}")

# Criar o frame à esquerda
left_frame = ttk.Frame(root, width=500)
left_frame.pack(side="left", padx=10, pady=10)

# Carregar a imagem
image_path = 'C:/Automacao_DECOM&DEMKT_BNDES/Automacao baseDados/assets/Logo CO.png'
image = PhotoImage(file=image_path)

# Obter a largura e altura da imagem
image_width = image.width()
image_height = image.height()

# Ajustar a altura do frame com base na altura da janela principal
root.update()
window_height = root.winfo_height()
frame_height = min(window_height, image_height)

# Criar o frame da imagem no lado esquerdo com a mesma largura e altura da imagem
image_frame = ttk.Frame(left_frame, width=image_width, height=frame_height)
image_frame.grid(row=0, column=0, padx=10, pady=10)

image_label = ttk.Label(image_frame, image=image)
image_label.pack()

# Criar o frame à direita para os menus
right_frame = ttk.Frame(root)
right_frame.pack(side="left", fill="both", expand=True)

# Criar o widget Notebook
tab_control = ttk.Notebook(right_frame)

# Menus
tab1()
tab2()
tab3()
tab5()
tab6()

tab_control.pack(expand=1, fill='both')

root.mainloop()


