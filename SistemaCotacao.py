import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry                                                                                            ##biblioteca precisa ser importada
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime
import numpy as np

requisicao = requests.get('https://economia.awesomeapi.com.br/last/USD-BRL,EUR-BRL,BTC-BRL')
dicionario_moedas = requisicao.json()                            ##transforma o dicionário json em dicionário python

lista_moedas =  list(dicionario_moedas.keys())                   ##retorna as chaves do dicionário
lista_moedas3 = [chave[:3] for chave in lista_moedas]             # Obtém as três primeiras letras de cada chave

def pegar_cotacao():
    moeda = combobox_selecionarmoeda.get()
    data_cotacao = calendario_moeda.get()
    ano = data_cotacao[-4:]
    mes = data_cotacao [3:5]
    dia = data_cotacao[:2]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}/10?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao= requisicao_moeda.json()
    valor_moeda =cotacao[0]['bid']
    label_textocotacao['text'] = f'A cotação da {moeda} no dia {data_cotacao} foi de ${valor_moeda}'


def selecionar_arquivo():
        caminho_arquivo = askopenfilename(title="Selecione o arquivo de moeda")
        var_caminhoarquivo.set(caminho_arquivo)
        if caminho_arquivo:
            label_arquivoselecionado['text'] = f"Arquivo selecionado: {caminho_arquivo}"


def atualizar_cotacoes():
    try:
        ## ler o dataframe de moedas
        df =  pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]

        ##pegar a data de início e fim das cotações
        data_inicial = calendario_datafinal.get()
        data_final = calendario_datafinal.get()
        ano_inicial = data_inicial[-4:]
        mes_inicial = data_inicial[3:5]
        dia_inicial = data_inicial[:2]

        ano_final = data_final[-4:]
        mes_final = data_final[3:5]
        dia_final = data_final[:2]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}/10" \
                   f"?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}"

            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao['timestamp'])
                bid = float(cotacao['bid'])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime('%d/%m/%Y')
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel("Teste.xlsx")
        label_atualizarcotacoes['text'] = "Arquivo Atualizado com Sucesso"

    except:
        label_atualizarcotacoes['text'] = "Selecione um arquivo Excel no Formato Correto"


janela = tk.Tk()
janela.title("Ferramenta de Cotação de Moedas")

label_cotacaomoeda = tk.Label(text= "Cotação de 1 Moeda Específica", borderwidth=2, relief="solid")
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky="NSEW", columnspan=3)                              ##padx - distância do label que estamps criando

label_selecionarmoeda = tk.Label(text= "Selecionar Moeda",  anchor='e')
label_selecionarmoeda.grid(row=1, column=0, padx=10, pady=10, sticky="NSEW", columnspan=2)

combobox_selecionarmoeda = ttk.Combobox(values=lista_moedas3)
combobox_selecionarmoeda.grid(row=1, column=2, padx=10, pady=10)

label_selecionardia = tk.Label(text= "Selecione o dia que deseja pegar a cotação", anchor='e')
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky="NSEW", columnspan=2)

calendario_moeda = DateEntry(year=2023, locale='pt_br')
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky="NSEW")

label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0,columnspan=2, padx=10, pady=10, sticky= "NSEW" )

botao_pegarcotacao =  tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky="NSEW")

#Cotação de várias moedas

label_cotacaovariasmoedas = tk.Label(text= "Cotação de Múltiplas Moedas", borderwidth=2, relief="solid")
label_cotacaovariasmoedas.grid(row=4, column=0, padx=10, pady=10, sticky="NSEW", columnspan=3)

label_selecionararquivo = tk.Label(text="Selecione um arquivo em Excel com as Moedas na Coluna A:")
label_selecionararquivo.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="NSEW")               ##columnspan eu passo a quantidade de colunas

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(text="Clique para Selecionar", command=selecionar_arquivo)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky="NSEW")

label_arquivoselecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor='e')
label_arquivoselecionado.grid(row=6, column=0, columnspan=3, padx=10, pady=10, sticky="NSEW")

label_datainicial = tk.Label(text="Data Inicial", anchor="e")
label_datafinal =  tk.Label(text="Data Final", anchor="e")
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky="NSEW")
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky="NSEW")

calendario_datainicial = DateEntry(year=2023, locale='pt_br')
calendario_datafinal = DateEntry(year=2023, locale='pt_br')
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky="NSEW")
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky="NSEW")

botao_atualizarcotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=9, column=0,  padx=10, pady=10, sticky="NSEW")

label_atualizarcotacoes = tk.Label(text="")
label_atualizarcotacoes.grid(row=9, column=1, columnspan=2,  padx=10, pady=10, sticky="NSEW")

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, sticky="NSEW")


janela.mainloop()