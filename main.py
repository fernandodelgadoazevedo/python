# Importa a biblioteca pandas e a renomeia para pd
import pandas as pd

# Importa a classe Client da biblioteca twilio.rest
from twilio.rest import Client

# Importa as credenciais de autenticação da conta do Twilio do arquivo credentials
from credentials import account_sid, auth_token

# Importa a classe Tk e a função filedialog da biblioteca tkinter
from tkinter import Tk, filedialog

# Cria uma instância da classe Client do Twilio com as credenciais fornecidas
client = Client(account_sid, auth_token)


# Define uma função para gerar uma lista de meses entre os índices de início e fim
def gerar_lista_meses(inicio, fim):
    meses = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho',
             'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro']
    return meses[inicio - 1:fim]


# Chama a função para gerar a lista de meses de janeiro a dezembro
lista_meses = gerar_lista_meses(1, 12)

# Imprime a lista de meses
print(lista_meses)

# Cria outra instância da classe Client do Twilio com as mesmas credenciais (redundante)
client = Client(account_sid, auth_token)


# Define uma função para selecionar arquivos Excel
def selecionar_arquivos_excel():
    root = Tk()
    root.withdraw()
    # Abre uma janela para selecionar arquivos Excel e retorna a lista de arquivos selecionados
    arquivos = filedialog.askopenfilenames(title="Selecione os arquivos Excel", filetypes=[("Excel Files", "*.xlsx")])
    return arquivos


# Chama a função para selecionar os arquivos Excel
arquivos_excel = selecionar_arquivos_excel()

# Itera sobre a lista de meses e arquivos Excel simultaneamente
for mes, arquivo_excel in zip(lista_meses, arquivos_excel):
    # Lê o arquivo Excel e cria um DataFrame usando a biblioteca pandas
    tabela_vendas = pd.read_excel(arquivo_excel)

    # Verifica se existe alguma venda maior que 55000 no DataFrame
    if (tabela_vendas['Vendas'] > 55000).any():
        # Obtém o nome do vendedor que bateu a meta de vendas
        vendedor = tabela_vendas.loc[tabela_vendas['Vendas'] > 55000, 'Vendedor'].values[0]
        # Obtém o valor total de vendas do vendedor que bateu a meta
        vendas = tabela_vendas.loc[tabela_vendas['Vendas'] > 55000, 'Vendas'].values[0]
        # Imprime uma mensagem indicando que o vendedor bateu a meta de vendas
        print(f'No mês de {mes} O vendedor: {vendedor}, bateu a meta de vendas no total de vendas {vendas}')

        # Envia uma mensagem via Twilio com as informações do vendedor que bateu a meta
        message = client.messages.create(
            body=f'No mês de {mes}, o vendedor {vendedor} bateu a meta de vendas, totalizando vendas de {vendas}',
            from_='+13343104434',
            to='+5511983078800'
        )

        # Imprime o SID da mensagem enviada pelo Twilio
        print(message.sid)
