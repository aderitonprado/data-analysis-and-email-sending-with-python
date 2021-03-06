# ---- importar as libs
import pandas as pd
import win32com.client as win32
import re

# ---- importar a base de dados
tabela_vendas = pd.read_excel('Vendas.xls')

# ---- visualizar a base de dados
pd.set_option('display.max_columns', None)

## metodo 1 para filtrar a tabela de dados (escolhendo as colunas)
# tabela_vendas = tabela_vendas[['ID Loja', 'Data']]

#print(tabela_vendas)

# ---- faturamento por loja

## metodo 2 agrupar por coluna
# tabela_vendas = tabela_vendas.groupby('ID Loja').sum()
# print(tabela_vendas)

## metodo 3: unir 2 metodos: filtro e agrupar
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
print(faturamento)
print('-' * 50)

# quantidade de produtos vendidos por loja
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)

print('-' * 50)

# ticket medio por produto em cada loja
# transformando o resultado da divisão em uma tabela com o .to_frame()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

# alterar o nome da coluna 0 (zero) para Ticket médio
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
print(ticket_medio)

## Função que vai solicitar ao usuario informar o e-mail que irá receber o relatorio
def informar_email():
    # recebe o e-mail que o usuario vai digitar
    email = input("Digite o email que vai receber o relatorio: ")
    # retornar o que o usuário digitou
    return email

# Função que vai verificar se o que o usuario digitou é um e-mail válido
def verifica_email(email):
    # define um padrão de validação de email usando expressao regular
    mypattern = re.search(r'^[a-zA-Z0-9._-]+@([a-z0-9]+)(\.[a-z]{2,3})+$', email)

    # se o email for valido retorna true
    if mypattern:
        return True
    else:
        return False

# chama a função para o usuario digitar o e-mail e armazena o resultado na variavel email
email = informar_email()
# chama a função de verificar o e-mail e armazena o resultado na variavel verifica
verifica = verifica_email(email)

# testa se o email do usuario é valido, se não for valido, solicita que o usuario digite novamente
while verifica == False:
    email = informar_email()
    verifica = verifica_email(email)

print('email validado: ', email)

# pede que o usuario digite o assunto do email e armazena na variavel subject
subject = input('Digite o assunto do e-mail: ')

# inicia o objeto outlook
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)

# função que vai montar o e-mail para enviar 
def montar_email(email, subject):
    # enviar um email com o relatório
    mail.To = email
    mail.Subject = subject
    mail.HTMLBody = f'''
    <p>Prezados,</p>

    <p>Segue o Relatório de Vendas por cada Loja.</p>

    <p>Faturamento:</p>
    {faturamento.to_html(formatters={'Valor Final': 'R$ {:,.2f}'.format})}

    <p>Quantidade Vendida:</p>
    {quantidade.to_html(formatters={'Quantidade': '{:,.0f}'.format})}

    <p>Ticket Médio dos Produtos em cada Loja:</p>
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R$ {:,.2f}'.format})}

    <p>Qualquer dúvida estou à disposição.</p>

    <p>Att.,</p>
    <p>Adériton Prado</p>
    '''

# função que invoca os metodos de montar email enviar e-mail, retornar a msg de e-mail enviado
def envia_email():
    montar_email(email, subject)
    mail.Send()
    print('Email Enviado')

# executa a função de mandar o e-mail
envia_email()