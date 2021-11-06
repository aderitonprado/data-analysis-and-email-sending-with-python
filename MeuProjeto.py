# ---- importar as libs
import pandas as pd
import win32com.client as win32

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

# enviar um e-mail com o relatorio
# enviar um email com o relatório
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'your-email@gmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
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

mail.Send()

print('Email Enviado')