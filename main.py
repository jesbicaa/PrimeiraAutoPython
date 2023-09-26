import pandas as pd
import win32com.client as win32

# importar a tabela
tabela_vendas = pd.read_excel('Vendas.xlsx')

# vizualizar a tabela
pd.set_option('display.max_columns', None)
#print(tabela_vendas)

# faturamneto por loja
faturamento = tabela_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()
#print(faturamento)

# quantidade de produtos vendidos
quantidade = tabela_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
#print(quantidade)

# ticket médio por produto em cada loja
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})
#print(ticket_medio)

# enviar tabelas no email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'detrazprafrente13@gmail.com'
mail.Subject = 'Relatório de Vendas por Lojas'
mail.HTMLBody = f'''
<p>Prezados,</p>

<p>Segue o Relaório de Vendas pos cada Loja.</p>

<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida:</p>
{quantidade.to_html()}

<p>Ticket Médio dos Produtos em cada Loja:</p>
{ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

<p>Qualquer coisa, estou a disposição.</p>

<p>att.,</p>
<p>Jesbica</p>
'''

try:
    mail.Send()
    print('email enviado')
except:
    print('erro')