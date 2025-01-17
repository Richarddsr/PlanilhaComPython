import pandas as pd

#IMPORTAR A BASE DE DADOS
tabela_de_vendas = pd.read_excel('Vendas.xlsx')



#visualizar base de dados
pd.set_option('display.max_columns', None)

#fTURAMENTO POR LOJA
faturamento = tabela_de_vendas[['ID Loja', 'Valor Final']].groupby('ID Loja').sum()

#QUNATIDADE DE PRODUTOS VENDIDOS POR LJA
quantidade = tabela_de_vendas[['ID Loja', 'Quantidade']].groupby('ID Loja').sum()
print(quantidade)
print('-' * 50)
# TICKETS MEDIOS POR produto em cada LOJKA
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()

# enviar emails e relatorios
import win32com.client as win32
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'richardrodrues10@gmail.com'
mail.Subject = 'Relatorio de vendas por loja'
mail.HTMLBody = f'''
<p>Prezados, bom dia</p>

<p>Segue relatorio de vendas</p>

<p>Faturamento</p>
{faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Quantidade Vendida</p>
{quantidade.to_html()}

<p>Ticket Medio</p>
{ticket_medio.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}

<p>Qualquer duvida estou a disposicao</p>

<p>Atenciosamente</p>
<p>Guilherme</p>
'''

mail.Send()
print("Email Enviado")