import pandas as pd
import win32com.client as win32
tabela_vendas = pd.read_excel('vendas.xlsx')

# visualizar a base de dados com pandas
pd.set_option('display.max_columns', None,)

faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()
print(faturamento)
quantidade = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()
print(quantidade)
#tikcet meio
ticket_medio = (faturamento['Valor Final'] / quantidade ['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Medio'})
print(ticket_medio)
#integraçao
outlook = win32.Dispatch('outlook.application')
# criar um email
email: object = outlook.CreateItem(0)
# configurar as informações do seu e-mail
email.To = "ychsfizy2@gmail.com"
email.Subject = "relatorio de vendas"
email.HTMLBody = f'''
<p>presados,</p> 
<p>segue a tabela com todas infos das vendas de todas as lojas</p>
<p>Faturamento:</p>
{faturamento.to_html(formatters={'Valor Final':'R${:,.2f}'.format})}
<p>Quantidades Vendidas</p>
<p>Ticket Medio dos Pacotes de cada Lojas</p>
{ticket_medio.to_html(formatters= {'Ticket Medio':'R${:,.2f}'.format})}
<p>Qualquer coisa entrarei em contato</p>
'''
email.Send()
print("Email Enviado")