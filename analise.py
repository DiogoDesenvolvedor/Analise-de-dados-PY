import pandas as pd
import win32com.client as win32
vendas_df = pd.read_excel("Vendas.xlsx")

faturamento = vendas_df[["ID Loja", "Valor Final"]].groupby('ID Loja').sum()
quantidade = vendas_df[["ID Loja", "Quantidade"]].groupby('ID Loja').sum()
ticket_medio = (faturamento['Valor Final'] / quantidade['Quantidade']).to_frame()
ticket_medio = ticket_medio.rename(columns={0: 'Ticket Médio'})

outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
mail.To = 'lunajv@hotmail.com'
mail.Subject = 'Relatório de Vendas por Loja'
mail.HTMLBody = f'''
    <p> Relatorio de vendas da semana </p> 
    
    <p> Faturamento por loja. </p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <p> Quantidade vendida por loja. </p>
    {quantidade.to_html()}
    <p> Tickt medio por vendas </p>    
    {ticket_medio.to_html(formatters={'Ticket Médio': 'R${:,.2f}'.format})}

    <p> Relatorio enviado por Diogo </p>
    
    <p> Ass: Diogo Quintana Luna </p>    

'''
mail.Send()