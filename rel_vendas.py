import email.message
import smtplib  # função que permite o envio de e-mail
import pandas as pd

# Importar base de dados

tabela_vendas = pd.read_excel('Vendas.xlsx')

#Visualizar base de dados

# 1 - processo para validar se todas as colunas foram lidas pelo pandas.
pd.set_option('display.max_columns',None)
#print(tabela_vendas) ---OBS: Tirar o comentário para realizar a leitura

# 1.1 - Filtrar colunas expecificas (informar a variável = [colocar colchestes pois é lista[colocar
#  o nome da coluna solicitada (deve ser igual o nome que está no arquivo excel) entre '']]
# ex: tabela_vendas[['id','valor total']]

# Faturamento por loja (criar uma variável para receber o filtro)
faturamento = tabela_vendas[['ID Loja','Valor Final']].groupby('ID Loja').sum()

#print(faturamento) ---OBS: Tirar o comentário para realizar a leitura

# Calcular quantidade e itens vendidos por loja
qntProd = tabela_vendas[['ID Loja','Quantidade']].groupby('ID Loja').sum()

#print(qntProd)

print('-'*50) # Separar tabelas visualmente
#Calcular o Ticket médio por produto em cada loja (É o resultado do valor do faturamento / quantidade)
ticketmedio = (faturamento['Valor Final']/qntProd['Quantidade']).to_frame() # OBS: to_frame-> Transforma os dados em tabela
ticketmedio = ticketmedio.rename(columns={0: 'Ticket Médio'})
#print(ticketmedio)

# Enviar e-mail com o relatório
def enviar_email(): # função para enviar e-mail

    corpo_email = f"""
    <p>Boa tarde, Caroline!</p>
    <p>Segue abaixo o relatório de Vendas.</p>
    <p><b> - Faturamento</b></p>
    {faturamento.to_html(formatters={'Valor Final': 'R${:,.2f}'.format})}
    <p><b> - Total de itens vendidos por loja</b></p>
    {qntProd.to_html()}
    <p><b> - Ticket Médio de Produto por Loja</b></p>
    {ticketmedio.to_html()}
    <p>Atenciosamente!</p>"""

# FUNÇÃO PARA ENVIAR E-MAIL (GMAIL)
    msg = email.message.Message()
    msg['Subject'] = "Relatório de Vendas"
    msg['From'] = 'carolsousaaraujo078@gmail.com'
    msg['To'] = 'caroline.rsc@outlook.com'
    password = '123' # inserir a senha gerada pela conta de e-mail
    msg.add_header('Content-Type','text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    s.login(msg['From'],password)
    s.sendmail(msg['From'],[msg['To']],msg.as_string().encode('utf-8'))
    print('Email enviado!')
print(enviar_email())

