import pandas as pd
import plotly.express as px
from babel.numbers import format_currency

#importando dados para o programa
tb_vendas = pd.read_excel("Vendas\Vendas_Base_de_Dados.xlsx")

#calculando produto mais vendido (quantidade)
tb_qtd = tb_vendas.groupby('Produto').sum()
tb_qtd = tb_qtd[['Quantidade']].sort_values(by='Quantidade', ascending=False)
print(tb_qtd)

#calculando produto mais vendido (faturamento)
tb_vendas['Faturamento'] = tb_vendas['Quantidade'] * tb_vendas['Valor Unitário']
tb_fatu = tb_vendas.groupby('Produto').sum()
tb_fatu = tb_fatu[['Faturamento']].sort_values(by='Faturamento', ascending=False)
print(tb_fatu)

#calculando a loja que mais vendeu (faturamento)
tb_Loja = tb_vendas.groupby('Loja').sum()
tb_Loja = tb_Loja[['Faturamento']].sort_values(by='Faturamento', ascending=False)
print(tb_Loja)

#ticket medio por loja
tb_vendas['Ticket Medio'] = tb_vendas['Valor Unitário']
tb_tic = tb_vendas.groupby('Loja').mean('Ticket Medio')
tb_tic = tb_tic[['Ticket Medio']]
print(tb_tic)

#formatação monetárias
tb_qtd_format = pd.DataFrame(tb_qtd['Quantidade'].apply(lambda x: format_currency(x, 'BRL', locale='pt_BR')))
tb_fatu_format = pd.DataFrame(tb_fatu['Faturamento'].apply(lambda x: format_currency(x, 'BRL', locale='pt_BR')))
tb_Loja_format = pd.DataFrame(tb_Loja['Faturamento'].apply(lambda x: format_currency(x, 'BRL', locale='pt_BR')))
tb_tic_format = pd.DataFrame(tb_tic['Ticket Medio'].apply(lambda x: format_currency(x, 'BRL', locale='pt_BR')))

#criando grafico de colunas (faturamento)

grafico = px.bar(tb_fatu, y='Faturamento', x=tb_fatu.index)
grafico.show()

import smtplib 
import email.message as em

#configurando o envio de email (gmail)
def enviar():
    corpo = f"""
    <p> Prezados, tudo bem? </p>
    <p>Segue o relatorio de vendas do mês</p>

    <p><b>Faturamento por loja: </b></p>
    <p>{tb_Loja_format.to_html()}</p>

    <p><b>Quantidade vendida por produto: </b></p>
    <p>{tb_fatu_format.to_html()}</p>

    <p><b>Faturamento por produto: </b></p>
    <p>{tb_qtd_format.to_html()}</p>

    <p><b>Ticket médio por loja: </b></p>
    <p>{tb_tic_format.to_html()}</p>

    <p>Qualquer dúvida, estou à disposição</p>
    <p>Atenciosamente</p>
    <p>Gustavo Grancieri</p>"""

    #configurações para envio
    msg = em.Message()
    msg['Subject'] = 'Relatório'#assunto do email
    msg['from'] = 'gustavograncieri2004@gmail.com'#email que vai enviar o email
    msg['to'] = 'gustavograncieri@outlook.com'#email que ira receber
    password = 'vrbt tfpp updo yuup'#senha do email que vai enviar
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo)
    #configurando servidor do Gmail
    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    #login do email
    s.login(msg['from'], password)
    s.sendmail(['from'], [msg['to']], msg.as_string().encode('utf-8'))
    print('Email enviado com sucesso')
enviar()