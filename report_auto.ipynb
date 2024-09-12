
#--------------------( Iportar bibliotecas )--------------------#

import pandas as pd
import pyodbc
import pymssql
import datetime
import smtplib
from smtplib import SMTP
import json
import re
import win32com.client as win32
import warnings
from pandas.core.common import SettingWithCopyWarning
warnings.filterwarnings("ignore", category=SettingWithCopyWarning)
import schedule
import time

##--------------------( Configurar o modo como eu vejo as tabelas )--------------------#

pd.set_option('display.max_columns', None)
pd.set_option('display.width', None)

#--------------------( Acessar o SQL para coletar os dados )--------------------#

#def ativar_atualizacao():
conn = pymssql.connect(server='', database='', user='', password='')

query = """

"""

tabela = pd.read_sql_query(query, conn)

#--------------------( Criar o Dataframe(tabela) e arrumar as informações )--------------------#

# Transformar em formato de tabela
df = pd.DataFrame(tabela)

# Criar coluna com data e hora atual
df['Hora atual'] = datetime.datetime.now()

# Criar coluna resultado sendo ele como tempo que finaliza a prévia que foi informada ao cliente
filtro = (df['Status ação nome'] == 'Chegada') & (df['Data Fim'].isnull())
df['ResultadoTempoChegada'] = ''
df.loc[filtro, 'ResultadoTempoChegada'] = (df.loc[filtro, 'Hora atual'].apply(lambda x: x.timestamp()) - df.loc[filtro, 'tempoFimPrevia'].apply(lambda x: x.timestamp())) / 60

# Formatar todos os campos de data em um único modelo
df['Data abertura da OS'] = pd.to_datetime(df['Data abertura da OS'])
df['Data Inicio'] = pd.to_datetime(df['Data Inicio'])
df['Data Fim'] = pd.to_datetime(df['Data Fim'])
df['Hora atual'] = pd.to_datetime(df['Hora atual'])
df['tempoDuracao'] = pd.to_datetime(df['tempoDuracao'])
df['tempoFimPrevia'] = pd.to_datetime(df['tempoFimPrevia'])

# Criar tempo de Experiência até o momento
df['tempoExperienciaAgora'] = df['Hora atual'] - df['Data abertura da OS']

# Tranformar em modelo datetime
tempo_timedelta = pd.to_timedelta(df['tempoExperienciaAgora'])
data_referencia = pd.Timestamp.now().normalize()
tempo_datetime = data_referencia + tempo_timedelta

# Criar um campo com tempo de atendimento ultrapassado, ou seja, vai mostrar em minutos, quanto tempos estamos a partir do final do acionamento
x = df['ResultadoTempoChegada']
def convert_to_hms(x):
    if x:
        hours = int(x // 3600)
        minutes = int((x % 3600) // 60)
        seconds = int(x % 60)
        return f"{hours:02d}:{minutes:02d}:{seconds:02d}"
    else:
        return ""

# Lançar informações nos campos que estão em branco, pois sem dados, não consigo fazer contas
df['ResultadoTempoChegada'] = pd.to_numeric(df['ResultadoTempoChegada'], errors='coerce')

# Criar o campo status que vai mostrar se estamos dentro da prévia ou fora
df['Status'] = df['ResultadoTempoChegada'].apply(lambda x: 'Fora da prévia' if x > 0 else 'Finalizado')

def formatar_tempo(tempo_minutos):
    horas = tempo_minutos // 60
    minutos = tempo_minutos % 60
    segundos = int((tempo_minutos - int(tempo_minutos)) * 60)
    return '{:02d}:{:02d}:{:02d}'.format(int(horas), int(minutos), segundos)


# Transformar o campo resultado que estava em minutos no modelo de dia, hora, minutos e segundos
import numpy as np
df['tempoForaPrevia'] = np.where(df['ResultadoTempoChegada'] > 0, df['ResultadoTempoChegada'].apply(lambda x: str(pd.to_timedelta(x, unit='m'))), 'Dentro da prévia')

# Tranformar a coluna tempoForaPrevia em datetime, conforme os demais campos.
df['tempoForaPrevia'] = df['tempoForaPrevia'].apply(lambda x: pd.Timestamp.now().normalize() + x if isinstance(x, pd.Timedelta) else x)

## Trocar valores em branco, com não finalizado
df['Data Fim'].fillna(value='Em andamento', inplace=True)

# Acionamento

# Função para calcular a duração
def calcular_duracao(row):
    if row['Status ação nome'] == 'Em Acionamento' and row['Data Fim'] == 'Em andamento':
        return row['Hora atual'] - row['Data Inicio']
    else:
        return np.nan

df['duracaoAcionamento'] = df.apply(calcular_duracao, axis=1)

# Tranformar a coluna Duração Acionamento em datetime
df['duracaoAcionamento'] = pd.to_datetime('1970-01-01') + df['duracaoAcionamento']

## Desejo visualizar os serviços que estão em andamento. Vou excluir todos que já foram finalizados
tabela = df.loc[df['Data Fim'] == 'Em andamento']

# Serviços somente com status em Acionamento
tabelaAcionamento = tabela.loc[tabela['Status ação nome'] == 'Em Acionamento']

# Serviços com status somente Chegada
tabelaChegada = tabela.loc[tabela['Status ação nome'] == 'Chegada']

# Tabela em acionamento
tabelaEmAcionamentoFinal = tabelaAcionamento[['Ordem Serviço', 'Data Inicio', 'duracaoAcionamento', 'tipoServiço', 'Cidade', 'Estado']]

# Tabela Chegada
tabelaChegadaFinal = tabelaChegada[['Ordem Serviço','Data Inicio','MSP_SLA_LIMITE_NEGOCIADO','tempoFimPrevia','Status','duracaoMinutos','tempoExperienciaAgora', 'Hora atual','tipoServiço', 'Cidade', 'Estado']]

tabelaChegadaFinal = tabelaChegadaFinal.loc[tabela['Status'] == 'Fora da prévia']

tabelaChegadaFinal['tempoExcedido'] = tabelaChegadaFinal['Hora atual'] - tabelaChegadaFinal['tempoFimPrevia']

# Tranformar em modelo datetime - Tempo fora da prévia
tempo_timedelta = pd.to_timedelta(tabelaChegadaFinal['tempoExcedido'])
data_referencia = pd.Timestamp.now().normalize()
tempo_datetime = data_referencia + tempo_timedelta

tabelaEmAcionamentoFinal['duracaoMinutos'] = tabelaEmAcionamentoFinal['duracaoAcionamento'].apply(lambda x: int(x.timestamp()))

# Criar tempo excedido depois da prévia que o prestador informou e deixar como número inteiro (minutos) 
tabelaChegadaFinal['tempoExcedido'] = round(tabelaChegadaFinal['tempoExcedido'].dt.total_seconds() / 60).astype(int)

## Vou formatar todo o processo para a tabela do Acionamento e depois para Chegada - De forma separada

## Acionamento:
# Criar uma coluna onde consigo saber quem sinalizar quando passar de um determinado tempo
tabelaEmAcionamentoFinal['Ação'] = ''

tabelaEmAcionamentoFinal.loc[tabelaEmAcionamentoFinal['duracaoMinutos'] <= 1800,'Ação'] = 'Sem sinalização'
tabelaEmAcionamentoFinal.loc[(tabelaEmAcionamentoFinal['duracaoMinutos'] > 1800) & (tabelaEmAcionamentoFinal['duracaoMinutos'] <= 3600), 'Ação'] = 'Coordenadores'
tabelaEmAcionamentoFinal.loc[(tabelaEmAcionamentoFinal['duracaoMinutos'] > 3600) & (tabelaEmAcionamentoFinal['duracaoMinutos'] <= 5400), 'Ação'] = 'Gerentes'
tabelaEmAcionamentoFinal.loc[tabelaEmAcionamentoFinal['duracaoMinutos'] > 5400,'Ação'] = 'Superintendente'

# Criar uma coluna somente com a hora do inicio do acionamento
tabelaEmAcionamentoFinal['Hora Inicio'] = tabelaEmAcionamentoFinal['Data Inicio']

## Transformar todas as colunas que contenha data, hora e modelo em somente hora

# Data Inicio
tabelaEmAcionamentoFinal['Data Inicio'] = pd.to_datetime(tabelaEmAcionamentoFinal['Data Inicio']) 

lista = []
for i in tabelaEmAcionamentoFinal['Data Inicio']:
    x = i.strftime("%d/%m")
    lista.append(x)

tabelaEmAcionamentoFinal['Data Inicio'] = lista

# Data Inicio
tabelaEmAcionamentoFinal['Hora Inicio'] = pd.to_datetime(tabelaEmAcionamentoFinal['Hora Inicio']) 

lista = []
for i in tabelaEmAcionamentoFinal['Hora Inicio']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaEmAcionamentoFinal['Hora Inicio'] = lista

# Duração do acionamento
tabelaEmAcionamentoFinal['duracaoAcionamento'] = pd.to_datetime(tabelaEmAcionamentoFinal['duracaoAcionamento']) 

lista = []
for i in tabelaEmAcionamentoFinal['duracaoAcionamento']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaEmAcionamentoFinal['duracaoAcionamento'] = lista

# Serviços de 30 à 45 minutos em acionamento - Avisar o Coordenador
filtro1 = (tabelaEmAcionamentoFinal['duracaoMinutos'] > 1800) & (tabelaEmAcionamentoFinal['duracaoMinutos'] <= 2700)
tabelaAcionamento0 = tabelaEmAcionamentoFinal.loc[filtro1]
tabelaAcionamento0 = tabelaAcionamento0[['Ordem Serviço', 'Data Inicio','Hora Inicio', 'duracaoAcionamento','tipoServiço', 'Cidade', 'Estado']]

# Serviços de 45 à 60 minutos em acionamento - Avisar o Coordenador
filtro1 = (tabelaEmAcionamentoFinal['duracaoMinutos'] > 2700) & (tabelaEmAcionamentoFinal['duracaoMinutos'] <= 3600)
tabelaAcionamento1 = tabelaEmAcionamentoFinal.loc[filtro1]
tabelaAcionamento1 = tabelaAcionamento1[['Ordem Serviço', 'Data Inicio','Hora Inicio', 'duracaoAcionamento','tipoServiço', 'Cidade', 'Estado']]

# Serviços de 60 à 90 minutos em acionamento - Avisar o Gerente
filtro2 = (tabelaEmAcionamentoFinal['duracaoMinutos'] > 3600) & (tabelaEmAcionamentoFinal['duracaoMinutos'] <= 5400)
tabelaAcionamento2 = tabelaEmAcionamentoFinal.loc[filtro2]
tabelaAcionamento2 = tabelaAcionamento2[['Ordem Serviço', 'Data Inicio','Hora Inicio', 'duracaoAcionamento','tipoServiço', 'Cidade', 'Estado']]

# Serviços acima de 90 minutos - Avisar o Superintendente
filtro3 = (tabelaEmAcionamentoFinal['duracaoMinutos'] > 5400)
tabelaAcionamento3 = tabelaEmAcionamentoFinal.loc[filtro3]
tabelaAcionamento3 = tabelaAcionamento3[['Ordem Serviço', 'Data Inicio','Hora Inicio', 'duracaoAcionamento','tipoServiço', 'Cidade', 'Estado']]


# Renomear a tabela do SLA
tabelaChegadaFinal = tabelaChegadaFinal.rename(columns={'MSP_SLA_LIMITE_NEGOCIADO': 'SLA'})

# Criar uma coluna somente com a hora do inicio do acionamento
tabelaChegadaFinal['Hora Inicio'] = tabelaChegadaFinal['Data Inicio']

tabelaChegadaFinal['ultimaAtualização'] = tabelaChegadaFinal['Hora atual']

# Prazo para finalizar a prévia informada pelo prestador
tabelaChegadaFinal['ultimaAtualização'] = pd.to_datetime(tabelaChegadaFinal['ultimaAtualização'])

# Tranformei o tempo exedido que estava em minutos para visualização de horas
tabelaChegadaFinal['tempoForaDaPrévia'] = pd.to_datetime(tabelaChegadaFinal['tempoExcedido'], unit='m').dt.strftime('%H:%M:%S')


## Transformar todas as colunas que contenha data, hora e modelo em somente hora
# Data Inicio
tabelaChegadaFinal['Data Inicio'] = pd.to_datetime(tabelaChegadaFinal['Data Inicio']) 

lista = []
for i in tabelaChegadaFinal['Data Inicio']:
    x = i.strftime("%d/%m")
    lista.append(x)

tabelaChegadaFinal['Data Inicio'] = lista

# Hora Inicio
tabelaChegadaFinal['Hora Inicio'] = pd.to_datetime(tabelaChegadaFinal['Hora Inicio']) 

lista = []
for i in tabelaChegadaFinal['Hora Inicio']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaChegadaFinal['Hora Inicio'] = lista

# Prazo para finalizar a prévia informada pelo prestador
tabelaChegadaFinal['tempoFimPrevia'] = pd.to_datetime(tabelaChegadaFinal['tempoFimPrevia']) 

lista = []
for i in tabelaChegadaFinal['tempoFimPrevia']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaChegadaFinal['tempoFimPrevia'] = lista

lista = []
for i in tabelaChegadaFinal['ultimaAtualização']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaChegadaFinal['ultimaAtualização'] = lista

# Tranformar a coluna tempo de experiência em datetime
tabelaChegadaFinal['tempoExperienciaAgora'] = tabelaChegadaFinal['tempoExperienciaAgora'].apply(lambda x: pd.Timestamp.now().normalize() + x if isinstance(x, pd.Timedelta) else x)

# Converter para somente hora
lista = []
for i in tabelaChegadaFinal['tempoExperienciaAgora']:
    x = i.strftime("%H:%M:%S")
    lista.append(x)

tabelaChegadaFinal['tempoExperienciaAgora'] = lista

# Serviços de 30 à 60 acima do tempo de chegada - Avisar o Coordenador
filtro4 = (tabelaChegadaFinal['tempoExcedido'] > 30) & (tabelaChegadaFinal['tempoExcedido'] <= 45)
tabelaChegada0 = tabelaChegadaFinal.loc[filtro4]
tabelaChegada0 = tabelaChegada0[['Ordem Serviço','Data Inicio','Hora Inicio','SLA', 'tempoFimPrevia','ultimaAtualização','tempoForaDaPrévia','tempoExperienciaAgora','tipoServiço', 'Cidade', 'Estado']]


# Serviços de 30 à 60 acima do tempo de chegada - Avisar o Coordenador
filtro4 = (tabelaChegadaFinal['tempoExcedido'] > 45) & (tabelaChegadaFinal['tempoExcedido'] <= 60)
tabelaChegada1 = tabelaChegadaFinal.loc[filtro4]
tabelaChegada1 = tabelaChegada1[['Ordem Serviço','Data Inicio','Hora Inicio','SLA', 'tempoFimPrevia','ultimaAtualização','tempoForaDaPrévia','tempoExperienciaAgora','tipoServiço', 'Cidade', 'Estado']]

# Serviços de 60 à 90 minutos em acionamento - Avisar o Gerente
filtro5 = (tabelaChegadaFinal['tempoExcedido'] > 60) & (tabelaChegadaFinal['tempoExcedido'] <= 90)
tabelaChegada2 = tabelaChegadaFinal.loc[filtro5]
tabelaChegada2 = tabelaChegada2[['Ordem Serviço','Data Inicio','Hora Inicio','SLA', 'tempoFimPrevia','ultimaAtualização','tempoForaDaPrévia','tempoExperienciaAgora','tipoServiço', 'Cidade', 'Estado']]

# Serviços acima de 90 minutos - Avisar o Superintendente
filtro6 = (tabelaChegadaFinal['tempoExcedido'] > 90)
tabelaChegada3 = tabelaChegadaFinal.loc[filtro6]
tabelaChegada3 = tabelaChegada3[['Ordem Serviço','Data Inicio','Hora Inicio','SLA', 'tempoFimPrevia','ultimaAtualização','tempoForaDaPrévia','tempoExperienciaAgora','tipoServiço', 'Cidade', 'Estado']]

# Aqui vou montar uma tabela que mostra a quantidade de serviços por faixa de tempo em momento de acionamento

# Tabela de acionamento
# Crie um dicionário com as opções e os resultados
resultadosAcionamento = {
    'Entre 30 e 45 minutos': tabelaAcionamento0['Ordem Serviço'].count(),
    'Entre 45 e 60 minutos': tabelaAcionamento1['Ordem Serviço'].count(),
    'Entre 60 e 90 minutos': tabelaAcionamento2['Ordem Serviço'].count(),
    'Acima de 90 minutos': tabelaAcionamento3['Ordem Serviço'].count()
}

# Crie a tabela a partir do dicionário
resultadosAcionamento = pd.DataFrame.from_dict(resultadosAcionamento, orient='index', columns=['Quantidade'])

# Inserir nome na coluna Index
resultadosAcionamento.rename_axis('Range',axis=0, inplace = True)

# Tabela de entrega

# Crie um dicionário com as opções e os resultados
resultadoChegada = {
    'Entre 30 e 45 minutos': tabelaChegada0['Ordem Serviço'].count(),
    'Entre 45 e 60 minutos': tabelaChegada1['Ordem Serviço'].count(),
    'Entre 60 e 90 minutos': tabelaChegada2['Ordem Serviço'].count(),
    'Acima de 90 minutos': tabelaChegada3['Ordem Serviço'].count()
}

# Crie a tabela a partir do dicionário
resultadoChegada = pd.DataFrame.from_dict(resultadoChegada, orient='index', columns=['Quantidade'])

# Inserir nome na coluna Index
resultadoChegada.rename_axis('Range',axis=0, inplace = True)
#--------------------( Processo para coletar dados da experiência do cliente)--------------------#

queryExperiencia = """
"""

experiencia = pd.read_sql_query(queryExperiencia, conn)

tmp = experiencia['Inicio acionamento'] - experiencia['Data abertura OS'] # Criação até acionamento, sem ninguém atuar
tmp.mean()

tmac = experiencia['Data Fim'] - experiencia['Inicio acionamento'] # Tempo em acionamento
tmac.mean()

queryExperienciaChegada = """

"""

experienciaChegada = pd.read_sql_query(queryExperienciaChegada, conn)

tmc = experienciaChegada['Data Fim'] - experienciaChegada['Inicio acionamento']

#tmc.mean()

tempototal = tmp + tmac + tmc

#tempototal.mean()

tempoExperiencia = tempototal.mean()

# Crie um objeto timedelta com o valor médio
tempo_total = pd.Series([pd.Timedelta(tempoExperiencia)])

# Defina uma função que formata um valor timedelta no formato de hora
def formatar_hora(delta):
    data_hora_atual = pd.to_datetime('now').replace(hour=0, minute=0, second=0, microsecond=0)
    data_hora_final = data_hora_atual + delta
    return data_hora_final.strftime('%H:%M:%S')

# Use o método apply() para aplicar a função formatar_hora em cada valor da Series
tempoExperienciaFinal = tempo_total.apply(formatar_hora)[0]
print(tempoExperienciaFinal)

#--------------------( Configurar o processo de 30 a 60 minutos de atuação )--------------------#

## Criar html para enviar no e-mail
# Adiciona a tabela em HTML ao corpo do e-mail
html = "<table border='1' cellpadding='5'>\n"
html += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>duracaoAcionamento</th><th>tipoServiço</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaAcionamento1.iterrows():
    html += "<tr>"
    html += f"<td>{row['Ordem Serviço']}</td>"
    html += f"<td>{row['Data Inicio']}</td>"
    html += f"<td>{row['Hora Inicio']}</td>"
    html += f"<td>{row['duracaoAcionamento']}</td>"
    html += f"<td>{row['tipoServiço']}</td>"
    html += f"<td>{row['Cidade']}</td>"
    html += f"<td>{row['Estado']}</td>"
    html += "</tr>\n"
html += "</table>"

# Adiciona a tabela em HTML ao corpo do e-mail
html1 = "<table border='1' cellpadding='5'>\n"
html1 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>SLA</th><th>tempoFimPrevia</th><th>ultimaAtualização</th><th>tempoForaDaPrévia</th><th>tipoServiço</th><th>tempoExperienciaAgora</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaChegada1.iterrows():
    html1 += "<tr>"
    html1 += f"<td>{row['Ordem Serviço']}</td>"
    html1 += f"<td>{row['Data Inicio']}</td>"
    html1 += f"<td>{row['Hora Inicio']}</td>"
    html1 += f"<td>{row['SLA']}</td>"
    html1 += f"<td>{row['tempoFimPrevia']}</td>"
    html1 += f"<td>{row['ultimaAtualização']}</td>"
    html1 += f"<td>{row['tempoForaDaPrévia']}</td>"
    html1 += f"<td>{row['tipoServiço']}</td>"
    html1 += f"<td>{row['tempoExperienciaAgora']}</td>"
    html1 += f"<td>{row['Cidade']}</td>"
    html1 += f"<td>{row['Estado']}</td>"
    html1 += "</tr>\n"
html1 += "</table>"

# Vou saber a quantidade de serviços em momento de acionamento
# Adiciona a tabela em HTML ao corpo do e-mail
html2 = "<table border='1' cellpadding='5'>\n"
html2 += "<tr><th>Range</th><th>Quantidade</th></tr>\n"
for k, row in enumerate(resultadosAcionamento['Quantidade']):
    html2 += "<tr>"
    df = resultadosAcionamento.index[k]
    html2 += f"<td>{df}</td>"
    html2 += f"<td>{row}</td>"
    html2 += "</tr>\n"
html2 += "</table>"

# Vou saber a quantidade de serviços em momento de entrega
# Adiciona a tabela em HTML ao corpo do e-mail
html3 = "<table border='1' cellpadding='5'>\n"
html3 += "<tr><th>Range</th><th>Quantidade</th></tr>\n"
for k, row2 in enumerate(resultadoChegada['Quantidade']):
    html3 += "<tr>"
    df2 = resultadoChegada.index[k]
    html3 += f"<td>{df2}</td>"
    html3 += f"<td>{row2}</td>"
    html3 += "</tr>\n"
html3 += "</table>"

## Processo de 60 à 90 minutos

# Adiciona a tabela em HTML ao corpo do e-mail
html4 = "<table border='1' cellpadding='5'>\n"
html4 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>duracaoAcionamento</th><th>tipoServiço</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaAcionamento2.iterrows():
    html4 += "<tr>"
    html4 += f"<td>{row['Ordem Serviço']}</td>"
    html4 += f"<td>{row['Data Inicio']}</td>"
    html4 += f"<td>{row['Hora Inicio']}</td>"
    html4 += f"<td>{row['duracaoAcionamento']}</td>"
    html4 += f"<td>{row['tipoServiço']}</td>"
    html4 += f"<td>{row['Cidade']}</td>"
    html4 += f"<td>{row['Estado']}</td>"
    html4 += "</tr>\n"
html4 += "</table>"

# Adiciona a tabela em HTML ao corpo do e-mail
html5 = "<table border='1' cellpadding='5'>\n"
html5 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>SLA</th><th>tempoFimPrevia</th><th>ultimaAtualização</th><th>tempoForaDaPrévia</th><th>tipoServiço</th><th>tempoExperienciaAgora</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaChegada2.iterrows():
    html5 += "<tr>"
    html5 += f"<td>{row['Ordem Serviço']}</td>"
    html5 += f"<td>{row['Data Inicio']}</td>"
    html5 += f"<td>{row['Hora Inicio']}</td>"
    html5 += f"<td>{row['SLA']}</td>"
    html5 += f"<td>{row['tempoFimPrevia']}</td>"
    html5 += f"<td>{row['ultimaAtualização']}</td>"
    html5 += f"<td>{row['tempoForaDaPrévia']}</td>"
    html5 += f"<td>{row['tipoServiço']}</td>"
    html5 += f"<td>{row['tempoExperienciaAgora']}</td>"
    html5 += f"<td>{row['Cidade']}</td>"
    html5 += f"<td>{row['Estado']}</td>"
    html5 += "</tr>\n"
html5 += "</table>"

## Processo acima de 90 minutos

# Adiciona a tabela em HTML ao corpo do e-mail
html6 = "<table border='1' cellpadding='5'>\n"
html6 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>duracaoAcionamento</th><th>tipoServiço</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaAcionamento3.iterrows():
    html6 += "<tr>"
    html6 += f"<td>{row['Ordem Serviço']}</td>"
    html6 += f"<td>{row['Data Inicio']}</td>"
    html6 += f"<td>{row['Hora Inicio']}</td>"
    html6 += f"<td>{row['duracaoAcionamento']}</td>"
    html6 += f"<td>{row['tipoServiço']}</td>"
    html6 += f"<td>{row['Cidade']}</td>"
    html6 += f"<td>{row['Estado']}</td>"
    html6 += "</tr>\n"
html6 += "</table>"

# Adiciona a tabela em HTML ao corpo do e-mail
html7 = "<table border='1' cellpadding='5'>\n"
html7 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>SLA</th><th>tempoFimPrevia</th><th>ultimaAtualização</th><th>tempoForaDaPrévia</th><th>tipoServiço</th><th>tempoExperienciaAgora</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaChegada3.iterrows():
    html7 += "<tr>"
    html7 += f"<td>{row['Ordem Serviço']}</td>"
    html7 += f"<td>{row['Data Inicio']}</td>"
    html7 += f"<td>{row['Hora Inicio']}</td>"
    html7 += f"<td>{row['SLA']}</td>"
    html7 += f"<td>{row['tempoFimPrevia']}</td>"
    html7 += f"<td>{row['ultimaAtualização']}</td>"
    html7 += f"<td>{row['tempoForaDaPrévia']}</td>"
    html7 += f"<td>{row['tipoServiço']}</td>"
    html7 += f"<td>{row['tempoExperienciaAgora']}</td>"
    html7 += f"<td>{row['Cidade']}</td>"
    html7 += f"<td>{row['Estado']}</td>"
    html7 += "</tr>\n"
html7 += "</table>"

# Adiciona a tabela em HTML ao corpo do e-mail
html8 = "<table border='1' cellpadding='5'>\n"
html8 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>duracaoAcionamento</th><th>tipoServiço</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaAcionamento0.iterrows():
    html8 += "<tr>"
    html8 += f"<td>{row['Ordem Serviço']}</td>"
    html8 += f"<td>{row['Data Inicio']}</td>"
    html8 += f"<td>{row['Hora Inicio']}</td>"
    html8 += f"<td>{row['duracaoAcionamento']}</td>"
    html8 += f"<td>{row['tipoServiço']}</td>"
    html8 += f"<td>{row['Cidade']}</td>"
    html8 += f"<td>{row['Estado']}</td>"
    html8 += "</tr>\n"
html8 += "</table>"

# Adiciona a tabela em HTML ao corpo do e-mail
html9 = "<table border='1' cellpadding='5'>\n"
html9 += "<tr><th>Ordem Serviço</th><th>Data Inicio</th><th>Hora Inicio</th><th>SLA</th><th>tempoFimPrevia</th><th>ultimaAtualização</th><th>tempoForaDaPrévia</th><th>tipoServiço</th><th>tempoExperienciaAgora</th><th>Cidade</th><th>Estado</th></tr>\n"
for i, row in tabelaChegada0.iterrows():
    html9 += "<tr>"
    html9 += f"<td>{row['Ordem Serviço']}</td>"
    html9 += f"<td>{row['Data Inicio']}</td>"
    html9 += f"<td>{row['Hora Inicio']}</td>"
    html9 += f"<td>{row['SLA']}</td>"
    html9 += f"<td>{row['tempoFimPrevia']}</td>"
    html9 += f"<td>{row['ultimaAtualização']}</td>"
    html9 += f"<td>{row['tempoForaDaPrévia']}</td>"
    html9 += f"<td>{row['tipoServiço']}</td>"
    html9 += f"<td>{row['tempoExperienciaAgora']}</td>"
    html9 += f"<td>{row['Cidade']}</td>"
    html9 += f"<td>{row['Estado']}</td>"
    html9 += "</tr>\n"
html9 += "</table>"

#--------------------( Configurar o processo de 30 à 45 minutos de atuação )--------------------#


# Primeiro vou verificar se os campos possuem dados, se os dois forem igual a 0, não irá enviar o e-mail.    
regex = r"Entre 30 e 45 minutos<\/td><td>(\d+)<\/td>"
valor_tabela_2 = re.search(regex, html2).group(1)
valor_tabela_3 = re.search(regex, html3).group(1)

# verifica se ambos os valores são zero
if int(valor_tabela_2) == 0 and int(valor_tabela_3) == 0:
    # não enviar o e-mail
    print("Não enviar e-mail")
else:
    # Inicializa o objeto do aplicativo Outlook
    outlook = win32.Dispatch('outlook.application')
    accounts = outlook.Session.Accounts

    # Verificar se existem pelo menos dois resultados
    if len(accounts) >= 2:
        account_segunda_opcao = accounts[1]
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = account_segunda_opcao

    # Define as informações do e-mail
    mail.To = ''
    mail.Subject = ''

    # Formatação da tabela em HTML
    html_table = tabelaAcionamento.to_html(index=False)
    html1_table = tabelaChegada.to_html(index=False)
    html2_table = resultadosAcionamento.to_html(index=False)
    html3_table = resultadoChegada.to_html(index=False)

    if valor_tabela_2 != '0' or valor_tabela_3 != '0':
        mail.HTMLBody =f""" 
        <html>
        <head>

        <title>Page Title</title>
        </head>
        <body>
        <h2>REPORT SERVIÇOS EM ANDAMENTO</h2>
        <p><h3>SERVIÇOS EM ACIONAMENTO</h3></p> 
        <p>Monitoramento - Consolidado</p> 
        {html2}
        <p> Monitoramento - Detalhado </p> 
        {html8}
        <p><font color="red"> Escalonamento - Coordenador - Serviços acima de 45 minutos e até 60 minutos de atuação</font></p> 
        {html}
        <p><font color="red"> Escalonamento - Gerente - Serviços acima de 60 minutos e até 90 minutos de atuação</font></p> 
        {html4}
        <p><font color="red"> Escalonamento - Superintendente - Serviços acima de 90 minutos de atuação</font></p> 
        {html6}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h3>SERVIÇOS EM TEMPO DE CHEGADA</h3></p> 
        <p>Monitoramento - Consolidado</p>
        {html3}
        <p> Monitoramento - Detalhado </p>
        {html9}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h2><font color="red">Tempo de Experiência do cliente - Visão dia:</font> {tempoExperienciaFinal}<h2></p>
        <p><h4> Obs.: Somente estamos considerando serviços aceitos e sem agendamentos </p><h4>
        </body>
        </html>"""

        # Envia o e-mail
        mail.Send()

        print('E-mail do primeiro processo enviado com sucesso!')

    else:
        print("Não enviar e-mail")


#--------------------( Configurar o processo de 45 à 60 minutos de atuação )--------------------#

# Primeiro vou verificar se os campos possuem dados, se os dois forem igual a 0, não irá enviar o e-mail.    
regex = r"Entre 45 e 60 minutos<\/td><td>(\d+)<\/td>"
valor_tabela_2 = re.search(regex, html2).group(1)
valor_tabela_3 = re.search(regex, html3).group(1)

# verifica se ambos os valores são zero
if int(valor_tabela_2) == 0 and int(valor_tabela_3) == 0:
    # não enviar o e-mail
    print("Não enviar e-mail")
else:
    # Inicializa o objeto do aplicativo Outlook
    outlook = win32.Dispatch('outlook.application')
    accounts = outlook.Session.Accounts

    # Verificar se existem pelo menos dois resultados
    if len(accounts) >= 2:
        account_segunda_opcao = accounts[1]
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = account_segunda_opcao

    # Define as informações do e-mail
    mail.To = ''
    mail.Subject = ''


    if valor_tabela_2 != '0' or valor_tabela_3 != '0':
        mail.HTMLBody =f""" 
        <html>
        <head>
        <title>Page Title</title>
        </head>
        <body>
        <h2>REPORT SERVIÇOS EM ANDAMENTO</h2>
        <p><h3>SERVIÇOS EM ACIONAMENTO</h3></p> 
        <p>Monitoramento - Consolidado</p> 
        {html2}
        <p> Monitoramento - Detalhado </p> 
        {html}
        <p><font color="red"> Escalonamento - Gerente - Serviços acima de 60 minutos e até 90 minutos de atuação</font></p> 
        {html4}
        <p><font color="red"> Escalonamento - Superintendente - Serviços acima de 90 minutos de atuação</font></p> 
        {html6}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h3>SERVIÇOS EM TEMPO DE CHEGADA</h3></p> 
        <p>Monitoramento - Consolidado</p>
        {html3}
        <p> Monitoramento - Detalhado </p>
        {html1}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h2><font color="red">Tempo de Experiência do cliente - Visão dia:</font> {tempoExperienciaFinal}<h2></p>
        <p><h4> Obs.: Somente estamos considerando serviços aceitos e sem agendamentos </p><h4>
        </body>
        </html>"""

        # Envia o e-mail
        mail.Send()

        print('E-mail do primeiro processo enviado com sucesso!')

    else:
        print("Não enviar e-mail")

#--------------------( Configurar o processo de 60 a 90 minutos de atuação )--------------------#

# extrai o valor para "Entre 60 e 90 minutos" em cada tabela
regex = r"Entre 60 e 90 minutos<\/td><td>(\d+)<\/td>"
valor_tabela_2 = re.search(regex, html2).group(1)
valor_tabela_3 = re.search(regex, html3).group(1)

# verifica se ambos os valores são zero
if int(valor_tabela_2) == 0 and int(valor_tabela_3) == 0:
    # não enviar o e-mail
    print("Não enviar e-mail")
else:
    # Inicializa o objeto do aplicativo Outlook
    outlook = win32.Dispatch('outlook.application')
    accounts = outlook.Session.Accounts

    # Verificar se existem pelo menos dois resultados
    if len(accounts) >= 2:
        account_segunda_opcao = accounts[1]
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = account_segunda_opcao

    # Define as informações do e-mail
    mail.To = ''
    mail.Subject = ''


    if valor_tabela_2 != '0' or valor_tabela_3 != '0':
        mail.HTMLBody =f""" 
        <html>
        <head>
        <title>Page Title</title>
        </head>
        <body>
        <h2>REPORT SERVIÇOS EM ANDAMENTO</h2>
        <p><h3>SERVIÇOS EM ACIONAMENTO</h3></p> 
        <p>Monitoramento - Consolidado</p> 
        {html2}
        <p> Monitoramento - Detalhado </p> 
        {html4}

        <p><font color="red"> Escalonamento - Superintendente - Serviços acima de 90 minutos de atuação</font></p> 
        {html6}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h3>SERVIÇOS EM TEMPO DE CHEGADA</h3></p> 
        <p>Monitoramento - Consolidado</p>
        {html3}
        <p> Monitoramento - Detalhado </p>
        {html5}

        <br> <br>
        <hr> <!-- Linha divisória -->

        <p><h2><font color="red">Tempo de Experiência do cliente - Visão dia:</font> {tempoExperienciaFinal}<h2></p>
        <p><h4> Obs.: Somente estamos considerando serviços aceitos e sem agendamentos </p><h4>
        </body>
        </html>"""

        # Envia o e-mail
        mail.Send()

        print('E-mail do segundo processo enviado com sucesso!')

    else:
        print("Não enviar e-mail")

#--------------------( Configurar o processo acima de 90 minutos de atuação )--------------------#

# extrai o valor para "Acima de 90 minutos" em cada tabela
regex = r"Acima de 90 minutos<\/td><td>(\d+)<\/td>"
valor_tabela_2 = re.search(regex, html2).group(1)
valor_tabela_3 = re.search(regex, html3).group(1)

# verifica se ambos os valores são zero
if int(valor_tabela_2) == 0 and int(valor_tabela_3) == 0:
    # não enviar o e-mail
    print("Não enviar e-mail")
else:
    # Inicializa o objeto do aplicativo Outlook
    outlook = win32.Dispatch('outlook.application')
    accounts = outlook.Session.Accounts

    # Verificar se existem pelo menos dois resultados
    if len(accounts) >= 2:
        account_segunda_opcao = accounts[1]
        mail = outlook.CreateItem(0)
        mail.SentOnBehalfOfName = account_segunda_opcao

    # Define as informações do e-mail
    mail.To = ''
    mail.Subject = ''

    # Formatação da tabela em HTML
    html6_table = tabelaAcionamento.to_html(index=False)
    html7_table = tabelaChegada.to_html(index=False)
    html2_table = resultadosAcionamento.to_html(index=False)
    html3_table = resultadoChegada.to_html(index=False)

    # Verifica se os valores das tabelas são diferentes de zero
    if valor_tabela_2 != '0' or valor_tabela_3 != '0':
        mail.HTMLBody = f""" 
        <html>
        <head>
        <title>Page Title</title>
        </head>
        <body>
        <h2>REPORT SERVIÇOS EM ANDAMENTO</h2>
        <p><h3>SERVIÇOS EM ACIONAMENTO</h3></p> 
        <p>Monitoramento - Consolidado</p> 
        {html2}
        <p> Monitoramento - Detalhado </p> 
        {html6}

        <br> <br>
        <hr> <!-- Linha divisória -->


        <p><h3>SERVIÇOS EM TEMPO DE CHEGADA</h3></p> 
        <p>Monitoramento - Consolidado</p>
        {html3}
        <p> Monitoramento - Detalhado </p>
        {html7}

        <br> <br>
        <hr> <!-- Linha divisória -->


        <p><h2><font color="red">Tempo de Experiência do cliente - Visão dia:</font> {tempoExperienciaFinal}<h2></p>
        <p><h4> Obs.: Somente estamos considerando serviços aceitos e sem agendamentos </p><h4>
        </body>
        </html>"""

        # Envia o e-mail
        mail.Send()

        print('E-mail do terceiro processo enviado com sucesso!')

    else:
        print("Não enviar e-mail")

conn.close()
