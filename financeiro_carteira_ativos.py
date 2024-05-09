#!/usr/bin/env python
# coding: utf-8

# ## Análise da performance de uma carteira de ativos (base de dados xlsx)
# 
#     Ao analisar a carteira de ativos:
#     - atualizar as cotações (dia atual em 2023);
#     - verificar o valor médio de cada ativo no ano de 2022;
#     - gerar gráfico para visualizar a variação de cada ativo;
#     - gerar gráfico comparativo da variação percentual de toda a carteira;
#     - comparar a carteira com o índice IBOVESPA;
#     - calcular o retorno atual da carteira em comparação ao IBOVESPA.
#     
#     Por fim, enviar as informações coletadas via email, juntamente com gráficos demonstrativos.

# In[1]:


import pandas as pd
from datetime import datetime
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
import win32com.client as win32
import os

# bibliotecas que acessam cotações na internet; retornam um df
import pandas_datareader.data as pdr
import yfinance as yf
yf.pdr_override()


# In[2]:


df = pd.read_excel('Carteira.xlsx')
display(df)


# ###### Atualizar as cotações (2023)

# In[3]:


# criar um df somente com as cotações do período desejado (início de 2022 a hoje)

df_cotacoes = pd.DataFrame()

start_date = datetime(2022, 1, 1)
dia = datetime.now().day
mes = datetime.now().month
ano = datetime.now().year
end_date = datetime(ano, mes, dia)

for ativo in df['Ativos']:
    df_cotacoes[ativo] = pdr.get_data_yahoo(f'{ativo}.SA', start=start_date, end=end_date)['Adj Close']

display(df_cotacoes)


# In[4]:


# tratamento dos dados: há cotações vazias/nulas?
print(df_cotacoes.isnull().sum())


# In[5]:


# criar uma coluna no df com as cotações atualizadas
data_hoje = end_date.strftime('%d/%m/%Y')

df[f'Cotação de hoje'] = 0

for i, ativo in enumerate(df['Ativos']):
    df.loc[df['Ativos']==ativo, 'Cotação de hoje'] = df_cotacoes.iloc[-1, i]
    
display(df)


# ###### Verificar o valor médio de cada ativo desde o ano de 2022

# In[6]:


# criar um df com o ticket médio das cotações desde 2022
df_media = df_cotacoes.mean()
df_media = df_media.reset_index()
display(df_media)


# In[7]:


# criar uma coluna no df original com o ticket médio das cotações desde 2022
df[f'Média Cotação desde 2022'] = 0

for i, ativo in enumerate(df['Ativos']):
    df.loc[df['Ativos']==ativo, 'Média Cotação desde 2022'] = df_media.loc[df_media['index']==ativo, 0]
    
display(df)


# ###### Gerar gráficos para visualizar a variação de cada ativo

# In[8]:


# gráfico de linha das cotações de cada ativo no período
for ativo in df['Ativos']:
    plt.figure(figsize=(10,5))
    plt.title(f'Cotações {ativo} desde 2022')
    sns.lineplot(data=df_cotacoes, y=df_cotacoes[ativo], x=df_cotacoes.index)


# In[9]:


# gráfico de caixa de diagrama das cotações de cada ativo no período
for ativo in df['Ativos']:
    plt.figure(figsize=(10,5))
    plt.title(f'Cotações {ativo} desde 2022')
    sns.boxplot(data=df_cotacoes, x=df_cotacoes[ativo])


# ###### Gerar gráfico comparativo da variação percentual de toda a carteira

# In[10]:


# gráfico de linha comparativo de todos os ativos no período

# normalizar = dividir todos os valores do df pelo valor inicial (1a linha)
df_normalizado = df_cotacoes / df_cotacoes.iloc[0, :]

plt.figure(figsize=(15,10))
plt.title('Comparativo dos ativos da carteira no período')
plt.legend(loc='best')
sns.lineplot(data=df_normalizado)
plt.savefig('comparativo_normalizado.png') 


# ###### Comparar a carteira com o índice IBOVESPA

# In[11]:


# histórico de cotações no período do IBOVESPA
df_ibovespa = pdr.get_data_yahoo('^BVSP', start=start_date, end=end_date)['Adj Close']

# adicionar as cotações ao df_atual
df_cotacoes['IBOVESPA'] = df_ibovespa

# criar um df com o valor total investido nos ativos por dia
df_valor_investimento = pd.DataFrame()

for ativo in df['Ativos']:
    df_valor_investimento[ativo] = df_cotacoes[ativo] * df.loc[df['Ativos']==ativo, 'Qtde'].values[0]

# criar neste df uma coluna com a soma do total investido por dia   
df_valor_investimento['Total'] = df_valor_investimento.sum(axis=1) #axis=1 soma a linha; axis=0 soma a coluna
    
display(df_valor_investimento)


# In[12]:


# criar dois gráficos de linha sobrepostos comparativos do total investido da carteira com o índice IBOVESPA (matplotlib)
plt.figure(figsize=(15,5))
plt.title('Comparação do valor total da carteira com o índice IBOVESPA')
plt.plot(df_valor_investimento['Total'], label='Carteira', c='Red')
plt.plot(df_cotacoes['IBOVESPA'], label='Ibovespa', c='Green')
plt.show()

# criar dois gráficos de linha sobrepostos comparativos do total investido da carteira com o índice IBOVESPA (plotly)
fig1 = px.line(df_valor_investimento['Total'])
fig1.update_traces(line_color='#0000FF')

fig2 = px.line(df_cotacoes['IBOVESPA'])
fig2.update_traces(line_color='#FF0000')

import plotly.graph_objects as go

fig3 = go.Figure(data = fig1.data + fig2.data)
fig3.show()


# In[13]:


#para melhorar a visualização da comparação usam-se gráficos normalizados (divide-se cada valor do ativo pelo seu valor inicial)
df_ibovespa_normalizado = df_cotacoes['IBOVESPA'] / df_cotacoes.iloc[0,-1]
df_valor_investimento_normalizado = df_valor_investimento['Total'] / df_valor_investimento.iloc[0,-1]

plt.figure(figsize=(15,5))
plt.title('Comparação da variação percentual do valor total da carteira com o índice IBOVESPA')
plt.plot(df_valor_investimento_normalizado, c='Red')
plt.plot(df_ibovespa_normalizado, c='Green')
plt.savefig('comparativo_ibovespa.png') 


# ###### Calcular o retorno atual da carteira em comparação ao IBOVESPA

# In[14]:


#cálculo do percentual = (valor mais atual / valor inicial) - 1
retorno_carteira = (df_valor_investimento.iloc[-1, -1] / df_valor_investimento.iloc[0, -1]) - 1
print(f'Retorno da carteira atual ({data_hoje}): {retorno_carteira:.2%}')

retorno_ibovespa = (df_cotacoes.iloc[-1, -1] / df_cotacoes.iloc[0, -1]) - 1
print(f'Retorno IBOVESPA atual ({data_hoje}): {retorno_ibovespa:.2%}')


# ##### Enviar email

# In[15]:


outlook = win32.Dispatch('outlook.application')
e = outlook.CreateItem(0)
e.To = 'bep_rafael@hotmail.com'
e.Subject = f'Análise da performance da carteira de ativos - {data_hoje}'
e.HTMLBody = f'''
<p>Retorno da carteira atual: <strong>{retorno_carteira:.2%}</strong></p>
<p>Retorno IBOVESPA atual: <strong>{retorno_ibovespa:.2%}</strong></p>
<br>
<p>{df.to_html(formatters={'Cotação de hoje':'R$ {:,.2f}'.format, 'Média Cotação desde 2022':'R$ {:,.2f}'.format})}</p>
'''
lista_arquivos = os.listdir(os.getcwd())
for arquivo in lista_arquivos:
    if '.png' in arquivo:
        e.Attachments.Add(fr'{os.getcwd()}\{arquivo}')
e.Send()


# In[ ]:




