<h1 align="center">
📄<br>README - Projeto Financeiro Carteira de Ativos
</h1>

## Índice 

* [Descrição do Projeto](#descrição-do-projeto)
* [Funcionalidades e Demonstração da Aplicação](#funcionalidades-e-demonstração-da-aplicação)
* [Pré requisitos](#pré-requisitos)
* [Execução](#execução)
* [Bibliotecas](#bibliotecas)

# Descrição do projeto
> Análise de dados financeiros de uma carteira de ativos. O objetivo deste projeto foi, a partir de uma base de dados contendo certos ativos financeiros e suas respectivas quantidades de cotas, encontrar e coletar as cotações atualizadas por meio de web-scrapping, atualizar a base de dados destes ativos, gerar gráficos comparativos de toda a carteira e desta com o índice Ibovespa e, então, enviar um e-mail com todas estas informações de maneira organizada e resumida.

# Funcionalidades e Demonstração da Aplicação

- Gráfico comparativo de toda a carteira de ativos:<br>
![Screenshot_1](https://user-images.githubusercontent.com/128300382/227935370-8bb1573c-551c-4742-a30e-b6f1e4d3d46a.png)

- E-mail enviado com as principais informações da carteira:<br>
![Screenshot_2](https://user-images.githubusercontent.com/128300382/227935395-0fc4d2d2-bf6f-4a0b-933b-91bf83fbabed.png)

## Pré requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Base de dados (arquivo excel com a carteira de ativos)

## Execução

O código é executado de maneira direta e deve estar no mesmo diretório da base de dados excel. Após sua execução, será enviado um email com 2 gráficos comparativos e a base de dados atualizada em forma de tabela html.

## Bibliotecas
* <strong>pandas:</strong> biblioteca que permite, no caso, a integração de arquivo excel<br>
* <strong>os:</strong> bibliotecas de integração de arquivos e pastas do computador<br>
* <strong>datetime:</strong> bibliotecas que permite o uso de data e hora<br>
* <strong>matplotlib.pyplot, seaborn, plotly.express:</strong> bibliotecas de geração de gráfico<br>
* <strong>win32com.client:</strong> biblioteca de integração de aplicativos do Windows, no caso, o Outlook<br>
* <strong>pandas_datareader.data:</strong> biblioteca de acesso remoto de dados do pandas<br>
* <strong>yfinance:</strong> biblioteca de integração de dados financeiro do Yahoo Finance<br>
