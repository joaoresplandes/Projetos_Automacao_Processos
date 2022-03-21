#!/usr/bin/env python
# coding: utf-8

### Passo 1 - Importar Arquivos e Bibliotecas

#importando bibliotecas
import pandas as pd
import pathlib
import yagmail
from Ipython import display

#importando bases de dados
emails = pd.read_csv(r'Bases de Dados/Emails.csv')
lojas = pd.read_csv(r'Bases de Dados/Lojas.csv', encoding = 'latin1', sep= ';')
vendas = pd.read_csv(r'Bases de Dados/Vendas.csv')
display(emails)
display(lojas)
display(vendas)

vendas = vendas.merge(lojas, on='ID Loja')
display(vendas)


### Passo 2 - Definir Criar uma Tabela para cada Loja e Definir o dia do Indicador

##### Será separado em relatórios individuais para cada loja
'''
Separa para cada loja um DataFrame diferente, tudo em um Dicionario só, ou seja, cada chave deste 
dicionário corresponde a uma loja, e seu valor é um dataframe correspondente ao mesmo.
 '''
dicionario_lojas = {}
for loja in lojas['Loja']:
    dicionario_lojas[loja] = vendas.loc[vendas['Loja'] == loja , :]

# retorna a última data do relatório
data_relatorio = vendas['Data'].max()

### Passo 3 - Salvar a planilha na pasta de backup

##### A cada dia de envio de relatório, será encaminhado às respectivas pastas, os relatórios de cada loja. Caso uma das lojas não exista ainda na pasta de backup, a mesma será criada para armazenamento do relatório

# Determinado o diretório padrão para o Backup
diretorio_backup = pathlib.Path('Backup Arquivos Lojas')
# Tornando o diretório padrão um iteravel
lista_arquivos_backup = diretorio_backup.iterdir()
# Colocando todas as pastas/arquivos existentes dentro de uma lista para posteriormente conferi-los.
lista_arquivos = [arquivo.name for arquivo in lista_arquivos_backup]

for loja in lojas['Loja']:
    if loja not in lista_arquivos:
        nome_diretorio = diretorio_backup / loja # Determina o caminho mas a pasta a ser criada.
        nome_diretorio.mkdir() # mkdir - cria a pasta com o caminho determinado
        
    # Separando os DataFrame para cada uma das lojas em cada pasta respectiva.
    local_arquivo = diretorio_backup/ loja / f'{data_relatorio.month}_{data_relatorio.day}_{loja}.xlsx'
    dicionario_lojas[loja].to_excel(local_arquivo, index = False)
    print(f'Backup {data_relatorio.month}_{data_relatorio.day}_{loja}.xlsx Realizado')

print('Backup Realizado com Sucesso!')

#### Metas a serem atingidas

#metas a serem atingidas:
meta_faturamento_dia = 1000
meta_faturamento_anual = 1650000
meta_diversidade_dia = 4
meta_diversidade_anual = 120
meta_ticket_dia = 500
meta_ticket_anual = 500

##### Assim que todo o processo de separação dos relatórios por loja e armazenamentos dos backups,
#  e-mails encaminhados para cada gerente serão enviados com seus respectivos indicadores calculados e relatórios para análise.

# Usando o provedor do Gmail para mandar os e-mail para os Gerentes
usuario = yagmail.SMTP(user = 'email_remetente', password = 'senha')
for loja in lojas['Loja']:

    # ### Passo 4 - Calcular os indicadores para as lojas
    #      - Faturamento Diario
    #      - Faturamento Anual
    #      - Diversidade de Produtos Vendidos no Dia
    #      - Diversidade de Produtos Vendidos no Ano
    #      - Ticket Médio por Venda no Dia
    #      - Ticket Médio por Venda no Ano

    df = dicionario_lojas[loja]  # DataFrame com todos os dias
    df_dia = df.loc[df['Data'] == data_relatorio, :]  # DataFrame filtrado só com o dia do relatório
    # - Calculando o Faturamento do dia do relatório por loja
    faturamento_loja_dia = df_dia['Valor Final'].sum()
    # - Calculando o Faturamento anual por loja
    faturamento_loja_anual = df['Valor Final'].sum()
    # - Calculando a diversidade de Produtos vendidos no dia do relatório por loja
    qtde_produtos_dia = len(df_dia['Produto'].unique())
    # - Calculando a diversidade de Produtos vendidos anualmente por loja
    qtde_produtos_anual = len(df['Produto'].unique())
    # - Calculando o ticket médio por Produto vendido no dia do relatório
    lista_ticket_dia = df_dia.groupby('Código Venda').sum()
    ticket_medio_dia = lista_ticket_dia['Valor Final'].mean()
    # - Calculando o ticket médio por Produto vendido anualmente por loja
    lista_ticket_anual = df.groupby('Código Venda').sum()
    ticket_medio_anual = lista_ticket_anual['Valor Final'].mean()

    gerente = emails.loc[emails['Loja'] == loja,'Gerente'].values[0]
    email = emails.loc[emails['Loja'] == loja, 'E-mail'].values[0]
    assunto = f'OnePage Dia {data_relatorio.day}/{data_relatorio.month}/{data_relatorio.year} - Loja {loja}'
    # verificação de indicadores atigiram a meta
    ## faturamento do dia:
    if faturamento_loja_dia >= meta_faturamento_dia:
        cor_fat_dia = 'green'
    else:
        cor_fat_dia = 'red'
    ## diversidade de produtos vendidos no dia:
    if qtde_produtos_dia >= meta_diversidade_dia:
        cor_div_dia = 'green'
    else:
        cor_div_dia = 'red'
    ## ticket médio do dia
    if ticket_medio_dia >= meta_ticket_dia:
        cor_ticket_dia = 'green'
    else:
        cor_ticket_dia = 'red'
    ## faturamento anual:
    if faturamento_loja_anual >= meta_faturamento_anual:
        cor_fat_anual = 'green'
    else:
        cor_fat_anual = 'red'
    ## diverside de produtos vendidos por ano
    if qtde_produtos_anual >= meta_diversidade_anual:
        cor_div_anual = 'green'
    else:
        cor_div_anual = 'red'
    ## ticket médio anual
    if ticket_medio_anual >= meta_ticket_anual:
        cor_ticket_anual = 'green'
    else:
        cor_ticket_anual = 'red'

    corpo_do_email = f'''
    <P>Bom dia, {gerente}</p>
    <p>O resultado de ontem <strong>(Dia {data_relatorio.day}/{data_relatorio.month})</strong> da <strong>Loja {loja}</strong> foi:</p>
    <table>
      <tr>
        <th>Indicadores</th>
        <th>Valor Dia</th>
        <th>Méta Dia</th>
        <th>Cenário Dia</th>
      </tr>
      <tr>
        <td>Faturamento Dia</td>
        <td style="text-align: center" >R$ {faturamento_loja_dia:.2f}</td>
        <td style="text-align: center" >R$ {meta_faturamento_dia:.2f}</td>
        <td style="text-align: center" ><font color = {cor_fat_dia}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center" >{qtde_produtos_dia}</td>
        <td style="text-align: center" >{meta_diversidade_dia}</td>
        <td style="text-align: center" ><font color = {cor_div_dia}>◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio Dia</td>
        <td style="text-align: center" >R$ {ticket_medio_dia:.2f}</td>
        <td style="text-align: center" >R$ {meta_ticket_dia:.2f}</td>
        <td style="text-align: center" ><font color = {cor_ticket_dia}>◙</font></td>
      </tr>
      <br>
      <tr>
        <th>Indicadores</th>
        <th>Valor Anual</th>
        <th>Méta Anual</th>
        <th>Cenário Anual</th>
      </tr>
      <tr>
        <td>Faturamento Anual</td>
        <td style="text-align: center" >R$ {faturamento_loja_anual:.2f}</td>
        <td style="text-align: center" >R$ {meta_faturamento_anual:.2f}</td>
        <td style="text-align: center" ><font color = {cor_fat_anual}>◙</font></td>
      </tr>
      <tr>
        <td>Diversidade de Produtos</td>
        <td style="text-align: center" >{qtde_produtos_anual}</td>
        <td style="text-align: center" >{meta_diversidade_anual}</td>
        <td style="text-align: center" ><font color = {cor_div_anual}>◙</font></td>
      </tr>
      <tr>
        <td>Ticket Médio Anual</td>
        <td style="text-align: center" >R$ {ticket_medio_anual:.2f}</td>
        <td style="text-align: center" >R$ {meta_ticket_anual:.2f}</td>
        <td style="text-align: center" ><font color = {cor_ticket_anual}>◙</font></td>
      </tr>
    </table>
    <p>Segue em anexo a planilha com todos os dados para mais analises.</p>
    <p>Qualquer dúvida, estou à disposição.</p>
    <p>Att.,</p>
    <p>João Pedro Resplandes</p>
    '''
    ### - Pass 5 - Automatizar todas as lojas e Enviar por e-mail para cada gerente

    anexo = diretorio_backup / loja / f'{data_relatorio.month}_{data_relatorio.day}_{loja}.xlsx'
    usuario.send(to = email, 
                 subject = assunto,
                 contents = corpo_do_email, 
                 attachments = anexo)
    print(F'Enviado para {gerente} no seguinte e-mail: {email}')
print('E-mails Enviados com Sucesso!!')

### Passo 6 - Criar ranking para diretoria

##### Realizando um Ranking tanto anual como para o dia do relatório indicando os rendimento de faturamento das lojas para os respectivas indicadores

# Ranking das melhores empresas para o faturamento anual
loja_valor_final = vendas[['Loja','Valor Final']].groupby('Loja').sum()
ranking_ano = loja_valor_final.sort_values(by = 'Valor Final', ascending = False)
#display(ranking_ano)
# - Realizando o backup do arquivo de ranking anual na pasta de backup
ranking_ano.to_excel(diretorio_backup / f'Ranking_anual.xlsx', index = False)

# Ranking das melhores empresas para o faturamento no dia do relatório
vendas_no_dia = vendas.loc[vendas['Data'] == data_relatorio, :]
lojas_valor_final_dia = vendas_no_dia[['Loja','Valor Final']].groupby('Loja').sum()
ranking_dia = lojas_valor_final_dia.sort_values(by = 'Valor Final', ascending = False)
#display(ranking_dia)
# - Realizando o backup do arquivo de ranking do dia do relatório
ranking_dia.to_excel(diretorio_backup / f'Ranking_dia.xlsx', index = False)

### Passo 7 - Enviar e-mail para diretoria

##### Por final, um e-mail será encaminhado para a diretoria com os rankings de faturamento anual e do dia do relatório,
# junto do indicador "Faturamento Anual" e "Faturamento do dia do Relatório" da melhor e pior loja de sua empresa

email_diretoria = emails.loc[emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
nome_diretoria = emails.loc[emails['Loja'] == 'Diretoria', 'Gerente'].values[0]
assunto = f'Faturamento da Melhor e pior Loja para o dia {data_relatorio.day}/{data_relatorio.month}'
corpo_email = f'''
Bom dia, {nome_diretoria}

A Loja que apresentou o melhor rendimento no dia {data_relatorio.day}/{data_relatorio.month} foi {ranking_dia.index[0]} com Faturamento de R$ {ranking_dia.iloc[0,0]:.2f} 
A Loja que apresentou o pior rendimento no dia {data_relatorio.day}/{data_relatorio.month} foi {ranking_dia.index[-1]} com Faturamento de R$ {ranking_dia.iloc[-1,0]:.2f}

A Loja que apresentou o melhor rendimento no ano de {data_relatorio.year} foi {ranking_ano.index[0]} com Faturamento de R$ {ranking_ano.iloc[0,0]:.2f}
A Loja que apresentou o pior rendimento no ano de {data_relatorio.year} foi {ranking_ano.index[-1]} com Faturamento R$ {ranking_ano.iloc[-1,0]:.2f}
'''
anexos = [diretorio_backup / 'Ranking_dia.xlsx', diretorio_backup / 'Ranking_anual.xlsx']
usuario.send(to = email_diretoria,
            subject = assunto,
            contents = corpo_email,
            attachments = anexos)
print(f'Email para a {nome_diretoria} enviado com Sucesso.')
