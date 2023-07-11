import pandas as pd
import numpy as np
import os

def gerar_relatorio(lista_plat, df):

  df_matricula = []
  df_nome = []
  df_funcao = []
  df_evento = []
  df_uop = []
  df_dias = []

  for plat in lista_plat:
    df_filtrado = df.query(f"Uop == '{plat}'")
    lista_nomes = list(df_filtrado['Nome'])
    lista_nomes = list(set(lista_nomes))
    for nome in lista_nomes:
      filtro_nome = df_filtrado.query(f"Nome == '{nome}'")
      filtro_nome = filtro_nome.reset_index(drop=True)
      lista_evento = list(set(filtro_nome['Evento']))
      for evento in lista_evento:
        cont = 0
        soma = 0
        filtro_evento = filtro_nome.query(f"Evento == '{evento}'")
        filtro_evento = filtro_evento.reset_index()
        for i in range(len(filtro_evento)):
          linha = filtro_evento.loc[cont]
          soma += linha['Dias']
          cont+=1
        matricula = linha['Matricula']
        nome = linha['Nome']
        funcao = linha['Funcao']
        uop = linha['Uop']
        df_matricula.append(matricula)
        df_nome.append(nome)
        df_funcao.append(funcao)
        df_evento.append(evento)
        df_uop.append(uop)
        df_dias.append(soma)

  df_sep = pd.DataFrame.from_dict(data={'Matricula': df_matricula,             #cria o dataframe
                                    'Nome': df_nome,
                                    'Função': df_funcao,
                                    'Evento': df_evento,
                                    'Uop': df_uop,
                                    'Dias': df_dias}, orient='columns')
  return df_sep

def tirar_espacos(df_junto, perenco):
  for a in range(len(perenco['Colaborador'])):
    perenco['Colaborador'][a] = perenco['Colaborador'][a].strip(' ')
  for a in range(len(df_junto['Nome'])):
    df_junto['Nome'][a] = df_junto['Nome'][a].strip(' ')

  return df_junto, perenco

def gerar_final(df_junto, perenco):
  lista_plat = ['PPG1', 'PCP-2', 'PCP-1/3', 'PVM1', 'PVM3', 'FSO', 'PCP-1/3 - WORKOVER']
  colaboradores = []
  plataforma = []
  diferenca = []
  for plat in lista_plat:
    cont = 0
    if plat == 'PCP-1/3 - WORKOVER':
      df_plat = df_junto.query(f"Uop == 'WORKOVER'")
      perenco_plat = perenco.query(f"Plataforma == 'WORKOVER'")
    else:
      df_plat = df_junto.query(f"Uop == '{plat}'")
      perenco_plat = perenco.query(f"Plataforma == '{plat}'")
    perenco_plat = perenco_plat.reset_index(drop=True)
    for i in perenco_plat['Colaborador']:
      filtro_df = df_plat.query(f"Nome == '{i}'")
      filtro_df = filtro_df.reset_index()
      linha_perenco = perenco_plat.loc[cont]
      try:
        linha_drake = filtro_df.loc[0]
        vlr_drake = linha_drake['Dias']
      except:
        vlr_drake = 0
        pass
      colaboradores.append(i)
      plataforma.append(linha_drake['Uop'])
      dif = int(linha_perenco['Dias de Trabalho']) - int(vlr_drake)
      diferenca.append(dif)
      cont += 1

  df_final = pd.DataFrame.from_dict(data={'Colaborador': colaboradores,  # cria o dataframe
                                          'Plataforma': plataforma,
                                          'Diferenca': diferenca}, orient='columns')
  return df_final


def comparar(lista_plat, df, bm):
  df_matricula = []
  df_nome = []
  df_funcao = []
  df_uop = []
  df_dias = []
  lista_num = []  # cria uma lista com todos os números que iniciam os colaboradores, para chamar o print
  cont = 1
  while cont < 650:
    lista_num.append(cont)
    cont += 3
  perenco = gerar_bm(lista_num, lista_plat, bm)
  perenco.to_excel('Perenco.xlsx', header = True, index=False)
  perenco = pd.read_excel('Perenco.xlsx')
  os.remove('Perenco.xlsx')
  df.rename(columns={'index': 'indice', 'Matrícula do Trabalhador': 'Matricula', 'Nome do Trabalhador': 'Nome',
                     'Uop do Trabalhador': 'PlatTrab',
                     'Situação do Trabalhador': 'Situacao', 'Função de Folha do Trabalhador': 'Funcao',
                     'Data de Início do Evento': 'Data_ini',
                     'Data de Término do Evento': 'Data_fin', 'Descrição do Evento': 'Evento',
                     'Uop do Evento': 'Uop', 'Quantidade de Dias no Período': 'Dias'}, inplace=True)
  df = df.replace('PCP-1/3 - WORKOVER', "WORKOVER")
  df = df.replace('PVM-1 - WORKOVER', "WORKOVER")
  df = df.replace('PCP-2 - WORKOVER', "WORKOVER")
  df = df.replace('PPG1 - WORKOVER', "WORKOVER")
  df.query(
    "Uop == 'PPG1' or Uop == 'PCP-2' or Uop == 'PCP-1/3' or Uop == 'PVM1' or Uop == 'PVM3' or Uop == 'FSO' or Uop == 'WORKOVER'",
    inplace=True)
  df.query("Evento != 'FOLGA'", inplace=True)
  df.query("Evento != 'RESERVA'", inplace=True)
  df.query("Evento != 'FALTA'", inplace=True)
  df.query("Evento != 'Desligamento'", inplace=True)
  df.query("Evento != 'ABONO ÓBITO'", inplace=True)
  df.query("Evento != 'LICENÇA MÉDICA VENCIDA'", inplace=True)
  df.drop(['Situacao'], inplace=True, axis=1)
  df = df.reset_index(drop=True)
  for plat in lista_plat:
    df_filtrado = df.query(f"Uop == '{plat}'")
    lista_nomes = list(df_filtrado['Nome'])
    lista_nomes = list(set(lista_nomes))
    for nome in lista_nomes:
      cont = 0
      soma = 0
      filtro_nome = df_filtrado.query(f"Nome == '{nome}'")
      filtro_nome = filtro_nome.reset_index(drop=True)
      # lista_evento = list(set(filtro_nome['Evento']))
      for i in range(len(filtro_nome)):
        linha = filtro_nome.loc[cont]
        soma += linha['Dias']
        cont+=1
      matricula = linha['Matricula']
      nome = linha['Nome']
      funcao = linha['Funcao']
      uop = linha['Uop']
      df_matricula.append(matricula)
      df_nome.append(nome)
      df_funcao.append(funcao)
      # df_evento.append(evento)
      df_uop.append(uop)
      df_dias.append(soma)

  df_junto = pd.DataFrame.from_dict(data={'Matricula': df_matricula,             #cria o dataframe
                                    'Nome': df_nome,
                                    'Função': df_funcao,
                                    # 'Evento': df_evento,
                                    'Uop': df_uop,
                                    'Dias': df_dias}, orient='columns')
  df_junto, perenco = tirar_espacos(df_junto, perenco)
  df_final = gerar_final(df_junto, perenco)

  return df_final

def verifica(plat, df, lista_num):
  cont = 0
  dias_trab = 0
  df_colaborador = []
  df_plat = []
  df_funcao = []
  df_ident = []
  df_trab = []
  for i in df.index + 1:
    if i == 547:            #como a planilha só vai até 546, se chegar em 547 ele printa o ultimo colaborador
      # print(f'{colaborador} - {plat} - {funcao} - {ident} - {dias_trab - 3}')
      df_colaborador.append(colaborador)
      df_plat.append(plat)
      df_funcao.append(funcao)
      df_ident.append(ident)
      df_trab.append(dias_trab)
    linha = df.loc[cont]         #define a linha atual
    if cont in lista_num:        #está dentro da lista então chama o print
      if cont == 1:              #primeiro valor está zerado, então pula
        pass
      else:
        # print(f'{colaborador} - {plat} - {funcao} - {ident} - {dias_trab - 3}')
        df_colaborador.append(colaborador)
        df_plat.append(plat)
        df_funcao.append(funcao)
        df_ident.append(ident)
        df_trab.append(dias_trab)
      colaborador = linha['Unnamed: 2']
      funcao = linha['Unnamed: 3']
      ident = linha['Unnamed: 4']
      plat = linha['Unnamed: 5']
      dias_trab = linha['Unnamed: 50']
    cont+=1

  df = pd.DataFrame.from_dict(data={'Colaborador': df_colaborador,             #cria o dataframe
                                    'Plataforma': df_plat,
                                    'Função': df_funcao,
                                    'Ident. Contratual': df_ident,
                                    'Dias de Trabalho': df_trab}, orient='columns')
  return df

def limpar(df):
  cont = 0
  for colaborador in df['Colaborador']:
    linha = df.loc[cont]
    if any(char.isdigit() for char in str(colaborador)):
      df = df.drop(cont)
    elif linha['Função'] == 0:
      df = df.drop(cont)
    cont+=1
  return df

def gerar_bm(lista_num, lista_plat, caminho):
  for i in lista_plat:
    if i == 'PPG1':
      plat = 'PPG1'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='PPG-1', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df1 = verifica(i, df, lista_num)
    elif i == 'PCP-2':
      plat = 'PCP-2'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='PCP-2', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df2 = verifica(i, df, lista_num)
    elif i == 'PCP-1/3':
      plat = 'PCP-1/3'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='PCP-1-3', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df3 = verifica(i, df, lista_num)
    elif i == 'PVM1':
      plat = 'PVM1'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='PVM-1', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df4 = verifica(i, df, lista_num)
    elif i == 'PVM3':
      plat = 'PVM3'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='PVM-3', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df5 = verifica(i, df, lista_num)
    elif i == 'FSO':
      plat = 'FSO'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='FSO', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df6 = verifica(i, df, lista_num)
    elif i == 'WORKOVER':
      plat = 'WORKOVER'
      # ALTERAR A COLUNA FINAL DE ACORDO COM A QUANTIDADE DE DIAS NO MES
      df = pd.read_excel(caminho, sheet_name='Workover', usecols='C,D,E,H:AY')  # define as colunas que vão ser lidas
      df = df.fillna(0)  # muda os valores nulos para 0
      df.insert(3, "Unnamed: 5", plat, True)  # insere a coluna com a plataforma
      df = df.loc[4:550]
      df.reset_index(inplace=True, drop=True)
      df6 = verifica(i, df, lista_num)

  data = pd.concat([df1, df2, df3, df4, df5, df6], ignore_index=True)
  final_data = limpar(data)

  return final_data
