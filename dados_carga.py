# Importando os pacotes necessários
import py_dss_interface
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
import datetime

# Carregando e tratando os dados de carga - Região Sudeste - 01/01/2020 até 31/12/2021

#lendo o arquivo
base_carga = pd.read_csv("C:\Projeto_Final\Base_Carga\Simples_Geração_de_Energia_Dia_data.csv",";")
#deslocando a primeira linha (dados vazios)
base_carga = base_carga.iloc[1:]
#removendo colunas que não iremos usar
base_carga.drop(["cod_aneel (tb_referenciacegusina (Usina))","cod_nucleoaneel (tb_referenciacegusina (Usina))","Data Dica","dsc_estado","nom_tipousinasite","nom_usina2","Período Exibido GE","id_subsistema"],axis = 1, inplace = True)
#renomeando as colunas
base_carga.rename(columns = {'Data Escala de Tempo 1 GE Simp 4':'Data/Hora', 'Selecione Tipo de GE Simp 4':'Carga'}, inplace = True)  
#alterando os tipos de dados das colunas e criando colunas de Hora, Data, Dia da Semana
base_carga['Data/Hora'] = pd.to_datetime(base_carga['Data/Hora'])
base_carga['Data'] = base_carga['Data/Hora'].dt.strftime('%d-%m-%Y')
base_carga['Hora'] = base_carga['Data/Hora'].dt.strftime('%H:%M:%S')
base_carga['Dia da Semana'] = base_carga['Data/Hora'].dt.strftime('%A')
base_carga['Data Hora'] = base_carga['Data/Hora'].dt.strftime('%d-%m-%Y %H:%M:%S')
base_carga['Mês'] = base_carga['Data/Hora'].dt.strftime('%m')
#mudando o tipo de dados para "float"
base_carga['Carga'] = base_carga['Carga'].str.replace(',','.').astype(float)
#dicionários para passar pro Português os dias da semana e definir os dias úteis
dias_da_semana = {'Sunday':'Domingo','Monday':'Segunda-Feira','Tuesday':'Terça-Feira','Wednesday':'Quarta-Feira','Thursday':'Quinta-Feira','Friday':'Sexta-Feira','Saturday':'Sábado'}
dias_uteis = {'Domingo':'Não Útil','Segunda-Feira':'Útil','Terça-Feira':'Útil','Quarta-Feira':'Útil','Quinta-Feira':'Útil','Sexta-Feira':'Útil','Sábado':'Não Útil'}
#adicionando as colunas de dias úteis e dias da semana em português
base_carga['Dia da Semana'] = base_carga['Dia da Semana'].replace(dias_da_semana)
base_carga['Dia Útil'] = base_carga['Dia da Semana'].replace(dias_uteis)
#reordenando as colunas
base_carga = base_carga[['Data/Hora','Data','Hora','Data Hora','Mês','Dia da Semana','Dia Útil','Carga']]

#exportando dados tratados para excel e csv
base_carga.to_excel("C:\Projeto_Final\Base_Carga\Dados_de_Carga.xlsx",index=False)
base_carga.to_csv("C:\Projeto_Final\Base_Carga\Dados_de_Carga.csv",index=False)

#Agrupando os dados por mês e hora para conseguir máx, mín, média e mediana dos
#valores de carga hora a hora de cada mês separados por dias úteis e não úteis
base_ = base_carga.groupby(['Mês','Hora','Dia Útil'])['Carga'].agg(['max','min','mean','median'])
base_['CargaMédia_pu'] = np.divide(base_['mean'],base_carga['Carga'].max())

# Valores para janeiro
jan = base_.loc[('01')]
jan_u = jan[jan.index.isin(['Útil'], level=1)]
jan_u = jan_u.reset_index(level="Hora")
jan_nu = jan[jan.index.isin(['Não Útil'], level=1)]
jan_nu = jan_nu.reset_index(level="Hora")
loadshape_jan_u = jan_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_jan_nu = jan_nu['CargaMédia_pu'].reset_index(drop=True)
jan_u.to_excel("C:\Projeto_Final\Base_Carga\\01jan_u.xlsx",index=False)
jan_nu.to_excel("C:\Projeto_Final\Base_Carga\\01jan_nu.xlsx",index=False)
loadshape_jan_u.to_csv("C:\Projeto_Final\\8500-Node\\01loadshape_jan_u.csv",index=False, header=False)
loadshape_jan_nu.to_csv("C:\Projeto_Final\\8500-Node\\01loadshape_jan_nu.csv",index=False, header=False)

# Valores para fevereiro
fev = base_.loc[('02')]
fev_u = fev[fev.index.isin(['Útil'], level=1)]
fev_u = fev_u.reset_index(level="Hora")
fev_nu = fev[fev.index.isin(['Não Útil'], level=1)]
fev_nu = fev_nu.reset_index(level="Hora")
loadshape_fev_u = fev_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_fev_nu = fev_nu['CargaMédia_pu'].reset_index(drop=True)
fev_u.to_excel("C:\Projeto_Final\Base_Carga\\02fev_u.xlsx",index=False)
fev_nu.to_excel("C:\Projeto_Final\Base_Carga\\02fev_nu.xlsx",index=False)
loadshape_fev_u.to_csv("C:\Projeto_Final\\8500-Node\\02loadshape_fev_u.csv",index=False, header=False)
loadshape_fev_nu.to_csv("C:\Projeto_Final\\8500-Node\\02loadshape_fev_nu.csv",index=False, header=False)

# Valores para março
mar = base_.loc[('03')]
mar_u = mar[mar.index.isin(['Útil'], level=1)]
mar_u = mar_u.reset_index(level="Hora")
mar_nu = mar[mar.index.isin(['Não Útil'], level=1)]
mar_nu = mar_nu.reset_index(level="Hora")
loadshape_mar_u = mar_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_mar_nu = mar_nu['CargaMédia_pu'].reset_index(drop=True)
mar_u.to_excel("C:\Projeto_Final\Base_Carga\\03mar_u.xlsx",index=False)
mar_nu.to_excel("C:\Projeto_Final\Base_Carga\\03mar_nu.xlsx",index=False)
loadshape_mar_u.to_csv("C:\Projeto_Final\\8500-Node\\03loadshape_mar_u.csv",index=False, header=False)
loadshape_mar_nu.to_csv("C:\Projeto_Final\\8500-Node\\03loadshape_mar_nu.csv",index=False, header=False)

# Valores para abril
abr = base_.loc[('04')]
abr_u = abr[abr.index.isin(['Útil'], level=1)]
abr_u = abr_u.reset_index(level="Hora")
abr_nu = abr[abr.index.isin(['Não Útil'], level=1)]
abr_nu = abr_nu.reset_index(level="Hora")
loadshape_abr_u = abr_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_abr_nu = abr_nu['CargaMédia_pu'].reset_index(drop=True)
abr_u.to_excel("C:\Projeto_Final\Base_Carga\\04abr_u.xlsx",index=False)
abr_nu.to_excel("C:\Projeto_Final\Base_Carga\\04abr_nu.xlsx",index=False)
loadshape_abr_u.to_csv("C:\Projeto_Final\\8500-Node\\04loadshape_abr_u.csv",index=False, header=False)
loadshape_abr_nu.to_csv("C:\Projeto_Final\\8500-Node\\04loadshape_abr_nu.csv",index=False, header=False)


# Valores para maio
mai = base_.loc[('05')]
mai_u = mai[mai.index.isin(['Útil'], level=1)]
mai_u = mai_u.reset_index(level="Hora")
mai_nu = mai[mai.index.isin(['Não Útil'], level=1)]
mai_nu = mai_nu.reset_index(level="Hora")
loadshape_mai_u = mai_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_mai_nu = mai_nu['CargaMédia_pu'].reset_index(drop=True)
mai_u.to_excel("C:\Projeto_Final\Base_Carga\\05mai_u.xlsx",index=False)
mai_nu.to_excel("C:\Projeto_Final\Base_Carga\\05mai_nu.xlsx",index=False)
loadshape_mai_u.to_csv("C:\Projeto_Final\\8500-Node\\05loadshape_mai_u.csv",index=False, header=False)
loadshape_mai_nu.to_csv("C:\Projeto_Final\\8500-Node\\05loadshape_mai_nu.csv",index=False, header=False)

# Valores para junho
jun = base_.loc[('06')]
jun_u = jun[jun.index.isin(['Útil'], level=1)]
jun_u = jun_u.reset_index(level="Hora")
jun_nu = jun[jun.index.isin(['Não Útil'], level=1)]
jun_nu = jun_nu.reset_index(level="Hora")
loadshape_jun_u = jun_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_jun_nu = jun_nu['CargaMédia_pu'].reset_index(drop=True)
jun_u.to_excel("C:\Projeto_Final\Base_Carga\\06jun_u.xlsx",index=False)
jun_nu.to_excel("C:\Projeto_Final\Base_Carga\\06jun_nu.xlsx",index=False)
loadshape_jun_u.to_csv("C:\Projeto_Final\\8500-Node\\06loadshape_jun_u.csv",index=False, header=False)
loadshape_jun_nu.to_csv("C:\Projeto_Final\\8500-Node\\06loadshape_jun_nu.csv",index=False, header=False)

# Valores para julho
jul = base_.loc[('07')]
jul_u = jul[jul.index.isin(['Útil'], level=1)]
jul_u = jul_u.reset_index(level="Hora")
jul_nu = jul[jul.index.isin(['Não Útil'], level=1)]
jul_nu = jul_nu.reset_index(level="Hora")
loadshape_jul_u = jul_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_jul_nu = jul_nu['CargaMédia_pu'].reset_index(drop=True)
jul_u.to_excel("C:\Projeto_Final\Base_Carga\\07jul_u.xlsx",index=False)
jul_nu.to_excel("C:\Projeto_Final\Base_Carga\\07jul_nu.xlsx",index=False)
loadshape_jul_u.to_csv("C:\Projeto_Final\\8500-Node\\07loadshape_jul_u.csv",index=False, header=False)
loadshape_jul_nu.to_csv("C:\Projeto_Final\\8500-Node\\07loadshape_jul_nu.csv",index=False, header=False)

# Valores para agosto
ago = base_.loc[('08')]
ago_u = ago[ago.index.isin(['Útil'], level=1)]
ago_u = ago_u.reset_index(level="Hora")
ago_nu = ago[ago.index.isin(['Não Útil'], level=1)]
ago_nu = ago_nu.reset_index(level="Hora")
loadshape_ago_u = ago_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_ago_nu = ago_nu['CargaMédia_pu'].reset_index(drop=True)
ago_u.to_excel("C:\Projeto_Final\Base_Carga\\08ago_u.xlsx",index=False)
ago_nu.to_excel("C:\Projeto_Final\Base_Carga\\08ago_nu.xlsx",index=False)
loadshape_ago_u.to_csv("C:\Projeto_Final\\8500-Node\\08loadshape_ago_u.csv",index=False, header=False)
loadshape_ago_nu.to_csv("C:\Projeto_Final\\8500-Node\\08loadshape_ago_nu.csv",index=False, header=False)

# Valores para setembro
set = base_.loc[('09')]
set_u = set[set.index.isin(['Útil'], level=1)]
set_u = set_u.reset_index(level="Hora")
set_nu = set[set.index.isin(['Não Útil'], level=1)]
set_nu = set_nu.reset_index(level="Hora")
loadshape_set_u = set_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_set_nu = set_nu['CargaMédia_pu'].reset_index(drop=True)
set_u.to_excel("C:\Projeto_Final\Base_Carga\\09set_u.xlsx",index=False)
set_nu.to_excel("C:\Projeto_Final\Base_Carga\\09set_nu.xlsx",index=False)
loadshape_set_u.to_csv("C:\Projeto_Final\\8500-Node\\09loadshape_set_u.csv",index=False, header=False)
loadshape_set_nu.to_csv("C:\Projeto_Final\\8500-Node\\09loadshape_set_nu.csv",index=False, header=False)

# Valores para Outubro
out = base_.loc[('10')]
out_u = out[out.index.isin(['Útil'], level=1)]
out_u = out_u.reset_index(level="Hora")
out_nu = out[out.index.isin(['Não Útil'], level=1)]
out_nu = out_nu.reset_index(level="Hora")
loadshape_out_u = out_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_out_nu = out_nu['CargaMédia_pu'].reset_index(drop=True)
out_u.to_excel("C:\Projeto_Final\Base_Carga\\10out_u.xlsx",index=False)
out_nu.to_excel("C:\Projeto_Final\Base_Carga\\10out_nu.xlsx",index=False)
loadshape_out_u.to_csv("C:\Projeto_Final\\8500-Node\\10loadshape_out_u.csv",index=False, header=False)
loadshape_out_nu.to_csv("C:\Projeto_Final\\8500-Node\\10loadshape_out_nu.csv",index=False, header=False)

# Valores para Novembro
nov = base_.loc[('11')]
nov_u = nov[nov.index.isin(['Útil'], level=1)]
nov_u = nov_u.reset_index(level="Hora")
nov_nu = nov[nov.index.isin(['Não Útil'], level=1)]
nov_nu = nov_nu.reset_index(level="Hora")
loadshape_nov_u = nov_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_nov_nu = nov_nu['CargaMédia_pu'].reset_index(drop=True)
nov_u.to_excel("C:\Projeto_Final\Base_Carga\\11nov_u.xlsx",index=False)
nov_nu.to_excel("C:\Projeto_Final\Base_Carga\\11nov_nu.xlsx",index=False)
loadshape_nov_u.to_csv("C:\Projeto_Final\\8500-Node\\11loadshape_nov_u.csv",index=False, header=False)
loadshape_nov_nu.to_csv("C:\Projeto_Final\\8500-Node\\11loadshape_nov_nu.csv",index=False, header=False)

# Valores para Dezembro
dez = base_.loc[('12')]
dez_u = dez[dez.index.isin(['Útil'], level=1)]
dez_u = dez_u.reset_index(level="Hora")
dez_nu = dez[dez.index.isin(['Não Útil'], level=1)]
dez_nu = dez_nu.reset_index(level="Hora")
loadshape_dez_u = dez_u['CargaMédia_pu'].reset_index(drop=True)
loadshape_dez_nu = dez_nu['CargaMédia_pu'].reset_index(drop=True)
dez_u.to_excel("C:\Projeto_Final\Base_Carga\\12dez_u.xlsx",index=False)
dez_nu.to_excel("C:\Projeto_Final\Base_Carga\\12dez_nu.xlsx",index=False)
loadshape_dez_u.to_csv("C:\Projeto_Final\\8500-Node\\12loadshape_dez_u.csv",index=False, header=False)
loadshape_dez_nu.to_csv("C:\Projeto_Final\\8500-Node\\12loadshape_dez_nu.csv",index=False, header=False)