# -*- coding: utf-8 -*-
# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.3'
#       jupytext_version: 0.8.6
#   kernelspec:
#     display_name: Python 3
#     language: python
#     name: python3
# ---

# +
import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('__file__'), '..')))

import pandas as pd

from cpm.variables import *

import cpm.functions as f

import cpm.agenda as a

import gspread

import gspread_dataframe as gs_df

from IPython.display import display

# Every time we change a module they are reload before executing 
# %reload_ext autoreload
# %autoreload 2
# -

# ## Carrega Planilhas

turmas = f.load_turmas()

alocação = f.load_df_from_sheet(ALOCACAO, "Agenda", col_names=COLS_AGENDA + ["Presente"])

alocação.iloc[235]

listas = f.carrega_listas() ; listas = listas[listas.Nome != ""]

listas.columns

listas.head()

# ## Define Aula a ser verificada e Filtra Lista de Presença

n, p, hw, cp = '2', "P2", "HW2", "CP2"

aula = listas[["Turma", "Nome", p, hw, cp]]
aula.tail()

# ## Analisa preenchimento de Aula

turmas = {}
for t in TURMAS:
    turmas[t] = aula[aula.Turma == t]    

# ### Presença em Branco 

for turma, df in turmas.items():
    display(df)  
        #display(alocação[alocação["Aula"] == n])

alocação.columns

# +
desist = (~listas.Desistência.isnull()) & (listas.Presença != '0')
obs = listas.Obs.isnull()

listas.loc[(desist & obs), ["Turma", "Nome", "Presença", "Obs"]]
# -

listas.loc[listas.Evasão == "Sim",["Turma", "Nome", "Presença"]]

# +
def preenche_lacunas(planilhas_de_presença):

    for k, (sh, df) in planilhas_de_presença.items():

        print("Processing: {}\n".format(k))

        print(110*"=")

        for item in f.AULAS:

            if isinstance(item, tuple):

                p, hw, cp = item

                t1 = ((df[p] == "NO") & (df[hw] == ""))

                t2 = ((df[p] == "NO") & (df[cp] == ""))                
                                
                if df.loc[t2].shape[0]:
                    
                    display(df.loc[t2])

                df.loc[t1, hw] = "N"
                
                if "SPK" in cp:
                    
                    df.loc[t2, cp] = "0"
                    
                else:

                    df.loc[t2, cp] = "N/A"
                    
                df.loc[df[hw] == "+/-", hw] = "'+/-"
                
                df.loc[df[hw] == "½", hw] = "'+/-"
                
                df.loc[df[hw] == "1/2", hw] = "'+/-"
        
        #corrige_preenchimento(sh, df, 7, 32)

# +
alunos = pd.DataFrame(columns=["Turma", "Nome", "P_15", "Nota_Final"])

for k, (sh, df) in planilhas_de_presença.items():
    print(k)
    p = ((df["P_15"] != "YES") & (df["Nome"] != ""))
    df["Turma"] = k
    subset = df.loc[:, ["Nome", "P_15","Nota_Final", "Turma"]][p]
    alunos = alunos.append(subset)    
# -

# ### Checar Presença

iterator = iter(planilhas_de_presença.items())

# +
k, (sh, df) = next(iterator)

print("Turma: " + k)
for item in f.AULAS:

    if isinstance(item, tuple) and len(item) == 3:

        p, hw, cp = item

        t1 = ((df[p] == "NO") & (df[hw] == ""))

        t2 = ((df[p] == "YES") & (~df[hw].isin(["Y", "N", "+/-", "N/H"])))

        t3 = ((df[p] == "NO") & (df[cp] == ""))

        t4 = ((df[p] == "YES") & (~df[cp].isin(["A", "B", "C", "N/A"])))
        
        
        if df.loc[t1].shape[0]:

            print('Falta sem nota de Homework\n')

            display(df.loc[t1, ["Nome",p,hw,cp]])

            df.loc[t1, hw] = "N/H"

        if df.loc[t2].shape[0]:

            print("Presença sem Nota de Homework\n")

            display(df.loc[t2, ["Nome",p,hw,cp]])

            df.loc[t2, hw] = "Y"

        if df.loc[t3].shape[0]:

            print("Falta sem nota de Participação\n")

            display(df.loc[t3, ["Nome",p,hw,cp]])

            if "SPK" not in cp:

                df.loc[t3, cp] = "N/A"
            

        if df.loc[t4].shape[0]:

            print("Presença sem nota de Participação")

            display(df.loc[t4, ["Nome",p,hw,cp]])

            if "SPK" not in cp:

                df.loc[t4, cp] = "A"
                

            
#t5 = ((df['P_8'] == "") | (df['P_8'] == 'NO') & (df['Nota_Mid'] != '0.00'))
              
#t6 = ((df['P_15'] == "") | (df['P_15'] == 'NO') & (df['Nota_Final'] != '0.00'))
                
                
#if df.loc[t5].shape[0]:

#    print(k)

#    display(df.loc[t5, ["Nome",p, "Nota_Mid"]])

#    df.loc[t5, ["P_8", "Nota_Mid"]] = ["NO", "0.00"]

#if df.loc[t6].shape[0]:

#    print(k)

#    display(df.loc[t6, ["Nome",p, "Nota_Final"]])

#    df.loc[t6, ["P_15", "Nota_Final"]] = ["NO", "0.00"]

df["Nota_Mid"] = df["Nota_Mid"].str.replace(",", ".")
df["Nota_Final"] = df["Nota_Final"].str.replace(",", ".")


corrige_preenchimento(sh, df, 7, 54)           

# +
notas_dos_alunos = {}

for nome, turma in turmas.items():

    sh, df = f.load_sheet_from_workbook(turma, 'Notas dos Alunos', skiprows=[1,2,3,4])
    
    display(df)

    lista = pd.DataFrame(columns=df.columns)

    registros = sh.get_all_values()

    #lista = f.nomeia_cols_lista(lista)

    notas_dos_alunos[nome.split("_")[-1]] = df
    
    break

# -

df.head()

# +
colunas_1 = ["Nome", "P_8", "Nota_Mid"]

colunas_2 = colunas_1 + ["Turma"]

notas_mid = pd.DataFrame(columns=colunas_2)

for k, (sh, df) in iterator:
    
    df["Turma"] = k
        
    notas_mid = notas_mid.append(df[df["Nome"] != ""][colunas_2])     
# -

notas_mid.to_csv("Relatorio_Mid.csv", index=False)

# +
aula = "Aula 11"

colunas = ["Nome", "Nota_Mid"]

k, (sh, df) = next(iterator)
    
print("Processing: {}\n".format(k))

filtro = (agenda["Aula"].str.contains(aula)) & (agenda["Turma"] == k)

professores = agenda[filtro]

print("Professores: \n", professores[["Nome", "Email CPM", "Email Pessoal"]])

criterio = (df["Nome"] != "")

display(df[criterio][colunas])
# -

agenda[agenda["Aula"].str.contains("11")]

alocacao[alocacao["Aula"].str.contains("12")]

for key, (sh, df) in planilhas_de_presença.items():
    print(key)
    display(df[["P_9", "HW_9", "CP_9"]])

presença = {}
for sh, wks in zip(planilhas, listas_de_presenças):
    presença[sh.title] = wks[1].range("Q4:S25")

presença = {}
for sh, wks in zip(planilhas, listas_de_presenças): 
    presença[sh.title] = wks[1]

for plan in presença.values():
    plan["Aula 4 (Speaking)"]

for key, value in presença.items():
    print(key)
    print("\n")
    for v in value:
        print(v.value)
    print("\n")

desistencia = pd.DataFrame(columns = listas_de_presenças[0].columns)
for lista in listas_de_presenças:
    desistencia = lista['Desistência'] == 'Sim'
    print(lista[desistencia][['Código', 'NOME COMPLETO', "Desistência", "Evasão"]])

# +
presença = []

for lista in listas_de_presenças:
    print("Planilha: {}, Coluna BC: {}".format(lista[0].title, lista[0].range("BC1:BC25")))        
# -

presença[0]

# +
count = []

for turma in presença:
    for i, aluno in enumerate(turma[0]):
        if aluno.value == r'':
            next
        elif turma[1][i].value == r'0%':
            count.append(aluno.value)            
# -

count

counts = agenda.iloc[:184]["Nome"].value_counts()

counts

# #### Merge the Teachers Info from all Spreadsheets into one

# +
#for i in skills[0]:
#    key = aula.cell(i, 1).value
#    value = aula.cell(i, 2).value
#    dict_[key] = value
    
#print(dict_exs)
# -


