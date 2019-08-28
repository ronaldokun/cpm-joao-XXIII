# -*- coding: utf-8 -*-
# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.4'
#       jupytext_version: 1.2.1
#   kernelspec:
#     display_name: Python [conda env:cpm]
#     language: python
#     name: conda-env-cpm-py
# ---

# +
from context import cpm


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

# +
#turmas = f.load_turmas("J23_2019_2S_Feedback")

prefix = r"https://drive.google.com/open?id="

for k,v in turmas.items():
    print(k+"\t", prefix+v.id)

# +
datas = """Data
04/08/2019
18/08/2019
25/08/2019
01/09/2019
15/09/2019
22/09/2019
29/09/2019
06/10/2019
20/10/2019
27/10/2019
03/11/2019
10/11/2019
24/11/2019
01/12/2019
08/12/2019
15/12/2019""".split("\n")

for name, sh in turmas.items():
    print(name)    
    wb = f.load_wb_from_sheet(sh.title, "Alocação")
    
    # Select a range
    cell_list = wb.range('A1:A17')

    for cell, value in zip(cell_list, datas):
        cell.value = value

    # Update in batch
    wb.update_cells(cell_list)
# -

for name, sh in turmas.items():
    print(name)    
    wb = f.load_wb_from_sheet(sh.title, "Alocação")
    
    # Select a range
    cell_list = wb.range('C3:C17')

    for cell in cell_list:
        cell.value = fr"='Parâmetros'!C{cell.row}"        

    # Update in batch
    wb.update_cells(cell_list, "USER_ENTERED")
    
    wb.update_acell("C2", "CANCELADO")

sheets = f.load_workbooks_from_drive()
sheets

# +
exp = "\'\"&turma&\"\'"
value = fr'=query(IMPORTRANGE(link_alunos, range_alunos), "select Col2,Col3,Col4,Col6,Col9,Col10,Col11,Col12,Col13,Col14, Col15 where Col1 = {exp}")'

for name, sh in turmas.items():
    print(name)
    wb = f.load_wb_from_sheet(sh.title, "Info_Students")
    wb.update_cell(1, 1, value)    
# -

alocação = f.load_df_from_sheet(ALOCACAO, "Agenda", col_names=COLS_AGENDA + ["Presente"])

alocação.head()

listas = f.carrega_listas()
listas.head()

listas = listas[listas.Evasão == "Não"]
listas = listas[listas.Desistência == ""]
listas = listas[listas.Obs == ""]

listas.head()

# ## Define Aula a ser verificada e Filtra Lista de Presença

def checa_coluna(lista: pd.DataFrame,  coluna: str)-> bool:
    return sum(lista[coluna] == "") > 0 

def checa_aula(lista: pd.DataFrame, aula: tuple)-> bool:
    check = 0
    for col in aula:
        if "SPK" not in col:
            check += int(checa_coluna(lista, col))
    return not bool(check)

for turma in TURMAS:
    df = listas[listas.Turma == turma]
    for i, aula in enumerate(AULAS[10:11]):
        if not checa_aula(df, aula):
            #display(alocação[alocação.Aula == str(i + 1)])
            display(df[["Nome", "Turma"] + list(aula)])

alocação[alocação.Aula == "1"]

for t in TURMAS:
    df = listas[listas.Turma ==t]
    if checa_coluna(listas[listas.Turma ==t], p):
        display(df[["Turma", "Nome", p]])

listas[listas.Turma == "A2"]

# ## Analisa preenchimento de Aula

turmas = {}
for t in TURMAS:
    turmas[t] = aula[aula.Turma == t]    

# ### Presença em Branco 

for df in turmas.values():
    print(df[p] == False)  
        #display(alocação[alocação["Aula"] == n])

df.p

# +
labels = ["B1", "D1", "F1", "Teens2"]

for t in labels:
    display(alocação[alocação.Turma == t])

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


