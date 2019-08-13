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
import os
import sys

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('__file__'), '..')))

import pandas as pd

from pathlib import Path

from cpm import functions as f

import gspread

from gspread import *

import gspread_dataframe as gs_df

#import openpyxl as xl

from IPython.display import display

# Every time we change a module they are reload before executing 
# %reload_ext autoreload
# %autoreload 2
# -

# ## Load Workbooks

planilhas = f.load_workbooks_from_drive()

planilhas

turmas = {k:v for k,v in planilhas.items() if '2019-1S' in k}

turmas = f.load_turmas()

turmas

# +
for title, sh in turmas.items():    
    print(title)
    wb = sh.worksheet("Lista de Presença")
    cells_1, cells_2 = wb.range("AP4:AP28"), wb.range("AQ4:AQ28")
    cells = wb.range("AP4:AQ28")
    for i, (c1, c2) in enumerate(zip(cells_1, cells_2), 4):
        c1.value = fr'=if(AO{i}="Absent", "Undone", "")'
        c2.value = fr'=if(AO{i}="Absent", 0, "")'
    #wb.update_cell(3,2, r"='Parâmetros'!$A$14")
    #wb.update_cell(6,2, r"=Class_Plan!$F$13")
    wb.update_cells(cells_1, "USER_ENTERED")
    wb.update_cells(cells_2, "USER_ENTERED")
    
#while i < 421:
#    wb.update_cell(i, 1, fr"=image(Info_Students!$K${j})")
#    i += 18
#    j += 1
# -




config = planilhas['Configuração_Planilhas'].worksheet("Class_Plan")

i = 6

# +
wb = sh.worksheet(f'Aula {i}')

aula = wb.get_all_values()

Theme = aula[5][1]

config.update_cell(i, 1, i)

config.update_cell(i, 6 , Theme)

i += 1
# -

aula[-3][0].split("\n")[1][23:54]

# ## ABA Parâmetros

param = {}
for title, wb in turmas.items():
    param[title] = f.load_sheet_from_workbook(wb, "Voluntarios_Geral", skiprows=1)

# ### Exportar Configurações do Semestre

config = f.load_sheet_from_workbook(planilhas['Configuração_Semestre'], 'Parâmetros', skiprows=1)


txt = r'=importrange("https://docs.google.com/spreadsheets/d/1zg8roK0-EFySIZivBaHkLeNcVsNJ1t41YrpMzlli6vQ", "voluntarios_2018_2S")'

for title, (sh, _) in param.items():
    print(title)
    sh.clear()
    sh.update_cell(1, 1, txt)

# ## Load Student Sheets

students = {}
for title, wb in turmas.items():
    students[title] = f.load_sheet_from_workbook(wb, 'Students')

# ## Load Listas de Presença

listas =  f.carrega_listas()

s = r'"Aula "'
for sheet, _ in listas.values():
    #sheet.update_acell("R1", fr"=CONCATENATE({s},'Parâmetros'!$B7 ," ",'Parâmetros'!$C7)")
    for i in range(4, 29):


# +
from time import sleep

sheet = students["A1"][0]

c = 2
for index in range(5, sheet.row_count+1, 18):
    print(index)
    i = index
    #cell_list = sheet.range(f"E{i}:F{i+3}")

    sheet.update_cell(i, 1, fr"=Info_Alunos!$B${c}")
    #sheet.cell(i, 1).value = fr"=Info_Alunos!$B${c}"
    i += 1
    sleep(2)
    sheet.update_cell(i, 6, fr"=Info_Alunos!$B${c}")
    #sheet.cell(i, 6).value = fr"=Info_Alunos!$B${c}"
    i += 1
    sleep(2)
    sheet.update_cell(i, 6, fr"=Info_Alunos!$D${c}")
    #sheet.cell(i, 6).value = fr"=Info_Alunos!$D${c}"    
    i += 1
    sleep(2)
    sheet.update_cell(i, 6, fr"=Info_Alunos!$F${c}")
    #sheet.cell(i, 6).value = fr"=Info_Alunos!$F${c}"
    i += 1
    sleep(2)
    sheet.update_cell(i, 6, fr"=Info_Alunos!$G${c}")
    #sheet.cell(i, 6).value = fr"=Info_Alunos!$G${c}"
    i += 1
    sleep(2)
    sheet.update_cell(i, 6, fr"=Info_Alunos!$A${c}")
    #sheet.cell(i, 6).value = fr"=Info_Alunos!$A${c}"
    i += 5 
    sleep(2)
    sheet.update_cell(i, 5, 'Attendance')
    sleep(2)
    #sheet.cell(i, 5).value = 'Attendance'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$B${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$B${c}"
    i +=1
    sheet.update_cell(i, 5, 'Homework')
    sleep(2)
    #sheet.cell(i, 5).value = 'Homework'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$C${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$C${c}"
    i +=1
    sheet.update_cell(i, 5, 'English Usage')
    sleep(2)
    #sheet.cell(i, 5).value = 'English Usage'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$D${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$D${c}"
    i +=1
    sheet.update_cell(i, 5, 'Speaking')
    sleep(2)
    #sheet.cell(i, 5).value = 'Speaking'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$E${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$E${c}"
    i +=1
    sheet.update_cell(i, 5, 'Midterm')
    sleep(2)
    #sheet.cell(i, 5).value = 'Midterm'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$F${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$F${c}"
    i +=1
    sheet.update_cell(i, 5, 'Final Exam')
    sleep(2)
    #sheet.cell(i, 5).value = 'Final Exam'
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$G${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$G${c}"
    i +=1
    sheet.update_cell(i, 5, "Final Grade")
    sleep(2)
    #sheet.cell(i, 5).value = "Final Grade"
    sheet.update_cell(i, 6, fr"='Notas dos Alunos'!$I${c}")
    sleep(2)
    #sheet.cell(i, 6).value = fr"='Notas dos Alunos'!$I${c}"
    #sheet.update_cells(cell_list)
    c += 1
    #sleep(100)

    #sheet.update_cells(sheet.range(f"A1:F{sheet.row_count}"))

# -

# ## Corrige células nulas das planilhas de presença

# +
p = ["P" + str(i) for i in range(1, 13)]

h = [c for c in df.columns if "HW" in c and "_" not in c]

cp = [c for c in df.columns if "CP" in c]
 
p, h, cp
# -

df.columns

# +
name = "A1"
df = listas[name][1]

cols = range(7,42)

df.iloc[:, cols].fillna("")
df.iloc[:, cols]
# -

df[p]

# +
#df[["P7", "P8", "P9", "P12"]] = df[["P7", "P8", "P9", "P12"]].replace("nan", "Ausente")
df[presença] = df[presença].replace("", "Ausente")

df[presença]
# -

df[h]

#df.HW9 = df.HW9.replace("Não Fez", "Não Houve")
#df[homework] = df[homework].replace('', "Não Fez")
df[h]

df[cp]

df[["SPK1", 'SPK2']]

for presença, participação in zip([i for i in p if i not in ["P1", "P6", "P7", "P12"]], cp):
    df[participação][df[presença] == "Ausente"] = ""
    df[participação][df[presença] == "Presente"] = 'Good'

# +
df["SPK1"][df['P6'] == "Ausente"] = 0.00
#df["SPK2"][df['P12'] == "Ausente"] = 0.00


df['Nota_Mid'][df['P7'] == "Ausente"] = 0.00
# -

df[["Nome", "Nota_Final"]][df.Nota_Final != 'nan']

f.salva_aba_no_drive(df.iloc[:, 7:40], 'J23_2018-2S_Feedback_'+name, aba_drive="Lista de Presença", row=4, col=8)

for title, wb in wbs.items():
    if "2018-2S_Feedback" in title:
        sheet = wb.worksheet('Students')
        #print(dir(sheet))
        for i in range(15,sheet.row_count+1, 18):
            print(sheet.cell(i, 5).value)
        break    

values[13]

templates = [wb for wb in wbs if "2018-2S" in wb.title]
templates

ws = templates[0]

delete = ["LEIA-ME", "Teachers", "Alocação", "Infos Matricula", "Students", "Lista_de_Presença"]

for ws in templates:
    for sheet in ws.worksheets():
        if sheet.title in delete:
            try:
                ws.del_worksheet(sheet)
            except Exception as e:
                print(repr(e))


