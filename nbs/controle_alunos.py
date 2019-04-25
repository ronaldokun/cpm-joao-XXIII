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

path = os.path.abspath(os.path.join(os.path.dirname('__file__'), '..'))
sys.path.insert(0, path)

os.chdir(path)

import pandas as pd

from cpm import functions as f

# Every time we change a module they are reload before executing 
# %reload_ext autoreload
# %autoreload 2
# -

listas = f.carrega_listas()

listas[["Total_Homework"]]

notas = ["Desistência", "Evasão"]

alunos = f.load_df_from_sheet("J23_Alunos_Consolidado", "Alunos_2019_1S")

for col in notas:
    df = listas[[col]]
    col_index = alunos.columns.get_loc(col) + 4
    f.atualizar_coluna_df(df, "J23_Alunos_Consolidado", "Alunos_2019_1S", row=2, col=col_index)



# +
trans = listas.Obs != ""

desistências = listas.Desistência == "Sim"

evasão = listas.Evasão == "Sim"
# -

listas[desistências & ~trans][["Nome", "Turma", "Ausencia"]]

listas[evasão & ~ trans][["Nome", "Turma", "Ausencia"]]

listas[~evasão & ~trans & ~desistências].shape

listas[trans]

## dfs = {}
for title, plan in turmas.items():
    dfs[title] = f.load_df_from_sheet("J23_2019_1S_Feedback_"+title, "Lista de Presença",skiprows=[1,2])
    dfs[title].Evasão = col_evasão

for title, plan in turmas.items():
    print(title)
    f.salva_aba_no_drive(dfs[title][["Evasão"]], plan, "Lista de Presença", row=4, col=5)

def atualizar_coluna_df(df: pd.DataFrame, sheet: Union[str, gspread.Spreadsheet], aba: str, row: int, col: int):
    salva_aba_no_drive(df, plan, , aba, row, col)


