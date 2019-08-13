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

#path = os.path.abspath(os.path.join(os.path.dirname('__file__')), '..')
path = '/home/ronaldo/projects/programming/cpm'
#sys.path.insert(0, path)

os.chdir(path)

import pandas as pd

from cpm import functions as f

from cpm.variables import *

# Every time we change a module they are reload before executing 
# %reload_ext autoreload
# %autoreload 2
# -

path

listas = f.carrega_listas()

turmas = f.load_turmas()

notas = f.carrega_notas()

notas

alunos = f.load_df_from_sheet("J23_Alunos_Consolidado", "Alunos_2019_1S")

listas[["Total_Homework"]]

notas = ["Desistência", "Evasão"]

for col in notas:
    df = listas[[col]]
    col_index = alunos.columns.get_loc(col) + 4
    f.atualizar_coluna_df(df, "J23_Alunos_Consolidado", "Alunos_2019_1S", row=2, col=col_index)

def atualizar_coluna_df(df: pd.DataFrame, sheet: Union[str, gspread.Spreadsheet], aba: str, row: int, col: int):
    salva_aba_no_drive(df, plan, , aba, row, col)


