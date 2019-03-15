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

import pandas as pd

import re
# Access and edit Google Sheets by gspread
import gspread

# Module to transform gsheets to data frame
import gspread_dataframe as gs_to_df

from oauth2client.service_account import  ServiceAccountCredentials

import datetime as dt

from pathlib import *

import sys

path = PurePath('__file__')
sys.path.insert(0, str(Path(path.parent).resolve().parent))

from cpm import functions as f


TEMPLATE = "Feedback_Template"

MATRICULA = "3. Planilha Matrículas 2019 - 1o sem"

MATR_ABA = "João XXIII"

MATR_CLEANED = "J23_Matrícula_2019-1S"


def check_feedback(gc, name):

    aloc = gc.open(name)

    # Convert gsheet to df
    aloc = gs_to_df.get_as_dataframe(aloc, dtype=str)


    # Transform String Dates to datetime
    f = lambda x : dt.datetime.strptime(x, "%d/%m/%Y")

    aloc['Data'] = aloc['Data'].map(f)

    # correct 'nan' strings to ''
    aloc.replace('nan', '', inplace=True)

def split_date_hour(col):

    return pd.Series(col.split(" "))

def concat_names(x,y):

    return x + " " + y

def split_celfone(col):

    if type(col) == str:

        pattern = ".*\(.*(\d{2})\).*(\d{5})(\d{4}).*"

        split = re.split(pattern, col)

        if len(split) >= 4:

            return "(" + split[1] + ")" + " "  + split[2] + "-" + split[3]

        return col

    return col

def split_fone(col):

    if type(col) == str:

        pattern = ".*\(.*(\d{2})\).*(\d{4}|\d{4})(\d{4}).*"

        split = re.split(pattern, col)

        if len(split) >= 4:

            return "(" + split[1] + ")" + " "  + split[2] + "-" + split[3]

        return col

    return col

def preprocess_df(df):

    presencial = df["Data/Hora preenchimento"] == "Presencial"

    espera = df["Data/Hora preenchimento"] == "Lista de Espera"

    pre = df[~ presencial & ~ espera]["Data/Hora preenchimento"]

    data_hora = pre.apply(split_date_hour)

    data = pd.Series.append(df[presencial]["Data/Hora preenchimento"],
                            df[espera]["Data/Hora preenchimento"])

    data = data.append(data_hora.iloc[:, 0]).sort_index()

    hora = pd.Series.append(df[presencial]["Data/Hora preenchimento"],
                            df[espera]["Data/Hora preenchimento"])

    hora = hora.append(data_hora.iloc[:, 1]).sort_index()

    df.rename(columns={"Data/Hora preenchimento": "Data_Pré_Matrícula"},
              inplace=True)

    df["Data_Pré_Matrícula"] = data

    df["Hora_Pré_Matrícula"] = hora

    df["Nome"] = df["Nome"].apply(str.upper).apply(str.strip)

    df["Sobrenome"] = df["Sobrenome"].apply(str.upper).apply(str.strip)

    df["Nome Responsável"] = df["Nome Responsável"].apply(str.upper).apply(str.strip)

    df["Sobrenome Responsável"] = df["Sobrenome Responsável"].apply(str.upper).apply(str.strip)

    df["Nome Responsável"] = concat_names(df["Nome Responsável"],
                                          df["Sobrenome Responsável"])

    del df["Sobrenome Responsável"]

    df["Nome"] = concat_names(df["Nome"], df["Sobrenome"])

    del df["Sobrenome"]

    df.rename(columns={"Telefone Celular ex: (011) 00000-0000": "Tel_Celular"},
              inplace=True)

    df["Tel_Celular"] = df["Tel_Celular"].apply(split_celfone)

    df.rename(columns={"Telefone Fixo ex: (011) 000-0000": "Tel_Fixo"},
              inplace=True)

    df["Tel_Fixo"] = df["Tel_Fixo"].apply(split_fone)

    df.rename(columns={"Celular do Responsável": "Celular_Responsável"},
              inplace=True)

    df["Celular_Responsável"] = df["Celular_Responsável"].apply(split_celfone)

    df.rename(columns={"RG \n(apenas números)" : "RG"}, inplace=True)

    return df

def main():

    gc = f.authenticate()
    
    wb = f.load_workbooks_from_drive()[MATRICULA]
    
    df = f.load_sheet_from_workbook(wb, MATR_ABA, skiprows=[1,2])[1]

    df = df.fillna('')

    df = preprocess_df(df)

    df.to_csv("matricula.csv", sep=",", index=False, columns=COLUNAS, na_rep='')

    df = pd.read_csv("matricula.csv", dtype=str, na_values='')

    matricula = gc.open(MATR_CLEANED)

    wks = matricula.worksheet("JoãoXXIII")

    wks.clear()

    gs_to_df.set_with_dataframe(worksheet=wks, dataframe=df)

main()

planilhas




