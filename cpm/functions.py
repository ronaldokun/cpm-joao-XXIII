import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('__file__'), '..')))

import datetime as dt
import gspread  # Access and edit Google Sheets by gspread
import pandas as pd

# Module to transform gsheets to data frame
import gspread_dataframe as gs_to_df
from oauth2client.service_account import ServiceAccountCredentials
from itertools import combinations
import random
from .variables import *
from typing import Sequence


path = os.path.dirname(__file__)

def authenticate(file: str ='../files/cpm_creds.json'):
    """
    Read the json file from google API and authenticate the access to the Google Sheets

    The Spreadsheets on the feed are the ones shared with the email created by Google API (json files)

    Return:

    The object to access the Spreadsheets
    """

    scope = ['https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']

    # Configurations necessary to gspread to work
    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        os.path.join(path, file), scope)

    gc = gspread.authorize(credentials)

    return gc

def load_workbooks_from_drive():
    """
    Receives: None

    Authenticate the access to Google Sheets Feed
    Open all authorized Spreadsheets and puts in a list wkbs
    Filter this list to contain only Spreadsheets with in tuple 'criteria'

    Returns:

    A list with Spreadsheet objects containing one of criteria in the title
    """
    gc = authenticate()

    planilhas = {x.title: x for x in gc.openall()}

    return planilhas

def load_sheet_from_drive(title: str)-> gspread.Spreadsheet:
    """Load Spreadsheet object of name <title> 
    
    Args:
        title (str): The Google Sheets' Name

    Raises:
        ValueError: If the Google Sheets is not found or not shared with the current user.
    
    Returns:
        gspread.Spreadsheet: The Google Sheet object
    """

    gc = authenticate()

    try:
        sh = gc.open(title=title)
    except gspread.SpreadsheetNotFound as e:
        raise ValueError(f"The Google Sheet {title} not found or not shared with this user. Check your Drive") from e

    return sh

def load_wb_from_sheet(title: str, aba: str):
    """Load Workbook <aba> from Spreadsheet <title>
    
    Args:
        title (str): The Google Sheets' Name
        aba (str): The Workbooks' Name
    Raises:
        ValueError: If the Workbook is not found in the Google Sheets.
    
    Returns:
        gspread.Worksheet: The Google Worksheet object
    """
    
    sh = load_sheet_from_drive(title=title)

    try:
        aba = sh.worksheet(title=aba)
    except gspread.WorksheetNotFound as e:
        raise ValueError(f"A aba {aba} não foi encontrada na planilha {title}. Verifique o seu drive") from e
    
    return aba

def load_df_from_sheet(title: str,aba: str, col_names: Sequence=None, skiprows: Sequence=None, na_values:Sequence=""):
    """Load Workbook <aba> from Spreadsheet <title> and returns it as a Dataframe
    
    Args:
        title (str): The Google Sheets' Name
        col_names ([type], optional): Defaults to None. List of columns' names
        skiprows ([type], optional): Defaults to None. List of rows to skip
        na_values (str, optional): Defaults to "". What values in the df to consider as NaN
    
    Returns:
        pandas.DataFrame: The Google Worksheet as a DataFrame
    """
    
    aba = load_wb_from_sheet(title=title, aba=aba)

    df = gs_to_df.get_as_dataframe(
        aba, evaluate_formulas=True,
        parse_dates=True,
        skiprows=skiprows, dtype=str,
        ignore=False, na_values=na_values)

    if col_names is not None:
        l1, l2 = len(col_names), len(df.columns)
        
        assert l1 == l2, f"There is a mismatch between the len(col_names): {l1} and the len(columns): {l2} in the DataFrame"
        df.columns = col_names



    return df.dropna(axis=0, how="all")

def load_turmas(semestre=None):

    sheets = load_workbooks_from_drive()

    if semestre is None:

        return {x.split('_')[-1]: sheets[x] for x in TURMAS}

    return {k.split('_')[-1]: v  for k,v in sheets.items() if semestre in k}

def salva_aba_no_drive(dataframe, planilha_drive, aba_drive, row=2, col=1, header=False, resize=False, clear=False):

    if isinstance(planilha_drive, gspread.models.Worksheet):

        workbook = planilha_drive

    elif isinstance(planilha_drive, str):

        workbook = authenticate().open(planilha_drive)

    worksheet = workbook.worksheet(aba_drive)

    if clear: worksheet.clear()


    gs_to_df.set_with_dataframe(dataframe=dataframe,
                                worksheet=worksheet,
                                row=row,
                                col=col,
                                include_column_header=header,
                                resize=resize)

def nomeia_cols_lista(df):

    Aulas = []

    # for i in range(1, 16):

    #    Aulas += [x + str(i) for x in AULAS]

    Aulas += LISTA[0:7]

    for aula in AULAS:

        Aulas += list(aula)

    df = df.iloc[:, 0:54]

    df.columns = Aulas

    return df

def carrega_listas(turmas=None):

    if turmas is None:
        turmas = TURMAS

    listas = pd.DataFrame(columns=LISTA)  

    for turma in turmas:

        df = load_df_from_sheet(turma, 'Lista de Presença',
                                             col_names=LISTA, skiprows=[1, 2], na_values="")

        df = df[df.Nome != 'nan']

        df.Turma = turma.split("_")[-1] 

        listas.append(df, ignore_index=True)

    return listas

def checa_presença(aula, lista):

    aula = list(aula)

    filtro = (lista.Nome != "")

    print(lista.aula.notnull().any().any())

def relatorio_alunos():

    #relatorio = pd.DataFrame(columns=)

    listas = {k:v[1] for k,v in carrega_listas()}

    for turma, lista in listas.items():
        rule = (lista.Nome != "") & (lista.Total_Presença < 0.3)
        lista["Turma"] = turma
        relatorio = relatorio.append(lista.loc[rule, ["Turma", "Nome", "Total_Presença"]])




