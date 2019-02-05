import pandas as pd
import sys, os
import datetime as dt
import gspread  # Access and edit Google Sheets by gspread

# Module to transform gsheets to data frame
import gspread_dataframe as gs_to_df
from oauth2client.service_account import ServiceAccountCredentials
from itertools import combinations
import random
from .variables import *

path = os.path.dirname(__file__)

def authenticate():
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
        os.path.join(path, '../files/cpm_creds.json'), scope)

    gc = gspread.authorize(credentials)

    return gc

def load_workbooks_from_drive():
    """
    Receives: tuple of strings

    Authenticate the access to Google Sheets Feed
    Open all authorized Spreadsheets and puts in a list wkbs
    Filter this list to contain only Spreadsheets with in tuple 'criteria'

    Returns:

    A list with Spreadsheet objects containing one of criteria in the title
    """
    gc = authenticate()

    planilhas = {x.title: x for x in gc.openall()}

    return planilhas

def load_sheet_from_workbook(wkb, aba, col_names=None, skiprows=None, na_values=""):
    """
    Receives:
    wkb: Spreadsheet object (gspread)
    aba: string - name of the worksheet to load from wkb

    Returns a tuple:

    aba as a gspread Worksheet object
    aba as Pandas DataFrame
    """

    wks = wkb.worksheet(aba)

    #col_names = [col for col in wks.row_values(1) if col != '']

    df = gs_to_df.get_as_dataframe(
        wks, evaluate_formulas=True,
        skiprows=skiprows, dtype=str,
        ignore=False, na_values=na_values)

    if col_names is not None:

        df.columns = col_names

    # df.fillna("", inplace=True)

    # df.replace("NaN", "", inplace=True)

    # df.replace('nan', "", inplace=True)

    return (wks, df)

def consolida_teachers(dict_teachers):

    teachers = pd.DataFrame(columns=COLS_TEACHERS)

    for k, (sh, df) in dict_teachers.items():

        print("Processing: ", k)

        df = df[COLS_TEACHERS]

        #df = df.assign(Turma=df.shape[0] * [k.split("_")[-1]])

        teachers = teachers.append(df)

    teachers = teachers[teachers["Nome"] != ""]

    teachers.sort_values(by=["Turma"], inplace=True)

    return teachers

def consolida_alocacao(dict_alocacao):

    alocacao = pd.DataFrame(columns=COLS_ALOCACAO)

    for k, (sh, df) in dict_alocacao.items():

        print("Processing: ", k)

        dict_aloc = {}

        for _, line in df.iterrows():

            dict_aloc["Aula"] = line.Aula
            dict_aloc["Data"] = line.Data
            dict_aloc['Turma'] = k.split("_")[-1]

            for nome in line[2:-1]:

                if nome:

                    dict_aloc["Nome"] = nome

                    alocacao = alocacao.append(dict_aloc, ignore_index=True)

    alocacao["Data"] = alocacao["Data"].apply(transform_date)

    alocacao.sort_values(by=["Data", "Turma"], inplace=True)

    return alocacao

def transform_date(date, str_format="%B %d, %Y"):

    return dt.datetime.strptime(str(date), str_format).date()

def cria_alocacao_geral(lista_aloc):

    alocacao_geral = pd.concat(lista_aloc)

    cols = [col for col in alocacao_geral.columns if "Sugestão" not in col]

    alocacao_geral = alocacao_geral[cols]

    alocacao_geral.fillna("", inplace=True)

    alocacao_geral["Data"] = alocacao_geral["Data"].apply(transform_date)

    alocacao_geral.sort_values(by=["Turma"], inplace=True)

    alocacao_geral.sort_values(by=["Data"], inplace=True)

    return alocacao_geral

def cria_agenda(alocacao, voluntarios):

    agenda = pd.merge(alocacao, voluntarios, on="Nome").dropna().drop(["Turma_y", "#Aulas"], axis=1)

    agenda.columns = COLS_AGENDA

    agenda.sort_values(by=["Data", "Turma"], inplace=True)

    agenda = agenda[COLS_AGENDA]

    return agenda

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

def load_turmas(semestre=None):

    if semestre is None:

        return {x.split('_')[-1]: load_workbooks_from_drive()[x] for x in TURMAS}

    return {k.split('_')[-1]: v  for k,v in load_workbooks_from_drive().items() if semestre in k}

def carrega_listas(turmas=None):

    if turmas is None:
        turmas = load_turmas()

    listas = {}

    for title, turma in turmas.items():

        lista, df = load_sheet_from_workbook(turma, 'Lista de Presença',
                                             col_names=LISTA, skiprows=[1, 2], na_values="")

        df = df[df.Nome != 'nan']

        listas[title] = (lista, df)

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




