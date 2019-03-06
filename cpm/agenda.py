import os, sys
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('__file__'), '..')))

import pandas as pd
import datetime as dt
from . import functions as f
from . import variables as v


def carrega_alocacao()-> dict: 
    """Carrega as planilhas de alocação de voluntários e retorna como um DataFrame
    
    Returns:
        dict: dicionário -- turma=DataFrame com alocação
    """
    aloc = {}
    
    for turma in v.TURMAS:
        
        aloc[turma.split("_")[-1]] = f.load_df_from_sheet(turma, v.ABA_ALOC)
        
    aloc["EA"] = f.load_df_from_sheet(v.ALOCACAO, v.ABA_ALOC_EAS)
    
    return aloc

def carrega_voluntarios()-> pd.DataFrame:
    """Carrega as informações dos voluntários do CPM, tanto Teachers quanto EAS;.
    
    Returns:
        pandas.DataFrame: Relação de Voluntários.
    """

    teachers = f.load_df_from_sheet(v.ALOCACAO, v.ABA_TEACHERS)

    eas = f.load_df_from_sheet(v.ALOCACAO, v.ABA_EAS)

    return teachers.append(eas)

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

    alocacao = pd.DataFrame(columns=v.COLS_ALOCACAO)

    for k, df in dict_alocacao.items():

        print("Processing: ", k)

        dict_aloc = {}

        for _, line in df.iterrows():

            dict_aloc["Aula"] = line.Aula
            dict_aloc["Data"] = line.Data
            dict_aloc['Turma'] = k

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

    agenda = pd.merge(alocacao, voluntarios) #, on="Nome").dropna().drop(["Turma_y", "#Aulas"], axis=1)

    agenda = agenda[v.COLS_AGENDA]

    agenda.sort_values(by=["Data", "Turma"], inplace=True)

    return agenda
