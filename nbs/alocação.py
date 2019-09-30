# -*- coding: utf-8 -*-
# ---
# jupyter:
#   jupytext:
#     text_representation:
#       extension: .py
#       format_name: light
#       format_version: '1.4'
#       jupytext_version: 1.2.3
#   kernelspec:
#     display_name: Python [conda env:cpm]
#     language: python
#     name: conda-env-cpm-py
# ---

# +
from context import cpm

import cpm.functions as functions
from cpm.variables import *
import cpm.agenda as agenda

# Every time we change a module they are reload before executing 
# %reload_ext autoreload
# %autoreload 2

AULA = '7'
# -

# ### Carrega a alocação
# Itera todas as planilhas de feedback, carrega a aba de alocação e formata de maneira tabular

alocacao = agenda.carrega_alocacao().dropna(subset=["Nome", "Aula"])
alocacao.head()

# Este método `.dropna(subset=["Nome", "Aula"])` ignora as linhas com a aula ou nome vazios

# ### Lê a agenda para verificar em qual linha termina

agenda_joao = functions.load_df_from_sheet(ALOCACAO, "Agenda").dropna(subset=["Nome"])
agenda_joao.tail()

agenda_joao.shape

linha_vazia = agenda_joao.shape[0] + 2

# ### Filtra a alocação somente para a aula atual

alocacao = alocacao[alocacao["Aula"] == AULA]

functions.salva_aba_no_drive(alocacao, ALOCACAO, "Agenda", row=linha_vazia)



