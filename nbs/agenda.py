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
import sys, os
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname('__file__'), '..')))

import cpm.functions as functions
import cpm.agenda as agenda
from cpm.variables import *
# -

aloc = agenda.carrega_alocacao()

voluntarios = agenda.carrega_voluntarios()

aloc = agenda.consolida_alocacao(aloc)

alocacao = agenda.cria_agenda(aloc, voluntarios)

alocacao = alocacao[alocacao.Aula == "5"]

row = functions.load_df_from_sheet(ALOCACAO, ABA_AGENDA, col_names=COLS_AGENDA + ["Presente?"]).shape[0] + 3


row

functions.salva_aba_no_drive(alocacao, ALOCACAO, 'Agenda', row=row)


