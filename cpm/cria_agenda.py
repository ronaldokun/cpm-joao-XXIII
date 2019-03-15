import sys

import functions
from cpm.variables import *
import agenda

def main(aula, row):

    print("Lendo Alocação de Professores e EAs\n")

    alocacao = agenda.carrega_alocacao()

    print("Lendo Informações dos EAs e Teachers\n")

    voluntarios = agenda.carrega_voluntarios()

    print("\nConsolidando Alocação Geral")

    alocacao = agenda.cria_agenda(alocacao, voluntarios)

    alocacao = alocacao[alocacao["Aula"] == str(aula)]

    print("Salvando Alocação no Drive")

    #row = functions.load_df_from_sheet(ALOCACAO, ABA_AGENDA, col_names=COLS_AGENDA + ["Presente?"]).shape[0] + 3
    
    functions.salva_aba_no_drive(alocacao, ALOCACAO, 'Agenda', row=int(row))

if __name__ == "__main__":

    main(*sys.argv[1:])
