def create_pair_teachers(teachers):
    """
    Receive the DataFrame teachers with a list of teacher alocated to a given class.

    It assumes it has at least the columns ["Novato", "Teacher"] on it

    Returns:

    A list with all pairs of teachers in this order "Mixed", "Veteran", "Novice"
    """

    teachers_df.sort_values(by='Novato', ascending=False)

    novatos = list(teachers_df["Teacher"][teachers_df["Novato"] == 'Sim'])

    veteranos = list(teachers_df["Teacher"][teachers_df["Novato"] == 'Não'])

    duplas = set(combinations(teachers_df["Teacher"], 2))

    duplas_novatos = {i for i in duplas if (
        i[0] in novatos and i[1] in novatos)}

    duplas_veteranos = {i for i in duplas if (
        i[0] in veteranos and i[1] in veteranos)}

    duplas_mistas = list(duplas.difference(
        duplas_novatos.union(duplas_veteranos)))

    duplas_novatos = list(duplas_novatos)

    duplas_veteranos = list(duplas_veteranos)

    duplas = []

    if duplas_mistas:

        random.shuffle(duplas_mistas)

        duplas += duplas_mistas

    if duplas_veteranos:

        random.shuffle(duplas_veteranos)

        duplas += duplas_veteranos

    if duplas_novatos:

        random.shuffle(duplas_veteranos)

        duplas += duplas_novatos

    return duplas

def alocar_professores(teachers, pares, num_aulas):

    count = {teacher: 0 for teacher in teachers["Teacher"]}

    alocacao = []

    while len(alocacao) < num_aulas:

        for pick in pares:

            test = False

            if count[pick[0]] == min(count.values()) and count[pick[1]] == min(count.values()):

                test = True

            elif (count[pick[0]] == min(count.values()) and count[pick[1]] == min(count.values()) + 1) or\
                 (count[pick[1]] == min(count.values()) and count[pick[0]] == min(count.values()) + 1):

                test = True

            if test:

                if len(alocacao):

                    if pick[0] in alocacao[-1] or pick[1] in alocacao[-1]:

                        next

                    else:

                        alocacao.append(pick)

                        count[pick[0]] += 1

                        count[pick[1]] += 1

                else:

                    alocacao.append(pick)

                    count[pick[0]] += 1

                    count[pick[1]] += 1

            if len(alocacao) == num_aulas:
                break

    return alocacao

