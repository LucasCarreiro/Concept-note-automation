import pandas as pd

class Aluno:
    AV1 = 0
    AV2 = 0
    AV3 = 0
    AV4 = 0

def media_by_bimestre(students_data):

    BIM1 = []
    BIM2 = []
    BIM3 = []

    for value in students_data.values:
        if value[3] == "1º BIM":
            aluno = procurarAluno(BIM1, value[0], value[2])
            if aluno is None:
                aluno = criarAluno(value)
                BIM1.append(aluno)
            setarNota(value[1],value[4],aluno)
        elif value[3] == "2º BIM":
            aluno = procurarAluno(BIM2, value[0], value[2])
            if aluno is None:
                aluno = criarAluno(value)
                BIM2.append(aluno)
            setarNota(value[1], value[4], aluno)
        elif value[3] == "3º BIM":
            aluno = procurarAluno(BIM3, value[0], value[2])
            if aluno is None:
                aluno = criarAluno(value)
                BIM3.append(aluno)
            setarNota(value[1], value[4], aluno)

    criarExcel(BIM1, "1º BIM")
    criarExcel(BIM2, "2º BIM")
    criarExcel(BIM3, "3º BIM")

    return "Finalizado"

def criarExcel(BIM, nomeBim):

    RA = []
    DESCPROVA = []
    DISCIPLINA = []
    MEDIA = []
    ETAPADESCRICAO = []
    CODTURMA = []
    SERIE = []
    CONCEITO = []

    for aluno in BIM:

        RA.append(aluno.RA)
        DESCPROVA.append(aluno.DESCPROVA)
        DISCIPLINA.append(aluno.DISCIPLINA)
        average_in_the_bimester_by_discipline = ((aluno.AV1 * 2) + (aluno.AV2 * 3) + (aluno.AV3 * 1) + (
                aluno.AV4 * 1)) / 7

        if average_in_the_bimester_by_discipline < 0.5:
            CONCEITO.append("F")
        elif 0.5 <= average_in_the_bimester_by_discipline < 0.75:
            CONCEITO.append("D-")
        elif 0.75 <= average_in_the_bimester_by_discipline < 1.25:
            CONCEITO.append("D")
        elif 1.25 <= average_in_the_bimester_by_discipline < 1.5:
            CONCEITO.append("D+")
        elif 1.5 <= average_in_the_bimester_by_discipline < 1.75:
            CONCEITO.append("C-")
        elif 1.75 <= average_in_the_bimester_by_discipline < 2.25:
            CONCEITO.append("C")
        elif 2.25 <= average_in_the_bimester_by_discipline < 2.5:
            CONCEITO.append("C+")
        elif 2.5 <= average_in_the_bimester_by_discipline < 2.75:
            CONCEITO.append("B-")
        elif 2.75 <= average_in_the_bimester_by_discipline < 3.2:
            CONCEITO.append("B")
        elif 3.2 <= average_in_the_bimester_by_discipline < 3.4:
            CONCEITO.append("B+")
        elif 3.4 <= average_in_the_bimester_by_discipline < 3.6:
            CONCEITO.append("A-")
        elif 3.6 <= average_in_the_bimester_by_discipline < 3.75:
            CONCEITO.append("A")
        elif 3.75 <= average_in_the_bimester_by_discipline:
            CONCEITO.append("A+")
        else:
            print(average_in_the_bimester_by_discipline)

        MEDIA.append(average_in_the_bimester_by_discipline)

        ETAPADESCRICAO.append(aluno.ETAPADESCRICAO)
        CODTURMA.append(aluno.CODTURMA)
        SERIE.append(aluno.SERIE)

    alunos = {"RA": RA, "DESCPROVA": DESCPROVA, "DISCIPLINA": DISCIPLINA, "MEDIA": MEDIA, "CONCEITO": CONCEITO,
              "ETAPADESCRICAO": ETAPADESCRICAO, "CODTURMA": CODTURMA, "SERIE": SERIE}

    dataframe = pd.DataFrame(alunos)
    dataframe.to_excel('media_alunos_conceito' + nomeBim + '.xlsx', sheet_name=nomeBim)

    print("Terminado: " + nomeBim)


def setarNota(DESCPROVA,NOTAFINAL, aluno):
    if DESCPROVA == 'PI':
        aluno.AV1 = NOTAFINAL
    elif DESCPROVA == 'ST':
        aluno.AV2 = NOTAFINAL
    elif DESCPROVA == 'BQ':
        aluno.AV3 = NOTAFINAL
    elif DESCPROVA == 'SP':
        aluno.AV4 = NOTAFINAL

def criarAluno(value):
    aluno = Aluno()
    aluno.RA = value[0]
    aluno.DESCPROVA = value[1]
    aluno.DISCIPLINA = value[2]
    aluno.ETAPADESCRICAO = value[3]
    aluno.CODTURMA = value[5]
    aluno.SERIE = value[6]

    return aluno


def procurarAluno(lista_alunos, RA, DISCIPLINA):
    for aluno in lista_alunos:
        if aluno.RA == RA:
            if aluno.DISCIPLINA == DISCIPLINA:
                return aluno

if __name__ == '__main__':
    data_students = pd.read_excel('Boletim.XLSX', engine='openpyxl',
                                  usecols='E,K,L,N,O,X,AE')

    media_by_bimestre(data_students)
