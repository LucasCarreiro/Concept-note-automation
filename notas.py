import pandas as pd


def media_by_bimestre(students_data, bim: str):
    students = []
    RA = []
    DESCPROVA = []
    DISCIPLINA = []
    MEDIA = []
    ETAPADESCRICAO = []
    CODTURMA = []
    SERIE = []
    CONCEITO = []
    AV1_BIM1 = 0
    AV2_BIM1 = 0
    AV3_BIM1 = 0
    AV4_BIM1 = 0
    for value in students_data.values:
        if value[3] == bim:
            if str(value[0]) + value[2] + value[3] not in students:
                for value2 in students_data.values:
                    if value2[0] == value[0]:
                        if value2[3] == bim:
                            if value2[2] == value[2]:
                                if value2[1] == 'PI':
                                    AV1_BIM1 = value2[4]
                                if value2[1] == 'ST':
                                    AV2_BIM1 = value2[4]
                                if value2[1] == 'BQ':
                                    AV3_BIM1 = value2[4]
                                if value2[1] == 'SP':
                                    AV4_BIM1 = value2[4]
                if value[3] == bim:
                    average_in_the_bimester_by_discipline = ((AV1_BIM1 * 2) + (AV2_BIM1 * 3) + (AV3_BIM1 * 1) + (
                            AV4_BIM1 * 1)) / 7
                    
                    if (average_in_the_bimester_by_discipline <=0):
                        CONCEITO.append("F")
                    elif (average_in_the_bimester_by_discipline >0 and average_in_the_bimester_by_discipline <=1):
                        CONCEITO.append("D")
                    elif (average_in_the_bimester_by_discipline >1 and average_in_the_bimester_by_discipline <=2):
                        CONCEITO.append("C")
                    elif (average_in_the_bimester_by_discipline >2 and average_in_the_bimester_by_discipline <=3):
                        CONCEITO.append("B")
                    elif (average_in_the_bimester_by_discipline >3 and average_in_the_bimester_by_discipline <=4):
                        CONCEITO.append("A")
                        
                    RA.append(value[0])
                    DESCPROVA.append(value[1])
                    DISCIPLINA.append(value[2])
                    MEDIA.append(average_in_the_bimester_by_discipline)
                    ETAPADESCRICAO.append(value[3])
                    CODTURMA.append(value[5])
                    SERIE.append(value[6])
                    students.append(str(value[0]) + value[2] + value[3])
                    print(bim)

    alunos = {"RA": RA, "DESCPROVA": DESCPROVA, "DISCIPLINA": DISCIPLINA, "MEDIA": MEDIA, "CONCEITO": CONCEITO,
              "ETAPADESCRICAO": ETAPADESCRICAO, "CODTURMA": CODTURMA, "SERIE": SERIE}

    dataframe = pd.DataFrame(alunos)
    dataframe.to_excel('MEDIA_CONCEITO_GREAT_'+bim+'.xlsx', sheet_name=bim)

    return "sucesso " + bim


if __name__ == '__main__':
    data_students = pd.read_excel('Boletim.XLSX', engine='openpyxl',
                                  usecols='E,K,L,N,O,X,AE')

    BIM1 = '1º BIM'
    BIM2 = '2º BIM'
    BIM3 = '3º BIM'

    result_bim1 = media_by_bimestre(data_students, BIM1)
    print(result_bim1)
    result_bim2 = media_by_bimestre(data_students, BIM2)
    print(result_bim2)
    result_bim3 = media_by_bimestre(data_students, BIM3)
    print(result_bim3)
