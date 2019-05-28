import openpyxl
from datetime import date


def data_num(d):
    num = date.toordinal(d)
    return num


def num_data(d):
    num = date.fromordinal(d)
    return num


def gera_nomes():
    lin = aba_mil_ind.max_row
    col = aba_mil_ind.max_column
    for i in range(1, lin + 1):
        nomes.append(aba_mil_ind.cell(i, 1).value)
    return nomes


def gera_periodo():
    i = data_num(aba_per_fer['A2'].value)
    f = data_num(aba_per_fer['B2'].value)
    while i <= f:
        periodo.append(i)
        i += 1


def gera_quadrinho():
    # Busca Roxa da Planilha
    for r in aba_per_fer['B4':'AZ4']:
        for c in r:
            if c.value != None:
                data = data_num(c.value)
                roxa.append(data)

    # Busca Vermelha da Planilha
    for r in aba_per_fer['B5':'AZ5']:
        for c in r:
            if c.value != None:
                data = data_num(c.value)
                vermelha.append(data)

    # Busca Marrom da Planilha
    for r in aba_per_fer['C6':'AZ6']:
        for c in r:
            if c.value != None:
                data = data_num(c.value)
                marrom.append(data)

    # Gera vermelha e Preta Automática
    for d in periodo:
        if date.weekday(num_data(d)) in (5, 6) and d not in vermelha and d not in roxa:
            vermelha.append(d)
        vermelha.sort()

    # Gera Marrom Automática
    for d in vermelha:
        if (d -1) not in vermelha and (d -1) not in roxa:
            marrom.append(d - 1)
        for d in roxa:
            if (d - 1) not in roxa:
                marrom.append(d - 1)
        marrom.sort()

    # Gera Preta Automática
    for d in periodo:
        if d not in vermelha and d not in roxa and d not in marrom:
            preta.append(d)


##########################################
# Lê o nome e as abas da planilha
wb = openpyxl.load_workbook('Escala.xlsx')
aba_mil_ind = wb['Militares.Indisponibilidades']
aba_per_fer = wb['Periodo.Feriados']
aba_ver = wb['Vermelha']
aba_pre = wb['Preta']
aba_mar = wb['Marrom']
aba_rox = wb['Roxa']

##########################################
# Gera os dias da Semana
diaSemana = ['SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO', 'DOMINGO']

##########################################
# Listas
nomes = []
periodo = []
vermelha = []
marrom = []
roxa = []
preta = []

##########################################
# Chamadas Funções
gera_nomes()
gera_periodo()
gera_quadrinho()

print(f'Nomes: {nomes}')
print(f'Período: {periodo}')
print(f'Roxa: {roxa}')
print(f'Vermelha: {vermelha}')
print(f'Marrom: {marrom}')
print(f'Preta: {preta}')

for a in periodo:
    if a in roxa:
        cor = 'Roxa'
    if a in vermelha:
        cor = 'Vermelha'
    if a in marrom:
        cor = 'Marrom'
    if a in preta:
        cor = 'Preta'
    print(f'{cor:>8} - {diaSemana[date.weekday(num_data(a))]:^13} - {num_data(a)}')

'''
a1 = aba_mil_Ind['A1']
a2 = aba_mil_Ind['A2']
a3 = aba_mil_Ind.cell(3, 1)

print(a1.value)
print(a2.value)
print(a3.value)

# mostra o número de linhas da planilha aba_mil_ind
print(aba_mil_ind.max_row)

# mostra o número de Colunas da planilha aba_mil_ind
print(aba_mil_ind.max_column)

# mostra o conteúdo da célula A1 até B10
for r in aba_mil_ind['A1':'B10']:
    for c in r:
        print(c.value)

# Converte data da planilha
import datetime
ws['A2'] = datetime.datetime.now()

#####################################################

# Edita a nova planilha
planilha = openpyxl.Workbook()
aba =  wb['Militares.Indisponibilidades'] # Seleciona a aba a ser modificada
aba.title = 'Aba1' # Altera o nome da aba selecionada
aba['A3'].value = 'Teste de escrita' # Escreve na célula escolhida
aba.create_sheet(title='NovaAba' # Cria uma nova aba
wb.save('Escala.final.xlsx') # Salva uma cópia com o novo nome ad planilha
'''