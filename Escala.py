import openpyxl
from datetime import date


def data_num(d):
    num = date.toordinal(d)
    return num


def num_data(d):
    num = date.fromordinal(d)
    return num


def gera_nomes():
    tmp = {}
    indisp = []
    lin = aba_inicio.max_row
    col = aba_inicio.max_column
    for i in range(8, lin - 8):
        tmp['Antig'] = i -8
        tmp['Nome'] = aba_inicio.cell(i, 1).value
        for c in range(2, col):
            d = aba_inicio.cell(i, c).value
            if d != None:
                e = data_num(d)
                indisp.append(e)
        tmp['Indisp'] = indisp.copy()
        indisp.clear()
        nomes.append(tmp.copy())
        tmp.clear()
    return nomes


def gera_periodo():
    i = data_num(aba_inicio['B1'].value)
    f = data_num(aba_inicio['C1'].value)
    while i <= f:
        periodo.append(i)
        i += 1


def gera_quadrinho():
    # Busca da Planilha
    for t in cores:
        for r in t['dias']:
            for c in r:
                if c.value != None:
                    data = data_num(c.value)
                    if data in periodo:
                        t['cor'].append(data)

    # Gera vermelha Automática
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


def fila(cor):
    fila = []
    z = {'roxa':lastro_roxa, 'vermelha':lastro_vermelha, 'marrom':lastro_marrom, 'preta':lastro_preta}
    for a in z[cor]:
        b = [a[0], a[1]]
        fila.append(b)
    fila.reverse()
    fila = sorted(fila, key=lambda x: x[1])
    return fila

'''
def busca_lastro_planilha():
    tmp = []
    tmp1 = []
    for a in cores:
        for i in range(3,a['linhas'] + 1):
            antiguidade = i - 3
            tmp.append(antiguidade)
            for j in range(1,(a['colunas'])+1):
                conteudo = a['conteudo'](row=i, column=j).value
                if conteudo != None:
                    if type(conteudo) is not str:
                        conteudo = data_num(conteudo)
                    tmp.append(conteudo)
            tmp1.append(tmp.copy())
            tmp.clear()
        for c in tmp1:
            c = [c, len(c)]
            a['lastro'].append(c)
        tmp1.clear()
'''


def busca_lastro_planilha():
    tmp = {}
    tmp1 = []
    for a in cores:
        for i in range(3,a['linhas'] + 1):
            tmp['cor'] = a['cor_texto']
            tmp['antig'] = i - 3
            tmp['nome'] = a['conteudo'](row=i, column=1).value
            for j in range(1,(a['colunas'])+1):
                conteudo = a['conteudo'](row=i, column=j+1).value
                if conteudo != None:
                    if type(conteudo) is not str:
                        conteudo = data_num(conteudo)
                    tmp1.append(conteudo)
            tmp['lastros'] = tmp1.copy()
            tmp['lastro_total'] = len(tmp1)
            a['lastro'].append(tmp.copy())
            tmp1.clear()
        tmp.clear()


def preenche_from_planilha():
    # Busca escala forçada da planilha
    tmp = {}
    for f in cores:
        for a in f['lastro']:
            for b in a['lastros']:
                if b in f['cor'] and b in periodo:
                    tmp['cor'] = f['cor_texto']
                    tmp['diaSemana'] = diaSemana[date.weekday(num_data(b))]
                    tmp['dia'] = b
                    tmp['nome'] = a['nome']
                    tmp['antig'] = a['antig']
                    escala_final.append(tmp.copy())
                    tmp.clear()
##########################################
# Lê o nome e as abas da planilha


wb = openpyxl.load_workbook('Escala.xlsx')
aba_inicio = wb['Inicio']
aba_ver = wb['Vermelha']
aba_pre = wb['Preta']
aba_mar = wb['Marrom']
aba_rox = wb['Roxa']

##########################################
# Gera os dias da Semana

##########################################
# Listas
nomes = []
periodo = []
vermelha = []
marrom = []
roxa = []
preta = []
lastro_roxa = []
lastro_vermelha = []
lastro_marrom = []
lastro_preta = []
escala_final = []
diaSemana = ['SEG', 'TER', 'QUA', 'QUI', 'SEX', 'SÁB', 'DOM']
cores = [{'cor_texto': 'ROXA','dias': aba_inicio['B3':'AZ3'], 'cor':roxa, 'linhas': aba_rox.max_row, 'colunas':aba_rox.max_column, 'conteudo': aba_rox.cell, 'lastro':lastro_roxa},
         {'cor_texto': 'VERMELHA','dias': aba_inicio['B4':'AZ4'], 'cor':vermelha, 'linhas': aba_ver.max_row, 'colunas':aba_ver.max_column, 'conteudo': aba_ver.cell, 'lastro':lastro_vermelha},
         {'cor_texto': 'MARROM','dias': aba_inicio['C5':'AZ5'], 'cor':marrom, 'linhas': aba_mar.max_row, 'colunas':aba_mar.max_column, 'conteudo': aba_mar.cell, 'lastro':lastro_marrom},
         {'cor_texto': 'PRETA','dias': aba_inicio['C6':'AZ6'], 'cor':preta, 'linhas': aba_pre.max_row, 'colunas':aba_pre.max_column, 'conteudo': aba_pre.cell, 'lastro':lastro_preta}]

##########################################
# Chamadas Funções
gera_nomes()
gera_periodo()
gera_quadrinho()
busca_lastro_planilha()
preenche_from_planilha()

##########################################
# Gerar lista sequencia da fila

#fila_roxa = fila('roxa')
#fila_vermelha = fila('vermelha')
#fila_marrom = fila('marrom')
#fila_preta = fila('preta')

##########################################

for a in vermelha:
    for b in escala_final:
        if a == b['dia']:
            vermelha.remove(a)
'''
cont = 0
for a in vermelha:

    tmp = {'cor': 'VERMELHA', 'diaSemana': diaSemana[date.weekday(num_data(a))], 'dia': a, 'nome':''}
    tmp['nome'] = fila_vermelha[cont][0][1]
    tmp['antig'] = fila_vermelha[cont][0][0]
    escala_final.append(tmp.copy())
    tmp.clear()
    fila_vermelha = fila('vermelha')
    cont += 1
'''
'''
    cont = 0
    cont_nomes = 0
    while tmp['nome'] == '':
        while cont_nomes <= len(nomes)-1:
            if nomes[cont_nomes]['Antig'] == fila_vermelha[cont][0][0] and a not in nomes[cont_nomes]['Indisp']:
                if escala_final:
                    for z in escala_final:
                        if z['antig'] == nomes[cont_nomes]['Antig']:
                            if (a == z['dia'] + 2) or (a == z['dia'] + 1) or (a == z['dia'] - 2) or (a == z['dia'] - 1) or (a == z['dia']):
                                break
                            else:
                                tmp['nome'] = fila_vermelha[cont][0][1]
                                tmp['antig'] = fila_vermelha[cont][0][0]
                                lastro_vermelha[nomes[cont_nomes]['Antig']][1] += 1
                                escala_final.append(tmp.copy())
                                cont = 0
                                break
                        else:
                            tmp['nome'] = fila_vermelha[cont][0][1]
                            tmp['antig'] = fila_vermelha[cont][0][0]
                            lastro_vermelha[nomes[cont_nomes]['Antig']][1] += 1
                            fila_vermelha = fila('vermelha')
                            escala_final.append(tmp.copy())
                            cont = 0
                            break
                else:
                    tmp['nome'] = fila_vermelha[cont][0][1]
                    tmp['antig'] = fila_vermelha[cont][0][0]
                    lastro_vermelha[nomes[cont_nomes]['Antig']][1] += 1
                    fila_vermelha = fila('vermelha')
                    escala_final.append(tmp.copy())
                    cont = 0
                    break
            cont_nomes += 1
        cont += 1
    tmp.clear()
'''
for a in periodo:
    for b in escala_final:
        if a == b['dia']:
            print(f'{b["cor"]:8} - {b["diaSemana"]:^5} - {num_data(b["dia"])} - {b["nome"]:10} - Antig: {b["antig"]}')

'''





lastro_roxa[25][1] += 1
lastro_roxa[31][1] += 1

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

# Mostra as células A3 até C10
for c1, c2, c3 in aba_rox['A3': 'C10']:
    print("{} {} {}".format(c1.value, c2.value, c3.value))


#####################################################

# Edita a nova planilha
planilha = openpyxl.Workbook()
aba =  wb['Militares.Indisponibilidades'] # Seleciona a aba a ser modificada
aba.title = 'Aba1' # Altera o nome da aba selecionada
aba['A3'].value = 'Teste de escrita' # Escreve na célula escolhida
aba.create_sheet(title='NovaAba' # Cria uma nova aba
wb.save('Escala.final.xlsx') # Salva uma cópia com o novo nome ad planilha
'''