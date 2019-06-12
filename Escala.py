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
        tmp['antig'] = i -8
        tmp['nome'] = aba_inicio.cell(i, 1).value
        for c in range(2, col):
            d = aba_inicio.cell(i, c).value
            if d != None:
                e = data_num(d)
                indisp.append(e)
                indisp.append(e + 1)
                indisp.append(e - 1)
        tmp['indisp'] = indisp.copy()
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

    # Gera Marrom Automática
    for d in vermelha:
        dia = d - 1
        if dia not in vermelha and dia not in roxa and dia in periodo:
            marrom.append(dia)
    for d in roxa:
        dia = d - 1
        if dia not in roxa and dia in periodo:
            marrom.append(dia)
    # Gera Preta Automática
    for d in periodo:
        if d not in vermelha and d not in roxa and d not in marrom:
            preta.append(d)


def fila_ver():
    fila = []
    for a in lastro_vermelha:
        c = [len(a['lastros']), a['antig'], a['nome']]
        fila.append(c)
    fila.reverse()
    fila = sorted(fila, key=lambda x: x[0])
    return fila


def fila_mar():
    fila = []
    for a in lastro_marrom:
        c = [len(a['lastros']), a['antig'], a['nome']]
        fila.append(c)
    fila.reverse()
    fila = sorted(fila, key=lambda x: x[0])
    return fila


def fila_pre():
    fila = []
    for a in lastro_preta:
        c = [len(a['lastros']), a['antig'], a['nome']]
        fila.append(c)
    fila.reverse()
    fila = sorted(fila, key=lambda x: x[0])
    return fila


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
aba_escala = wb['Escala']

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
diaSemana = ['SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO', 'DOMINGO']
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
for a in nomes:
    a['lastro_total'] = []
    for b in lastro_roxa:
        if a['antig'] == (b['antig']):
            for c in b['lastros']:
                a['lastro_total'].append(c)
    for b in lastro_vermelha:
        if a['antig'] == (b['antig']):
            for c in b['lastros']:
                a['lastro_total'].append(c)
    for b in lastro_marrom:
        if a['antig'] == (b['antig']):
            for c in b['lastros']:
                a['lastro_total'].append(c)
    for b in lastro_preta:
        if a['antig'] == (b['antig']):
            for c in b['lastros']:
                a['lastro_total'].append(c)

vermelha_copy = vermelha.copy()
marrom_copy = marrom.copy()
preta_copy = preta.copy()

for a in escala_final:
    if a['dia'] in preta_copy:
        preta_copy.remove(a['dia'])
    if a['dia'] in vermelha_copy:
        vermelha_copy.remove(a['dia'])
    if a['dia'] in marrom_copy:
        marrom_copy.remove(a['dia'])

cont = 0
for a in marrom_copy:
    fila_marrom = fila_mar()
    tmp = {'cor': 'MARROM', 'diaSemana': diaSemana[date.weekday(num_data(a))], 'dia': a, 'nome': ''}
    while True:
        for b in nomes:
            if b['antig'] == fila_marrom[cont][1]:
                if a not in b['lastro_total'] \
                        and a - 1 not in (b['lastro_total']) \
                        and a + 1 not in (b['lastro_total']) \
                        and a + 2 not in (b['lastro_total']) \
                        and a - 2 not in (b['lastro_total'])\
                        and a not in b['indisp']:
                    tmp['nome'] = fila_marrom[cont][2]
                    tmp['antig'] = fila_marrom[cont][1]
                    lastro_marrom[b['antig']]['lastros'].append(a)
                    b['lastro_total'].append(a)
                    escala_final.append(tmp.copy())
                    tmp.clear()
                    cont = 0
                    break
                else:
                    cont += 1
        if not tmp:
            break

cont = 0
for a in vermelha_copy:
    fila_vermelha = fila_ver()
    tmp = {'cor': 'VERMELHA', 'diaSemana': diaSemana[date.weekday(num_data(a))], 'dia': a, 'nome': ''}
    while True:
        for b in nomes:
            if b['antig'] == fila_vermelha[cont][1]:
                if a not in b['lastro_total'] \
                        and a - 1 not in (b['lastro_total']) \
                        and a + 1 not in (b['lastro_total']) \
                        and a + 2 not in (b['lastro_total']) \
                        and a - 2 not in (b['lastro_total'])\
                        and a not in b['indisp']:
                    tmp['nome'] = fila_vermelha[cont][2]
                    tmp['antig'] = fila_vermelha[cont][1]
                    lastro_vermelha[b['antig']]['lastros'].append(a)
                    b['lastro_total'].append(a)
                    escala_final.append(tmp.copy())
                    tmp.clear()
                    cont = 0
                    break
                else:
                    cont += 1
        if not tmp:
            break

cont = 0
for a in preta_copy:
    fila_preta = fila_pre()
    tmp = {'cor': 'PRETA', 'diaSemana': diaSemana[date.weekday(num_data(a))], 'dia': a, 'nome': ''}
    while True:
        for b in nomes:
            if b['antig'] == fila_preta[cont][1]:
                if a not in b['lastro_total'] \
                        and a - 1 not in (b['lastro_total']) \
                        and a + 1 not in (b['lastro_total']) \
                        and a + 2 not in (b['lastro_total']) \
                        and a - 2 not in (b['lastro_total'])\
                        and a not in b['indisp']:
                    tmp['nome'] = fila_preta[cont][2]
                    tmp['antig'] = fila_preta[cont][1]
                    lastro_preta[b['antig']]['lastros'].append(a)
                    b['lastro_total'].append(a)
                    escala_final.append(tmp.copy())
                    tmp.clear()
                    cont = 0
                    break
                else:
                    cont += 1
        if not tmp:
            break

for a in periodo:
    for b in escala_final:
        if a == b['dia']:
            print(f'{b["cor"]:8} - {b["diaSemana"]:^15} - {num_data(b["dia"])} - {b["nome"]:10}')

escala_planilha = [(), (), ('Data', 'Dia da Semana', 'Militar', 'Cor')]
for a in periodo:
    for b in escala_final:
        if a == b['dia']:
            tmp = (str(date.strftime(num_data(b["dia"]), "%d/%m/%Y")), str(b["diaSemana"]), str(b["nome"]), str(b['cor']))
            escala_planilha.append(tmp)


for a in escala_planilha:
    aba_escala.append(a)
wb.save('Escala.final.xlsx')

