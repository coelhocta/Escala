from xlrd import open_workbook, xldate_as_tuple
from datetime import date


def converteData(d):
    data = xldate_as_tuple(d, book.datemode)
    formatado = date(data[0], data[1], data[2])
    return formatado


def geraPeriodo(i, f):
    tmp = {}
    while i <= f:
        tmp.clear()
        if i in roxa:
            quadrinho = 'ROXA'
        elif i in vermelha:
            quadrinho = 'VERMELHA'
        elif i in marrom:
            quadrinho = 'MARROM'
        else:
            quadrinho = 'PRETA'

        tmp['cor'] = quadrinho
        tmp['diaDaSemana'] = diaSemana[date.weekday(i)]
        tmp['diaDoMes'] = i
        resultado.append(tmp.copy())
        periodo.append(f'{quadrinho} - {date.strftime(i, "%d/%m/%Y")} - {diaSemana[date.weekday(i)]}')
        i = date.fromordinal(i.toordinal() + 1)


def geradorNomes():
    for d in range(aba1.nrows):
        nomes.append(aba1.cell_value(rowx=d, colx=0))


def geradorCores():

    for n in range(1, aba2.ncols):
        try:
            roxa.append(converteData(aba2.cell_value(3, n)))
        except:
            print()

        try:
            vermelha.append(converteData(aba2.cell_value(4, n)))
        except:
            print()

        try:
            marrom.append(converteData(aba2.cell_value(5, n)))
        except:
            print()

        try:
            preta.append(converteData(aba2.cell_value(6, n)))
        except:
            print()


def removeRepetidos(entrada):
    limpa = []
    for a in entrada:
        if a not in limpa:
            limpa.append(a)
    return limpa


nomes = []
vermelha = []
preta = []
marrom = []
roxa = []
periodo = []
resultado = []
diaSemana = ['SEGUNDA-FEIRA', 'TERÇA-FEIRA', 'QUARTA-FEIRA', 'QUINTA-FEIRA', 'SEXTA-FEIRA', 'SÁBADO', 'DOMINGO']

book = open_workbook("Escala.xlsx")
aba1 = book.sheet_by_index(0)
aba2 = book.sheet_by_index(1)
aba3 = book.sheet_by_index(2)

inicio = converteData(aba2.cell_value(rowx=1, colx=0))
fim = converteData(aba2.cell_value(rowx=1, colx=1))

geradorNomes()
geradorCores()

n = aba2.ncols

d = inicio
while d <= fim:
    if date.weekday(d) in (5, 6):
        vermelha.append(d)
    vermelha.sort()
    #print(f'{date.strftime(d,"%d/%m/%Y")} - {diaSemana[date.weekday(d)]}')
    d = date.fromordinal(d.toordinal() + 1)

for data in vermelha:
    if date.fromordinal(data.toordinal() - 1) not in vermelha:
        marrom.append(date.fromordinal(data.toordinal() - 1))
        marrom.sort()

for data in roxa:
    if date.fromordinal(data.toordinal() - 1) not in roxa:
        marrom.append(date.fromordinal(data.toordinal() - 1))
        marrom.sort()

#Verificar
for data in preta:
    if date.fromordinal(data.toordinal() - 1) not in roxa:
        preta.append(date.fromordinal(data.toordinal()))
        preta.sort()

for a in roxa:
    if a in vermelha:
        vermelha.remove(a)
    if a in marrom:
        marrom.remove(a)
    if a in preta:
        preta.remove(a)

marrom = removeRepetidos(marrom)
vermelha = removeRepetidos(vermelha)
roxa = removeRepetidos(roxa)
preta = removeRepetidos(preta)

geraPeriodo(inicio, fim)

print(f'Militares: {nomes}')
print(f'Escala Vermelha: {vermelha}')
print(f'Escala Marrom: {marrom}')
print(f'Escala Preta: {preta}')
print(f'Escala Roxa: {roxa}')
print(f'Período: {periodo}')
print(f'Resultado {resultado}')

tmp = {}
escalaRoxa = []
for k, v in enumerate(roxa):
    tmp.clear()
    tmp['Dia'] = v
    tmp['Militar'] = nomes[-k-1]
    escalaRoxa.append(tmp.copy())

escalaVermelha = []
for k, v in enumerate(vermelha):
    tmp.clear()
    tmp['Dia'] = v
    tmp['Militar'] = nomes[-k-1]
    escalaVermelha.append(tmp.copy())

escalaMarrom = []
for k, v in enumerate(marrom):
    tmp.clear()
    tmp['Dia'] = v
    tmp['Militar'] = nomes[-k-1]
    escalaMarrom.append(tmp.copy())

escalaPreta = []
for k, v in enumerate(preta):
    tmp.clear()
    tmp['Dia'] = v
    tmp['Militar'] = nomes[k]
    escalaPreta.append(tmp.copy())


for z in escalaPreta:
    print(z)

print(escalaRoxa[0]['Dia'])
print(resultado[0]['diaDoMes'])


for n in range(len(resultado)):
    d = resultado[n]['diaDoMes']
    for p in range(len(escalaRoxa)):
        rox = escalaRoxa[p]['Dia']
        if rox == d:
            resultado[n]['nome'] = escalaRoxa[p]['Militar']

    for p in range(len(escalaVermelha)):
        ver = escalaVermelha[p]['Dia']
        if ver == d:
            resultado[n]['nome'] = escalaVermelha[p]['Militar']

    for p in range(len(escalaMarrom)):
        mar = escalaMarrom[p]['Dia']
        if mar == d:
            resultado[n]['nome'] = escalaMarrom[p]['Militar']

    for p in range(len(escalaPreta)):
        pre = escalaPreta[p]['Dia']
        if pre == d:
            resultado[n]['nome'] = escalaPreta[p]['Militar']

for l in resultado:
    print(l)