import pandas as pd

import locale
try:
    locale.setlocale(locale.LC_ALL, 'pt_BR')
except:
    locale.setlocale(locale.LC_ALL, 'Portuguese_Brazil')

file = "mensal-data.xlsx"
df = pd.read_excel(file, sheet_name = "Planilha1",)


maior = [0, 0]
menor = [[], 0]
media = []

for i, row in df.iterrows():
    som = 0
    id = row[0]

    for c in range(1, len(row)):
        som += row[c]

    if som > maior[1]:
        maior[0] = id
        maior[1] = som

    if som < menor[1] or i == 0:
        menor[0] = [id]
        menor[1] = som
    elif som == menor[1]:
        menor[0].append(id)

    media.append([id, som/(len(row)-1)])


maiorMes = [0, 0]
menorMes = [0, 0]
ocupacao = []

for index, column in df.iteritems():
    if index != "Unnamed: 0":
        rent = 0
        som = 0

        for i, row in df.iterrows():
            som += row[index]

            if row[index] > 0:
                rent += 1

        if som > maiorMes[1]:
            maiorMes[0] = index
            maiorMes[1] = som

        if som < menorMes[1] or menorMes[1] == 0:
            menorMes[0] = index
            menorMes[1] = som

        ocupacao.append(rent/500)


print(f'- O Maior faturamento foi do ID \033[30:42m{maior[0]}\033[m em um total de \033[4mR${maior[1]:.2F}\033[m')

print()

print(f'- O Menor faturamento foi do ID \033[30:41m{menor[0][0]}\033[m em um total de \033[4mR${menor[1]:.2F}\033[m')
print(f'\033[m{"":<4}Entretanto os seguintes IDs obtiveram o mesmo faturamento(R$0.00):')
for c in menor[0]:
    if c != menor[0][0]:
        print(f'{"":<6}\033[7m{c}\033[m')

print()

print(f'- IDs e suas respectivas médias:')
for c in media:
    print(f'{"":<4}\033[7m{c[0]}\033[m: \033[4mR${c[1]:.2f}\033[m')

print()

print(f'- O Melhor mês foi \033[30:42m{maiorMes[0].strftime("%B")}\033[m com \033[4mR${maiorMes[1]:.2f}\033[m de faturamento')
print(f'- O Pior mês foi \033[30:41m{menorMes[0].strftime("%B")}\033[m com \033[4mR${menorMes[1]:.2f}\033[m de faturamento')

print()

print(f'- A taxa de ocupação de \033[30:42mJaneiro\033[m de 2020 foi de \033[4m{ocupacao[5]*100:.1f}%\033[m')
print(f'- A taxa de ocupação de \033[30:41mOutubro\033[m de 2019 foi de \033[4m{ocupacao[2]*100:.1f}%\033[m')