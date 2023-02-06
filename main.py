from calendar import monthrange
import datetime
import random
import os
import sys


def generate_entry_minutes(qtd: int, init: int, end: int) -> [int]:
    minutes_list = []
    for _ in range(qtd):
        minute = random.randint(init, end)
        if len(minutes_list) != 0:
            while minute == minutes_list[-1]:
                minute = random.randint(init, end)

        minutes_list.append(minute)

    return minutes_list


def generate_departure_minutes(minutes_list: [int], variation: {int}) -> [int]:
    departure_minutes_list = []
    for minute in minutes_list:
        departure_minute = minute + random.randint(variation[0], variation[1])
        if len(departure_minutes_list) != 0:
            while departure_minute == departure_minutes_list[-1]:
                departure_minute = minute + random.randint(0, 3)

        departure_minutes_list.append(departure_minute)

    return departure_minutes_list


try:
    MONTH_NUMBER = int(input('Digite o mês para geração dos pontos (ex: 3): '))

    INIT_HOUR_RANGE = int(input("Digite a hora que você costuma entrar (ex: 10): "))
    INIT_MIN_RANGE = input("Digite o range de minutos que você costuma entrar (ex: 0, 45): ")
    INIT_MIN_RANGE = INIT_MIN_RANGE.strip().replace(",", " ").split()
    INIT_MIN_RANGE = [int(item) for item in INIT_MIN_RANGE]

    MIDDLE_HOUR_RANGE = int(input("Digite a hora que você costuma sair para almoçar (ex: 12): "))
    MIDDLE_MIN_RANGE = input("Digite o range de minutos que você costuma sair para almoçar (ex: 0, 35): ")
    MIDDLE_MIN_RANGE = MIDDLE_MIN_RANGE.strip().replace(",", " ").split()
    MIDDLE_MIN_RANGE = [int(item) for item in MIDDLE_MIN_RANGE]

    NUMBER_OF_DAYS = monthrange(datetime.date.today().year, MONTH_NUMBER)[1]
except Exception as e:
    print(f'Erro: {e}')
    sys.exit()

try:
    minutos_entrada = generate_entry_minutes(NUMBER_OF_DAYS, INIT_MIN_RANGE[0], INIT_MIN_RANGE[1])
    minutos_almoco = generate_entry_minutes(NUMBER_OF_DAYS, MIDDLE_MIN_RANGE[0], MIDDLE_MIN_RANGE[1])
    minutos_volta_almoco = generate_departure_minutes(minutos_almoco, (0,3))
    minutos_saida = generate_departure_minutes(minutos_entrada, (0,7))

    pontos_list = []
    for i in range(NUMBER_OF_DAYS):
        date = datetime.date(datetime.date.today().year, MONTH_NUMBER, i+1)

        if date.weekday() in [0, 1, 2, 3, 4]:
            entrada = datetime.time(INIT_HOUR_RANGE, minutos_entrada[i], 0, 0)
            almoco = datetime.time(MIDDLE_HOUR_RANGE, minutos_almoco[i], 0, 0)
            volta_almoco = datetime.time(MIDDLE_HOUR_RANGE + 1, minutos_volta_almoco[i], 0, 0)
            saida = datetime.time(INIT_HOUR_RANGE + 9, minutos_saida[i], 0, 0)

            pontos_list.append((entrada, almoco, volta_almoco, saida))
        else:
            pontos_list.append(None)
except Exception as e:
    print(f'Erro: {e}')
    sys.exit()

try:
    with open('planilha_pontos.csv', 'w+') as file:
        for ponto in pontos_list:
            if ponto is not None:
                file.write((';'.join([str(item) for item in ponto]) + '\n'))
            else:
                file.write("\n")

    os.system("start EXCEL.EXE planilha_pontos.csv")
except Exception as e:
    print(f'Erro: {e}')
    sys.exit()


