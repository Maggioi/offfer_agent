cena = 12321321321


cena = str(cena)

i = 1
cena_tablica = []

for _ in range(len(cena)):
    if i == 4:
        cena_tablica.append(" ")
    cena_tablica.append(cena[_])
    i += 1

for _ in cena_tablica:
    print(_, end="")

print()