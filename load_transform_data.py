import popups

def load_data(sheet, nazwa_oferty, i):
    """Loading data from the calculator"""

    # Loading data from row nr i
    dlugosc = sheet[f'F{i}'].value        
    wysokosc = sheet[f'G{i}'].value      
    akustyka = sheet[f'J{i}'].value       
    parkowanie_raw = sheet[f'H{i}'].value 
    liczba_modulow = sheet[f'N{i}'].value
    plyta_raw = sheet[f'P{i}'].value
    klasa_palnosci_raw = sheet[f'K{i}'].value
    ukryte_krawedzie_raw = sheet[f'L{i}'].value
    kolor_toru_raw = sheet[f'O{i}'].value
    szklo_raw = sheet[f'R{i}'].value 
    liczba_DE = sheet[f'T{i}'].value
    liczba_DE2 = sheet[f'U{i}'].value
    polautomat_raw = sheet[f'V{i}'].value
    dodatkowy_tor_raw = sheet[f'W{i}'].value
    obnizenie_raw = sheet[f'X{i}'].value
    lakierowane_profile_raw = sheet[f'Y{i}'].value
    cena = sheet[f'AF{i}'].value
    liczba_scian = sheet[f'C{i}'].value
    cena_wszystkich = sheet[f'AF2'].value


    data = {
        "dlugosc" : dlugosc,
        "wysokosc" : wysokosc,
        "akustyka" : akustyka,
        "parkowanie_raw" : parkowanie_raw,
        "liczba_modulow": liczba_modulow,
        "plyta_raw" : plyta_raw,
        "klasa_palnosci_raw" : klasa_palnosci_raw,
        "ukryte_krawedzie_raw" : ukryte_krawedzie_raw,
        "kolor_toru_raw" : kolor_toru_raw,
        "szklo_raw" : szklo_raw,
        "liczba_DE" : liczba_DE,
        "liczba_DE2" : liczba_DE2,
        "polautomat_raw" : polautomat_raw,
        "dodatkowy_tor_raw" : dodatkowy_tor_raw,
        "obnizenie_raw" : obnizenie_raw,
        "lakierowane_profile_raw" : lakierowane_profile_raw,
        "cena" : cena,
        "liczba_scian" : liczba_scian,
        "cena_wszystkich" : cena_wszystkich,
        "nazwa_oferty" : nazwa_oferty
    }
    return data

def transform_data(data):
    '''Transforming the data from calculator'''  

    # Formatting length and height
    dlugosc = f'{data["dlugosc"]:,}'.replace(",", " ")
    wysokosc = f'{data["wysokosc"]:,}'.replace(",", " ")
    dlugosc, wysokosc = str(dlugosc), str(wysokosc)

    # Mapping acoustics level
    if str(data["akustyka"]).strip().lower() == "ei":
        akustyka = 52
    elif data["szklo_raw"]:
        akustyka = 50
    else:
        akustyka = data["akustyka"]
    
    # Checking if it's not OP50
    if akustyka == 45:
        items = [
            "Optimal 110",
            "Optimal 50",
        ]
        if data['plyta_raw'] == "BEZ PŁYT":
            items.append("Optimal 50 glass")
        op50 = popups.wybor_popup(items)

        if "Optimal 50" in op50:
            system = "Optimal 50"
            certyfikacja_BRI = ""
            if "glass" in op50:
                data["cena"] += data["dlugosc"] / 1000 * data['wysokosc'] / 1000 * 275 / 4.3
                akustyka = 33
            else:
                akustyka = 32
        else:
            system = "Optimal 110"
            certyfikacja_BRI = " (Certified by BRI)"
    else:
        op50 = False
        system = "Optimal 110"
        certyfikacja_BRI = " (Certified by BRI)"

    # Liczba modułów
    liczba_modulow = str(data["liczba_modulow"])
    
    # Mapping suspension type to text
    if str(data["parkowanie_raw"]).strip().lower() == 'j':
        parkowanie = "axis (-J-) / 1-point suspension"
    elif str(data["parkowanie_raw"]).strip().lower() == 'npn':
        parkowanie = "lateral (-NPN-) / 2-point"

    # Mapping panel types to text
    if data["plyta_raw"] == "BEZ PŁYT":
        plyta = "finishing panels sourced locally"
        if op50 and "glass" in op50:
            plyta = "-"
    elif data["plyta_raw"].startswith("laminowana"):
        if data["klasa_palnosci_raw"] == 1 or akustyka == 52:
            plyta = "Stopfire Melamine faced chipboard M1 (B - s2,d0)"
        else:
            plyta = "Melamine faced chipboard M3 class (D - s2,d0)"
    elif data["plyta_raw"] == "tablica suchościeralna 1-str":
        plyta = "Onesided magnetic dry erasable board"
    elif data["plyta_raw"] == "tablica suchościeralna 2-str":
        plyta = "Twosided magnetic dry erasable board"
    elif data["plyta_raw"].startswith("DOWOLNA"):
        if not data["plyta_wybrana"]:
            items = [
            "CPL on M3 class chipboard (D-s2,d0)", 
            "CPL on M1 class (s/fire) chipboard (B-s2,d0)", 
            "HPL on M3 class chipboard (D-s2,d0)", 
            "HPL on M1 class (s/fire) chipboard (B-s2,d0)",
            "Veneer on M3 chipboard (D-s2,d0)", 
            "Veneer on M1 (s/fire) chipboard (B-s2,d0)"
        ]
            plyta = popups.wybor_popup(items)
            data["plyta_wybrana"] = plyta
        else:
            plyta = data["plyta_wybrana"]

    if akustyka == 52:
        plyta += " EI30"


    additional_equipment_list = []
    # Mapping concealed profiles
    if data["ukryte_krawedzie_raw"] == 1 or akustyka >= 52:
        ukryte_krawedzie = "Concealed profiles"
        additional_equipment_list.append(ukryte_krawedzie)

    # Mapping rail's color
    if not data["kolor_toru_raw"] or data["kolor_toru_raw"] == 0:
        kolor_toru = "Raw"
    elif data["kolor_toru_raw"] == 1:
        kolor_toru = "RAL 9010"
    elif data["kolor_toru_raw"] == 2:
        kolor_toru = "Other RAL"
    
    # Mapping glass
    if not data["szklo_raw"]:
        szklo = "-"
        if op50 and "glass" in op50:
            szklo = "OPT 50 33.1 VSG"
    else:
        akustyka = 50
    if data["szklo_raw"] == "ESG 8 mm":
        szklo = "Surface glass 8mm ESG (toughened)"
    elif ["szklo_raw"] == "33.1":
        szklo = "OPT 50 33.1 VSG"
    elif data["szklo_raw"] == "44.1":
        szklo = "OPT 50 44.1 VSG/ Internal glass 44.2mm VSG (laminated)"
    elif data["szklo_raw"]:
        szklo = "Without glass - sourced locally"
    
    # Mapowanie steering type
    if data["polautomat_raw"] == 1:
        polautomat = "Semi-automatic"
    else:
        polautomat = "Manual"
    
    # Mapping additional track
    if data["dodatkowy_tor_raw"]:
        dodatkowy_tor = f"Additional track {data["dodatkowy_tor_raw"]}mm"
        additional_equipment_list.append(dodatkowy_tor)

    # Mapping suspension
    if data["obnizenie_raw"] and data["obnizenie_raw"] >= 500:
        obnizenie = f"Steel suspension {data["obnizenie_raw"]}mm"
        additional_equipment_list.append(obnizenie)

    # Mapping color powded profiles
    if data["lakierowane_profile_raw"] == 1:
        lakierowane_profile = "Powder coated RAL"
    else:
        lakierowane_profile = "Anodised"

    # Mapping doors
    drzwi = ""
    if data["liczba_DE"] and data["liczba_DE"] > 0:
        drzwi += "Single door"
        if data["liczba_DE"] > 1:
            drzwi += f" x {data["liczba_DE"]}"
        if data["liczba_DE2"] and data["liczba_DE2"] > 0:
            drzwi += ", "

    if data["liczba_DE2"] and data["liczba_DE2"] > 0:
        drzwi += "Double door"
        if data["liczba_DE2"] > 1:
            drzwi += f" x {data["liczba_DE2"]}"
    
    if (not data["liczba_DE"] and not data["liczba_DE2"]) or (data["liczba_DE"] == 0 and data["liczba_DE2"] == 0):
        drzwi = "-"


    # Mapping additional equipment:
    additional_equipment = ""
    for i in range(len(additional_equipment_list)):
        if i < len(additional_equipment_list) - 1:
            additional_equipment += additional_equipment_list[i]
            additional_equipment += ", "
        else:
            additional_equipment += additional_equipment_list[i]
    if len(additional_equipment) == 0:
        additional_equipment = "-"

    
    akustyka = f"Rw = {akustyka} dB"

    # Mapping number of walls
    liczba_scian = data["liczba_scian"]

    # Mapping the price
    cena = data["cena"]

    if liczba_scian > 1:
        cena /= liczba_scian
        cena = int(cena)
        cena = str(cena)
        cena = str(liczba_scian) + " x " + cena
    else:
        cena = f'{cena:,}'.replace(",", " ")
        cena = str(cena).split(".")[0]

    liczba_scian = str(liczba_scian)
    # Mapping the client's and project's name, number of offer

    nazwa_oferty = data["nazwa_oferty"]
    nr_oferty = nazwa_oferty.split(",")[0]
    klient = nazwa_oferty.split(",")[1].strip()
    if len(nazwa_oferty.split(",")) > 2:
        projekt = nazwa_oferty.split(",")[2]
        projekt = projekt.replace(".xlsx", "")
    else:
        projekt = ""

    # Mapping price of all walls:

    cena_wszystkich = data["cena_wszystkich"]
    cena_wszystkich = f'{cena_wszystkich:,}'.replace(",", " ")
    cena_wszystkich = str(cena_wszystkich)
    cena_wszystkich = cena_wszystkich.split(".")[0]


    data = {
        "dlugosc" : dlugosc,
        "wysokosc" : wysokosc,
        "akustyka" : str(akustyka),
        "parkowanie" : parkowanie,
        "liczba_modulow": liczba_modulow,
        "plyta" : plyta,
        "kolor_toru" : kolor_toru,
        "szklo" : szklo,
        "drzwi" : drzwi,
        "polautomat" : polautomat,
        "lakierowane_profile" : lakierowane_profile,
        "cena" : cena,
        "liczba_scian" : liczba_scian,
        "cena_wszystkich" : cena_wszystkich,
        "additional_equipment" : additional_equipment,
        "cena_wszystkich" : cena_wszystkich,
        "nr_oferty" : nr_oferty,
        "klient" : klient,
        "projekt" : projekt,
        "plyta_wybrana" : data["plyta_wybrana"],
        "certyfikacja_BRI" : certyfikacja_BRI,
        "system" : system
    }

    return data