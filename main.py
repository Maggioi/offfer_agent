import openpyxl
import docx
import io
import msoffcrypto
import os
import popups
from dtu import duplicate_table_underneath
import copy

def open_calculator():
    '''Opening a passworded calculator'''
    f = open("password.txt")
    password = f.read()
    decrypted_workbook = io.BytesIO()
    with open("kalkulator_official.xlsx", "rb") as encrypted_workbook:
        file = msoffcrypto.OfficeFile(encrypted_workbook)
        file.load_key(password=password)
        file.decrypt(decrypted_workbook)

    wb = openpyxl.load_workbook(decrypted_workbook, data_only=True)
    sheet = wb['KALKULATOR']
    return sheet

def load_transform_data(sheet, i):
    '''Loading data from calculator'''
    
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


    # Mapping suspension type to text
    if parkowanie_raw and str(parkowanie_raw).strip().lower() == 'j':
        zawieszenie = "axis (-J-) / 1-point suspension"
    elif parkowanie_raw and str(parkowanie_raw).strip().lower() == 'npn':
        zawieszenie = "lateral (-NPN-) / 2-point"
    else:
        zawieszenie = "-"


    # Mapping panel types to text
    if plyta_raw == "BEZ PŁYT":
        plyta = "finishing panels sourced locally"

    elif plyta_raw.startswith("laminowana"):
        if klasa_palnosci_raw == 1:
            plyta = "Stopfire Melamine faced chipboard M1 (B - s2,d0)"
        elif str(klasa_palnosci_raw).lower() == "ei":
            plyta = "Stopfire Melamine faced chipboard M1 (B - s2,d0)"
        else:
            plyta = "Melamine faced chipboard M3 class (D - s2,d0)"
        
    elif plyta_raw == "tablica suchościeralna 1-str":
        plyta = "Onesided magnetic dry erasable board"
    elif plyta_raw == "tablica suchościeralna 2-str":
        plyta = "Twosided magnetic dry erasable board"
    
    elif plyta_raw.startswith("DOWOLNA"):
        plyta = popups.wybor_plyty_popup()

    
    # Mapping of additional equipment
    additional_equipment_list = []

    
    # Mapping concealed profiles
    if ukryte_krawedzie_raw == 1 or akustyka >= 54:
        ukryte_krawedzie = "Concealed profiles"
        additional_equipment_list.append(ukryte_krawedzie)

    # Mapping rail's color
    if not kolor_toru_raw or kolor_toru_raw == 0:
        kolor_toru = "Raw"
    elif kolor_toru_raw == 1:
        kolor_toru = "RAL 9010"
    elif kolor_toru_raw == 2:
        kolor_toru = "Other RAL"
    
    # Mapping glass
    if not szklo_raw:
        szklo = "-"
    elif szklo_raw == "ESG 8 mm":
        szklo = "Surface glass 8mm ESG (toughened)"
    elif szklo_raw == "33.1":
        szklo = "OPT 50 33.1 VSG"
    elif szklo_raw == "44.1":
        szklo = "OPT 50 44.1 VSG/ Internal glass 44.2mm VSG (laminated)"
    else:
        szklo = "Without glass - sourced locally"
    
    # Mapowanie steering type
    if polautomat_raw == 1:
        polautomat = "Semi-automatic"
    else:
        polautomat = "Manual"
    
    # Mapping additional track
    if dodatkowy_tor_raw and dodatkowy_tor_raw > 0:
        dodatkowy_tor = "Additional track"
        additional_equipment_list.append(dodatkowy_tor)

    # Mapping suspension
    if obnizenie_raw and dodatkowy_tor_raw >= 500:
        obnizenie = "Steel suspension"
        additional_equipment_list.append(obnizenie)

    # Mapping color powded profiles
    if lakierowane_profile_raw == 1:
        lakierowane_profile = "Powder coated RAL"
    else:
        lakierowane_profile = "Anodised"

    # Mapping doors
    drzwi = ""
    if liczba_DE and liczba_DE > 0:
        drzwi += "Single door"
        if liczba_DE > 1:
            drzwi += f" x {liczba_DE}"
        if liczba_DE2 and liczba_DE2 > 0:
            drzwi += ", "

    if liczba_DE2 and liczba_DE2 > 0:
        drzwi += "Double door"
        if liczba_DE2 > 1:
            drzwi += f" x {liczba_DE2}"
    else:
        drzwi = "-"
    # Mapping additional equipment:
    additional_equipment = ""
    for i in range(len(additional_equipment_list)):
        if i < len(additional_equipment_list) - 1:
            additional_equipment += additional_equipment_list[i]
            additional_equipment += ", "
        else:
            additional_equipment += additional_equipment_list[i]
    
    akustyka = f"Rw = {akustyka} dB"

    return (dlugosc, wysokosc, zawieszenie, liczba_modulow, akustyka,
            plyta, polautomat, kolor_toru, drzwi, szklo, lakierowane_profile,
            additional_equipment, cena)

def open_offer(doc, tabela, dlugosc, wysokosc, zawieszenie, liczba_modulow, akustyka,
            plyta, polautomat, kolor_toru, drzwi, szklo, lakierowane_profile,
            additional_equipment, cena):
    '''Making a table with a wall'''
    # 2. Opening of a Word Offer template

    for cell in tabela._cells:
        for paragraph in cell.paragraphs:
            # Check the text at the paragraph level first to avoid unnecessary iteration over runs
            p_text = paragraph.text
            
            # If the paragraph is empty, skip to the next one
            if not p_text:
                continue

            # For all other standard replacements, safely iterate through the text runs
            for run in paragraph.runs:
                if not run.text:
                    continue
                    
                # Safe replacement for Dimensions (L)
                if "Podaj długość" in run.text:
                    run.text = run.text.replace("Podaj długość", str(dlugosc))
                    
                # Safe replacement for Dimensions (H)
                if "Podaj wysokość" in run.text:
                    run.text = run.text.replace("Podaj wysokość", str(wysokosc))
                    
                # Safe replacement for Parking / Suspension setup
                if "Kac" in run.text:
                    run.text = run.text.replace("Kac", zawieszenie)
                   
                # Replacement for the number of modules
                if "Liczba modułów" in run.text:
                    run.text = run.text.replace("Liczba modułów", str(liczba_modulow))
                
                if "akustyka" in run.text:
                    run.text = run.text.replace("akustyka", akustyka)

                # Replacement for the board/panel type
                if "płyta" in run.text:
                    run.text = run.text.replace("płyta", plyta)

                # Replacement for the operation type
                if "operacja" in run.text:
                    run.text = run.text.replace("operacja", polautomat)
                
                # Replacement for the track color
                if "tory" in run.text:
                    run.text = run.text.replace("tory", kolor_toru)
                
                # Replacement for the door type
                if "drzwi" in run.text:
                    run.text = run.text.replace("drzwi", drzwi)

                # Replacement for the glass type
                if "szkło" in run.text:
                    run.text = run.text.replace("szkło", szklo)
                
                # Replacement for the profile powder coating / finish
                if "aluminium" in run.text:
                    run.text = run.text.replace("aluminium", lakierowane_profile)
                
                # Replacement for additional equipment:
                if "adyszynal" in run.text:
                    run.text = run.text.replace("adyszynal", additional_equipment)

                # Replacement for the price
                if "Cena" in run.text:
                    run.text = run.text.replace("Cena", str(cena))
    return doc
            
def save_offer(doc):
    # 4. Saving filled offer document.
    nowa_nazwa = "Nowy_krolik.docx"
    doc.save(nowa_nazwa)
    print(f"Sukces! Wygenerowano plik: {nowa_nazwa}")
    os.startfile("Nowy_krolik.docx")
    
if __name__ == "__main__":
    # Opening calculator and offer files.    
    calculator_sheet = open_calculator()
    doc = docx.Document("krolik_doswiadczalny.docx")

    # Iterating through walls
    i = 6
    nowa_tabela = doc.tables[2]

    while True:  
        transformed_data = load_transform_data(calculator_sheet, i)
        
        # 1. Sprawdzamy, czy w kalkulatorze jest kolejna ściana do przerobienia
        kolejne_i = i + 2
        ma_kolejna_sciane = calculator_sheet[f'C{kolejne_i}'].value and calculator_sheet[f'C{kolejne_i}'].value > 0
        
        # 2. Jeśli tak, duplikujemy tabelę ZANIM ją wypełnimy (kiedy ma czyste znaczniki)
        if ma_kolejna_sciane:
            nastepna_tabela = duplicate_table_underneath(nowa_tabela, nowa_tabela, doc)
            
        # 3. Dopiero teraz wypełniamy obecną tabelę danymi
        open_offer(doc, nowa_tabela, *transformed_data)
        
        # 4. Przechodzimy do kolejnej ściany
        if ma_kolejna_sciane:
            nowa_tabela = nastepna_tabela
            i = kolejne_i
        else:
            break

    save_offer(doc)
    
