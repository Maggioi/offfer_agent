import openpyxl
import docx
import io
import msoffcrypto
import os
import popups

def uzupelnij_dokument():
    # 1. Wczytanie danych z kalkulatora Excel
    f = open("password.txt")
    password = f.read()
    decrypted_workbook = io.BytesIO()
    with open("kalkulator_official.xlsx", "rb") as encrypted_workbook:
        file = msoffcrypto.OfficeFile(encrypted_workbook)
        file.load_key(password=password)
        file.decrypt(decrypted_workbook)

    wb = openpyxl.load_workbook(decrypted_workbook, data_only=True)
    sheet = wb['KALKULATOR']
    
    
    # Pobranie wartości z wiersza nr 6
    dlugosc = sheet['F6'].value        
    wysokosc = sheet['G6'].value      
    akustyka = sheet['J6'].value       
    parkowanie_raw = sheet['H6'].value 
    liczba_modulow = sheet['N6'].value
    plyta_raw = sheet['P6'].value
    klasa_palnosci_raw = sheet['K6'].value
    ukryte_krawedzie_raw = sheet['L6'].value
    kolor_toru_raw = sheet['O6'].value
    szklo_raw = sheet['R6'].value 
    liczba_DE = sheet['T6'].value
    liczba_DE2 = sheet['U6'].value
    polautomat_raw = sheet['V6'].value
    dodatkowy_tor_raw = sheet['W6'].value
    obnizenie_raw = sheet['X6'].value
    lakierowane_profile_raw = sheet['Y6'].value
    cena = sheet['AF6'].value


    # Mapowanie typu zawieszenia na tekst
    if parkowanie_raw and str(parkowanie_raw).strip().lower() == 'j':
        zawieszenie = "axis (-J-) / 1-point suspension"
    elif parkowanie_raw and str(parkowanie_raw).strip().lower() == 'npn':
        zawieszenie = "lateral (-NPN-) / 2-point"
    else:
        zawieszenie = "-"


    # Mapowanie płyt na tekst
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

    


    
    # Mapowanie ukrytych krawędzi
    if ukryte_krawedzie_raw == 1 or akustyka >= 54:
        ukryte_krawedzie = "Concealed profiles"


    # Mapowanie koloru toru
    if not kolor_toru_raw or kolor_toru_raw == 0:
        kolor_toru = "raw"
    elif kolor_toru_raw == 1:
        kolor_toru = "RAL 9010"
    elif kolor_toru_raw == 2:
        kolor_toru = "Other RAL"
    
    # Mapowanie szkła
    if not szklo_raw:
        szklo = "-"
    elif szklo_raw == "ESG 8 mm":
        szklo = "Surface glass 8mm ESG (toughened)"
    elif szklo_raw == "33.1":
        szklo = "OPT 50 33.1 VSG"
    elif szklo_raw == "44.1":
        szklo = "OPT 50 44.1 VSG/ Internal glass 44.2mm VSG (laminated)"
    else:
        szklo == "Without glass - sourced locally"
    
    # Mapowanie sterowania
    if polautomat_raw == 1:
        polautomat = "semi-automatic"
    else:
        polautomat = "manual"
    
    # Mapowanie dodatkowego toru
    if dodatkowy_tor_raw and dodatkowy_tor_raw > 0:
        dodatkowy_tor = "Additional track"

    # Mapowanie obniżenia
    if obnizenie_raw and dodatkowy_tor_raw >= 500:
        obnizenie = "Steel suspension"

    # Mapowanie lakierowanych profili
    if lakierowane_profile_raw == 1:
        lakierowane_profile = "Powder coated RAL"
    else:
        lakierowane_profile = "Anodised"

    # Mapowanie drzwi
    drzwi = "-"
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

    # Mapowanie dodatkowego wyposażenia
    additional_equipment = ""

    print(f"Pobrane dane z Excela:")
    print(f" - L = {dlugosc} mm")
    print(f" - H = {wysokosc} mm")
    print(f" - Rw = {akustyka} dB")
    print(f" - Zawieszenie = {zawieszenie}\n")

    # 2. Otwarcie szablonu dokumentu Word
    doc = docx.Document("krolik_doswiadczalny.docx")

    # 3. Przetwarzanie tabel w pliku Word
    if len(doc.tables) > 2:
        tabela = doc.tables[2]
        
        # Iterujemy po unikalnych komórkach (odporne na ValueError przy scaleniach)
        for cell in tabela._cells:
            
            # Bezpieczna podmiana Wymiarów (L)
            if "Podaj długość" in cell.text:
                cell.text = cell.text.replace("Podaj długość", str(dlugosc))
                
            # Bezpieczna podmiana Wymiarów (H)
            if "Podaj wysokość" in cell.text:
                cell.text = cell.text.replace("Podaj wysokość", str(wysokosc))
                
            # Bezpieczna podmiana Akustyki (Rw)
            if "Rw = " in cell.text and "dB" in cell.text:
                if f"Rw = {akustyka} dB" not in cell.text:
                    cell.text = f"Rw = {akustyka} dB"

            # Bezpieczna podmiana parkingu
            if "Kac bałagane" in cell.text:
                cell.text = cell.text.replace("Kac bałagane", zawieszenie)
            
            # Podmiana liczby modułów
            if "Liczba modułów" in cell.text:
                cell.text = cell.text.replace("Liczba modułów", str(liczba_modulow))

            # Podmiana rodzaju pyty
            if "pyta" in cell.text:
                cell.text = cell.text.replace("pyta", plyta)

            # Podmiana operacji
            if "operacja" in cell.text:
                cell.text = cell.text.replace("operacja", polautomat)
            
            # Podmiana koloru torów
            if "tory" in cell.text:
                cell.text = cell.text.replace("tory", kolor_toru)
            
            # Podmiana dźwi
            if "drzwi" in cell.text:
                cell.text = cell.text.replace("drzwi", drzwi)

            # Podmiana szkła
            if "szkło" in cell.text:
                cell.text = cell.text.replace("szkło", szklo)
            
            # Podmiana lakieru profili
            if "aluminium" in cell.text:
                cell.text = cell.text.replace("aluminium", lakierowane_profile)
            
            # Podmiana dodatkowego wyposażenia
            ...

            # Podmiana ceny
            if "Cena" in cell.text:
                cell.text = cell.text.replace("Cena", str(cena))


    # 4. Zapisanie uzupełnionego dokumentu
    nowa_nazwa = "Nowy_krolik.docx"
    doc.save(nowa_nazwa)
    print(f"Sukces! Wygenerowano plik: {nowa_nazwa}")
    os.startfile("Nowy_krolik.docx")
    
if __name__ == "__main__":
    uzupelnij_dokument()