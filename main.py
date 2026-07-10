import docx
from dtu import duplicate_table_underneath
from pathlib import Path
import sys
from load_transform_data import load_data, transform_data
from calculations import calculate_nr_of_walls, calculate_additions
from files_managment import open_calculator, open_offer, update_summary_table, save_offer

def run_program():
    """Running the program"""
    # Opening calculator and offer files.
    directory = Path(__file__).parent
    nazwa_oferty = next(directory.glob("*.xlsx")).name

    calculator_sheet = open_calculator(nazwa_oferty)
    doc = docx.Document("krolik_doswiadczalny.docx")
    
    # Iterating through walls
    i = 6
    nowa_tabela = doc.tables[2]
    calkowita_liczba_scian = 0
    
    plyta_wybrana = False

    while True:
        data = load_data(calculator_sheet, nazwa_oferty, i)
        data["plyta_wybrana"] = plyta_wybrana
        transformed_data = transform_data(data) 
        plyta_wybrana = transformed_data['plyta_wybrana']
        liczba_scian = transformed_data["liczba_scian"]

        liczby_scian = calculate_nr_of_walls(calkowita_liczba_scian, int(liczba_scian))
        calkowita_liczba_scian = liczby_scian["calkowita_liczba_scian"]
        nr_sciany = liczby_scian["nr_sciany"]

        # 1. Checking if there is next wall
        kolejne_i = i + 2
        ma_kolejna_sciane = calculator_sheet[f'C{kolejne_i}'].value and calculator_sheet[f'C{kolejne_i}'].value > 0
        
        # 2. If so, we duplicate the table BEFORE filling it in (when its pointers are empty)
        if ma_kolejna_sciane:
            nastepna_tabela = duplicate_table_underneath(nowa_tabela, nowa_tabela, doc)
            
        # 3. We are only now filling in the current table with data
        open_offer(doc, nowa_tabela, transformed_data, nr_sciany)
        
        # 4. Going to another wall
        if ma_kolejna_sciane:
            nowa_tabela = nastepna_tabela
            i = kolejne_i
        else:
            break

    update_summary_table(doc, calkowita_liczba_scian, transformed_data["cena_wszystkich"])
    save_offer(doc, nazwa_oferty)
    sys.exit()

if __name__ == "__main__":

    run_program()
    # directory = Path(__file__).parent
    # nazwa_oferty = next(directory.glob("*.xlsx")).name
    # calculate_additions('587-MM-26, Espaces Mobiles, DP4138.xlsx')
    
    
  
