import popups
import win32com.client
from files_managment import decrypt_excel_file
from popups import wybor_i_tekst_popup

def calculate_nr_of_walls(calkowita_liczba_scian, liczba_scian):
    """Calculating number of walls"""
    calkowita_liczba_scian += liczba_scian
    if liczba_scian == 1:
        nr_sciany = f'W{calkowita_liczba_scian}'
    else:
        nr_sciany = f'W{calkowita_liczba_scian-liczba_scian+1}-W{calkowita_liczba_scian}'

    data = {
        "calkowita_liczba_scian" : calkowita_liczba_scian,
        "nr_sciany" : nr_sciany
    }

    return data

def calculate_markup_bottom_top_lines(initial_markup):
    """fixing markup based on initial one"""

    modulo = initial_markup % 0.01
    if modulo >= 0.005:
        markup_bottom_line = initial_markup - modulo + 0.005
        markup_top_line = initial_markup - modulo + 0.015
    else:
        markup_bottom_line = initial_markup - modulo - 0.005
        markup_top_line = initial_markup - modulo + 0.005

    return markup_bottom_line, markup_top_line

def adjust_markup(sheet, i, markup_bottom_line, markup_top_line):

    while True:
        # markup
        if sheet.Range(f"AG{i}").value < markup_bottom_line:
        # discount
            sheet.Range(f'AF{i+1}').value -= 0.01
        elif sheet.Range(f"AG{i}").value >= markup_top_line:
            sheet.Range(f'AF{i+1}').value += 0.01
        else:
            break
        
def calculate_additions(directory, nazwa_oferty):
    """Calculating additions"""

    items = [
        "No additions", 
        "Single door",
        "Double door",
        "Powder coated profiles in RAL",
        "Concealed profiles",
        "Semiautomatic",
        "Modules factory assembled",
        "J",
        "NPN",
        "Rails in RAL 9010",
        "Other RAL of rails",
        "CPL",
        "HPL",
        "M1 Melamine",
        "M1 CPL",
        "M1 HPL",
        "Other covering panels",
        "45dB",
        "49dB",
        "51dB",
        "53dB",
        "54dB",
        "57dB"
    ]

    wybor = popups.wybor_popup(items, wielokrotny=True)

    def calculate_addition(addition_type, sheet, cena_przed):
        """returns addition_type, doplata"""

        kolumna2 = False

        if addition_type == "Single door":
            addition_value = 1
            kolumna = "T"
        elif addition_type == "Double door":
            addition_value = 1
            kolumna = "U"
        elif addition_type == "Powder coated profiles in RAL":
            addition_value = 1
            kolumna = "Y"
        elif addition_type == "Concealed profiles":
            addition_value = 1
            kolumna = "L"
        elif addition_type == "Semiautomatic":
            addition_value = 1
            kolumna = "V"
        elif "dB" in addition_type:
            addition_value = addition_type[:2]
            kolumna = "j"
        elif addition_type == "NPN":
            addition_value = "npn"
            kolumna = "H"
            kolumna2 = "I"
        elif addition_type == "J":
            addition_value = "j"
            kolumna = "H"
            kolumna2 = "I"
        elif addition_type == "Rails in RAL 9010":
            addition_value = 1
            kolumna = "O"
        elif addition_type == "Other RAL of rails":
            addition_value = 2
            kolumna = "O"

        # Calculating each wall:
        calkowita_liczba_scian = 0
        i = 6
        i_to_omit = []
        # discounts, values to save to write them back at the end
        discounts = {}
        values = {}
        values2 = {}

        # nr of walls
        while sheet.Range(f'C{i}').value and sheet.Range(f'C{i}').value >= 1:
            liczba_scian = int(sheet.Range(f'C{i}').value)
            values[i] = sheet.Range(f'{kolumna}{i}').value
            values2[i] = False
            discounts[i] = sheet.Range(f'AF{i+1}').value
            markup_bottom_line, markup_top_line = calculate_markup_bottom_top_lines(sheet.Range(f"AG{i}").value)
            sheet.Range(f'{kolumna}{i}').value = addition_value

            # number of parking stacks (when suspension change)
            if kolumna2:

                if sheet.Range(f'{kolumna2}{i}').value:
                    values2[i] = sheet.Range(f'{kolumna2}{i}').value
                else:
                    values2[i] = 0
                sheet.Range(f'{kolumna}{i}').value = addition_value
                if addition_value == "npn":
                    ilosc_stosow = sheet.Range(f'N{i}').value/8
                    if ilosc_stosow % 1 == 0:
                        ilosc_stosow = int(ilosc_stosow)
                    else:
                        ilosc_stosow = int(ilosc_stosow) + 1

                elif addition_value == "j":
                    ilosc_stosow = 0

                sheet.Range(f'{kolumna2}{i}').value = ilosc_stosow

            calkowita_liczba_scian += liczba_scian
            adjust_markup(sheet, i, markup_bottom_line, markup_top_line)
            i += 2

        # calculating the supplementary prices
        cena_po = sheet.Range('AF2').value
        if calkowita_liczba_scian != 0:
            if "door" in addition_type:
                doplata = int((cena_po-cena_przed)/calkowita_liczba_scian)
                doplata = f'{doplata:,}'.replace(",", " ")
                doplata = str(doplata) + " EUR / door"
            else:
                doplata = int(cena_po-cena_przed)
                doplata = f'{doplata:,}'.replace(",", " ")
                doplata = str(doplata) + " EUR"
        else:
            doplata = 0

        # getting back to previous values
        i = 6
        while sheet.Range(f'C{i}').value and sheet.Range(f'C{i}').value >= 1:
            sheet.Range(f'{kolumna}{i}').value = values[i]
            sheet.Range(f'AF{i+1}').value = discounts[i]
            if values2[i]:
                sheet.Range(f'{kolumna2}{i}').value = values2[i]
            i += 2
        print(f'{addition_type} : {doplata}')
        return addition_type, doplata

    if "No additions" not in wybor:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.ReferenceStyle = -4150
        excel.visible = False
        excel.ScreenUpdating = False
        excel.DisplayAlerts = False

        decrypt_excel_file(nazwa_oferty, f'decrypted{nazwa_oferty}')
        wb = excel.Workbooks.Open(str(directory) + "\\" + f"decrypted{nazwa_oferty}")
        sheet = wb.Sheets("KALKULATOR")
        cena_przed = sheet.Range('AF2').value
        additional_items = []

        if "Single door" in wybor:
            additional_item, doplata = calculate_addition("Single door", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
            additional_items.append(additionals)
        if "Double door" in wybor:
            additional_item, doplata = calculate_addition("Double door", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
        if "Powder coated profiles in RAL" in wybor:
            additional_item, doplata = calculate_addition("Powder coated profiles in RAL", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
        if "Concealed profiles" in wybor:
            additional_item, doplata = calculate_addition("Concealed profiles", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
        if "Semiautomatic" in wybor:
            additional_item, doplata = calculate_addition("Semiautomatic", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
        if "Modules factory assembled" in wybor:
            i = 6
            liczba_modulow = 0
            liczba_drzwi = 0

            while sheet.Range(f"C{i}").value and sheet.Range(f"C{i}").value > 0:
                liczba_modulow += sheet.Range(f"N{i}").value * sheet.Range(f"C{i}").value
                if sheet.Range(f"T{i}").value:
                    liczba_drzwi += sheet.Range(f"T{i}").value
                if sheet.Range(f"U{i}").value:
                    liczba_drzwi += sheet.Range(f"U{i}").value
                i += 2

            liczba_skrzyn = liczba_modulow / 10
            if liczba_skrzyn % 1 != 0:
                liczba_skrzyn = int(liczba_skrzyn) + 1
            else:
                liczba_skrzyn = int(liczba_skrzyn)
            doplata = liczba_modulow * 100 + liczba_skrzyn * 200 + 100 * liczba_drzwi
            doplata = int(doplata)
            doplata = str(doplata) + " EUR"
            additionals = {
                "additional_item" : "Modules factory assembled",
                "doplata" : doplata
            }
            print(f'Modules factory assembled: {doplata}')
        if "J" in wybor:
            additional_item, doplata = calculate_addition("J", sheet, cena_przed)
            additionals = {
                "additional_item" : additional_item,
                "doplata" : doplata
            }
        if "NPN" in wybor:
                    additional_item, doplata = calculate_addition("NPN", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "Rails in RAL 9010" in wybor:
                    additional_item, doplata = calculate_addition("Rails in RAL 9010", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "Other RAL of rails" in wybor:
                    additional_item, doplata = calculate_addition("Other RAL of rails", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "45dB" in wybor:
                    additional_item, doplata = calculate_addition("45dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "49dB" in wybor:
                    additional_item, doplata = calculate_addition("49dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "51dB" in wybor:
                    additional_item, doplata = calculate_addition("51dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "53dB" in wybor:
                    additional_item, doplata = calculate_addition("53dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "54dB" in wybor:
                    additional_item, doplata = calculate_addition("54dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }
        if "57dB" in wybor:
                    additional_item, doplata = calculate_addition("57dB", sheet, cena_przed)
                    additionals = {
                        "additional_item" : additional_item,
                        "doplata" : doplata
                    }

        wb.Save()
        wb.Close()
        excel.Quit()