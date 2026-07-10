import popups

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

def calculate_additions(wb):
    """Calculating additions"""

    items = [ 
        "Single door",
        "Double door",
        "Powder coated profiles in RAL"
    ]

    wybor = popups.wybor_popup(items, wielokrotny=True)

    if "Single door" in wybor:
        ...

   