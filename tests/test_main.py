import pytest
from main import transform_data, calculate_nr_of_walls

input_data = {
    "dlugosc" : 10000,
    "wysokosc" : 5000,
    "akustyka" : "ei",
    "parkowanie_raw" : "j",
    "liczba_modulow": 10,
    "plyta_raw" : "laminowana - STANDARD (Egger)",
    "klasa_palnosci_raw" : "ei",
    "ukryte_krawedzie_raw" : None,
    "kolor_toru_raw" : 0,
    "szklo_raw" : None,
    "liczba_DE" : 0,
    "liczba_DE2" : 0,
    "polautomat_raw" : 0,
    "dodatkowy_tor_raw" : 0,
    "obnizenie_raw" : None,
    "lakierowane_profile_raw" : None,
    "cena" : 9500,
    "liczba_scian" : 1,
    "cena_wszystkich" : 9500,
    "nazwa_oferty" : "2137-MM-26, Scianex"
}

@pytest.mark.parametrize(
        "updates, key, expected_value",
        [
                ({}, "dlugosc", "10 000"),
                ({}, "wysokosc", "5 000"),
                ({},"akustyka" , "Rw = 52 dB"),
                ({}, "parkowanie" , "axis (-J-) / 1-point suspension"),
                ({"parkowanie_raw" : "npn"}, "parkowanie", "lateral (-NPN-) / 2-point"),
                ({},"liczba_modulow", "10"),
                ({},"plyta" , "Stopfire Melamine faced chipboard M1 (B - s2,d0) EI30"),
                ({},"kolor_toru" , "Raw"),
                ({},"szklo" , "-"),
                ({},"drzwi" , "-"),
                ({},"polautomat" , "Manual"),
                ({},"lakierowane_profile" , "Anodised"),
                ({},"cena" , "9 500"),
                ({},"liczba_scian" , "1"),
                ({},"cena_wszystkich" , "9 500"),
                ({},"additional_equipment" , "Concealed profiles"),
                ({},"nr_oferty" , "2137-MM-26"),
                ({},"klient" , "Scianex"),
                ({},"projekt" , "")
        ]
)
def test_transform_data(updates, key, expected_value):
    test_input = input_data.copy()  
    test_input.update(updates)

    actual_output = transform_data(test_input)
    assert actual_output.get(key) == expected_value

def test_calculate_nr_of_walls():
    data = calculate_nr_of_walls(10, 2)
    data_expected = {
        "calkowita_liczba_scian" : 12,
        "nr_sciany" : "W11-W12"
    }
    assert data == data_expected