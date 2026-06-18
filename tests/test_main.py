import pytest
from main import transform_data

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
        "key, expected_value",
        [
                ("dlugosc", "10 000"),
                ("wysokosc", "5 000"),
                "dlugosc" : "10 000",
    "wysokosc" : "5 000",
    "akustyka" : "Rw = 52",
    "parkowanie" : "axis (-J-) / 1-point suspension",
    "liczba_modulow": "10",
    "plyta" : "Stopfire Melamine faced chipboard M1 (B - s2,d0) EI30",
    "kolor_toru" : "Raw",
    "szklo" : "-",
    "drzwi" : "-",
    "polautomat" : "Manual",
    "lakierowane_profile" : "Anodised",
    "cena" : "9 500",
    "liczba_scian" : 1,
    "cena_wszystkich" : "9 500",
    "additional_equipment" : "Concealed profiles",
    "nr_oferty" : "2137-MM-26",
    "klient" : "Scianex",
    "projekt" : ""
        ]
)

def test_transform_data(key, expected_value):

    actual_output = transform_data(input_data)
    assert actual_output.get(key) == expected_value