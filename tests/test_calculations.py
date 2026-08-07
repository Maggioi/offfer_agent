import pytest
from calculations import calculate_nr_of_walls, calculate_markup_bottom_top_lines

def test_calculate_nr_of_walls():
    data = calculate_nr_of_walls(10, 2)
    data_expected = {
        "calkowita_liczba_scian" : 12,
        "nr_sciany" : "W11-W12"
    }
    assert data == data_expected

@pytest.mark.parametrize(
        "value, expected",
        [
            # rounds down
            (0.5345005, (0.525, 0.535)),
            (0.481, (0.475, 0.485)),
            (0.5421312, (0.535, 0.545)),
            (0.52, (0.515, 0.525)),
            (0.563421412, (0.555, 0.565)),
            # rounds up
            (0.535, (0.535, 0.545)),
            (0.5153248923, (0.515, 0.525)),
            (0.5476, (0.545, 0.555)),
            (0.558, (0.555, 0.565)),
            (0.539312312, (0.535, 0.545))
        ]
)

def test_calculate_markup_bottom_top_lines(value, expected):
    assert calculate_markup_bottom_top_lines(value) == pytest.approx(expected)

        