"""Тесты разбора цен и количеств: python -m pytest из корня проекта."""

import pytest

from merge_pricelists import parse_number


@pytest.mark.parametrize(
    "value, expected",
    [
        (590, 590),
        ("12150.00", 12150),
        ("1 340,00 ₽", 1340),
        ("1 340,00 руб.", 1340),
        ("1\xa0200\xa0000,50 р.", 1200000.5),
        ("1.340,00", 1340),
        ("1,200.00", 1200),
        ("1.200.000", 1200000),
        ("450,5", 450.5),
        ("0,125", 0.125),
        ("60 шт", 60),
        ("60 шт.", 60),
    ],
)
def test_parsed(value, expected):
    assert parse_number(value) == pytest.approx(expected)


@pytest.mark.parametrize(
    "value",
    [
        # Неоднозначные: то ли тысячи, то ли дробь
        "1,200",
        "1.200",
        # Числа нет
        "под заказ",
        "",
        None,
        # Несколько чисел в строке — какое из них имелось в виду, неизвестно
        "10-20 шт",
        "1.2e3",
        # Разделители тысяч не по 3 цифры
        "1.2.3",
        "1 34,00",
    ],
)
def test_not_parsed(value):
    assert parse_number(value) is None
