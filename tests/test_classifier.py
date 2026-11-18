import pytest

from statistika.catalog import Catalog
from statistika.classifier import Classifier


@pytest.mark.parametrize("name", ["НОСКИ мужские", "Носок спортивный", "Колготок набор", "ЧУЛКИ", "Термоноски", "Гольфы детские"])
def test_hosiery_builtin_group(name):
    assert Classifier(Catalog()).group_for(name) == 100


def test_catalog_words_and_user_rules():
    # Учебные описания, не значения из документа заказчика.
    catalog = Catalog({35: "Кофе молотый; кофейные напитки", 36: "Чай", 44: "Мыло"}, {36: ("чай", "сенча")})
    classifier = Classifier(catalog)
    assert classifier.group_for("Кофе молотого 250 г") == 35
    assert classifier.group_for("кофейного напитка 200 мл") == 35
    assert classifier.group_for("Сенча японская") == 36
    assert classifier.group_for("Мыло") is None
    assert classifier.group_for("Чайник") is None
    assert classifier.group_for("Набор: чай и кофе молотый") is None
    assert classifier.group_for("ЧаЙ и носки (подарок)") is None


def test_exclusions_and_generic_definitions_are_not_automatically_guessed():
    original = {35: "Кофе, кроме молотого", 36: "Изделия", 37: "Сахар (белый)", 38: "мёд"}
    classifier = Classifier(Catalog(original))
    assert classifier.group_for("Кофе") is None
    assert classifier.group_for("Кофе молотый") is None
    assert classifier.group_for("Изделия из теста") is None
    assert classifier.group_for("Сахар белый") is None
    assert classifier.group_for("МЕД цветочный") == 38
    assert Classifier(Catalog(original, {35: ("кофе зерновой",)})).group_for("кофе зерновой") == 35


def test_phrase_markers_are_editable_literal_conditions():
    catalog = Catalog(rules={53: ("# Сверло не входит в описание этой строки", "^шпатель + фасадный", "!для депиляции", "^сверло SDS+")})
    classifier = Classifier(catalog)
    assert classifier.group_for("1 Шпатель 100 мм фасадный") == 53
    assert classifier.group_for("Набор: шпатель фасадный") is None
    assert classifier.group_for("Шпатель фасадный для депиляции") is None
    assert classifier.group_for("Шпатель 100 мм") is None
    assert classifier.group_for("Сверло SDS+ GEPARD") == 53
    assert classifier.group_for("Сверло не входит в описание этой строки") is None
