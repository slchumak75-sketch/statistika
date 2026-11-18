import pytest

from statistika.catalog import Catalog, PDF_GROUPS, builtin_catalog
from statistika.classifier import Classifier


@pytest.mark.parametrize(("name", "category", "number"), [
    ('Вино "Саперави" п/сух. 0,75л', "", 35),
    ('Вино "Вермут Бьянко" 0,75л', "", 35),
    ("Вино фрукт. Абрикос 0,75л", "", 36),
    ("Вино плодовое игристое", "", 36),
    ("Вино игр. Просекко 0,75л", "", 37),
    ("Советское шампанское п/сл.", "", 37),
    ("Водка Хлебное вино 0,5л", "", 38),
    ("Напиток коньяч. 0,5л", "", 39),
    ("Бренди 0,5л", "", 39),
    ("Пиво светлое 5,0%", "", 40),
    ("Напиток лик. Амаретто 0,5л", "", 41),
    ("Джин 0,5л", "", 41),
    ("Бальзам Белорусский 0,5л", "", 41),
    ("Сидр яблочный 1,5л", "", 42),
    ("Напиток пивн. Garage 0,45л", "", 42),
    ("Джин-тоник 0,33л", "", 42),
    ("Виски ирландское 0,7л", "", 43),
    ("Водка виноградная Чача", "", 43),
    ("Вода пит. Аура негаз. 1,5л", "", 45),
    ("Минеральная вода газированная 1л", "", 45),
    ("Сигареты с фильтром", "", 46),
    ("Стики NEO BERRY", "", 46),
    ("1 Колено PPR 20*45", "3 полипропилен", 53),
    ("Муфта 40 мм", "3 металлопласт", 53),
    ("Лейка для душа", "3 смесители душевое сопутствующее", 53),
    ("Замок врезной", "", 53),
    ("Автошампунь с воском", "", 57),
    ("Бальзам д/мытья посуды 450мл", "", 57),
    ("Порошок Optim color 3 кг", "", 57),
    ("Шампунь для чистки ковров", "", 57),
    ("Вода мицелярная 200мл", "", 63),
    ("Бальзам д/волос 360мл", "", 63),
    ("Гель д/умывания 100мл", "", 63),
    ("Прокладки гиг. Cleanic", "", 63),
    ("Колготки детск. 700К-001", "", 100),
    ("Получулки(гольфы)женские", "", 100),
    ("Вино газ. п.сух. 0,75л", "", None),
    ("Детское шампанское", "", None),
    ("Пиво безалк. 0,5 л", "", None),
    ("Пиво 0,0 % 0,5 л", "", None),
    ("Бокал для шампанского", "", None),
    ("Сигарета электронная", "", None),
    ("Вода дистиллированная", "", None),
    ("Напиток виноградный без сахара", "", None),
    ("Корм для кошек", "2 Обои", None),
    ("Сверло для дерева", "2 стройматериалы", None),
    ("Зубная щётка электрическая", "", None),
    ("Губка для мытья посуды", "", None),
])
def test_known_products_and_exclusions(name, category, number):
    assert Classifier(builtin_catalog()).group_for(name, category) == number


def test_official_rules_not_attached_to_unrelated_pdf_numbering():
    official = builtin_catalog()
    assert all(official.is_official_group(number) for number in PDF_GROUPS)
    assert official.short_label(35) == "№35 — Вина виноградные"
    assert "По строке 22" in official.label(35)
    other = Catalog({35: "Холодильники"})
    assert Classifier(other).group_for("Вино Саперави") is None
    assert Classifier(other).group_for("Холодильник") == 35


def test_conflicting_custom_alias_leaves_official_match_unassigned():
    official = builtin_catalog()
    catalog = Catalog(official.points, {100: ("водка",)}, official.source)
    assert Classifier(catalog).group_for("Водка 0,5л") is None
