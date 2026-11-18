import pytest

from dataclasses import replace

from statistika.catalog import GROUPS, Catalog, PDF_GROUPS, RecognitionRule, builtin_catalog
from statistika.classifier import Classifier


@pytest.mark.parametrize(("name", "number"), [
    ('Вино "Саперави" п/сух. 0,75л', 35),
    ('Вино "Вермут Бьянко" 0,75л', 35),
    ("Вино фрукт. Абрикос 0,75л", 36),
    ("Вино плодовое игристое", 36),
    ("Вино игр. Просекко 0,75л", 37),
    ("Советское шампанское п/сл.", 37),
    ("Водка Хлебное вино 0,5л", 38),
    ("Напиток коньяч. 0,5л", 39),
    ("Бренди 0,5л", 39),
    ("Пиво светлое 5,0%", 40),
    ("Напиток лик. Амаретто 0,5л", 41),
    ("Джин 0,5л", 41),
    ("Бальзам Белорусский 0,5л", 41),
    ("Сидр яблочный 1,5л", 42),
    ("Напиток пивн. Garage 0,45л", 42),
    ("Джин-тоник 0,33л", 42),
    ("Виски ирландское 0,7л", 43),
    ("Водка виноградная Чача", 43),
    ("Вода пит. Аура негаз. 1,5л", 45),
    ("Минеральная вода газированная 1л", 45),
    ("Сигареты с фильтром", 46),
    ("Стики NEO BERRY", 46),
    ("1 Колено PPR 20*45", 53),
    ("Муфта 40 мм", None),
    ("Лейка для душа", 53),
    ("Замок врезной", 53),
    ("Автошампунь с воском", 57),
    ("Бальзам д/мытья посуды 450мл", 57),
    ("Порошок Optim color 3 кг", 57),
    ("Шампунь для чистки ковров", 57),
    ("Вода мицелярная 200мл", 63),
    ("Бальзам д/волос 360мл", 63),
    ("Гель д/умывания 100мл", 63),
    ("Прокладки гиг. Cleanic", 63),
    ("Колготки детск. 700К-001", 100),
    ("Получулки(гольфы)женские", 100),
    ("Вино газ. п.сух. 0,75л", None),
    ("Детское шампанское", None),
    ("Пиво безалк. 0,5 л", None),
    ("Пиво 0,0 % 0,5 л", None),
    ("Бокал для шампанского", None),
    ("Сигарета электронная", None),
    ("Вода дистиллированная", None),
    ("Напиток виноградный без сахара", None),
    ("Корм для кошек", None),
    ("Сверло для дерева", 53),
    ("Зубная щётка электрическая", None),
    ("Губка для мытья посуды", None),
])
def test_known_products_and_exclusions(name, number):
    assert Classifier(builtin_catalog()).group_for(name) == number


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
    catalog = replace(official, rules={**official.rules, 100: (RecognitionRule("водка"),)})
    assert Classifier(catalog).group_for("Водка 0,5л") is None


@pytest.mark.parametrize(("name", "expected"), [
    ("Антисептик", 53),
    ("Антисептик Белмастер Экобио 5л.", 53),
    ("Защитное средство Белмастер Эко-Био 10 л", 53),
    ("Антисептик Антиплесень 0,45л", 53),
    ("Антисептик для древесины ХМ-11 5л.", 53),
    ("Антисептик Биокиллер 5л. GOLDBASTIK", 53),
    ("Антисептик для рук спиртовой", None),
    ("Антисептик кожный", None),
    ("Антисептик медицинский", None),
    ("Крем-мыло Антисептическое 5л", 63),
    ("Антисептическое средство Медный купорос 100г", None),
    ("Гайка прижимная для УШМ М14", None),
    ("Краска для садовых деревьев", None),
    ("Краска БТ-177 Бронза 0,5л", 53),
    ("Клей Бустилат 0,9кг", 53),
    ("Сверло SDS+ 10*100*160мм GEPARD", 53),
    ("Шпатель 100мм фасадный с пл.руч. ВОЛАТ", 53),
    ("Шпатель для депиляции деревянный", None),
    ("Муфта 3/4 БИФИП", 53),
    ("Картридж для принтера", None),
    ("Картридж PP 5мкм для мех.очистки воды Terwa 10SL", 53),
    ("Порог дуб айсберг 30мм 0,9м", 53),
    ("Жидкие обои 3 кг", 53),
    ("Обои Зара к-91", None),
    ("Ламинат Дуб 33 класс", None),
])
def test_construction_by_researched_purpose_and_scope(name, expected):
    assert Classifier(builtin_catalog()).group_for(name) == expected


@pytest.mark.parametrize(("group", "sample"), [
    (35, "Вино Саперави"), (36, "Вино плодовое"), (37, "Вино игристое брют"), (38, "Водка"),
    (39, "Бренди"), (40, "Пиво"), (41, "Ликёр"), (42, "Сидр"), (43, "Виски"), (45, "Питьевая вода"),
    (46, "Сигареты"), (53, "Антисептик"), (57, "Порошок для стирки"), (63, "Зубная паста"), (100, "Носки"),
])
def test_every_builtin_group_is_backed_by_editable_phrases(group, sample):
    catalog = builtin_catalog()
    assert tuple(sorted(catalog.rules)) == GROUPS
    assert Classifier(catalog).group_for(sample) == group
    updated = replace(catalog, rules={**catalog.rules, group: ()})
    assert Classifier(updated).group_for(sample) is None


@pytest.mark.parametrize(("name", "expected"), [
    ("Американка 1/2 прямая Kalde", 53),
    ("Американка ПП нар. 32х3/4", 53),
    ("Кран шар латун Ду 15 с америк.", 53),
    ("Ключ для разъемных соед. американок 1/2-1 1/4", 53),
    ("Труба канализ. ПВХ 110*2,2*1000мм", 53),
    ("Колено кан. наруж.ПП 110/45", 53),
    ("Тройник 15 БФИП внутр рез", 53),
    ("Кран Маевского 1/2 д/радиатора", 53),
    ("Уголок крепеж.52*52*45*2,0", 53),
    ("Пластина креп. 140*55*2,0", 53),
    ("Петля рояльная L 250 цинк", 53),
    ("Кисть Economy + 100*15", 53),
    ("Кисть плоская 2-50мм SM", 53),
    ("Ролик нитяной 10см D15 Baulux под руч.д6", 53),
    ("Валик прижимной резиновый 180мм", 53),
    ("Шпательная лопатка 100мм ПРОФИ нерж.ст.", 53),
    ("Клещи-автомат для снятия изоляции", 53),
    ("Дрель электрическая WERKER ID 810", 53),
    ("Лопата совковая цельномет.1150мм STARTUL", 53),
    ("Угол 10*10 2,7м Идеал белый", 53),
    ("Стык береза 38мм 0,9", 53),
    ("Плита ППТ 15Н-А-Р 1000*1000*50мм", 53),
    ("Плита OSB 12*1250*2500", 53),
    ("Подкровельная пленка B 30м2 ONDUTISS", 53),
    ("Вентрешетка 135х135мм П Storm", 53),
    ("Пена монтаж. быт. WUNDER Foam 335мл.всес.", 53),
    ("Клей монт. черн. POINT 290мл. 105 морской", 53),
    ("Клей ПВА строит. GOLDBASTIK BF 07 1л.", 53),
    ("Краска ВД акриловая для кухни и ван. 3кг. FARBITEX", 53),
    ("Лак битумный БТ-577 0,5л", 53),
    ("Кисть д/рисования ArtSpace №1", None),
    ("Кисть силиконовая кондитерская", None),
    ("Ролик для одежды 50л. MF", None),
    ("Валик для массажа", None),
    ("Решетка для барбекю ECOS-RD-103С", None),
    ("Плитка эл. Blackton Bt HP101W", None),
    ("Плита газовая 2-х конф.", None),
    ("Клей ПВА Kores 60мл", None),
    ("Клей для грызунов и насек. 135гр.", None),
    ("Клей ПВА канцелярский DEKOR", None),
    ("Лак для волос POWER 225мл Тафт", None),
    ("Грунт для рассады 25л.", None),
    ("Лопата совковая игрушечная для песка", None),
    ("Тройник 4гн.б/з", None),
])
def test_whole_inventory_audit_adds_context_specific_construction_families(name, expected):
    classifier = Classifier(builtin_catalog())
    assert classifier.group_for(name) == expected


@pytest.mark.parametrize("name", [
    "Автомат ВА 47-29 1п 16А IEK", "Автомат АЕ 2045М 25А", "Диф.авт.выкл.АД-12 16А",
    "Выключатель дифф.авт.АВДТ-63/2Р С25", "УЗО ВД1-63 2Р 25А 30мА IEK",
    "Кабель АВВГ-4*16", "Кабель-канал 25*16", "Провод ПВС 3*1,5", "Розетка 7-202-01 с/з",
    "Выключатель 2 клав. 6А бел.Пралеска Аква",
])
def test_recognized_installation_goods_use_the_normal_construction_group(name):
    catalog = builtin_catalog()
    assert Classifier(catalog).group_for(name) == 53
    removed = replace(catalog, rules={**catalog.rules, 53: ()})
    assert Classifier(removed).group_for(name) is None


@pytest.mark.parametrize("name", [
    "Автомат игрушечный", "Автомат для продажи кофе", "Автомат", "Розетка для варенья 100мл",
    "Триммерная головка автомат", "Лампа галог. GU10 50W", "Стиральный порошок автомат", "Ламинат Дуб 33 класс",
])
def test_construction_rules_do_not_capture_unrelated_homonyms(name):
    assert Classifier(builtin_catalog()).group_for(name) != 53
