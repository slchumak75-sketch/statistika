import hashlib
import json
from pathlib import Path

import pytest
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

from statistika.catalog import (
    GROUPS, PDF_GROUPS, Catalog, CatalogError, RecognitionRule, builtin_catalog, extract_pdf_text, format_points,
    load_catalog, parse_points, save_catalog,
)


def test_catalog_keeps_non_target_points_and_does_not_overwrite_pdf_point_100():
    assert GROUPS == (35, 36, 37, 38, 39, 40, 41, 42, 43, 45, 46, 53, 57, 63, 100)
    text = "УЧЕБНЫЙ ПРИМЕР\n1. Первый пункт\n35. Первая строка\nпродолжение\n35.1 Подпункт\n2\n"
    for number in PDF_GROUPS[1:]:
        text += f"№ {number} — Тестовый пункт {number}\n"
    text += "44     Дополнительный пункт\n100) Отдельный исходный пункт"
    points = parse_points(text)
    catalog = Catalog(points)
    assert catalog.missing == ()
    assert set(points) == {*PDF_GROUPS, 1, 44, 100}
    assert points[35] == "Первая строка продолжение 35.1 Подпункт"
    assert points[44] == "Дополнительный пункт"
    assert points[100] == "Отдельный исходный пункт"
    assert catalog.label(100) == "№100 — Чулочно-носочные изделия"
    assert catalog.label(None) == ""
    assert parse_points(format_points(points)) == points


def test_duplicate_requires_review():
    with pytest.raises(CatalogError, match="35 повторяется"):
        parse_points("35. Описание\n35. Другое описание")


def test_default_loads_all_confirmed_points(tmp_path):
    catalog = load_catalog(tmp_path / "catalog.json")
    assert catalog.label(100) == "№100 — Чулочно-носочные изделия"
    assert catalog.source == "4-торг инст..pdf"
    assert set(catalog.points) == set(range(1, 75))
    assert catalog.missing == ()
    assert "граппа, чача" in catalog.points[43]  # Продолжение на следующей странице.
    assert parse_points(format_points(catalog.points)) == catalog.points


def test_instruction_parser_keeps_wrapping_and_drops_chapter_titles():
    text = (
        "УКАЗАНИЯ\nпо  заполнению  формы  4-торг\n\nГЛАВА 1\nОБЩИЕ ПОЛОЖЕНИЯ\n"
        "      1. Первый текст\n4-торг (продажа)\n2\nпродолжение пункта.\n"
        "ГЛАВА 2\nСЛУЖЕБНЫЙ ЗАГОЛОВОК\n      2. Второй текст\n4 (текст о форме)"
    )
    assert parse_points(text) == {1: "Первый текст 4-торг (продажа) продолжение пункта.", 2: "Второй текст 4 (текст о форме)"}


def test_empty_prototype_catalog_does_not_hide_builtin_points(tmp_path):
    local = tmp_path / "catalog.json"
    save_catalog(Catalog(rules={35: (RecognitionRule("особое наименование"),)}), local)
    upgraded = load_catalog(local)
    assert upgraded.missing == () and upgraded.rules[35] == (RecognitionRule("особое наименование"),)


def test_saved_catalog_is_loaded_on_restart(tmp_path):
    path = tmp_path / "Статистика" / "catalog.json"
    catalog = Catalog({2: "Другой пункт", 35: "Чай", 36: "Описание"}, {35: (RecognitionRule("чай зелёный"),)}, "Справочник.pdf")
    save_catalog(catalog, path)
    assert load_catalog(path) == catalog
    assert "чай зелёный" in path.read_text(encoding="utf-8")
    assert list(path.parent.iterdir()) == [path]
    path.write_text('{"version": 20}', encoding="utf-8")
    with pytest.raises(CatalogError, match="неподдерживаемая версия"):
        load_catalog(path)


def _make_pdf(path: Path, with_text: bool) -> None:
    writer = PdfWriter()
    page = writer.add_blank_page(width=400, height=300)
    if with_text:
        # Стандартный шрифт PDF; обработку кириллицы проверяет test_catalog_keeps_… .
        page[NameObject("/Resources")] = DictionaryObject({
            NameObject("/Font"): DictionaryObject({
                NameObject("/F1"): DictionaryObject({
                    NameObject("/Type"): NameObject("/Font"),
                    NameObject("/Subtype"): NameObject("/Type1"),
                    NameObject("/BaseFont"): NameObject("/Helvetica"),
                }),
            }),
        })
        stream = DecodedStreamObject()
        stream.set_data(b"BT /F1 12 Tf 40 258 Td (7. Other point) Tj 0 -18 Td (35. Sample group) Tj ET")
        page[NameObject("/Contents")] = stream
    writer.write(path)


def test_pdf_with_text_layer(tmp_path):
    path = tmp_path / "пример.pdf"
    _make_pdf(path, True)
    text, blank = extract_pdf_text(path)
    assert blank == 0
    assert parse_points(text) == {7: "Other point", 35: "Sample group"}


def test_pdf_without_text_explains_ocr(tmp_path):
    path = tmp_path / "no-text.pdf"
    _make_pdf(path, False)
    with pytest.raises(CatalogError, match="OCR"):
        extract_pdf_text(path)


def test_legacy_rules_extend_defaults_and_cleared_lists_stay_empty(tmp_path):
    from statistika.classifier import Classifier

    path = tmp_path / 'catalog.json'
    path.write_text(json.dumps({
        "version": 1, "points": builtin_catalog().points, "source": builtin_catalog().source,
        "rules": {"53": ["монтажная смесь опытная"]},
    }), encoding="utf-8")
    catalog = load_catalog(path)
    assert Classifier(catalog).group_for("Антисептик") == 53
    assert Classifier(catalog).group_for("Монтажная смесь опытная 10 кг") == 53
    changed = Catalog(catalog.points, {**catalog.rules, 53: (), 100: ()}, catalog.source)
    save_catalog(changed, path)
    reloaded = load_catalog(path)
    assert reloaded.rules[53] == reloaded.rules[100] == ()
    classifier = Classifier(reloaded)
    assert classifier.group_for("Антисептик") is None
    assert classifier.group_for("Носки") is None
    assert classifier.group_for("Пиво") == 40


def test_v2_marked_rules_migrate_without_broadening_matches(tmp_path):
    from statistika.classifier import Classifier

    path = tmp_path / "catalog.json"
    path.write_text(json.dumps({
        "version": 2, "points": builtin_catalog().points, "source": builtin_catalog().source,
        "rules": {"53": ["# Черновик не работает", "^шпатель + фасадный", "!для депиляции",
                         "^сверло SDS+", "водосток + ^комплект"], "100": []},
    }), encoding="utf-8")
    catalog = load_catalog(path)
    assert catalog.recognition_rules(53) == (
        RecognitionRule("шпатель", at_start=True, also=("фасадный",)),
        RecognitionRule("для депиляции", exclude=True),
        RecognitionRule("сверло SDS+", at_start=True),
        RecognitionRule("водосток", also_at_start=("комплект",)),
    )
    classifier = Classifier(catalog)
    assert classifier.group_for("Шпатель фасадный") == 53
    assert classifier.group_for("Шпатель фасадный для депиляции") is None
    assert classifier.group_for("Черновик не работает") is None
    assert classifier.group_for("Комплект для водостока") == 53
    assert classifier.group_for("Лейка, комплект для водостока") is None
    assert classifier.group_for("Американка 1/2 Kalde") is None  # Изменённый список — не исходный.
    assert classifier.group_for("Носки") is None
    save_catalog(catalog, path)
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["version"] == 3 and data["rules"]["100"] == []
    assert data["rules"]["53"][0] == {"phrase": "шпатель", "at_start": True, "also": ["фасадный"]}
    assert load_catalog(path) == catalog


def test_unchanged_builtin_v2_list_can_follow_new_defaults_without_touching_custom_rules(tmp_path, monkeypatch):
    from statistika import catalog as catalog_module

    original = [{"phrase": "старый термин"}]
    fingerprint = hashlib.sha256(json.dumps(original, ensure_ascii=False, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    monkeypatch.setattr(catalog_module, "_ORIGINAL_53_V2", fingerprint)
    data = {"version": 2, "points": builtin_catalog().points, "rules": {"53": ["старый термин"], "35": ["свой термин"], "100": []}}
    path = tmp_path / "catalog.json"
    path.write_text(json.dumps(data), encoding="utf-8")
    catalog = load_catalog(path)
    assert 53 not in catalog.rules
    assert catalog.recognition_rules(53) == builtin_catalog().recognition_rules(53)
    assert catalog.rules[35] == (RecognitionRule("свой термин"),) and catalog.rules[100] == ()
    data["rules"]["53"].append("свой второй термин")
    path.write_text(json.dumps(data), encoding="utf-8")
    assert load_catalog(path).rules[53] == (RecognitionRule("старый термин"), RecognitionRule("свой второй термин"))


def test_legacy_v3_flag_is_ignored_without_losing_saved_recognition_conditions(tmp_path):
    from statistika.classifier import Classifier

    path = tmp_path / "catalog.json"
    name = "автомат"
    path.write_text(json.dumps({
        "version": 3, "points": builtin_catalog().points,
        "rules": {"53": [{"phrase": name, "at_start": True, "also": ["ва47"], "internal_only": True}], "100": []},
    }), encoding="utf-8")
    catalog = load_catalog(path)
    assert catalog.recognition_rules(53) == (RecognitionRule(name, at_start=True, also=("ва47",)),)
    classifier = Classifier(catalog)
    assert classifier.group_for("Автомат ВА47 16А") == 53
    assert classifier.group_for("Автомат игрушечный") is None
    assert classifier.group_for("Носки") is None
    save_catalog(catalog, path)
    assert "internal_only" not in path.read_text(encoding="utf-8")
    assert load_catalog(path) == catalog


def test_unchecked_phrases_stay_saved_but_do_not_classify(tmp_path):
    from dataclasses import replace

    from statistika.classifier import Classifier

    base = builtin_catalog()
    positive = RecognitionRule("шпатель", at_start=True, also=("фасадный",))
    exception = RecognitionRule("игрушечный", exclude=True)
    catalog = replace(base, rules={53: (replace(positive, disabled=True), exception), 100: ()})
    path = tmp_path / "catalog.json"
    save_catalog(catalog, path)
    saved = load_catalog(path)
    assert saved == catalog and saved.points == base.points
    data = json.loads(path.read_text(encoding="utf-8"))
    assert data["rules"]["53"][0] == {
        "phrase": "шпатель", "at_start": True, "also": ["фасадный"], "disabled": True,
    }
    assert "disabled" not in data["rules"]["53"][1]
    assert Classifier(saved).group_for("Шпатель фасадный") is None
    assert Classifier(saved).group_for("Американка 1/2 Kalde") is None  # Не восстанавливать встроенные фразы.
    assert Classifier(saved).group_for("Носки") is None

    enabled = replace(saved, rules={53: (replace(saved.rules[53][0], disabled=False), exception)})
    assert Classifier(enabled).group_for("Шпатель фасадный") == 53
    assert Classifier(enabled).group_for("Ручка, шпатель фасадный") is None
    assert Classifier(enabled).group_for("Шпатель фасадный игрушечный") is None
    no_exception = replace(enabled, rules={53: (positive, replace(exception, disabled=True))})
    assert Classifier(no_exception).group_for("Шпатель фасадный игрушечный") == 53


def test_saved_phrase_checkbox_must_be_boolean(tmp_path):
    path = tmp_path / "catalog.json"
    path.write_text(json.dumps({
        "version": 3, "points": builtin_catalog().points,
        "rules": {"53": [{"phrase": "антисептик", "disabled": "false"}]},
    }), encoding="utf-8")
    with pytest.raises(CatalogError, match="некорректные поля правила"):
        load_catalog(path)


def test_switching_phrase_to_exclusion_only_blocks_its_own_group(tmp_path):
    from dataclasses import replace

    from statistika.classifier import Classifier

    phrase = RecognitionRule("для макета", also=("пластик",))
    catalog = Catalog({53: "Стройматериалы", 57: "Бытовая химия"}, {
        53: (RecognitionRule("шпатель"), phrase),
        57: (RecognitionRule("средство для чистки"),),
    })
    path = tmp_path / "catalog.json"
    test_item = "Шпатель для макета, пластик"
    assert Classifier(catalog).group_for(test_item) == 53
    with_exclusion = replace(catalog, rules={**catalog.rules, 53: (catalog.rules[53][0], replace(phrase, exclude=True))})
    save_catalog(with_exclusion, path)
    saved = load_catalog(path)
    assert saved == with_exclusion
    classifier = Classifier(saved)
    assert classifier.group_for(test_item) is None
    assert classifier.group_for("Шпатель для макета, металл") == 53  # Условие фразы не потеряно.
    assert classifier.group_for(test_item + " средство для чистки") == 57

    restored = replace(saved, rules={**saved.rules, 53: (saved.rules[53][0], replace(saved.rules[53][1], exclude=False))})
    assert restored == catalog and Classifier(restored).group_for(test_item) == 53
