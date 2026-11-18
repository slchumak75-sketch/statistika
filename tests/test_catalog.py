from pathlib import Path

import pytest
from pypdf import PdfWriter
from pypdf.generic import DecodedStreamObject, DictionaryObject, NameObject

from statistika.catalog import (
    GROUPS, PDF_GROUPS, Catalog, CatalogError, extract_pdf_text, format_points, load_catalog, parse_points, save_catalog,
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
    save_catalog(Catalog(rules={35: ("особое наименование",)}), local)
    upgraded = load_catalog(local)
    assert upgraded.missing == () and upgraded.rules[35] == ("особое наименование",)


def test_saved_catalog_is_loaded_on_restart(tmp_path):
    path = tmp_path / "Статистика" / "catalog.json"
    catalog = Catalog({2: "Другой пункт", 35: "Чай", 36: "Описание"}, {35: ("чай зелёный",)}, "Справочник.pdf")
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
