import os

import pytest

from engine import docimages
from engine.model import DocImage, DocImageError


def scan(path):
    return docimages.scan_images(path)


def test_excel_floating_pictures_and_chart(samples):
    result = scan(samples["xlsx_floating"])
    assert [(i.sheet, i.cell, i.alt) for i in result] == [("Pictures", "B3", "Red square"),
                                                          ("Pictures", "F3", "Blue circle")]
    assert all(i.shape_name for i in result)  # used to find the shape in the open workbook
    assert result.charts == 1 and len(result.chart_items) == 1
    chart = result.chart_items[0]
    assert (chart.kind, chart.cell, chart.title, chart.alt) == ("chart", "H12", "Quarterly sales", "Column chart of sales")
    assert not chart.chartsheet and chart.shape_name


def test_excel_header_background_and_in_cell_pictures(samples):
    header, = scan(samples["xlsx_header"])
    assert (header.kind, header.hf, header.sheet) == ("headerfooter", "LH", "Header")
    background, = scan(samples["xlsx_background"])
    assert (background.kind, background.sheet) == ("background", "Background")
    incell = scan(samples["xlsx_incell"])
    assert (incell[0].kind, incell[0].cell, incell[0].alt) == ("incell", "A1", "Green triangle in a cell")
    assert incell[1].cell == "A3" and incell[1].alt == ""
    assert len(incell) == 52


def test_repeated_pictures_are_written_once(samples, tmp_path):
    images = scan(samples["xlsx_incell"])
    docimages.extract_files(samples["xlsx_incell"], images, str(tmp_path))
    assert len({i.path for i in images}) == 3 and len(os.listdir(tmp_path)) == 3


def test_excel_chart_sheet_is_a_chart_not_a_picture(samples):
    result = scan(samples["xlsx_chartsheet"])
    assert len(result) == 0 and result.charts == 1
    chart, = result.chart_items
    assert (chart.chartsheet, chart.sheet, chart.title) == (True, "Chart1", "Growth over time")


def test_picture_used_as_a_shape_fill_is_reported_as_unlisted(samples):
    path = samples["xlsx_shapefill"]
    result = scan(path)
    assert len(result) == 1
    assert docimages.unlisted_count(path, result) == 1


def test_word_body_table_header_and_footer(samples):
    images = scan(samples["docx"])
    assert [i.part for i in images] == ["body", "body", "body", "header", "footer"]
    red, table, second = images[:3]
    assert red.alt == "Red square" and table.context == "In a table:"
    assert second.page_est == 2  # after the page break
    assert docimages.unlisted_count(samples["docx"], images) == 0


def test_word_chart(samples):
    result = scan(samples["docx_chart"])
    picture, = result
    chart, = result.chart_items
    assert result.charts == 1
    assert (chart.kind, chart.title, chart.alt) == ("chart", "Sales by region", "Column chart of sales")
    assert chart.paragraph == picture.paragraph - 1


def test_powerpoint_pictures_and_chart(samples):
    result = scan(samples["pptx"])
    assert [i.slide for i in result] == [1, 2, 2]
    assert result[0].alt == "Blue circle logo" and result[0].context == "Company overview"
    assert result[1].alt == "" and result[2].alt == ""
    chart, = result.chart_items
    assert (chart.slide, chart.title, chart.alt) == (2, "Results by group", "Column chart of results")


def test_opendocument(samples):
    text, = scan(samples["odt"])
    assert text.alt == "A long description" and text.context == "Hello world"
    slides = scan(samples["odp"])
    assert [i.slide for i in slides] == [1, 2] and slides[1].alt == "Second slide"


def test_bad_files(tmp_path):
    junk = tmp_path / "junk.docx"
    junk.write_bytes(b"not a zip")
    with pytest.raises(DocImageError) as e:
        scan(str(junk))
    assert e.value.code == "corrupt"
    text = tmp_path / "notes.txt"
    text.write_text("hello")
    with pytest.raises(DocImageError) as e:
        scan(str(text))
    assert e.value.code == "unsupported"


def test_export_names_carry_the_location(samples, tmp_path):
    images = scan(samples["xlsx_floating"])
    docimages.extract_files(samples["xlsx_floating"], images, str(tmp_path / "x"))
    folder = tmp_path / "out"
    folder.mkdir()
    assert docimages.export_all(images, str(folder)) == 2
    assert sorted(os.listdir(folder)) == ["Pictures_B3_image_001.png", "Pictures_F3_image_002.png"]
    docx_images = scan(samples["docx"])
    names = [docimages.export_name(i) for i in docx_images]
    assert names[0].startswith("p001_") and names[3].startswith("header_") and names[4].startswith("footer_")
    chart = scan(samples["pptx"]).chart_items[0]
    assert docimages.export_name(chart) == "slide02_chart_001.png"


def test_without_pillow_only_the_formats_that_need_it_fail(tmp_path, monkeypatch):
    from PIL import Image
    png = tmp_path / "a.png"
    Image.new("RGB", (4, 4)).save(png)
    tif = tmp_path / "b.tif"
    Image.new("RGB", (4, 4)).save(tif)
    monkeypatch.setitem(__import__("sys").modules, "PIL", None)  # as if NVDA did not ship Pillow
    assert docimages.prepare_for_sharing(str(png)) == str(png)
    with pytest.raises(DocImageError) as e:
        docimages.prepare_for_sharing(str(tif))
    assert e.value.code == "nopillow"


def test_pictures_are_prepared_for_sharing(tmp_path):
    from PIL import Image
    png = tmp_path / "a.png"
    Image.new("RGB", (4, 4)).save(png)
    assert docimages.prepare_for_sharing(str(png)) == str(png)
    tif = tmp_path / "b.tif"
    Image.new("RGBA", (4, 4)).save(tif)
    converted = docimages.prepare_for_sharing(str(tif))
    assert converted.endswith(".png") and os.path.exists(converted)
