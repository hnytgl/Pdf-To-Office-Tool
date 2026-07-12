from pathlib import Path

import pytest

from pdf2word_tool.cli import validate_page_range
from pdf2word_tool.converter import ConversionJob, convert_excel, discover_jobs, normalize_format, suffix_for_format
from pdf2word_tool.excel_extractor import extract_best_tables, is_probable_stamp_object


def test_discover_single_pdf_default_output(tmp_path: Path):
    source = tmp_path / "demo.pdf"
    source.write_bytes(b"%PDF-1.7")

    jobs = discover_jobs(source, output_format="word")

    assert len(jobs) == 1
    assert jobs[0].source == source.resolve()
    assert jobs[0].target == source.with_suffix(".docx").resolve()
    assert jobs[0].output_format == "word"


def test_discover_directory_recursive_preserves_relative_paths(tmp_path: Path):
    source_dir = tmp_path / "pdfs"
    nested = source_dir / "nested"
    nested.mkdir(parents=True)
    (source_dir / "a.pdf").write_bytes(b"%PDF-1.7")
    (nested / "b.pdf").write_bytes(b"%PDF-1.7")
    output_dir = tmp_path / "out"

    jobs = discover_jobs(source_dir, output_dir, recursive=True, output_format="ppt")

    targets = {job.target for job in jobs}
    assert output_dir.resolve() / "a.pptx" in targets
    assert output_dir.resolve() / "nested" / "b.pptx" in targets


def test_validate_page_range_rejects_invalid_values():
    with pytest.raises(ValueError):
        validate_page_range(0, None)
    with pytest.raises(ValueError):
        validate_page_range(3, 2)


def test_normalize_format_and_suffixes():
    assert normalize_format("docx") == "word"
    assert normalize_format("pptx") == "ppt"
    assert normalize_format("xlsx") == "excel"
    assert suffix_for_format("excel") == ".xlsx"


def test_red_stamp_color_detection_supports_normalized_and_byte_rgb():
    assert is_probable_stamp_object({"stroking_color": (1, 0, 0)})
    assert is_probable_stamp_object({"non_stroking_color": (220, 30, 20)})
    assert not is_probable_stamp_object({"stroking_color": (0, 0, 0)})


class FakePage:
    def __init__(self, filtered: bool = False):
        self.filtered = filtered
        self.objects = {"curve": [{"stroking_color": (1, 0, 0)}]}

    def filter(self, predicate):
        assert not predicate(self.objects["curve"][0])
        return FakePage(filtered=True)

    def extract_tables(self, table_settings):
        if self.filtered and table_settings["vertical_strategy"] == "lines_strict":
            return [[['名称', '金额'], ['服务费', '123.45']]]
        return [[['章']]]


def test_best_table_extraction_prefers_stamp_filtered_business_table():
    result = extract_best_tables(FakePage())

    assert result.strategy == "seal-filtered/lines-strict"
    assert result.stamp_objects_removed == 1
    assert result.tables[0][1] == ["服务费", "123.45"]


def test_convert_stamped_table_to_excel_preserves_literal_data(tmp_path: Path):
    fitz = pytest.importorskip("fitz")
    openpyxl = pytest.importorskip("openpyxl")
    source = tmp_path / "stamped.pdf"
    target = tmp_path / "stamped.xlsx"

    document = fitz.open()
    page = document.new_page(width=420, height=240)
    xs = [40, 190, 360]
    ys = [40, 85, 130, 175]
    for x in xs:
        page.draw_line((x, ys[0]), (x, ys[-1]), color=(0, 0, 0), width=1)
    for y in ys:
        page.draw_line((xs[0], y), (xs[-1], y), color=(0, 0, 0), width=1)
    values = [["Item", "Value"], ["Service", "123.45"], ["Literal", "=1+1"]]
    for row, row_values in enumerate(values):
        for column, value in enumerate(row_values):
            page.insert_text((xs[column] + 8, ys[row] + 27), value, fontsize=11)
    # A vector red seal overlaps table borders and cells.
    page.draw_circle((270, 108), 48, color=(1, 0, 0), width=4)
    page.draw_line((235, 108), (305, 108), color=(1, 0, 0), width=3)
    document.save(source)
    document.close()

    result = convert_excel(ConversionJob(source, target, "excel"))

    assert result.status == "converted", result.message
    workbook = openpyxl.load_workbook(target, data_only=False)
    sheet = workbook["第1页"]
    assert sheet["A2"].value == "Service"
    assert sheet["B2"].value == "123.45"
    assert sheet["B3"].value == "=1+1"
    assert sheet["B3"].data_type == "s"
    assert workbook["_转换报告"].sheet_state == "hidden"
