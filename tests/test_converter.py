from pathlib import Path

import pytest

from pdf2word_tool.cli import validate_page_range
from pdf2word_tool.converter import discover_jobs, normalize_format, suffix_for_format


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
