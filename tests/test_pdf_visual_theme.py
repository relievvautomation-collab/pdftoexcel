"""Unit tests for PDF-derived Excel theme."""

from __future__ import annotations

import unittest

from converter.pdf_parser import ParsedPDF, TextBlock
from converter.pdf_visual_theme import (
    build_theme_for_page,
    default_sheet_theme,
    guess_paper_size,
)


def _tb(
    text: str,
    page: int = 0,
    *,
    size: float = 10.0,
    bold: bool = False,
    rgb: tuple[int, int, int] = (0, 0, 0),
) -> TextBlock:
    return TextBlock(
        text=text,
        page=page,
        x0=0.0,
        y0=0.0,
        x1=10.0,
        y1=10.0,
        font="Arial",
        size=size,
        color_rgb=rgb,
        bold=bold,
        italic=False,
        underline=False,
        align="left",
    )


class TestPdfVisualTheme(unittest.TestCase):
    def test_default_theme_body_11(self) -> None:
        t = default_sheet_theme()
        self.assertEqual(t.body_pt, 11.0)
        self.assertEqual(t.row_scale, 1.0)

    def test_build_theme_median_body_and_scale(self) -> None:
        parsed = ParsedPDF(
            page_widths=[595.0],
            page_heights=[842.0],
            text_blocks=[
                _tb("a", size=10.0),
                _tb("b", size=12.0),
                _tb("c", size=12.0),
            ],
            images=[],
            drawings=[],
            camelot_tables=[],
            plumber_tables=[],
            used_ocr=False,
        )
        th = build_theme_for_page(parsed, 0, is_traces=True)
        self.assertEqual(th.body_pt, 12.0)
        self.assertAlmostEqual(th.row_scale, 12.0 / 11.0, places=5)
        self.assertEqual(th.paper_size, 9)  # A4

    def test_empty_blocks_uses_page_metrics(self) -> None:
        parsed = ParsedPDF(
            page_widths=[612.0],
            page_heights=[792.0],
            text_blocks=[],
            images=[],
            drawings=[],
            camelot_tables=[],
            plumber_tables=[],
            used_ocr=False,
        )
        th = build_theme_for_page(parsed, 0)
        self.assertEqual(th.body_pt, 11.0)
        self.assertEqual(th.paper_size, 1)

    def test_guess_paper_size_letter(self) -> None:
        parsed = ParsedPDF(
            page_widths=[612.0],
            page_heights=[792.0],
            text_blocks=[],
            images=[],
            drawings=[],
            camelot_tables=[],
            plumber_tables=[],
            used_ocr=False,
        )
        self.assertEqual(guess_paper_size(parsed, 0), 1)


if __name__ == "__main__":
    unittest.main()
