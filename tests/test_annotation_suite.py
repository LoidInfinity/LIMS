"""Contracts for content annotations and the per-purpose style catalog."""
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / "shared")]
from research_style.annotation import (DEFAULT_TEXT_STYLES, classify, citation_segments,
                                       derive_text_styles, merge_styles, parse_citation,
                                       run_annotation, style_for)


def make_deck(path):
    from pptx import Presentation
    from pptx.util import Inches, Pt
    presentation = Presentation()
    presentation.slide_width, presentation.slide_height = Inches(13.333), Inches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    first = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(6), Inches(0.5))
    run = first.text_frame.paragraphs[0].add_run()
    run.text = "Nat. Mater. 2021, 21, 503."
    run.font.size = Pt(10)
    second = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(6), Inches(0.5))
    run = second.text_frame.paragraphs[0].add_run()
    run.text = "Chem. Rev. 2020, 120, 7399."
    run.font.size = Pt(14)
    heading = slide.shapes.add_textbox(Inches(1), Inches(3), Inches(6), Inches(0.6))
    run = heading.text_frame.paragraphs[0].add_run()
    run.text = "研究背景"
    run.font.size = Pt(16)
    run.font.bold = True
    presentation.save(path)
    return path


class AnnotationTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_parse_citation_variants(self):
        cases = {
            "Nat. Mater. 2021, 21, 503.": ("Nat. Mater.", "2021", "21", "503"),
            "Therm. Sci. Eng. Prog. 2023, 46, 102244.": ("Therm. Sci. Eng. Prog.", "2023", "46", "102244"),
            "EnergyChem 2024, 6, 100123.": ("EnergyChem", "2024", "6", "100123"),
            "Science 2020, 368, 1091.": ("Science", "2020", "368", "1091"),
            # Article-style references use 7-digit article numbers.
            "Adv. Funct. Mater. 2024, 34, 2314419.": ("Adv. Funct. Mater.", "2024", "34", "2314419"),
            "Adv. Funct. Mater. 2022, 32, 2208286.": ("Adv. Funct. Mater.", "2022", "32", "2208286"),
            "Energy Environ. Sci. 2023, 16, 1234-1245": ("Energy Environ. Sci.", "2023", "16", "1234-1245"),
            "Matter 2024, 7, e2304567": ("Matter", "2024", "7", "e2304567"),
        }
        for text, expected in cases.items():
            parsed = parse_citation(text)
            self.assertIsNotNone(parsed, text)
            self.assertEqual((parsed["journal"], parsed["year"], parsed["volume"], parsed["pages"]), expected)
        for text in ("研究背景", "See Fig. 2 for details", "2021 年总结", "doi:10.1000/xyz"):
            self.assertIsNone(parse_citation(text), text)

    def test_classify_by_purpose(self):
        self.assertEqual(classify("Nat. Mater. 2021, 21, 503.", 12, False), "citation")
        self.assertEqual(classify("研究背景", 16, True), "heading")
        self.assertEqual(classify("研究背景", 14, True), "subheading")
        self.assertEqual(classify("谢谢", 28, True), "display")
        self.assertEqual(classify("一句说明", 12, False, False, True), "caption")
        self.assertEqual(classify("一段正文", 18, False), "body_large")
        self.assertEqual(classify("一段正文", 14, False), "body")
        self.assertEqual(classify("小注释", 10.5, False), "note")
        self.assertEqual(classify("9 kPa", 18, True), "data_label")
        self.assertEqual(classify("1.82 mV/K", 18, True), "data_label")
        self.assertEqual(classify("2024", 18, True), "data_label")
        self.assertIsNone(classify("   ", 12, False))

    def test_run_annotation(self):
        self.assertEqual(run_annotation("#C00000", "#000000", None, False, False), "emphasis")
        self.assertEqual(run_annotation("#000000", "#000000", 30000, False, False), "superscript")
        self.assertEqual(run_annotation("#000000", "#000000", -25000, False, False), "subscript")
        self.assertIsNone(run_annotation("#000000", "#000000", None, False, False))

    def test_derived_catalog_normalises_same_purpose(self):
        deck = make_deck(self.root / "deck.pptx")
        from pptx import Presentation
        catalog = merge_styles(derive_text_styles(Presentation(deck)))
        citation = catalog["citation"]
        # Citation boxes at 10/12/14 pt collapse to one canonical size, including
        # the 7-digit article-number reference.
        self.assertIn(citation["font_pt"], (10.0, 12.0, 14.0))
        self.assertIn("parts", citation)
        self.assertEqual(citation["parts"]["year"], {"bold": True})
        self.assertEqual(citation["parts"]["journal"], {"italic": True})
        self.assertEqual(catalog["heading"]["font_pt"], 16.0)
        self.assertTrue(catalog["heading"]["bold"])

    def test_style_resolution_and_segments(self):
        catalog = merge_styles({"heading": {"font_pt": 18, "color": "#FF0000"},
                                "citation": {"parts": {"year": {"italic": False}}}})
        style = style_for(catalog, "heading", {"align": "center"})
        self.assertEqual(style["font_pt"], 18)
        self.assertEqual(style["color"], "#FF0000")
        self.assertEqual(style["align"], "center")
        self.assertEqual(catalog["citation"]["parts"]["year"], {"bold": True, "italic": False})
        self.assertEqual(catalog["citation"]["parts"]["journal"], {"italic": True})
        self.assertIn("heading", DEFAULT_TEXT_STYLES)
        parsed = parse_citation("Nat. Mater. 2021, 21, 503.")
        parts = [part for part, _ in citation_segments(parsed)]
        self.assertEqual(parts, ["journal", "separator", "year", "separator",
                                 "volume", "separator", "pages", "separator"])

    def test_emphasis_color_roles(self):
        from research_style.annotation import (DEFAULT_TEXT_COLORS, color_role_for,
                                               derive_text_colors, merge_text_colors)
        palette = merge_text_colors({"weak": "#002060", "strong": "#C00000", "accent": "#1F8DCF"})
        self.assertEqual(palette["normal"], "#000000")
        self.assertEqual(color_role_for("#000000", palette), "normal")
        self.assertEqual(color_role_for("#002060", palette), "weak")
        self.assertEqual(color_role_for("#C00000", palette), "strong")
        self.assertEqual(color_role_for("#1F8DCF", palette), "accent")
        # An orange correspondence colour has no role yet: keep it raw.
        self.assertIsNone(color_role_for("#ED7D31", palette))
        self.assertIsNone(DEFAULT_TEXT_COLORS["corresponding"])
        deck = make_deck(self.root / "deck.pptx")
        from pptx import Presentation
        colors = derive_text_colors(Presentation(deck))
        for role in ("normal", "weak", "strong", "accent"):
            self.assertIn(role, colors)


if __name__ == "__main__":
    unittest.main()
