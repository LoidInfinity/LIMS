"""Contracts for the slide analyzer that drafts content-page specs."""
from pathlib import Path
import sys
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / "shared"), str(ROOT / "tools")]
import analyze_page


def make_deck(path, image):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.enum.dml import MSO_LINE_DASH_STYLE
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.oxml import parse_xml
    from pptx.oxml.ns import nsdecls, qn
    from pptx.util import Inches, Pt
    presentation = Presentation()
    presentation.slide_width, presentation.slide_height = Inches(13.333), Inches(7.5)
    slide = presentation.slides.add_slide(presentation.slide_layouts[5])  # title only
    slide.shapes.title.text = "分析标题\v"

    frame = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2), Inches(2), Inches(5), Inches(3))
    frame.fill.background()
    frame.line.color.rgb = RGBColor(0xB8, 0xB8, 0xB7)
    frame.line.width = Pt(2)
    frame.line.dash_style = MSO_LINE_DASH_STYLE.ROUND_DOT
    frame.adjustments[0] = 0

    container = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(8), Inches(1), Inches(4), Inches(2))
    container.fill.solid()
    container.fill.fore_color.rgb = RGBColor(0x44, 0x72, 0xC4)
    srgb = container._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
    srgb.append(parse_xml('<a:alpha %s val="15000"/>' % nsdecls("a")))
    container.line.fill.background()

    slide.shapes.add_picture(str(image), Inches(2.2), Inches(2.2), Inches(2), Inches(1))

    box = slide.shapes.add_textbox(Inches(2), Inches(5.4), Inches(6), Inches(1))
    properties = box.text_frame.paragraphs[0]._p.get_or_add_pPr()
    properties.append(parse_xml('<a:buFont %s typeface="Wingdings"/>' % nsdecls("a")))
    properties.append(parse_xml('<a:buChar %s char="&#216;"/>' % nsdecls("a")))
    paragraph = box.text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "要点一"
    run.font.size = Pt(18)
    run.font.bold = True
    run_properties = run._r.get_or_add_rPr()
    run_properties.append(parse_xml(
        '<a:solidFill %s><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'
        % nsdecls("a")))
    for tag, face in (("latin", "黑体"), ("ea", "黑体"), ("cs", "Calibri")):
        run_properties.append(parse_xml('<a:%s %s typeface="%s"/>' % (tag, nsdecls("a"), face)))

    superscript = paragraph.add_run()
    superscript.text = "3-"
    superscript.font.size = Pt(18)
    superscript._r.get_or_add_rPr().set("baseline", "30000")

    citation = slide.shapes.add_textbox(Inches(4), Inches(6.5), Inches(5), Inches(0.5))
    citation_paragraph = citation.text_frame.paragraphs[0]
    for text, italic, bold in (("Nat. Mater.", True, False), (" ", False, False), ("2021", False, True),
                               (", ", False, False), ("21", True, False), (", 503.", False, False)):
        citation_run = citation_paragraph.add_run()
        citation_run.text = text
        citation_run.font.size = Pt(12)
        citation_run.font.italic = italic
        citation_run.font.bold = bold
        citation_run._r.get_or_add_rPr().append(parse_xml('<a:latin %s typeface="Calibri"/>' % nsdecls("a")))

    presentation.save(path)
    return path


class AnalyzeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def all_items(self, spec):
        return [item for page in spec["pages"] for module in page["modules"]
                for item in module["items"]]

    def analysed_items(self, spec):
        return [item for item in self.all_items(spec) if "paragraphs" in item]

    def test_analysis_captures_geometry_colors_and_assets(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (40, 30), (200, 10, 10)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        assets = self.root / "assets"
        spec = analyze_page.build_spec(deck, [1], assets, "assets", 5, template="demo")
        self.assertFalse(spec["role_styles"], "reproduction must not inject role styling")
        page = spec["pages"][0]
        self.assertEqual(page["title"].strip(), "分析标题")
        self.assertIn("title_xml", page)
        self.assertTrue(page["layout_part"].startswith("ppt/slideLayouts/"))

        modules = page["modules"]
        frames = [module for module in modules if module.get("border")]
        self.assertTrue(frames, "dashed frame must be captured as a module")
        self.assertEqual(frames[0]["border"]["dash"], "round_dot")
        self.assertEqual(frames[0]["border"]["color"], "#B8B8B7")
        self.assertEqual(frames[0].get("radius_mm", 0), 0)

        filled = [module for module in modules if isinstance(module.get("fill"), dict)]
        self.assertTrue(filled, "solid fill must be captured as a fill module")
        self.assertEqual(filled[0]["fill"]["color"], "#4472C4")
        self.assertEqual(filled[0]["fill"]["alpha_pct"], 15.0)

        items = self.all_items(spec)
        images = [item["image"] for item in items if "image" in item]
        self.assertTrue(images)
        self.assertTrue(images[0].startswith("assets/"))
        self.assertTrue((assets / Path(images[0]).name).is_file())

    def test_runs_keep_bullets_fonts_and_baseline(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (10, 10), (0, 0, 0)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        spec = analyze_page.build_spec(deck, [1], self.root / "assets", "assets", 5)
        items = self.analysed_items(spec)
        heading = next(item for item in items if "要点一" in analyze_page.payload_plain_text(item))
        self.assertEqual(heading["bullet"], {"char": "\u00d8", "font": "Wingdings"})
        self.assertEqual(heading["font_latin"], "黑体")
        self.assertEqual(heading["font_cjk"], "黑体")
        self.assertEqual(heading["font_cs"], "Calibri")
        self.assertIn("insets_mm", heading)
        runs = heading["paragraphs"][0]["runs"]
        self.assertEqual(runs[0]["text"], "要点一")
        self.assertTrue(runs[0].get("color", "").startswith("#"))
        self.assertNotEqual(runs[0]["color"], "#4472C4", "lumMod must darken the theme colour")
        superscript = next(run for run in runs if run["text"] == "3-")
        self.assertEqual(superscript["baseline"], 30000)

    def test_roles_and_citation_runs(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (10, 10), (0, 0, 0)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        spec = analyze_page.build_spec(deck, [1], self.root / "assets", "assets", 5)
        items = self.analysed_items(spec)
        citation = next(item for item in items if "Nat. Mater." in analyze_page.payload_plain_text(item))
        self.assertEqual(citation["role"], "citation")
        runs = citation["paragraphs"][0]["runs"]
        self.assertTrue(next(run for run in runs if run["text"] == "Nat. Mater.").get("italic"))
        self.assertTrue(next(run for run in runs if run["text"] == "2021").get("bold"))
        self.assertTrue(next(run for run in runs if run["text"] == "21").get("italic"))
        heading = next(item for item in items if "要点一" in analyze_page.payload_plain_text(item))
        self.assertIn(heading["role"], ("heading", "subheading"))

    def test_catalog_mode_emits_annotations(self):
        from PIL import Image
        from research_style.annotation import merge_styles
        image = self.root / "figure.png"
        Image.new("RGB", (10, 10), (0, 0, 0)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        spec = analyze_page.build_spec(deck, [1], self.root / "assets", "assets", 5,
                                       catalog=merge_styles(None))
        items = self.all_items(spec)
        citation = next(item for item in items if item.get("annotation") == "citation")
        self.assertEqual(citation["citation"],
                         {"journal": "Nat. Mater.", "year": "2021", "volume": "21", "pages": "503"})
        self.assertEqual(citation["font_pt"], 12.0)
        self.assertNotIn("paragraphs", citation)
        annotated = [item for item in items if item.get("annotation") and "paragraphs" in item]
        self.assertTrue(annotated, "ordinary text must carry a purpose annotation")
        self.assertIn(annotated[0]["annotation"],
                      ("heading", "subheading", "body", "body_large", "caption", "note", "label"))

    def test_analysis_ignores_text_free_decorations(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (10, 10), (0, 0, 0)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        spec = analyze_page.build_spec(deck, [1], self.root / "assets", "assets", 5)
        for item in self.analysed_items(spec):
            self.assertTrue("image" in item or "paragraphs" in item)
            for paragraph in item.get("paragraphs", []):
                for run in paragraph["runs"]:
                    self.assertIsInstance(run["text"], str)


if __name__ == "__main__":
    unittest.main()
