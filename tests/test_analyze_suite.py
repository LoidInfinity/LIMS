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
    box = slide.shapes.add_textbox(Inches(2), Inches(6), Inches(6), Inches(1))
    paragraph = box.text_frame.paragraphs[0]
    properties = paragraph._p.get_or_add_pPr()
    properties.append(parse_xml('<a:buFont %s typeface="Wingdings"/>' % nsdecls("a")))
    properties.append(parse_xml('<a:buChar %s char="&#216;"/>' % nsdecls("a")))
    run = paragraph.add_run()
    run.text = "要点一"
    run.font.size = Pt(18)
    run_properties = run._r.get_or_add_rPr()
    run_properties.append(parse_xml(
        '<a:solidFill %s><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:solidFill>'
        % nsdecls("a")))
    for tag, face in (("latin", "黑体"), ("ea", "黑体"), ("cs", "Calibri")):
        run_properties.append(parse_xml('<a:%s %s typeface="%s"/>' % (tag, nsdecls("a"), face)))
    presentation.save(path)
    return path


class AnalyzeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)

    def tearDown(self):
        self.temp.cleanup()

    def test_analysis_captures_geometry_colors_and_assets(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (40, 30), (200, 10, 10)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        assets = self.root / "assets"
        spec = analyze_page.build_spec(deck, [1], assets, "assets", 5, template="demo")
        page = spec["pages"][0]
        self.assertEqual(page["title"].strip(), "分析标题")
        self.assertIn("title_xml", page)
        self.assertTrue(page["layout_part"].startswith("ppt/slideLayouts/"))
        self.assertEqual(spec["template"], "demo")

        modules = page["modules"]
        frames = [module for module in modules if module.get("border")]
        self.assertTrue(frames, "dashed frame must be captured as a module")
        self.assertEqual(frames[0]["border"]["dash"], "round_dot")
        self.assertEqual(frames[0]["border"]["color"], "#B8B8B7")
        self.assertEqual(frames[0]["border"]["width_pt"], 2.0)
        self.assertEqual(frames[0].get("radius_mm", 0), 0)

        filled = [module for module in modules if isinstance(module.get("fill"), dict)]
        self.assertTrue(filled, "solid fill must be captured as a fill module")
        self.assertEqual(filled[0]["fill"]["color"], "#4472C4")
        self.assertEqual(filled[0]["fill"]["alpha_pct"], 15.0)

        items = [item for module in modules for item in module["items"]]
        images = [item["image"] for item in items if "image" in item]
        self.assertTrue(images)
        self.assertTrue(images[0].startswith("assets/"))
        self.assertTrue((assets / Path(images[0]).name).is_file())
        texts = [item for item in items if "text" in item]
        self.assertTrue(texts)
        sample = next(item for item in texts if "要点一" in item["text"])
        self.assertNotIn("• ", sample["text"], "bullet glyphs belong to the paragraph, not the text")
        self.assertEqual(sample["bullet"]["char"], "\u00d8")
        self.assertEqual(sample["bullet"]["font"], "Wingdings")
        self.assertEqual(sample["font_latin"], "黑体")
        self.assertEqual(sample["font_cjk"], "黑体")
        self.assertEqual(sample["font_cs"], "Calibri")
        colored = [item for item in items if item.get("color")]
        self.assertTrue(colored)
        # accent1 with lumMod 50% must resolve to the darker shade, not #4472C4.
        self.assertNotEqual(colored[0]["color"], "#4472C4")

    def test_analysis_ignores_text_free_decorations(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (10, 10), (0, 0, 0)).save(image)
        deck = make_deck(self.root / "deck.pptx", image)
        spec = analyze_page.build_spec(deck, [1], self.root / "assets", "assets", 5)
        for page in spec["pages"]:
            for module in page["modules"]:
                for item in module["items"]:
                    self.assertTrue("image" in item or "text" in item)
                    for key in ("text", "font_pt"):
                        if key in item:
                            self.assertIsInstance(item[key], (str, int, float))


if __name__ == "__main__":
    unittest.main()
