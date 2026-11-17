"""Contracts for page prototypes and template-driven composition."""
from io import BytesIO
import json
from pathlib import Path
import sys
import tempfile
import unittest
from xml.etree import ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / "shared")]
from research_style.compose import _text_width_mm
from research_style.compose import build as compose
from research_style.core import read_json
from research_style.ingest import ingest

TITLE_TOKEN = "SECRET_TITLE_TOKEN"


def make_deck(path, image=None):
    from pptx import Presentation
    from pptx.dml.color import RGBColor
    from pptx.util import Inches, Pt
    presentation = Presentation()
    presentation.slide_width, presentation.slide_height = Inches(13.333), Inches(7.5)
    cover = presentation.slides.add_slide(presentation.slide_layouts[0])
    cover.shapes.title.text = TITLE_TOKEN
    cover.placeholders[1].text = "SECRET_SUBTITLE_TOKEN"
    if image:
        cover.shapes.add_picture(str(image), Inches(1), Inches(1), Inches(2), Inches(2))
    content = presentation.slides.add_slide(presentation.slide_layouts[1])
    content.shapes.title.text = "SECOND_TOKEN"
    for run in content.shapes.title.text_frame.paragraphs[0].runs:
        run.font.size, run.font.bold = Pt(40), True
        run.font.color.rgb = RGBColor(0x00, 0x20, 0x60)
    content.notes_slide.notes_text_frame.text = "SECRET_NOTE_TOKEN"
    presentation.save(path)
    return path


def package_parts(path):
    with zipfile.ZipFile(path) as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def shape_geometry(slide):
    return [(shape.shape_id, shape.left, shape.top, shape.width, shape.height) for shape in slide.shapes]


class ComposeTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.library = self.root / "templates"
        self.library.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def template(self, image=None):
        source = make_deck(self.root / "deck.pptx", image)
        descriptor, target, _, _ = ingest(source, self.library, school="测试大学", school_slug="test-university",
                                          tags=["测试"], contributor="unit-test", license_name="test-only",
                                          keep_slides=[(1, "cover")])
        return descriptor, target

    def layout_index(self, descriptor, name):
        return next(layout["index"] for layout in descriptor["layouts"] if layout["name"] == name)

    def test_prototype_package_keeps_design_and_drops_notes(self):
        from PIL import Image
        image = self.root / "logo.png"
        Image.new("RGB", (32, 32), (10, 20, 30)).save(image)
        descriptor, target = self.template(image=image)
        package = target / descriptor["files"]["package"]
        parts = package_parts(package)
        slides = [name for name in parts if name.startswith("ppt/slides/") and "/_rels/" not in name]
        self.assertEqual(slides, ["ppt/slides/slide1.xml"])
        self.assertIn("ppt/slides/_rels/slide1.xml.rels", parts)
        self.assertFalse([name for name in parts if name.startswith("ppt/notesSlides/")])
        media = [parts[name] for name in parts if name.startswith("ppt/media/")]
        self.assertIn(image.read_bytes(), media, "prototype pictures must survive")
        self.assertNotIn(b"SECRET_NOTE_TOKEN", b"".join(parts.values()))
        prototype = descriptor["prototypes"][0]
        self.assertEqual((prototype["id"], prototype["source_slide"]), ("cover", 1))
        self.assertIn(TITLE_TOKEN, [slot["text"] for slot in prototype["text_slots"]])
        self.assertEqual(descriptor["content_style"]["title"]["font_pt"], 40)
        self.assertTrue(descriptor["content_style"]["title"]["bold"])
        self.assertEqual(descriptor["content_style"]["title"]["color"], "#002060")
        self.assertEqual(descriptor["content_style"]["source_slide"], 2)

    def test_compose_rewrites_text_and_adds_layout_pages(self):
        from pptx import Presentation
        descriptor, target = self.template()
        title_slot = next(slot for slot in descriptor["prototypes"][0]["text_slots"]
                          if slot["text"] == TITLE_TOKEN)
        spec = {"template": descriptor["id"], "pages": [
            {"prototype": "cover", "texts": {str(title_slot["shape_id"]): "全新的标题"}},
            {"layout": self.layout_index(descriptor, "Title and Content"),
             "title": "内容页", "body": ["第一条", "第二条"]},
        ]}
        output = self.root / "out.pptx"
        manifest, _ = compose(target, spec, output)
        self.assertEqual([page["kind"] for page in manifest["pages"]], ["prototype", "layout"])
        self.assertEqual(manifest["pages"][0]["texts_changed"], [title_slot["shape_id"]])
        self.assertTrue((self.root / "out.pptx.compose.json").is_file())

        source = Presentation(str(target / descriptor["files"]["package"]))
        composed = Presentation(str(output))
        self.assertEqual(len(composed.slides), 2)
        self.assertEqual(shape_geometry(composed.slides[0]), shape_geometry(source.slides[0]))
        self.assertEqual(composed.slides[0].shapes.title.text_frame.text, "全新的标题")
        self.assertEqual(composed.slides[1].slide_layout.name, "Title and Content")
        self.assertEqual(composed.slides[1].shapes.title.text_frame.text, "内容页")

    def test_identity_compose_preserves_prototype_xml(self):
        descriptor, target = self.template()
        output = self.root / "identity.pptx"
        compose(target, {"template": descriptor["id"], "pages": [{"prototype": "cover"}]}, output)
        before = package_parts(target / descriptor["files"]["package"])["ppt/slides/slide1.xml"]
        after = package_parts(output)["ppt/slides/slide1.xml"]
        self.assertEqual(ET.tostring(ET.fromstring(before)), ET.tostring(ET.fromstring(after)),
                         "an untouched prototype must round-trip without XML changes")

    def test_compose_guards(self):
        descriptor, target = self.template()
        layout = self.layout_index(descriptor, "Title and Content")
        cases = [
            ({"template": descriptor["id"], "pages": [{"prototype": "nope"}]}, "unknown prototype"),
            ({"template": descriptor["id"], "pages": [{"prototype": "cover"}, {"prototype": "nope"}]},
             "unknown prototype"),
            ({"template": "other-template", "pages": [{"prototype": "cover"}]}, "targets template"),
            ({"template": descriptor["id"], "pages": [{"prototype": "cover", "layout": layout}]}, "exactly one"),
            ({"template": descriptor["id"], "pages": []}, "nonempty"),
            ({"template": descriptor["id"], "pages": [{"layout": 999, "title": "x"}]}, "No layout with index"),
        ]
        for spec, message in cases:
            with self.subTest(message=message):
                with self.assertRaisesRegex(ValueError, message):
                    compose(target, spec, self.root / "guard.pptx", force=True)

    def test_compose_cli(self):
        import os
        import subprocess
        descriptor, target = self.template()
        spec = self.root / "spec.json"
        spec.write_text(json.dumps({"template": descriptor["id"],
                                    "pages": [{"prototype": "cover"}],
                                    "extra": True}, ensure_ascii=False), encoding="utf-8")
        environment = os.environ.copy()
        environment["PYTHONUTF8"] = "1"
        script = ROOT / "ppt-deck" / "scripts" / "compose_deck.py"
        result = subprocess.run([sys.executable, str(script), "--template", str(target),
                                 "--spec", str(spec), "--output", str(self.root / "cli.pptx")],
                                cwd=self.root, env=environment, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("1 pages", result.stdout)

    def test_content_title_uses_captured_shape(self):
        from pptx import Presentation
        descriptor, target = self.template()
        self.assertIn("title_xml", descriptor["content_style"])
        spec = {"template": descriptor["id"], "pages": [
            {"layout": self.layout_index(descriptor, "Title and Content"), "title": "新的页面标题"}]}
        output = self.root / "titled.pptx"
        compose(target, spec, output)
        title = Presentation(str(output)).slides[0].shapes.title
        self.assertEqual(title.text_frame.text, "新的页面标题")
        run = title.text_frame.paragraphs[0].runs[0]
        self.assertEqual(run.font.size.pt, 40)
        self.assertTrue(run.font.bold)
        self.assertEqual(str(run.font.color.rgb), "002060")
        self.assertIn("ph type=\"title\"", descriptor["content_style"]["title_xml"])

    def test_modules_page_draws_borders_and_places_content(self):
        from PIL import Image
        from pptx import Presentation
        from pptx.enum.dml import MSO_LINE_DASH_STYLE
        from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
        descriptor, target = self.template()
        figure = self.root / "fig.png"
        Image.new("RGB", (100, 50), (200, 10, 10)).save(figure)
        spec = {"template": descriptor["id"], "base_dir": str(self.root), "pages": [{
            "layout": self.layout_index(descriptor, "Title and Content"), "title": "内容页",
            "modules": [{
                "rect": [10, 45, 150, 60],
                "border": {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"},
                "items": [
                    {"box": [0, 0, 1, 0.24], "text": "模块标题", "font_pt": 16, "bold": True, "align": "center"},
                    {"box": [0.05, 0.3, 0.9, 0.64], "image": "fig.png"},
                ]}]}]}
        output = self.root / "modules.pptx"
        manifest, _ = compose(target, spec, output)
        self.assertEqual(manifest["pages"][0]["modules"], [{"index": 1, "rect_mm": [10, 45, 150, 60], "items": 2}])
        slide = Presentation(str(output)).slides[0]
        border = next(shape for shape in slide.shapes
                      if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                      and shape.auto_shape_type == MSO_SHAPE.RECTANGLE)
        self.assertEqual(border.line.dash_style, MSO_LINE_DASH_STYLE.ROUND_DOT)
        self.assertEqual(str(border.line.color.rgb), "B8B8B7")
        picture = next(shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.PICTURE)
        self.assertGreaterEqual(picture.left, border.left)
        self.assertLessEqual(picture.left + picture.width, border.left + border.width)
        self.assertLessEqual(picture.top + picture.height, border.top + border.height)

    def test_modules_reject_border_crossing_and_overlap(self):
        descriptor, target = self.template()
        layout = self.layout_index(descriptor, "Title and Content")

        def page(module):
            return {"template": descriptor["id"], "pages": [{"layout": layout, "title": "t", "modules": [module]}]}

        module = {"rect": [10, 45, 150, 60], "items": [
            {"box": [0, 0, 1, 0.2], "text": "标题", "font_pt": 14}]}
        out = self.root / "guard.pptx"
        with self.assertRaisesRegex(ValueError, "leaves the module"):
            compose(target, page({"rect": [10, 45, 150, 60], "items": [
                {"box": [-0.05, 0, 1, 0.2], "text": "越界", "font_pt": 14}]}), out, force=True)
        with self.assertRaisesRegex(ValueError, "items overlap"):
            compose(target, page({"rect": [10, 45, 150, 60], "items": [
                {"box": [0, 0, 0.6, 0.4], "text": "一", "font_pt": 14},
                {"box": [0.4, 0.2, 0.6, 0.4], "text": "二", "font_pt": 14}]}), out, force=True)
        with self.assertRaisesRegex(ValueError, "needs .* mm but its box"):
            compose(target, page({"rect": [10, 45, 150, 60], "items": [
                {"box": [0, 0, 1.0, 0.05], "text": "很长很长的一段文字" * 6, "font_pt": 20}]}), out, force=True)
        with self.assertRaisesRegex(ValueError, "missing image"):
            compose(target, page({"rect": [10, 45, 150, 60], "items": [
                {"box": [0, 0, 1, 0.5], "image": "nope.png"}]}), out, force=True)
        with self.assertRaisesRegex(ValueError, "too small for a"):
            compose(target, page({"rect": [10, 45, 150, 60], "padding_mm": 0.5,
                                 "border": {"color": "#B8B8B7", "width_pt": 2.0}, "items": [
                {"box": [0, 0, 1, 0.5], "text": "标题", "font_pt": 14}]}), out, force=True)

    def test_module_radius_is_explicit(self):
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE, MSO_SHAPE_TYPE
        descriptor, target = self.template()
        spec = {"template": descriptor["id"], "pages": [{
            "layout": self.layout_index(descriptor, "Title and Content"), "title": "圆角",
            "modules": [{"rect": [10, 45, 100, 50], "radius_mm": 5,
                         "border": {"color": "#B8B8B7", "width_pt": 2.0},
                         "items": [{"box": [0, 0, 1, 0.3], "text": "标题", "font_pt": 14}]}]}]}
        output = self.root / "radius.pptx"
        compose(target, spec, output)
        slide = Presentation(str(output)).slides[0]
        rounded = next(shape for shape in slide.shapes
                       if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE
                       and shape.auto_shape_type == MSO_SHAPE.ROUNDED_RECTANGLE)
        self.assertAlmostEqual(rounded.adjustments[0], 5 / 50, places=3)

    def test_panels_split_golden_ratio_and_equal_whitespace(self):
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        descriptor, target = self.template()
        layout = self.layout_index(descriptor, "Title and Content")
        border = {"color": "#B8B8B7", "width_pt": 2.0, "dash": "round_dot"}
        spec = {"template": descriptor["id"], "pages": [{
            "layout": layout, "title": "双板块",
            "panels": {"area": [8.0, 26.0, 322.0, 156.0], "direction": "horizontal", "ratio": "golden",
                       "gap_mm": 8.0, "padding_mm": 4.0,
                       "panels": [
                           {"border": border, "items": [
                               {"box": [0, 0, 1, 0.9], "text": "左", "font_pt": 16, "align": "center"}]},
                           {"border": border, "items": [
                               {"box": [0, 0, 1, 0.9], "text": "右", "font_pt": 16, "align": "center"}]}]}}]}
        output = self.root / "panels.pptx"
        manifest, _ = compose(target, spec, output)
        panels = manifest["pages"][0]["modules"]
        self.assertEqual([panel["panel"] for panel in panels], [1, 2])
        first, second = (panel["rect_mm"] for panel in panels)
        self.assertAlmostEqual(first[2] / second[2], 1.618, places=2)
        self.assertLessEqual(abs(panels[0]["whitespace"] - panels[1]["whitespace"]), 0.08)
        slide = Presentation(str(output)).slides[0]
        borders = [shape for shape in slide.shapes
                   if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE]
        self.assertEqual(len(borders), 2)

        unbalanced = {"template": descriptor["id"], "pages": [{
            "layout": layout, "title": "t",
            "panels": {"area": [8.0, 26.0, 322.0, 156.0], "ratio": "golden", "gap_mm": 8.0,
                       "panels": [
                           {"items": [{"box": [0, 0, 1, 0.95], "text": "满", "font_pt": 14}]},
                           {"items": [{"box": [0, 0, 1, 0.4], "text": "空", "font_pt": 14}]}]}}]}
        with self.assertRaisesRegex(ValueError, "panel whitespace differs too much"):
            compose(target, unbalanced, self.root / "unbalanced.pptx", force=True)
        with self.assertRaisesRegex(ValueError, "unknown panels.ratio"):
            compose(target, {"template": descriptor["id"], "pages": [{
                "layout": layout, "title": "t",
                "panels": {"ratio": "fibonacci", "panels": [{"items": []}, {"items": []}]}}]},
                self.root / "ratio.pptx", force=True)

    def test_module_fill_keeps_alpha(self):
        from pptx import Presentation
        from pptx.enum.shapes import MSO_SHAPE_TYPE
        from pptx.oxml.ns import qn
        descriptor, target = self.template()
        spec = {"template": descriptor["id"], "pages": [{
            "layout": self.layout_index(descriptor, "Title and Content"), "title": "底色",
            "modules": [{"rect": [10, 45, 150, 60], "padding_mm": 2.0,
                         "fill": {"color": "#4472C4", "alpha_pct": 15},
                         "items": [{"box": [0, 0, 1, 0.3], "text": "标题", "font_pt": 14}]}]}]}
        output = self.root / "filled.pptx"
        compose(target, spec, output)
        slide = Presentation(str(output)).slides[0]
        shape = next(shape for shape in slide.shapes if shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE)
        alpha = shape._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr")).find(qn("a:alpha"))
        self.assertIsNotNone(alpha)
        self.assertEqual(alpha.get("val"), "15000")

    def test_text_item_carries_fonts_and_bullet(self):
        from pptx import Presentation
        from pptx.oxml.ns import qn
        descriptor, target = self.template()
        item = {"box": [0, 0, 1, 0.5], "text": "要点一", "font_pt": 18,
                "font_latin": "黑体", "font_cjk": "黑体", "font_cs": "Calibri",
                "bullet": {"char": "\u00d8", "font": "Wingdings", "size_pct": 80,
                           "mar_l_mm": 5.0, "indent_mm": -5.0}}
        spec = {"template": descriptor["id"], "pages": [{
            "layout": self.layout_index(descriptor, "Title and Content"), "title": "字体",
            "modules": [{"rect": [10, 45, 150, 60], "padding_mm": 2.0, "items": [item]}]}]}
        output = self.root / "fonts.pptx"
        compose(target, spec, output)
        slide = Presentation(str(output)).slides[0]
        shape = next(shape for shape in slide.shapes
                     if shape.has_text_frame and shape.text_frame.text.strip() == "要点一")
        run = shape.text_frame.paragraphs[0].runs[0]
        self.assertEqual(run.font.name, "黑体")
        properties = run._r.find(qn("a:rPr"))
        self.assertEqual(properties.find(qn("a:ea")).get("typeface"), "黑体")
        self.assertEqual(properties.find(qn("a:cs")).get("typeface"), "Calibri")
        paragraph = shape.text_frame.paragraphs[0]._p.find(qn("a:pPr"))
        self.assertEqual(paragraph.find(qn("a:buChar")).get("char"), "\u00d8")
        self.assertEqual(paragraph.find(qn("a:buFont")).get("typeface"), "Wingdings")
        self.assertEqual(paragraph.find(qn("a:buSzPct")).get("val"), "80000")
        self.assertEqual(paragraph.get("marL"), str(int(5.0 * 36000)))
        self.assertEqual(paragraph.get("indent"), str(int(-5.0 * 36000)))

    def test_rich_runs_citation_and_baseline(self):
        from pptx import Presentation
        from pptx.enum.text import PP_ALIGN
        from pptx.oxml.ns import qn
        descriptor, target = self.template()
        layout = self.layout_index(descriptor, "Title and Content")
        page = {"layout": layout, "title": "富文本", "modules": [
            {"rect": [10, 45, 300, 80], "padding_mm": 2.0, "items": [
                {"box": [0, 0, 1, 0.42],
                 "paragraphs": [{"runs": [{"text": "Fe(CN)"}, {"text": "6", "baseline": -25000},
                                          {"text": "3-", "baseline": 30000}]}],
                 "font_pt": 20, "font_latin": "黑体", "font_cjk": "黑体"},
                {"box": [0, 0.48, 1, 0.42], "role": "citation",
                 "citation": {"journal": "Nat. Mater.", "year": "2021",
                              "volume": "21", "pages": "503"}}]}]}
        output = self.root / "rich.pptx"
        compose(target, {"template": descriptor["id"], "pages": [page]}, output)
        slide = Presentation(str(output)).slides[0]
        rich = next(shape for shape in slide.shapes if shape.has_text_frame and "Fe(CN)" in shape.text_frame.text)
        baselines = [run._r.find(qn("a:rPr")).get("baseline")
                     for run in rich.text_frame.paragraphs[0].runs]
        self.assertEqual(baselines, [None, "-25000", "30000"])
        citation = next(shape for shape in slide.shapes
                        if shape.has_text_frame and "Nat. Mater." in shape.text_frame.text)
        segments = citation.text_frame.paragraphs[0].runs
        self.assertTrue(segments[0].font.italic)      # journal italic
        self.assertTrue(segments[2].font.bold)        # year bold
        self.assertTrue(segments[4].font.italic)      # volume italic
        self.assertEqual(segments[0].font.size.pt, 12.0)
        self.assertEqual(citation.text_frame.paragraphs[0].alignment, PP_ALIGN.CENTER)

        disabled = dict(page)
        compose(target, {"template": descriptor["id"], "role_styles": False, "pages": [disabled]},
                self.root / "norole.pptx", force=True)
        plain = Presentation(str(self.root / "norole.pptx")).slides[0]
        plain_citation = next(shape for shape in plain.shapes
                              if shape.has_text_frame and "Nat. Mater." in shape.text_frame.text)
        # Catalog styles still apply; role defaults are what role_styles turns off.
        self.assertEqual(plain_citation.text_frame.paragraphs[0].runs[0].font.size.pt, 12.0)

    def test_prototype_reuse_clones_the_slide(self):
        from PIL import Image
        from pptx import Presentation
        image = self.root / "logo.png"
        Image.new("RGB", (40, 20), (10, 120, 200)).save(image)
        descriptor, target = self.template(image=image)
        title_slot = next(slot for slot in descriptor["prototypes"][0]["text_slots"]
                          if slot["text"] == TITLE_TOKEN)
        spec = {"template": descriptor["id"], "pages": [
            {"prototype": "cover"},
            {"prototype": "cover", "texts": {str(title_slot["shape_id"]): "第二份封面"}},
        ]}
        output = self.root / "cloned.pptx"
        manifest, _ = compose(target, spec, output)
        self.assertEqual([page["clone"] for page in manifest["pages"]], [False, True])
        presentation = Presentation(str(output))
        self.assertEqual(len(presentation.slides), 2)
        self.assertEqual(presentation.slides[1].shapes.title.text_frame.text, "第二份封面")
        pictures = [shape for shape in presentation.slides[1].shapes if shape.shape_type == 13]
        self.assertTrue(pictures, "cloned slide must keep its picture relationship")
        self.assertEqual(pictures[0].image.size, (40, 20))

    def test_emphasis_color_roles(self):
        from pptx import Presentation
        descriptor, target = self.template()
        layout = self.layout_index(descriptor, "Title and Content")
        spec = {"template": descriptor["id"],
                "text_colors": {"weak": "#002060", "strong": "#C00000",
                                "normal": "#000000", "accent": "#1F8DCF"},
                "pages": [{"layout": layout, "title": "颜色", "modules": [
                    {"rect": [10, 45, 300, 80], "padding_mm": 2.0, "items": [
                        {"box": [0, 0, 1, 0.3], "annotation": "heading", "text": "小标题"},
                        {"box": [0, 0.35, 1, 0.3], "annotation": "body", "text": "正文"},
                        {"box": [0, 0.7, 1, 0.3], "annotation": "body", "color": "#ED7D31",
                         "text": "对应色占位"}]}]}]}
        output = self.root / "colors.pptx"
        compose(target, spec, output)
        slide = Presentation(str(output)).slides[0]
        colors = {}
        for shape in slide.shapes:
            if not shape.has_text_frame or not shape.text_frame.text.strip():
                continue
            run = shape.text_frame.paragraphs[0].runs[0]
            colors[shape.text_frame.text] = "#" + str(run.font.color.rgb)
        self.assertEqual(colors["小标题"], "#002060")   # weak emphasis
        self.assertEqual(colors["正文"], "#000000")     # normal
        self.assertEqual(colors["对应色占位"], "#ED7D31")  # raw colour wins

    def test_fitted_textbox_centers_rewritten_text(self):
        from pptx import Presentation
        from pptx.enum.text import PP_ALIGN
        from pptx.util import Inches, Mm, Pt
        presentation = Presentation()
        presentation.slide_width, presentation.slide_height = Inches(13.333), Inches(7.5)
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        text = "研究内容"
        # PowerPoint-era decks size TOC boxes to the text and leave them left
        # aligned; the visual design reads as centered. Rewriting must re-center.
        box = slide.shapes.add_textbox(Mm(100), Mm(80), Mm(_text_width_mm(text, 32) + 5.1), Mm(20))
        run = box.text_frame.paragraphs[0].add_run()
        run.text = text
        run.font.size = Pt(32)
        source = self.root / "fitted.pptx"
        presentation.save(source)
        descriptor, target, _, _ = ingest(source, self.library, school="测试大学", school_slug="test-university",
                                          contributor="unit-test", license_name="test-only",
                                          keep_slides=[(1, "cover")])
        slot = descriptor["prototypes"][0]["text_slots"][0]
        output = self.root / "fitted-out.pptx"
        compose(target, {"template": descriptor["id"], "pages": [
            {"prototype": "cover", "texts": {str(slot["shape_id"]): "研究结论与展望"}}]}, output)
        changed = Presentation(str(output)).slides[0].shapes[0]
        self.assertEqual(changed.text_frame.text, "研究结论与展望")
        self.assertEqual(changed.text_frame.paragraphs[0].alignment, PP_ALIGN.CENTER)


if __name__ == "__main__":
    unittest.main()
