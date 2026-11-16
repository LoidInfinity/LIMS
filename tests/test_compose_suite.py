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
            ({"template": descriptor["id"], "pages": [{"prototype": "cover"}, {"prototype": "cover"}]},
             "cannot be used twice"),
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


if __name__ == "__main__":
    unittest.main()
