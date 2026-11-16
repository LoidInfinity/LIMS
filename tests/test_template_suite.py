"""Contracts for the template library: content-free packages, catalog and queries."""
from io import BytesIO
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
sys.path[:0] = [str(ROOT / "shared")]
from research_style.core import read_json
from research_style.ingest import aspect_of, ingest, preview_png, slim
from research_style.template import catalog_drift, load_catalog, query, rebuild_catalog, slugify, validate_template

SECRETS = ("SECRET_TITLE_TOKEN", "SECRET_BODY_TOKEN", "SECRET_NOTE_TOKEN")


def make_source(path, image=None, size=(13.333, 7.5)):
    from pptx import Presentation
    from pptx.util import Inches
    presentation = Presentation()
    presentation.slide_width = Inches(size[0])
    presentation.slide_height = Inches(size[1])
    first = presentation.slides.add_slide(presentation.slide_layouts[0])
    first.shapes.title.text = SECRETS[0]
    first.placeholders[1].text = SECRETS[1]
    second = presentation.slides.add_slide(presentation.slide_layouts[1])
    second.shapes.title.text = "Second"
    if image:
        second.shapes.add_picture(str(image), Inches(1), Inches(1), Inches(2), Inches(2))
    second.notes_slide.notes_text_frame.text = SECRETS[2]
    presentation.save(path)
    return path


class TemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.library = self.root / "templates"
        self.library.mkdir()

    def tearDown(self):
        self.temp.cleanup()

    def base(self, name="deck.pptx", image=None, size=(13.333, 7.5)):
        return make_source(self.root / name, image, size)

    def tags(self, **extra):
        return dict(school="清华大学", school_slug="tsinghua", tags=["学术汇报", "答辩"],
                    contributor="unit-test", license_name="authorized by the author", **extra)

    def test_ingest_produces_content_free_package(self):
        from PIL import Image
        image = self.root / "figure.png"
        Image.new("RGB", (64, 64), (255, 0, 0)).save(image)
        source = self.base(image=image)
        descriptor, target, source_size, package_size = ingest(source, self.library, **self.tags())

        self.assertEqual(descriptor["id"], "tsinghua-16x9")
        self.assertEqual(descriptor["aspect"], "16:9")
        self.assertEqual(descriptor["school"]["slug"], "tsinghua")
        self.assertLess(package_size, source_size)
        for name in ("template.pptx", "preview.png", "template.json"):
            self.assertTrue((target / name).is_file(), name)
        self.assertEqual((target / "preview.png").read_bytes()[:8], b"\x89PNG\r\n\x1a\n")

        package = (target / "template.pptx").read_bytes()
        for secret in SECRETS:
            self.assertNotIn(secret.encode(), package)
        with zipfile.ZipFile(BytesIO(package)) as archive:
            names = archive.namelist()
            self.assertEqual(names[0], "[Content_Types].xml")
            self.assertFalse([name for name in names if name.startswith("ppt/slides/")])
            self.assertFalse([name for name in names if name.startswith("ppt/notesSlides/")])
            media = [archive.read(name) for name in names if name.startswith("ppt/media/")]
            self.assertNotIn(image.read_bytes(), media, "slide-only picture must be pruned")
            relationships = archive.read("ppt/_rels/presentation.xml.rels").decode("utf-8")
            self.assertNotIn('relationships/slide"', relationships)
            content_types = archive.read("[Content_Types].xml").decode("utf-8")
            self.assertNotIn("/ppt/slides/", content_types)
            self.assertTrue(any(name.startswith("ppt/slideMasters/") for name in names))

        from pptx import Presentation
        presentation = Presentation(BytesIO(package))
        self.assertEqual(len(presentation.slides), 0)
        layouts = [layout for master in presentation.slide_masters for layout in master.slide_layouts]
        self.assertEqual(len(layouts), len(descriptor["layouts"]))
        validate_template(read_json(target / "template.json"))

    def test_catalog_query_and_drift(self):
        ingest(self.base("a.pptx"), self.library, **self.tags())
        ingest(self.base("b.pptx", size=(10, 7.5)), self.library, school="北京大学", school_slug="peking",
               template_id="peking-4x3", tags=["组会"], contributor="unit-test")
        catalog = rebuild_catalog(self.library)
        self.assertEqual([entry["id"] for entry in catalog["templates"]], ["peking-4x3", "tsinghua-16x9"])
        self.assertEqual(catalog_drift(self.library), [])

        entries = load_catalog(self.library)["templates"]
        self.assertEqual([entry["id"] for entry in query(entries, school="清华")], ["tsinghua-16x9"])
        self.assertEqual([entry["id"] for entry in query(entries, school="peking")], ["peking-4x3"])
        self.assertEqual([entry["id"] for entry in query(entries, tags=["答辩"])], ["tsinghua-16x9"])
        self.assertEqual(query(entries, tags=["答辩", "组会"]), [])
        self.assertEqual([entry["id"] for entry in query(entries, aspect="16:9")], ["tsinghua-16x9"])
        self.assertEqual([entry["id"] for entry in query(entries, aspect="4:3")], ["peking-4x3"])
        self.assertEqual([entry["id"] for entry in query(entries, layout_type="title")],
                         ["peking-4x3", "tsinghua-16x9"])
        self.assertEqual(query(entries, layout_type="chart"), [])
        self.assertEqual([entry["id"] for entry in query(entries, text="清华 title")], ["tsinghua-16x9"])
        self.assertEqual(query(entries, text="no-such-token"), [])

        descriptor_path = self.library / "tsinghua" / "tsinghua-16x9" / "template.json"
        document = read_json(descriptor_path)
        document["name"] = "Renamed after the catalog was written"
        descriptor_path.write_text(json.dumps(document, ensure_ascii=False), encoding="utf-8")
        self.assertTrue(any("tsinghua-16x9" in problem for problem in catalog_drift(self.library)))

    def test_duplicate_and_slug_guards(self):
        source = self.base()
        ingest(source, self.library, **self.tags())
        with self.assertRaisesRegex(ValueError, "--force"):
            ingest(source, self.library, **self.tags())
        ingest(source, self.library, force=True, **self.tags())
        with self.assertRaisesRegex(ValueError, "--school-slug"):
            ingest(source, self.library, school="清华大学")
        with self.assertRaisesRegex(ValueError, r"\.pptx"):
            ingest(self.root / "deck.txt", self.library, **self.tags())

    def test_geometry_and_aspect(self):
        descriptor, _, _, _ = ingest(self.base(), self.library, **self.tags())
        self.assertAlmostEqual(descriptor["size"]["width_mm"], 338.667, delta=0.05)
        self.assertAlmostEqual(descriptor["size"]["height_mm"], 190.5, delta=0.05)
        title = descriptor["layouts"][0]
        self.assertEqual(title["type"], "title")
        geometry = [item for item in title["placeholders"] if "x_mm" in item]
        self.assertTrue(geometry)
        self.assertTrue(all(item["w_mm"] > 0 for item in geometry))
        self.assertEqual(aspect_of(9144000, 6858000), "4:3")
        self.assertEqual(aspect_of(12192000, 7620000), "16:10")
        self.assertEqual(slugify("清华大学"), "")
        self.assertEqual(slugify("Zhejiang University"), "zhejiang-university")

    def test_preview_is_deterministic_and_slim_keeps_masters(self):
        descriptor, target, _, _ = ingest(self.base(), self.library, **self.tags())
        first = preview_png(descriptor)
        self.assertEqual(first, preview_png(descriptor))
        self.assertEqual(first, (target / "preview.png").read_bytes())
        package = (target / "template.pptx").read_bytes()
        with zipfile.ZipFile(BytesIO(package)) as archive:
            parts = {name: archive.read(name) for name in archive.namelist()}
        trimmed = slim(parts)
        self.assertEqual(sorted(trimmed), sorted(parts))
        self.assertIn("ppt/slideMasters/slideMaster1.xml", trimmed)

    def test_cli_ingest_and_query(self):
        source = self.base()
        environment = os.environ.copy()
        environment["PYTHONUTF8"] = "1"
        ingest_script = ROOT / "ppt-deck" / "scripts" / "ingest_template.py"
        query_script = ROOT / "ppt-deck" / "scripts" / "query_template.py"
        result = subprocess.run([sys.executable, str(ingest_script), str(source), "--root", str(self.library),
                                 "--school", "浙江大学", "--school-slug", "zju", "--tags", "学术汇报,答辩",
                                 "--contributor", "unit-test", "--license", "authorized"],
                                cwd=self.root, env=environment, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertIn("zju-16x9", result.stdout)

        listing = subprocess.run([sys.executable, str(query_script), "--root", str(self.library), "list"],
                                 cwd=self.root, env=environment, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(listing.returncode, 0, listing.stdout + listing.stderr)
        self.assertIn("zju-16x9", listing.stdout)
        self.assertIn("浙江大学", listing.stdout)

        # Redirected output must stay UTF-8 even without PYTHONUTF8, otherwise
        # Chinese school names turn into mojibake in logs and pipes.
        legacy = {key: value for key, value in environment.items() if key != "PYTHONUTF8"}
        redirected = subprocess.run([sys.executable, str(query_script), "--root", str(self.library), "list"],
                                    cwd=self.root, env=legacy, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(redirected.returncode, 0, redirected.stderr)
        self.assertIn("浙江大学", redirected.stdout)

        check = subprocess.run([sys.executable, str(query_script), "--root", str(self.library), "index"],
                               cwd=self.root, env=environment, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(check.returncode, 0, check.stdout + check.stderr)

        detail = subprocess.run([sys.executable, str(query_script), "--root", str(self.library), "show", "zju-16x9"],
                                cwd=self.root, env=environment, text=True, encoding="utf-8", capture_output=True)
        self.assertEqual(detail.returncode, 0, detail.stdout + detail.stderr)
        self.assertIn("Title Slide", detail.stdout)


if __name__ == "__main__":
    unittest.main()
