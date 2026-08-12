"""A rebrand the packaged app can run on itself.

Until now the frozen build was only ever checked by "it compiled" and "it opens
a window". That leaves the part that matters untested: whether a *packaged* app
can actually brand a PDF. The interesting failures there are all invisible from
outside —

* a missing hidden import, so `reportlab` or `openpyxl` isn't in the bundle;
* the bundled Poppins not resolving under `sys._MEIPASS`, which makes
  `_ensure_font()` return False and silently drops **every page stamp** with no
  error at all;
* poppler or the brand assets not shipping.

None of those show up in a boot smoke test, and none can be reached from outside
the process, because the Windows build is windowed and has no console to run
code from. So the app tests itself.

It builds its own source PDFs and its own brand kit, so there are no fixtures to
ship or paths to configure — which also means CI can run it. Results go to a
report file *and* the exit code, because a windowed build has nowhere to print.
"""
import sys
import traceback
from pathlib import Path

REPORT_NAME = "_self_test_report.txt"


class _Report:
    def __init__(self):
        self.lines = []
        self.failures = 0

    def check(self, name, ok, detail=""):
        ok = bool(ok)
        if not ok:
            self.failures += 1
        self.lines.append(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f"  {detail}" if detail else ""))
        return ok

    def note(self, msg):
        self.lines.append(f"       {msg}")

    def text(self, version, frozen):
        head = [f"DocRefine Pro {version} — packaged rebrand self-test",
                f"frozen: {frozen}",
                f"executable: {sys.executable}", ""]
        tail = ["", "=" * 52,
                f"RESULT: {len(self.lines_checks()) - self.failures}/{len(self.lines_checks())} passed"]
        return "\n".join(head + self.lines + tail)

    def lines_checks(self):
        return [l for l in self.lines if l.startswith("[")]


def _make_source_pdfs(src: Path):
    """Three documents, deliberately awkward: portrait, landscape, multi-page."""
    from reportlab.pdfgen import canvas as rlc
    src.mkdir(parents=True, exist_ok=True)
    made = []
    specs = [
        ("portrait-guide.pdf", (612, 792), 1),
        ("landscape-drawing.pdf", (1224, 792), 1),   # wide, like a cut sheet
        ("multipage-manual.pdf", (612, 792), 3),     # stamps must land on every page
    ]
    for name, size, pages in specs:
        p = src / name
        c = rlc.Canvas(str(p), pagesize=size)
        for i in range(pages):
            c.setFont("Helvetica", 14)
            c.drawString(40, size[1] - 60, f"SELF TEST DOCUMENT {name} page {i + 1}")
            c.drawString(40, size[1] - 90, "Original content that must survive rebranding.")
            c.showPage()
        c.save()
        made.append((name, pages))
    return made


def _make_brand_kit(kit: Path):
    """A minimal but real kit: both orientations, all five assets, plus brand.json."""
    import json
    from PIL import Image, ImageDraw
    navy = (32, 14, 101, 255)

    def png(path, w, h, alpha=255, band=True):
        img = Image.new("RGBA", (w, h), (255, 255, 255, 0))
        d = ImageDraw.Draw(img)
        if band:
            d.rectangle([0, 0, w, h], fill=navy[:3] + (alpha,))
        img.save(path)

    for orient, (w, h) in (("Portrait", (2550, 3300)), ("Landscape", (3508, 2480))):
        d = kit / orient
        d.mkdir(parents=True, exist_ok=True)
        png(d / "Header.png", w, int(h * 0.05))
        png(d / "Footer.png", w, int(h * 0.06))
        png(d / "Cover.png", w, h)
        png(d / "Back Cover.png", w, h)
        png(d / "Watermark.png", int(w * 0.5), int(h * 0.5), alpha=40)

    (kit / "brand.json").write_text(json.dumps({
        "name": "Self Test Brand",
        "slug": "self-test-brand",
        "tagline": "Trusted By The Nation",
        "disclaimer": "This guide is provided for reference. Always consult "
                      "manufacturer specifications for complete details.",
        "version_label": "Version 1.0",
        # Deliberately blank, exactly as the real Batch 4 kit leaves it: the
        # engine then composes "Last Updated: <Month Year>" from the run date.
        # Filling this in overrides that whole suffix, so a value here would
        # quietly stop the test covering the SOP's wording.
        "last_updated": "",
    }, indent=2), encoding="utf-8")


def run(workdir, src_dir=None, kit_dir=None):
    """Brand a few PDFs with every stamp on, then inspect the result.

    Returns (ok, report_text). Never raises: a crash is itself a result, and the
    caller is a command line that has to report *something*.
    """
    from .config import SystemUtils
    r = _Report()
    frozen = bool(getattr(sys, "frozen", False))
    work = Path(workdir)
    work.mkdir(parents=True, exist_ok=True)

    try:
        # --- the imports that a bad spec breaks -------------------------------
        try:
            import reportlab  # noqa: F401
            r.check("reportlab is importable in this build", True)
        except Exception as e:
            r.check("reportlab is importable in this build", False, str(e))
        try:
            import openpyxl  # noqa: F401
            r.check("openpyxl is importable in this build", True)
        except Exception as e:
            r.check("openpyxl is importable in this build", False, str(e))

        from . import rebrand, reviews
        from .worker import Worker
        from pypdf import PdfReader
        r.check("the rebrand engine is importable in this build", True)

        # --- the silent one: can the bundled font be found? -------------------
        fp = rebrand._find_font()
        r.check("the bundled brand font resolves", bool(fp), str(fp))
        r.check("_ensure_font() succeeds, so page stamps will be drawn",
                rebrand._ensure_font() is True)
        if frozen and fp:
            r.check("the font came from inside the bundle",
                    str(SystemUtils.get_resource_dir()) in str(fp), str(fp))

        # --- build inputs ------------------------------------------------------
        src = Path(src_dir) if src_dir else work / "src"
        kit = Path(kit_dir) if kit_dir else work / "kit"
        out = work / "out"
        if not src_dir:
            made = _make_source_pdfs(src)
            r.note(f"generated {len(made)} source PDFs in {src}")
        else:
            made = [(p.name, len(PdfReader(str(p)).pages)) for p in sorted(src.glob("*.pdf"))[:3]]
            r.note(f"using {len(made)} supplied PDFs from {src}")
        if not kit_dir:
            _make_brand_kit(kit)
            r.note(f"generated a brand kit in {kit}")

        bk = rebrand.BrandKit(kit)
        r.check("the brand kit resolves its required art", bk.has("portrait") or bk.has("landscape"),
                f"portrait={bk.has('portrait')} landscape={bk.has('landscape')}")

        # --- a review sheet, which also exercises openpyxl ---------------------
        rows = [{"file": name, "action": "rebrand", "doc_type": "installation guide",
                 "product": Path(name).stem, "asset_type": "installation-guide",
                 "manufacturer": "Salsbury Industries", "title": "",
                 "pages": pages, "confidence": 1.0, "source": "selftest", "notes": ""}
                for name, pages in made]
        plan = work / "self_test_plan.xlsx"
        reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS, src_root=src)
        r.check("a review sheet can be written", plan.is_file(), f"{plan.stat().st_size} bytes")
        r.check("and read back", len(reviews.read_plan(plan)) == len(rows))

        # --- the actual run, every stamp on -----------------------------------
        log = []
        w = Worker(callback=lambda e: None)
        w.log = lambda m, err=False: log.append(("ERR " if err else "") + str(m))
        w.run_rebrand_apply(str(src), str(kit), str(plan), out_dir=str(out),
                            complete_set=False, show_attribution=False,
                            keep_original_names=True,
                            stamp_opts={"footer_attribution": True, "stamp_tagline": True,
                                        "stamp_version": True, "stamp_disclaimer": True})

        # --- inspect what came out --------------------------------------------
        produced = sorted(out.rglob("*.pdf"))
        r.check("every document was produced", len(produced) == len(made),
                f"{len(produced)} of {len(made)}")

        for name, pages in made:
            got = out / name
            if not r.check(f"{name} exists in the output", got.is_file()):
                continue
            rd = PdfReader(str(got))
            r.check(f"{name} gained a front and back cover", len(rd.pages) == pages + 2,
                    f"{len(rd.pages)} pages, source had {pages}")
            r.check(f"{name} is under the 50 MB limit",
                    got.stat().st_size < 50 * 1024 * 1024,
                    f"{got.stat().st_size / 1e6:.1f} MB")
            meta = rd.metadata or {}
            r.check(f"{name} carries the brand as Author",
                    "Self Test Brand" in str(meta.get("/Author", "")) or
                    "Budget Mailboxes" in str(meta.get("/Author", "")),
                    str(meta.get("/Author")))

            # the document's own words must survive
            body = " ".join((p.extract_text() or "") for p in rd.pages).replace("\n", " ")
            r.check(f"{name} keeps its original text", "SELF TEST DOCUMENT" in body or src_dir)

            # and every stamp must be present, on a content page
            # "Last Updated:" carries its colon per the SOP (v149), so assert the
            # colon too — a silent loss of it is exactly the kind of drift that
            # got fixed once already.
            for label, needle in (("attribution", "Salsbury Industries"),
                                  ("tagline", "Trusted By The Nation"),
                                  ("version", "Version 1.0"),
                                  ("last updated", "Last Updated:"),
                                  ("disclaimer", "consult manufacturer specifications")):
                r.check(f"{name} shows the {label} stamp", needle in body)

            # a stamp drawn in a non-embedded font would break the brief's rules
            embedded = []
            for pg in rd.pages:
                for _, f in ((pg.get("/Resources") or {}).get("/Font") or {}).items():
                    d = (f.get_object().get("/FontDescriptor") or {})
                    d = d.get_object() if hasattr(d, "get_object") else d
                    if any(k in d for k in ("/FontFile", "/FontFile2", "/FontFile3")):
                        embedded.append(str(f.get_object().get("/BaseFont", "")))
            r.check(f"{name} embeds the stamp font", bool(embedded), ", ".join(sorted(set(embedded))))

        if r.failures:
            r.note("run log follows:")
            for line in log:
                r.note(line)

    except Exception:
        r.check("the self-test ran to completion", False, "exception below")
        for line in traceback.format_exc().splitlines():
            r.note(line)

    from .config import SystemUtils as SU
    text = r.text(SU.CURRENT_VERSION, frozen)
    try:
        (work / REPORT_NAME).write_text(text, encoding="utf-8")
    except Exception:
        pass
    return r.failures == 0, text
