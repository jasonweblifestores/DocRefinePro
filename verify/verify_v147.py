"""v147: per-page SOP stamps (each its own toggle) + the rebrand run history.

Run: python verify_v147.py <repo> "<...>\BrandKit\Sample Files"
"""
import sys, json, inspect, tempfile, shutil
from datetime import date
from pathlib import Path

REPO, BK = Path(sys.argv[1]), Path(sys.argv[2])
sys.path.insert(0, str(REPO))

results = []
def check(n, ok, d=""):
    results.append(bool(ok)); print(f"[{'PASS' if ok else 'FAIL'}] {n}  {d}")

from PySide6.QtWidgets import QApplication
app = QApplication.instance() or QApplication([])

from pypdf import PdfReader
from reportlab.pdfgen import canvas as rlc

from docrefine import stamps as S
from docrefine import runs, reviews, rebrand
from docrefine.rebrand import BrandKit, rebrand_pdf, page_size
from docrefine.worker import Worker
from docrefine.config import CFG, ConfigData, USER_DIR
from docrefine.gui.dialogs import RebrandDialog, PipelineDialog, StampOptions

work = Path(tempfile.mkdtemp(prefix="drp_v147_"))
REAL_RUNS_PATH = runs.RUNS_PATH   # captured before the tests redirect it
rebrand._ensure_font()
FONT = rebrand._FONT_NAME

TAGLINE = "Trusted by the Nation"
DISCLAIMER = ("This document is provided for reference only. Specifications are "
              "subject to change without notice.")

def text_pdf(path, body="INSTALLATION AND ASSEMBLY GUIDE", pages=1, size=(612, 792)):
    path.parent.mkdir(parents=True, exist_ok=True)
    c = rlc.Canvas(str(path), pagesize=size)
    for i in range(pages):
        c.setFont("Helvetica", 20); c.drawString(70, size[1] - 92, f"{body} PAGE {i + 1}")
        c.showPage()
    c.save()

def all_text(pdf):
    return " ".join(" ".join((p.extract_text() or "").split()) for p in PdfReader(str(pdf)).pages)

# a kit with wording, and the stock kit (which has no brand.json at all)
KIT_TXT = work / "kit_with_text"
shutil.copytree(BK, KIT_TXT, ignore=shutil.ignore_patterns("*.pdf"))
(KIT_TXT / "brand.json").write_text(json.dumps({
    "name": "Budget Mailboxes", "slug": "budget-mailboxes",
    "tagline": TAGLINE, "disclaimer": DISCLAIMER,
    "manufacturer_aliases": {"Florencemailboxes.com": "Florence Corporation"},
}), encoding="utf-8")

kit_plain = BrandKit(BK)
kit_text = BrandKit(KIT_TXT)
ALL_OFF = {k: False for k in StampOptions.KEYS}
ALL_ON = {k: True for k in StampOptions.KEYS}

# =====================================================================
#  A. Everything is a toggle, and every toggle starts off
# =====================================================================
cfg = ConfigData()
for key in StampOptions.KEYS:
    check(f"A1 config default off: rebrand_{key}", getattr(cfg, f"rebrand_{key}") is False)
for fn in (Worker.run_rebrand_apply, Worker.run_pipeline):
    p = inspect.signature(fn).parameters
    check(f"A2 {fn.__name__} takes stamp_opts, default None",
          "stamp_opts" in p and p["stamp_opts"].default is None)
check("A3 rebrand_pdf takes stamps, default None",
      inspect.signature(rebrand_pdf).parameters["stamps"].default is None)
check("A4 all toggles off yields no stamp object",
      Worker._stamps_for(kit_text, "Salsbury Industries", ALL_OFF, date.today()) is None)
check("A5 no stamp_opts at all yields none",
      Worker._stamps_for(kit_text, "Salsbury Industries", None, date.today()) is None)

# =====================================================================
#  B. The wording is data: a kit without it prints nothing
# =====================================================================
s_plain = Worker._stamps_for(kit_plain, "", {"stamp_tagline": True, "stamp_disclaimer": True},
                             date.today())
check("B1 kit with no brand.json produces no tagline/disclaimer stamp", s_plain is None)
check("B2 the gap is reported, naming both stamps",
      S.missing_wording(kit_plain.brand, tagline=True, disclaimer=True) == ["tagline", "disclaimer"])
check("B3 nothing is reported for stamps that are off",
      S.missing_wording(kit_plain.brand, tagline=False, disclaimer=False) == [])
check("B4 wording is read from the kit's brand.json",
      kit_text.brand.get("tagline") == TAGLINE and kit_text.brand.get("disclaimer") == DISCLAIMER)
check("B5 no disclaimer wording is hardcoded anywhere in the engine",
      not any("subject to change" in (REPO / "docrefine" / f).read_text(encoding="utf-8").lower()
              for f in ("stamps.py", "rebrand.py", "worker.py")))
s_tag = Worker._stamps_for(kit_text, "", {"stamp_tagline": True}, date.today())
check("B6 with wording present, the stamp exists", bool(s_tag) and s_tag.tagline == TAGLINE)

# =====================================================================
#  C. The manufacturer is cleaned before it is ever printed
# =====================================================================
BRAND = "Budget Mailboxes"
check("C1 a real manufacturer is kept",
      S.clean_manufacturer("Salsbury Industries", BRAND) == "Salsbury Industries")
check("C2 a website is rejected", S.clean_manufacturer("Florencemailboxes.com", BRAND) == "")
check("C3 the brand itself is rejected", S.clean_manufacturer("Budget Mailboxes", BRAND) == "")
check("C4 the seller is rejected", S.clean_manufacturer("WebLife Stores LLC", BRAND) == "")
check("C5 blank stays blank", S.clean_manufacturer("", BRAND) == "")
check("C6 a kit alias overrides the rejection",
      S.clean_manufacturer("Florencemailboxes.com", BRAND,
                           kit_text.brand.get("manufacturer_aliases")) == "Florence Corporation")
good = S.for_document(kit_text.brand, "Salsbury Industries", attribution=True)
bad = S.for_document(kit_text.brand, "Budget Mailboxes", attribution=True)
check("C7 a good value renders the SOP line",
      good.attribution == "Manufactured by Salsbury Industries | Sold by Budget Mailboxes",
      good.attribution)
check("C8 an unusable value renders no line at all, rather than a wrong one",
      bad.attribution == "" and not bad)

# =====================================================================
#  D. Version / last-updated line
# =====================================================================
v = S.for_document(kit_text.brand, version=True, today=date(2026, 8, 5))
check("D1 version and month in the SOP's format (note the colon)",
      v.version_line == "Version 1.0  ·  Last Updated: August 2026", v.version_line)
v2 = S.for_document({"version_label": "Version 2.3", "last_updated": "Last Updated Q3 2026"},
                    version=True, today=date(2026, 8, 5))
check("D2 the kit can override both", v2.version_line == "Version 2.3  ·  Last Updated Q3 2026",
      v2.version_line)

# =====================================================================
#  E. Nothing is drawn over the document
# =====================================================================
src = work / "src"
text_pdf(src / "guide.pdf", pages=3)
SRC_W, SRC_H = 612, 792
plain_out = work / "plain.pdf"
rebrand_pdf(src / "guide.pdf", plain_out, kit_text, "Installation Manual")
stamped_out = work / "stamped.pdf"
st = Worker._stamps_for(kit_text, "Salsbury Industries", ALL_ON, date(2026, 8, 5))
band = st.height(SRC_W, FONT)
rebrand_pdf(src / "guide.pdf", stamped_out, kit_text, "Installation Manual", stamps=st)

rp, rs = PdfReader(str(plain_out)), PdfReader(str(stamped_out))
check("E1 page count is unchanged (source + 2 covers)",
      len(rp.pages) == len(rs.pages) == 5, f"{len(rp.pages)} / {len(rs.pages)}")
check("E2 the stamp band has real height", band > 6, f"{band:.1f}pt")
h_plain = page_size(rp.pages[1])[1]
h_stamp = page_size(rs.pages[1])[1]
check("E3 the page grew by exactly the band height",
      abs((h_stamp - h_plain) - band) < 0.5, f"{h_plain:.1f} -> {h_stamp:.1f} (+{band:.1f})")
check("E4 the document's own page is never scaled or cropped",
      all(abs(page_size(p)[0] - SRC_W) < 0.5 for p in rs.pages[1:-1]))
check("E5 covers keep the document's page size",
      abs(page_size(rs.pages[0])[1] - SRC_H) < 0.5 and abs(page_size(rs.pages[-1])[1] - SRC_H) < 0.5)
for i in range(1, 4):
    body = rs.pages[i].extract_text() or ""
    check(f"E6.{i} original page {i} text survives", f"PAGE {i}" in body.upper())

# the content is lifted clear of BOTH the footer art and the stamp band
raw = rs.pages[1].get_contents().get_data().decode("latin-1")
lifts = [ln.split()[-2] for ln in raw.splitlines()
         if ln.strip().endswith(" cm") and ln.strip().startswith("1 0")]
check("E7 the content is translated above footer + stamps",
      any(float(x) >= band for x in lifts), f"lifts={lifts[:3]} band={band:.1f}")

# =====================================================================
#  F. Every stamp lands on every content page
# =====================================================================
for i in range(1, 4):
    t = " ".join((rs.pages[i].extract_text() or "").split())
    check(f"F1.{i} attribution on page {i}", "Manufactured by Salsbury Industries" in t)
    check(f"F2.{i} tagline on page {i}", TAGLINE in t)
    check(f"F3.{i} version line on page {i}", "Version 1.0" in t and "August 2026" in t)
    check(f"F4.{i} disclaimer on page {i}", "reference only" in t)
covers = (rs.pages[0].extract_text() or "") + (rs.pages[-1].extract_text() or "")
check("F5 stamps are not on the covers", TAGLINE not in covers and "Version 1.0" not in covers)
check("F6 metadata author still set", (rs.metadata or {}).get("/Author") == "Budget Mailboxes")
check("F7 fonts still embedded",
      b"FontFile2" in stamped_out.read_bytes() or b"FontFile" in stamped_out.read_bytes())
check("F8 still far under the 50 MB cap", stamped_out.stat().st_size < 50e6,
      f"{stamped_out.stat().st_size/1e6:.2f} MB")

# each toggle acts on its own
only_tag = Worker._stamps_for(kit_text, "Salsbury Industries", {"stamp_tagline": True}, date.today())
one_out = work / "only_tagline.pdf"
rebrand_pdf(src / "guide.pdf", one_out, kit_text, "Installation Manual", stamps=only_tag)
t1 = all_text(one_out)
check("F9 one toggle prints only its own stamp",
      TAGLINE in t1 and "Manufactured by" not in t1 and "Version 1.0" not in t1
      and "reference only" not in t1)
check("F10 a smaller band for fewer stamps",
      only_tag.height(SRC_W, FONT) < band, f"{only_tag.height(SRC_W, FONT):.1f} < {band:.1f}")

# =====================================================================
#  G. Landscape, oversize pages, and a runaway disclaimer
# =====================================================================
land = work / "land.pdf"
text_pdf(land, body="CLUSTER BOX UNIT DRAWING", size=(3168, 2448))
land_out = work / "land_out.pdf"
rebrand_pdf(land, land_out, kit_text, "Technical Drawing",
            stamps=Worker._stamps_for(kit_text, "Florence Corporation", ALL_ON, date(2026, 8, 5)))
lt = all_text(land_out)
check("G1 landscape pages carry the stamps too",
      TAGLINE in lt and "Manufactured by Florence Corporation" in lt)
check("G2 stamps scale with the page rather than staying 6pt on a 44in sheet",
      Worker._stamps_for(kit_text, "x", ALL_ON, date.today()).height(3168, FONT) > band * 3)
check("G3 the drawing's own content survives", "CLUSTER BOX UNIT" in lt.upper())

# Alignment is deliberate: attribution/version pair off left and right, while the
# tagline and disclaimer stack centred so they read as one block.
al = S.for_document(kit_text.brand, "Salsbury Industries", attribution=True, version=True,
                    tagline=True, disclaimer=True, today=date(2026, 8, 5))
al_rows, _ = al.layout(SRC_W, FONT)
kinds = {r[2] if r[0] != "pair" else "PAIR": r[0] for r in al_rows}
check("G5 attribution and version share one left/right row",
      al_rows[0][0] == "pair", str(al_rows[0][0]))
check("G6 the tagline is centred", kinds.get(TAGLINE) == "center", str(kinds.get(TAGLINE)))
disc_kinds = {r[0] for r in al_rows if "reference only" in str(r[2])}
check("G7 the disclaimer is centred too, square under the tagline",
      disc_kinds == {"center"}, str(disc_kinds))

# Each element keeps its own edge whether or not the other is there — otherwise a
# batch where only some files carry an attribution gets two footer layouts.
only_v = S.for_document(kit_text.brand, "", attribution=True, version=True, today=date(2026, 8, 5))
ov_rows, _ = only_v.layout(SRC_W, FONT)
check("G8 with no attribution, the version line still sits RIGHT",
      len(ov_rows) == 1 and ov_rows[0][0] == "right", str([r[0] for r in ov_rows]))
only_a = S.for_document(kit_text.brand, "Salsbury Industries", attribution=True, today=date(2026, 8, 5))
oa_rows, _ = only_a.layout(SRC_W, FONT)
check("G9 with no version, the attribution still sits LEFT",
      len(oa_rows) == 1 and oa_rows[0][0] == "left", str([r[0] for r in oa_rows]))

long_kit = dict(kit_text.brand); long_kit["disclaimer"] = ("word " * 400).strip()
ls = S.for_document(long_kit, disclaimer=True)
rows, _ = ls.layout(SRC_W, FONT)
check("G4 an over-long disclaimer is capped, not allowed to swallow the page",
      len(rows) == S.DISCLAIMER_MAX_LINES and rows[-1][2].endswith("…"), f"{len(rows)} lines")

# =====================================================================
#  H. Off means byte-for-byte what v146 produced
# =====================================================================
a = work / "off_a.pdf"; b = work / "off_b.pdf"
rebrand_pdf(src / "guide.pdf", a, kit_text, "Installation Manual")
rebrand_pdf(src / "guide.pdf", b, kit_text, "Installation Manual",
            stamps=Worker._stamps_for(kit_text, "Salsbury Industries", ALL_OFF, date.today()))
check("H1 stamps off changes nothing about the output",
      a.read_bytes() == b.read_bytes())

# =====================================================================
#  I. The run history: store
# =====================================================================
check("I1 the history lives in the app data folder, never inside a delivery tree",
      REAL_RUNS_PATH.parent == USER_DIR and REAL_RUNS_PATH.name == "rebrand_runs.jsonl",
      str(REAL_RUNS_PATH))
runs.RUNS_PATH = work / "runs.jsonl"          # never touch the real history in a test
runs.record("Rebrand (Apply)", r"C:\Batch 4\_unique-to-rebrand",
            output=r"C:\Batch 4\_unique-to-rebrand_rebranded", sheet=str(work / "s.xlsx"),
            kit=str(BK), brand="Budget Mailboxes", seconds=3480,
            counts={"rebranded": 874, "left": 0, "copied": 1300},
            settings={"keep_original_names": True, "stamp_tagline": True})
runs.record("Analyze", r"C:\Batch 4\_unique-to-rebrand", counts={"analyzed": 2174})
hist = runs.load()
check("I2 both runs recorded", len(hist) == 2)
check("I3 newest first", hist[0]["kind"] == "Analyze")
check("I4 the version is stamped on the record", hist[0]["version"].startswith("v"))
rec = hist[1]
check("I5 label names the folder WITH its parent, so 01_Master_Files is not ambiguous",
      runs.label(rec) == "Batch 4\\_unique-to-rebrand", runs.label(rec))
check("I6 label disambiguates a masters folder",
      runs.label({"source": r"C:\ws\Batch 4_20260709\01_Master_Files"})
      == "Batch 4_20260709\\01_Master_Files")
check("I7 result reads plainly", runs.result_text(rec) == "874 branded · 1,300 copied",
      runs.result_text(rec))
check("I8 the toggles that shaped the output are named",
      runs.settings_text(rec) == "original names, tagline", runs.settings_text(rec))
check("I9 defaults say so rather than listing nothing",
      runs.settings_text({"settings": {}}) == "defaults")
check("I10 duration is human", runs.duration_text(rec) == "58.0 min", runs.duration_text(rec))
check("I11 when is formatted", len(runs.when_text(rec)) == 16, runs.when_text(rec))

with open(runs.RUNS_PATH, "a", encoding="utf-8") as f:
    f.write("{ this is not json\n")
check("I12 a corrupt line is skipped, not fatal", len(runs.load()) == 2)
check("I13 two runs in the same second are still distinct records",
      hist[0]["ts"] != hist[1]["ts"], f'{hist[0]["ts"]} vs {hist[1]["ts"]}')
runs.remove(rec["ts"])
left = runs.load()
check("I14 forgetting a run removes exactly that one",
      len(left) == 1 and left[0]["kind"] == "Analyze", str([x["kind"] for x in left]))

runs.RUNS_PATH = work / "trim.jsonl"
old_max, runs.MAX_RECORDS = runs.MAX_RECORDS, 5
for i in range(9):
    runs.record("Analyze", f"C:\\f{i}")
check("I15 the history is capped, oldest dropped", len(runs.load()) == 5)
check("I16 the newest survive the trim", runs.load()[0]["source"] == "C:\\f8")
runs.MAX_RECORDS = old_max

# =====================================================================
#  J. Runs are recorded by the real code paths
# =====================================================================
runs.RUNS_PATH = work / "e2e.jsonl"
rows = [{"file": "guide.pdf", "action": "rebrand", "product": "Salsbury Mailbox",
         "asset_type": "installation-guide", "manufacturer": "Salsbury Industries",
         "title": "", "pages": 3, "confidence": 0.95, "source": "llm", "notes": ""}]
plan = work / "plan.xlsx"
reviews.write_plan(plan, rows, Worker.REBRAND_PLAN_COLUMNS, src_root=src)
out_e2e = work / "out_e2e"
Worker(callback=lambda e: None).run_rebrand_apply(
    str(src), str(KIT_TXT), str(plan), out_dir=str(out_e2e),
    stamp_opts=dict(ALL_ON), keep_original_names=True)
h = runs.load()
check("J1 Apply records a run", len(h) == 1 and h[0]["kind"] == "Rebrand (Apply)")
if h:
    r0 = h[0]
    check("J2 it records what it produced", (r0["counts"] or {}).get("rebranded") == 1,
          str(r0.get("counts")))
    check("J3 it records the brand kit used", r0.get("brand") == "Budget Mailboxes")
    check("J4 it records the output folder", Path(r0.get("output", "")).name == "out_e2e")
    check("J5 it records the sheet", Path(r0.get("sheet", "")).name == "plan.xlsx")
    check("J6 it records every stamp toggle",
          all(r0["settings"].get(k) is True for k in StampOptions.KEYS), str(r0["settings"]))
    check("J7 it records the naming choice", r0["settings"].get("keep_original_names") is True)
    check("J8 duration recorded", isinstance(r0.get("seconds"), float))
made = list(out_e2e.rglob("*.pdf"))
check("J9 the run really produced the branded file", len(made) == 1 and made[0].name == "guide.pdf",
      str([p.name for p in made]))
check("J10 and it carries the stamps", TAGLINE in all_text(made[0]))

runs.RUNS_PATH = work / "pipe.jsonl"
src2 = work / "src2"; text_pdf(src2 / "a.pdf")
Worker(callback=lambda e: None).run_pipeline(
    str(src2), False, True, False, kit_dir=str(KIT_TXT), out_dir=str(work / "out_pipe"),
    complete_set=False, stamp_opts={"stamp_tagline": True})
hp = runs.load()
check("J11 the pipeline records a run too",
      len(hp) == 1 and hp[0]["kind"].startswith("Pipeline:"), str([x["kind"] for x in hp]))
if hp:
    check("J12 with the stage that defined the output",
          (hp[0]["counts"] or {}).get("rebranded") == 1, str(hp[0].get("counts")))
    check("J13 and the stamp it used", hp[0]["settings"].get("stamp_tagline") is True)
piped = list((work / "out_pipe").rglob("*.pdf"))
check("J14 pipeline output carries the stamp",
      len(piped) == 1 and TAGLINE in all_text(piped[0]), str([p.name for p in piped]))
src_txt = (REPO / "docrefine" / "worker.py").read_text(encoding="utf-8")
check("J15 Analyze records a run as well", '_record_run("Analyze"' in src_txt)

# =====================================================================
#  K. Wired into the interface, and named so runs can be told apart
# =====================================================================
from docrefine.gui.main_window import MainWindow
win = MainWindow()
check("K1 the dashboard has a runs list", hasattr(win, "run_tree"))
check("K2 with its own columns",
      [win.run_tree.headerItem().text(i) for i in range(3)] == ["Run", "Result", "When"])
labels = [w.text() for w in win.findChildren(type(win.lbl_status))]
check("K3 both lists are labelled, so runs are not mistaken for ingest jobs",
      "Ingest Jobs" in labels and "Rebrand & Processing Runs" in labels)
check("K4 run details start hidden", win.gb_run.isVisibleTo(win) is False)
for b in ("btn_run_open", "btn_run_sheet", "btn_run_forget"):
    check(f"K5 {b} exists", hasattr(win, b))

runs.RUNS_PATH = work / "gui.jsonl"
runs.record("Rebrand (Apply)", r"C:\Batch 4\_unique-to-rebrand",
            output=str(out_e2e), kit=str(BK), brand="Budget Mailboxes", seconds=120,
            counts={"rebranded": 874, "copied": 1300},
            settings={"keep_original_names": True, "footer_attribution": True})
runs.record("Rebrand (Apply)", r"C:\Batch 4\_unique-to-rebrand",
            output=str(out_e2e), kit=str(BK), brand="Budget Mailboxes", seconds=90,
            counts={"rebranded": 874}, settings={})
win.refresh_run_list()
check("K6 the list loads the history", win.run_tree.topLevelItemCount() == 2)
i0, i1 = win.run_tree.topLevelItem(0), win.run_tree.topLevelItem(1)
check("K7 two runs of the SAME folder are distinguishable",
      i0.text(0) == i1.text(0) and i0.toolTip(0) != i1.toolTip(0))
check("K8 the settings are what tell them apart",
      "defaults" in i0.toolTip(0) and "footer attribution" in i1.toolTip(0))
i1.setSelected(True)
check("K9 selecting a run shows its details", win.gb_run.isVisibleTo(win) is True)
check("K10 including the brand kit", "Budget Mailboxes" in win.lbl_run_kit.text())
check("K11 and the settings", "footer attribution" in win.lbl_run_settings.text())
check("K12 the output button is live when the folder is there", win.btn_run_open.isEnabled())

# dialogs: the toggles are present, off, and carried out
prev = {k: CFG.get(f"rebrand_{k}") for k in StampOptions.KEYS}
try:
    for k in StampOptions.KEYS:
        setattr(CFG._data, f"rebrand_{k}", False)
    d = RebrandDialog(None, default_kit=str(BK), default_source=str(src))
    check("K13 RebrandDialog has the stamp group", hasattr(d, "stamps"))
    check("K14 every stamp starts unticked",
          not any(d.stamps.values().values()), str(d.stamps.values()))
    d.stamps.chk_tagline.setChecked(True)
    d.plan_path = str(plan); d.txt_plan.setText(str(plan)); d._sync_apply()
    d.on_apply()
    check("K15 ticking one is carried out of the dialog",
          d.stamp_opts == {"footer_attribution": False, "stamp_tagline": True,
                           "stamp_version": False, "stamp_disclaimer": False}, str(d.stamp_opts))
    d.deleteLater()
    p = PipelineDialog(None, default_kit=str(BK))
    check("K16 PipelineDialog has the same group", hasattr(p, "stamps"))
    check("K17 and it follows the Rebrand step being on/off",
          p.stamps.isEnabled() is p.chk_rebrand.isChecked())
    p.chk_rebrand.setChecked(False)
    check("K18 stamps grey out with no Rebrand step", p.stamps.isEnabled() is False)
    p.deleteLater()
finally:
    for k, v in prev.items():
        setattr(CFG._data, f"rebrand_{k}", v)

# =====================================================================
#  L. Inconsistent manufacturer spellings are reported, not rewritten
#     (found by running the real batch through this module: 9 such groups)
# =====================================================================
V = S.spelling_variants(["Venia Products LLC", "VENIA PRODUCTS LLC", "Venia Products LLC.",
                         "Salsbury", "Salsbury Industries", "", None, "dVault", "DVault"])
flat = {g[0].lower() for g in V}
check("L1 case-only variants are grouped", any("venia" in f for f in flat), str(V))
check("L2 punctuation-only variants join the same group",
      any(len(g) == 3 for g in V), str([len(g) for g in V]))
check("L3 genuinely different names are NOT merged",
      not any("salsbury" in g[0].lower() for g in V), str(V))
check("L4 blanks are ignored", all(all(x for x in g) for g in V))
check("L5 a single spelling reports nothing",
      S.spelling_variants(["Salsbury Industries", "Salsbury Industries"]) == [])
check("L6 the grouping is deterministic",
      S.spelling_variants(["dVault", "DVault"]) == S.spelling_variants(["DVault", "dVault"]))
check("L7 nothing is rewritten — every original spelling is preserved",
      sorted(x for g in V for x in g) == sorted(
          ["Venia Products LLC", "VENIA PRODUCTS LLC", "Venia Products LLC.", "dVault", "DVault"]))

logged = []
w2 = Worker(callback=lambda e: None)
w2.log = lambda m, err=False: logged.append(m)
w2._warn_stamp_gaps(kit_text, [
    {"action": "rebrand", "manufacturer": "Venia Products LLC"},
    {"action": "rebrand", "manufacturer": "VENIA PRODUCTS LLC"},
    {"action": "rebrand", "manufacturer": ""},
    {"action": "rebrand", "manufacturer": "Budget Mailboxes"},
    {"action": "leave", "manufacturer": "ignored spelling"},
], dict(ALL_ON))
blob = " ".join(logged)
check("L8 the omission count is reported before the run",
      "omitted on 2 file(s)" in blob and "1 with no manufacturer" in blob, blob)
check("L9 the variant warning names the fix", "manufacturer_aliases" in blob)
check("L10 rows we are not branding are left out of the warning",
      "ignored spelling" not in blob)
logged2 = []
w2.log = lambda m, err=False: logged2.append(m)
w2._warn_stamp_gaps(kit_text, [{"action": "rebrand", "manufacturer": "x.com"}], dict(ALL_OFF))
check("L12 all stamps off warns about nothing", logged2 == [], str(logged2))

win.deleteLater()
shutil.rmtree(work, ignore_errors=True)
print("\n" + "=" * 56)
print(f"RESULT: {sum(results)}/{len(results)} passed")
sys.exit(0 if all(results) else 1)
