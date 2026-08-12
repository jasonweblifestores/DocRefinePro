# SAVE AS: docrefine/classify.py
"""
Local document classification for the rebranding pipeline.

Uses a locally-running Ollama model to read a PDF's text and decide whether it
should be rebranded (installation guides, spec sheets, manuals) or left as-is
(CAD/technical drawings, UL/safety certifications), and to draft the product,
asset type, manufacturer, and cover title. Everything runs on the machine — no
network calls beyond localhost. Degrades gracefully to filename-based guesses
when Ollama is unavailable or a PDF has no extractable text.
"""
import json
import os
import re
import shutil
import subprocess
import time
import urllib.request
from pathlib import Path

from pypdf import PdfReader

# 127.0.0.1 (not "localhost") — on Windows localhost resolves to IPv6 ::1 first,
# which Ollama does not listen on, causing a spurious "connection refused".
OLLAMA_HOST = "http://127.0.0.1:11434"
OLLAMA_URL = OLLAMA_HOST  # back-compat alias
DEFAULT_MODEL = "llama3.2:3b"
OLLAMA_DOWNLOAD_URL = "https://ollama.com/download"


def find_ollama_exe():
    """Locate the ollama executable if installed."""
    exe = shutil.which("ollama")
    if exe:
        return exe
    candidates = [
        Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Ollama" / "ollama.exe",
        Path(os.environ.get("PROGRAMFILES", "")) / "Ollama" / "ollama.exe",
        Path("/usr/local/bin/ollama"),
        Path("/opt/homebrew/bin/ollama"),
        Path.home() / ".ollama" / "bin" / "ollama",
    ]
    for c in candidates:
        try:
            if str(c) and Path(c).exists():
                return str(c)
        except Exception:
            pass
    return None


def server_up(url=OLLAMA_HOST, timeout=3):
    try:
        with urllib.request.urlopen(url + "/api/tags", timeout=timeout) as r:
            json.loads(r.read())
        return True
    except Exception:
        return False


def list_models(url=OLLAMA_HOST, timeout=5):
    try:
        with urllib.request.urlopen(url + "/api/tags", timeout=timeout) as r:
            data = json.loads(r.read())
        return [m.get("name", "") for m in data.get("models", [])]
    except Exception:
        return []


def has_model(model=DEFAULT_MODEL, url=OLLAMA_HOST):
    base = model.split(":")[0]
    return any(n == model or n.startswith(base) for n in list_models(url))


def start_server(timeout=25):
    """Start the local Ollama server if it's installed but not running. Returns True if up."""
    if server_up():
        return True
    exe = find_ollama_exe()
    if not exe:
        return False
    try:
        kwargs = {"stdout": subprocess.DEVNULL, "stderr": subprocess.DEVNULL}
        if os.name == "nt":
            kwargs["creationflags"] = 0x08000000  # CREATE_NO_WINDOW
        subprocess.Popen([exe, "serve"], **kwargs)
    except Exception:
        return False
    deadline = time.time() + timeout
    while time.time() < deadline:
        if server_up():
            return True
        time.sleep(0.6)
    return False


def ollama_status(model=DEFAULT_MODEL):
    """Readiness state: 'ready' | 'no_model' | 'not_running' | 'not_installed'."""
    if server_up():
        return {"state": "ready" if has_model(model) else "no_model",
                "exe": find_ollama_exe(), "models": list_models()}
    exe = find_ollama_exe()
    return {"state": "not_running" if exe else "not_installed", "exe": exe, "models": []}


def pull_model(model=DEFAULT_MODEL, on_progress=None, url=OLLAMA_HOST):
    """Download a model via Ollama's HTTP API, streaming progress. Returns True on success."""
    body = json.dumps({"model": model, "stream": True}).encode()
    req = urllib.request.Request(url + "/api/pull", data=body, headers={"Content-Type": "application/json"})
    try:
        with urllib.request.urlopen(req) as resp:
            for line in resp:
                line = line.strip()
                if not line:
                    continue
                try:
                    msg = json.loads(line)
                except Exception:
                    continue
                if on_progress:
                    on_progress(msg)
                if msg.get("error"):
                    return False
                if msg.get("status") == "success":
                    return True
        return has_model(model, url)
    except Exception:
        return False


def ollama_available(model=DEFAULT_MODEL):
    """True if the model is usable — auto-starting the local server if needed."""
    if not server_up():
        start_server()
    return server_up() and has_model(model)

_PROMPT = """You classify a manufacturer's product PDF for a rebranding pipeline.

Choose "action":
- "rebrand" if it is an installation guide, spec sheet, product manual, assembly
  instructions, or a similar customer-facing document.
- "leave" if it is a CAD or technical dimensional drawing, a UL or safety
  certification, or a compliance listing.

Respond with ONLY a JSON object with these keys:
- action: "rebrand" or "leave"
- doc_type: short label (e.g. "installation guide", "spec sheet", "manual", "cad drawing", "certification")
- product: short product name or model number
- asset_type: hyphenated slug (e.g. "installation-guide", "spec-sheet", "manual", "drawing", "certification")
- manufacturer: the company that manufactured the product, or ""
- title: a SHORT cover title of 2-4 words saying what the document IS, in plain
  language a mailbox customer would use. Never include model, part or drawing
  numbers. Good: "Installation Manual", "Specification Sheet", "Product Warranty",
  "Product Care & Cleaning". Bad: "Installing 1570 F Series Mailboxes",
  "Drawing No. WEB-1932", "140055OU Side View Configuration Details".
- confidence: a number from 0 to 1

DOCUMENT TEXT:
"""


def extract_text(pdf_path, max_pages=4, max_chars=2200):
    """Text from the first few pages. Stops early once there's plenty to classify on.

    Reads past page 2 because documents that open with a full-page image would
    otherwise be misread as having no text at all.
    """
    try:
        r = PdfReader(str(pdf_path))
        parts = []
        for pg in r.pages[:max_pages]:
            parts.append(pg.extract_text() or "")
            if sum(len(p) for p in parts) >= max_chars:
                break
        return " ".join(" ".join(parts).split())[:max_chars]
    except Exception:
        return ""


# A PDF with no extractable text tells us nothing, so its filename is the only
# evidence we have. In this corpus that group is overwhelmingly CAD drawings —
# correctly left alone — but it also hides genuine installation guides that are
# drawn/outlined rather than typeset. Those would otherwise be skipped in
# silence, so a filename that names itself as instructions flips the row to
# "rebrand" for a human to confirm. A false positive costs one review decision;
# a false negative silently drops a document from the delivery.
# Note: no \b before INS — the real filenames run it straight onto a part number
# ("206550INS-1400.pdf"), so a word boundary would never match.
_INSTRUCTION_NAME_RE = re.compile(
    r"instruction|install|assembly|user[-_ ]?guide|owners?[-_ ]?manual|INS[-_]?\d", re.I)


def filename_suggests_instructions(name):
    """True if a filename names the document as instructions rather than a drawing."""
    return bool(_INSTRUCTION_NAME_RE.search(Path(name).stem))


# The mirror of the rule above. These names are how the manufacturers label
# dimensional drawings, and the brief says to leave drawings alone. This now
# decides the action outright rather than only breaking a tie: on the real batch
# the model called 213 of these "installation guide" with full confidence, and a
# sample of six turned out to be six CAD drawings.
_DRAWING_NAME_RE = re.compile(
    r"^tech[-_]|drawing|elevation|cut[-_]?sheet|[-_]cs$|bolt[-_]pattern|foundation|pad[-_]spec", re.I)


def filename_suggests_drawing(name):
    """True if a filename labels the document as a technical drawing."""
    stem = Path(name).stem
    if filename_suggests_instructions(stem):
        return False          # "…-installation-cut-sheet" is instructions first
    return bool(_DRAWING_NAME_RE.search(stem))


# A dimensioned drawing is a *visual* artifact, and the model only ever sees
# extracted text. That is the root of the misreads: on a Florence cut sheet the
# extractable text is dimension labels plus a note reading "Designed to mount …
# For use with front loading modules only", which is a fair description of an
# installation guide. The model answered the evidence it was given.
#
# These two numbers are that missing evidence, measured on the real batch:
#
#   text density (characters per square inch of page)
#       drawings  median  4.6   (90th percentile 4.6)
#       documents median 24.4   (10th percentile 4.6)
#   orientation
#       drawings  99% landscape
#       documents  6% landscape
#
# Together they separate 211 of 215 known drawings while touching 4 of 111
# documents — and on inspection 3 of those 4 were drawings the filename missed.
DRAWING_MAX_DENSITY = 10.0     # chars per square inch


def page_shape_suggests_drawing(pdf_path, text=None):
    """Structural evidence that a PDF is a dimensioned drawing, not a document.

    A landscape page carrying almost no text is a drawing; a document that reads
    like prose is not. Cheap — no rendering, no inference — and deterministic, so
    the review sheet can explain itself.
    """
    from .rebrand import page_size
    try:
        rd = PdfReader(str(pdf_path))
        w, h = page_size(rd.pages[0])
    except Exception:
        return False
    if not (w > h):
        return False
    area = (w * h) / (72.0 * 72.0)
    if area <= 0:
        return False
    if text is None:
        text = extract_text(pdf_path)
    return (len(text) / area) < DRAWING_MAX_DENSITY


# --------------------------------------------------------------------------
#  Looking at the page
# --------------------------------------------------------------------------
# Everything above reasons about a document without ever seeing it, and that is
# the ceiling on how well it can do. On this corpus 48% of files have no
# extractable text at all, so their fate was decided by filename alone: 204 were
# left as-is because the name said nothing, and 130 were branded sight-unseen
# because the name said "instructions". A drawing, a scanned guide and a UL
# certificate are all indistinguishable to a text model when there is no text.
#
# A local vision model closes that gap by actually looking. It is slower — a few
# seconds a page instead of a few hundred milliseconds — so it is used as a
# second opinion where the text pass is blind or hedging, not as a replacement.
DEFAULT_VISION_MODEL = "qwen2.5vl:7b"
VISION_DPI = 110               # legible small print without huge images
VISION_MAX_PAGES = 2
# Cap the long edge. A 44x34in drawing renders to 4840x3740 at 110dpi, which the
# model refuses outright — so the very documents this exists to identify were the
# ones it could not see. Capping also cuts the work: a vision model's cost scales
# with pixels, so this is both the fix and the speed-up.
VISION_MAX_PX = 1200
# The vision model is ~6GB of weights and an 8GB card also has to hold the KV
# cache. Measured with `ollama ps`, the default context left it at 29% CPU / 71%
# GPU — partly on the processor, which is several times slower. A smaller context
# is plenty for one short JSON reply and keeps the whole model on the GPU.
VISION_NUM_CTX = 2048
VISION_UNSURE_BELOW = 0.9      # consult vision when the text model hedges
VISION_TIMEOUT = 300

_VISION_PROMPT = """You are looking at page images from a manufacturer's product PDF.
Decide what kind of document it is FROM WHAT YOU SEE.

Tell these apart:
- A DIMENSIONED TECHNICAL DRAWING: line art of a product with measurement arrows
  and dimension labels, often a title block in a corner listing SCALE, REV,
  DRAWING NUMBER or DRAWN BY. Usually landscape, mostly white space.
- An INSTALLATION GUIDE or ASSEMBLY INSTRUCTIONS: numbered steps, exploded
  diagrams, tools or parts lists, photographs of the work being done.
- A SPEC SHEET: a table or list of specifications, dimensions and model numbers,
  often with a product photograph.
- A MANUAL: many pages of prose with headings.
- A CERTIFICATION or COMPLIANCE LISTING: a UL or similar certificate, a letter or
  a seal, naming a standard.

Choose "action":
- "rebrand" for installation guides, assembly instructions, spec sheets, manuals
  and similar customer-facing documents.
- "leave" for dimensioned technical drawings, certifications and compliance
  listings.

Respond with ONLY a JSON object with these keys:
- action: "rebrand" or "leave"
- doc_type: short label of what you SEE (e.g. "installation guide", "cad drawing",
  "spec sheet", "certification")
- product: the product name or model number shown, or ""
- asset_type: hyphenated slug (e.g. "installation-guide", "spec-sheet", "drawing")
- manufacturer: the company that MADE the product if a logo or name is visible, or ""
- title: a SHORT cover title of 2-4 words saying what the document IS, in plain
  language. Never include model, part or drawing numbers. Good: "Installation
  Manual", "Specification Sheet". Bad: "Drawing No. WEB-1932".
- confidence: a number from 0 to 1
- visual_evidence: one short phrase naming what you saw that decided it
"""


def render_pages_b64(pdf_path, pages=VISION_MAX_PAGES, dpi=VISION_DPI):
    """The first pages as base64 PNGs, for a vision model. [] if unrenderable.

    The resolution is chosen from the page size rather than fixed. A 44x34in
    drawing at 110dpi is an 18-megapixel render that then has to be thrown away
    down to 1200px — on the real batch that was 38 seconds a file, nearly all of
    it spent rasterising pixels we discard. Asking poppler for the size we want
    costs a fraction of that.
    """
    import base64
    import io
    from .processing import POPPLER_BIN, convert_from_path
    try:
        from .rebrand import page_size
        w, h = page_size(PdfReader(str(pdf_path)).pages[0])
        long_pt = max(w, h)
        if long_pt > 0:
            dpi = max(40, min(dpi, int(VISION_MAX_PX * 72.0 / long_pt)))
    except Exception:
        pass
    try:
        imgs = convert_from_path(str(pdf_path), dpi=dpi, first_page=1,
                                 last_page=pages, poppler_path=POPPLER_BIN)
    except Exception:
        return []
    out = []
    for im in imgs[:pages]:
        try:
            im = im.convert("RGB")
            if max(im.size) > VISION_MAX_PX:
                from PIL import Image as _I
                r = VISION_MAX_PX / max(im.size)
                im = im.resize((max(1, round(im.width * r)), max(1, round(im.height * r))),
                               _I.LANCZOS)
            buf = io.BytesIO()
            im.save(buf, "PNG", optimize=True)
            out.append(base64.b64encode(buf.getvalue()).decode("ascii"))
        except Exception:
            continue
    return out


def needs_a_look(info):
    """True if the text pass could not settle this file, so it is worth looking at.

    Two cases: there was no text to read (the classification came from the
    filename), or there was text but the model hedged.
    """
    src = str((info or {}).get("source", "")).lower()
    if src == "fallback":
        return True
    if src != "llm":
        return False
    try:
        return float(info.get("confidence") or 0) < VISION_UNSURE_BELOW
    except (TypeError, ValueError):
        return True


def unload_model(model, url=OLLAMA_URL):
    """Release a model from memory now, instead of waiting for Ollama's timeout.

    Ollama keeps a model resident for minutes after the last request. That is the
    right default mid-run and the wrong one once a run is over: 6GB of weights sit
    in VRAM while the user does something else entirely. `keep_alive: 0` asks for
    it back. Also used between phases, because the text and vision models
    together want more memory than an 8GB card has — held at once, the larger one
    gets pushed partly onto the CPU and runs several times slower.
    """
    try:
        payload = json.dumps({"model": model, "keep_alive": 0}).encode()
        req = urllib.request.Request(url + "/api/generate", data=payload,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            resp.read()
        return True
    except Exception:
        return False


def loaded_models(url=OLLAMA_HOST):
    """Models currently held in memory, as [(name, size_bytes)]."""
    try:
        with urllib.request.urlopen(url + "/api/ps", timeout=5) as r:
            data = json.loads(r.read()) or {}
        return [(m.get("name", ""), int(m.get("size", 0) or 0))
                for m in (data.get("models") or [])]
    except Exception:
        return []


# A smaller vision model for machines that cannot hold a 7B one on the GPU.
SMALL_VISION_MODEL = "granite3.2-vision:2b"


def model_placement(model, url=OLLAMA_HOST):
    """How much of a loaded model is on the GPU: (total_bytes, vram_bytes, pct).

    Returns None if it is not loaded. This is the only honest way to talk about
    speed, because it depends entirely on the machine: the same model is fully
    resident on a workstation card, two-thirds resident on an 8GB laptop, and
    entirely on the CPU on a machine with no usable GPU at all. Rather than assume
    any of those, ask.
    """
    try:
        with urllib.request.urlopen(url + "/api/ps", timeout=5) as r:
            data = json.loads(r.read()) or {}
        for m in (data.get("models") or []):
            if str(m.get("name", "")) == model:
                total = int(m.get("size", 0) or 0)
                vram = int(m.get("size_vram", 0) or 0)
                if total <= 0:
                    return None
                return total, vram, round(100.0 * vram / total)
    except Exception:
        pass
    return None


def has_vision_model(model=DEFAULT_VISION_MODEL, url=OLLAMA_HOST):
    """True if the vision model is downloaded and the server is up."""
    try:
        return any(str(m).startswith(model.split(":")[0]) and model in str(m)
                   for m in list_models(url))
    except Exception:
        return False


def classify_visually(pdf_path, model=DEFAULT_VISION_MODEL, url=OLLAMA_URL):
    """Classify by looking at the rendered page. None if it could not be done."""
    images = render_pages_b64(pdf_path)
    if not images:
        return None
    payload = json.dumps({
        "model": model,
        "prompt": _VISION_PROMPT,
        "images": images,
        "stream": False,
        "format": "json",
        "keep_alive": "15m",          # do not reload 6GB of weights per document
        # The reply is a short JSON object; without a cap the model can ramble for
        # thousands of tokens and each one costs wall-clock time.
        "options": {"temperature": 0.1, "num_predict": 220, "num_ctx": VISION_NUM_CTX},
    }).encode()
    try:
        req = urllib.request.Request(url + "/api/generate", data=payload,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=VISION_TIMEOUT) as resp:
            raw = json.loads(resp.read())["response"]
        d = json.loads(raw)
    except Exception:
        return None
    action = str(d.get("action", "")).lower().strip()
    if action not in ("rebrand", "leave"):
        return None
    try:
        conf = round(float(d.get("confidence", 0) or 0), 2)
    except (TypeError, ValueError):
        conf = 0.0
    from .rebrand import DEFAULT_TITLE
    seen = str(d.get("visual_evidence", "")).strip()[:80]
    return {
        "action": action,
        "doc_type": str(d.get("doc_type", "")).strip(),
        "product": str(d.get("product", "")).strip(),
        "asset_type": str(d.get("asset_type", "")).strip(),
        "manufacturer": str(d.get("manufacturer", "")).strip(),
        "title": str(d.get("title", "")).strip() or DEFAULT_TITLE,
        "confidence": conf,
        "source": "vision",
        "notes": (f"read the page: {seen}" if seen else "read the page"),
    }


def _fallback(pdf_path, note=""):
    from .rebrand import DEFAULT_TITLE
    stem = Path(pdf_path).stem
    return {
        "action": "rebrand",
        "doc_type": "",
        "product": stem.replace("_", " ").replace("-", " ").strip(),
        "asset_type": "document",
        "manufacturer": "",
        # A filename makes a poor cover title ("4C11D 09Cs"). Without any text to
        # read, the honest title is the generic one; the review sheet is where a
        # human upgrades it.
        "title": DEFAULT_TITLE,
        "confidence": 0.0,
        "source": "fallback",
        "notes": note,
    }


def classify_document(pdf_path, text=None, model=DEFAULT_MODEL, url=OLLAMA_URL,
                      vision_model=None):
    """Return a dict describing how to handle a PDF (see _PROMPT for keys).

    With `vision_model` set, a local vision model is consulted wherever the text
    pass cannot see: a PDF with no extractable text, or one the text model was
    unsure about. Where it is consulted its answer stands, because the filename
    and page-shape rules below exist only to compensate for *not* being able to
    look — once we can, the proxy is redundant.
    """
    if text is None:
        text = extract_text(pdf_path)

    if vision_model and len(text) < 15:
        seen = classify_visually(pdf_path, model=vision_model, url=url)
        if seen:
            return seen
        # fall through to the filename guesses if the model could not answer

    if len(text) < 15:
        # No extractable text — likely a drawing or an outlined/scanned page.
        # Default to LEAVE (safer: don't rebrand a cert or CAD drawing sight
        # unseen), unless the filename names it as instructions.
        if filename_suggests_instructions(pdf_path):
            from .rebrand import title_for
            fb = _fallback(pdf_path, note="no extractable text — filename says instructions; confirm")
            fb["action"] = "rebrand"
            fb["doc_type"] = "(no text — filename suggests instructions)"
            fb["asset_type"] = "installation-guide"
            fb["title"] = title_for("installation-guide")
            return fb
        fb = _fallback(pdf_path, note="no extractable text — review (may need OCR)")
        fb["action"] = "leave"
        fb["doc_type"] = "(no text)"
        return fb
    payload = json.dumps({
        "model": model,
        "prompt": _PROMPT + text,
        "stream": False,
        "format": "json",
        "options": {"temperature": 0.1},
    }).encode()
    try:
        req = urllib.request.Request(url + "/api/generate", data=payload,
                                     headers={"Content-Type": "application/json"})
        with urllib.request.urlopen(req, timeout=180) as resp:
            raw = json.loads(resp.read())["response"]
        d = json.loads(raw)
        action = str(d.get("action", "")).lower().strip()
        if action not in ("rebrand", "leave"):
            action = "rebrand"
        title = str(d.get("title", "")).strip() or _fallback(pdf_path)["title"]
        try:
            conf = round(float(d.get("confidence", 0) or 0), 2)
        except (TypeError, ValueError):
            conf = 0.0
        # The text model hedged. Looking at the page beats guessing from the
        # filename, so ask the vision model and take its answer.
        if vision_model and conf < VISION_UNSURE_BELOW:
            seen = classify_visually(pdf_path, model=vision_model, url=url)
            if seen:
                seen["notes"] = (f"text model was unsure ({conf}); " + seen["notes"])
                return seen

        note = ""
        # A landscape page with almost no text on it is a drawing, whatever its
        # filename says and whatever the model concluded from the words alone.
        # The instructions rule still wins: a sparse landscape mounting template
        # that names itself as instructions is a guide.
        if (action == "rebrand" and not filename_suggests_instructions(pdf_path)
                and page_shape_suggests_drawing(pdf_path, text)):
            action = "leave"
            note = ("looks like a dimensioned drawing — landscape page with very little "
                    "text; the brief leaves these as-is. Flip to rebrand if it is a document")
        if action == "rebrand" and filename_suggests_drawing(pdf_path):
            # The brief leaves CAD drawings alone, and the filename is the more
            # reliable signal here than the model's own label.
            #
            # This used to fire only when the model was unsure (below
            # DRAWING_TIEBREAK_BELOW). On the real batch that let 213 dimensioned
            # CAD drawings through, because the model called them "installation
            # guide" with confidence 1.0 — sampling six of them found six
            # drawings and no guides. Confidence measured how sure the model was,
            # not whether it was right.
            #
            # Leaving is also the reversible direction: the original ships
            # untouched, and the row is flagged so a human can send it back.
            action = "leave"
            note = ("filename says technical drawing — the brief leaves these as-is; "
                    "flip to rebrand if this one is really a guide or spec sheet")
        return {
            "action": action,
            "doc_type": str(d.get("doc_type", "")).strip(),
            "product": str(d.get("product", "")).strip(),
            "asset_type": str(d.get("asset_type", "")).strip(),
            "manufacturer": str(d.get("manufacturer", "")).strip(),
            "title": title,
            "confidence": conf,
            "source": "llm",
            "notes": note,
        }
    except Exception as e:
        return _fallback(pdf_path, note=f"model error: {str(e)[:40]}")
