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


def classify_document(pdf_path, text=None, model=DEFAULT_MODEL, url=OLLAMA_URL):
    """Return a dict describing how to handle a PDF (see _PROMPT for keys)."""
    if text is None:
        text = extract_text(pdf_path)
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
        return {
            "action": action,
            "doc_type": str(d.get("doc_type", "")).strip(),
            "product": str(d.get("product", "")).strip(),
            "asset_type": str(d.get("asset_type", "")).strip(),
            "manufacturer": str(d.get("manufacturer", "")).strip(),
            "title": title,
            "confidence": conf,
            "source": "llm",
            "notes": "",
        }
    except Exception as e:
        return _fallback(pdf_path, note=f"model error: {str(e)[:40]}")
