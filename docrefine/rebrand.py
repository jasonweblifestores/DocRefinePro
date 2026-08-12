# SAVE AS: docrefine/rebrand.py
"""
Vector rebranding engine.

Composes a branded PDF *without rasterizing* the source, so the original vector
content and its text layer are preserved (output stays searchable). Each page is
extended with a header strip on top and a footer strip on bottom (the original
page is never overlapped or cropped), a watermark is laid over the content, and a
titled front/back cover are added.

Key efficiency detail: every branding image (header, footer, watermark, covers)
is embedded ONCE and shared across all pages, so a long document stays small
(well under the 50 MB delivery cap) instead of re-embedding art per page.
"""
import io
import json
import os
import re
import tempfile
from pathlib import Path
from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation
from pypdf.generic import NameObject
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from .config import SystemUtils

# --- Layout constants (tuned against the Canva reference) ---
TITLE_FRAC = 0.027      # cover title size as a fraction of the cover-art height
TITLE_TOP_FRAC = 0.11   # title vertical position from the top of the cover art
TITLE_RGB = (1, 1, 1)   # white
TITLE_UPPERCASE = True  # house style (see the Batch 1 hand-made covers)
TITLE_MAX_LINES = 2     # wrap to at most this many lines, then shrink to fit
TITLE_MIN_SCALE = 0.55  # never shrink below this fraction of the nominal size
TITLE_SIDE_MARGIN = 0.12  # keep this much of the page width clear either side
WATERMARK_WIDTH_FRAC = 0.60
WATERMARK_MAX_PX = 1400  # cap watermark resolution (it is embedded once, but stay lean)
COVER_MAX_PX = 2200      # cap cover art (~260dpi on Letter) — crisp on screen, smaller files

_FONT_NAME = "BrandTitle"
_font_ready = False


def _find_font():
    base = SystemUtils.get_resource_dir()
    candidates = [
        base / "docrefine" / "assets" / "fonts" / "Poppins-Bold.ttf",
        base / "assets" / "fonts" / "Poppins-Bold.ttf",
        Path(__file__).parent / "assets" / "fonts" / "Poppins-Bold.ttf",
    ]
    for p in candidates:
        if p.exists():
            return str(p)
    return None


def _ensure_font():
    global _font_ready
    if not _font_ready:
        fp = _find_font()
        if fp:
            pdfmetrics.registerFont(TTFont(_FONT_NAME, fp))
            _font_ready = True
    return _font_ready


def _orientation(w, h):
    return "landscape" if w > h else "portrait"


def page_rotation(page):
    """The page's /Rotate value normalised to 0/90/180/270."""
    try:
        return int(page.get("/Rotate") or 0) % 360
    except (TypeError, ValueError):
        return 0


def page_size(page):
    """A page's width and height **as a reader sees them**.

    Two things have to be normalised, and both have bitten us:

    * A PDF rectangle is defined by any two opposite corners, so a perfectly
      legal page box can be written top-down — [0, 792, 612, 0] — and pypdf then
      reports the height as -792. Taken literally that flipped the page to
      "landscape", inverted the cover scale and pushed the content off the page.
    * /Rotate turns the page when it is displayed. A 612x792 box with /Rotate 270
      is a *landscape* page to everyone who opens it. Reading the raw box called
      it portrait, so it was given portrait covers and strips and its content was
      squeezed into a portrait frame and cropped — on a quarter of the current
      batch.

    Returning display dimensions keeps orientation, cover sizing and the page
    strips all agreeing with what the document actually looks like.
    """
    box = page.mediabox
    w, h = abs(float(box.width)), abs(float(box.height))
    if page_rotation(page) in (90, 270):
        w, h = h, w
    return w, h


def _load_asset(path, max_px=None, opaque=False):
    """Load a branding image into an encoded-bytes 'spec' (safe to cache & share).

    opaque=True flattens onto the art's own background and stores JPEG (much smaller,
    for full-page covers that need no transparency). Otherwise keeps PNG alpha
    (for strips and the watermark, which overlay content). Returns only immutable
    data — reportlab ImageReaders are NOT thread-safe, so they're built per-call
    from this spec instead of being shared across worker threads.
    """
    im = Image.open(path).convert("RGBA")
    if max_px and max(im.size) > max_px:
        r = max_px / max(im.size)
        im = im.resize((round(im.width * r), round(im.height * r)), Image.LANCZOS)
    corner = im.convert("RGB").getpixel((2, 2))
    buf = io.BytesIO()
    if opaque:
        flat = Image.new("RGB", im.size, corner)
        flat.paste(im, (0, 0), im)
        flat.save(buf, "JPEG", quality=88)
    else:
        im.save(buf, "PNG")
    return {"bytes": buf.getvalue(), "px_size": im.size,
            "corner_rgb": tuple(v / 255 for v in corner)}


def _make_reader(spec):
    """Build a fresh reportlab ImageReader from a cached spec (one per call/thread)."""
    if spec is None:
        return None
    reader = ImageReader(io.BytesIO(spec["bytes"]))
    reader._px_size = spec["px_size"]
    reader._corner_rgb = spec["corner_rgb"]
    return reader


class BrandKit:
    """Resolves the branding assets for each orientation from a brand-kit folder.

    Expects orientation sub-folders (e.g. 'Portrait' / 'Landscape') each holding
    PNGs whose names contain: header, footer, watermark, cover, back.
    Assets are loaded once and reused (shared across every page/document).
    """
    _PORTRAIT_DIRS = ("portrait",)
    _LANDSCAPE_DIRS = ("landscape", "landscaape")  # tolerate the source folder's spelling
    REQUIRED = ("cover", "back")   # every other asset is optional; these are not

    def __init__(self, root):
        self.root = Path(root)
        self._paths = {}
        self._readers = {}
        for orient, names in (("portrait", self._PORTRAIT_DIRS), ("landscape", self._LANDSCAPE_DIRS)):
            d = self._match_dir(names)
            if d:
                self._paths[orient] = self._resolve_assets(d)
        self.brand = self._load_brand()

    # Wording for the per-page stamps lives in the kit, never in code: the task
    # briefs name a "tagline" and a "standard disclaimer" without giving the
    # words. A kit that doesn't supply them simply gets no stamp — which also
    # stops one brand's tagline appearing on another brand's documents.
    _BRAND_TEXT_KEYS = ("tagline", "disclaimer", "version_label", "last_updated",
                        "attribution", "stamp_ink", "stamp_bg")

    def _load_brand(self):
        """Brand wording for this kit, from an optional brand.json beside the assets.

        Keeps the kit genuinely swappable: the images and the brand *wording*
        travel together. Defaults preserve the Budget Mailboxes behaviour of
        earlier builds.
        """
        cfg = {"name": "Budget Mailboxes", "slug": "budget-mailboxes"}
        f = self.root / "brand.json"
        try:
            if f.is_file():
                data = json.loads(f.read_text(encoding="utf-8"))
                if isinstance(data, dict):
                    for k in ("name", "slug"):
                        if str(data.get(k, "")).strip():
                            cfg[k] = str(data[k]).strip()
                    cfg["slug"] = slugify(cfg["slug"]) or "budget-mailboxes"
                    for k in self._BRAND_TEXT_KEYS:
                        if str(data.get(k, "")).strip():
                            cfg[k] = str(data[k]).strip()
                    aliases = data.get("manufacturer_aliases")
                    if isinstance(aliases, dict):
                        cfg["manufacturer_aliases"] = aliases
        except Exception:
            pass
        return cfg

    @property
    def brand_name(self):
        return self.brand["name"]

    @property
    def brand_slug(self):
        return self.brand["slug"]

    def subtitle_for(self, manufacturer=""):
        """The attribution line under the cover title."""
        manufacturer = (manufacturer or "").strip()
        if manufacturer:
            return f"Manufactured by {manufacturer} | Sold by {self.brand_name}"
        return f"Sold by {self.brand_name}"

    def has_folder(self, orientation):
        """True if an orientation folder was found (regardless of what's inside)."""
        return orientation in self._paths

    def missing(self, orientation):
        """Required assets absent for an orientation (empty tuple when usable)."""
        paths = self._paths.get(orientation)
        if paths is None:
            return self.REQUIRED
        return tuple(k for k in self.REQUIRED if not paths.get(k))

    def _match_dir(self, names):
        for child in self.root.iterdir() if self.root.exists() else []:
            if child.is_dir() and child.name.lower() in names:
                return child
        return None

    @staticmethod
    def _resolve_assets(d):
        def find(*kw, exclude=()):
            for p in sorted(d.iterdir()):
                n = p.name.lower()
                if p.suffix.lower() == ".png" and all(k in n for k in kw) and not any(x in n for x in exclude):
                    return p
            return None
        return {
            "header": find("header"),
            "footer": find("footer"),
            "watermark": find("watermark"),
            "cover": find("cover", exclude=("back",)),
            "back": find("back"),
        }

    def has(self, orientation):
        """Usable for this orientation — the folder exists AND its required art resolved.

        Folder-presence alone used to be enough, so a kit missing cover.png passed
        validation and then failed on every single document.
        """
        return orientation in self._paths and not self.missing(orientation)

    def specs(self, orientation):
        """Cached asset specs (encoded bytes) for an orientation, loaded once.

        Safe to call for warm-up before a thread pool starts; the specs are
        immutable and shared, while live ImageReaders are built per-call below.
        """
        if orientation not in self._readers:
            paths = self._paths.get(orientation) or {}
            out = {}
            for key, path in paths.items():
                if path is None:
                    out[key] = None
                    continue
                cap = None
                opaque = False
                if key == "watermark":
                    cap = WATERMARK_MAX_PX
                elif key in ("cover", "back"):
                    cap = COVER_MAX_PX
                    opaque = True   # full-page art → JPEG, no transparency needed
                out[key] = _load_asset(path, max_px=cap, opaque=opaque)
            self._readers[orientation] = out
        return self._readers[orientation]

    # Back-compat name used by warm-up call sites.
    readers = specs

    def live_readers(self, orientation):
        """Fresh ImageReaders for a single rebrand_pdf call (never shared across threads)."""
        return {k: _make_reader(spec) for k, spec in self.specs(orientation).items()}


def _dominant_orientation(pages):
    counts = {"portrait": 0, "landscape": 0}
    for pg in pages:
        counts[_orientation(*page_size(pg))] += 1
    if counts["landscape"] > counts["portrait"]:
        return "landscape"
    if counts["portrait"] > counts["landscape"]:
        return "portrait"
    return _orientation(*page_size(pages[0]))


def _strip_height(reader, page_w):
    pw, ph = reader._px_size
    return page_w * ph / pw


def _wrap_to_width(words, font, size, max_w, max_lines):
    """Greedily wrap words into at most max_lines. Returns lines, or None if it won't fit."""
    lines, cur = [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if pdfmetrics.stringWidth(trial, font, size) <= max_w or not cur:
            cur = trial
        else:
            lines.append(cur)
            cur = w
            if len(lines) == max_lines:
                return None
    if cur:
        lines.append(cur)
    if len(lines) > max_lines:
        return None
    return lines if all(pdfmetrics.stringWidth(l, font, size) <= max_w for l in lines) else None


def fit_title(title, nominal_pt, max_w, font=None, max_lines=TITLE_MAX_LINES):
    """Lay a cover title out as (lines, font_size) that actually fits the page.

    Wraps to at most `max_lines`, shrinking the type until it fits rather than
    letting a long title run off both edges. Falls back to an ellipsis only if
    even the smallest permitted size can't contain it.
    """
    font = font or _FONT_NAME
    words = (title or "").split()
    if not words:
        return [], nominal_pt
    size = nominal_pt
    floor = nominal_pt * TITLE_MIN_SCALE
    while size >= floor:
        lines = _wrap_to_width(words, font, size, max_w, max_lines)
        if lines:
            return lines, size
        size *= 0.94
    # Still too long (e.g. one unbroken 400-character "word"): truncate to fit.
    size = floor
    cur = ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if pdfmetrics.stringWidth(trial + "…", font, size) > max_w:
            break
        cur = trial
    if not cur:
        # Not even the first word fits — cut it character by character.
        cur = words[0]
        while cur and pdfmetrics.stringWidth(cur + "…", font, size) > max_w:
            cur = cur[:-1]
    return [cur + "…"] if cur else [], size


def _draw_cover(c, cover_reader, page_w, page_h, title, subtitle=None):
    """Draw a cover page: art fitted (aspect preserved) on its own bg, optional title/subtitle."""
    c.setPageSize((page_w, page_h))
    cw, ch = cover_reader._px_size
    scale = min(page_w / cw, page_h / ch)
    dw, dh = cw * scale, ch * scale
    ox, oy = (page_w - dw) / 2, (page_h - dh) / 2
    # background = cover's own corner colour so any letterbox stays seamless
    c.setFillColorRGB(*cover_reader._corner_rgb)
    c.rect(0, 0, page_w, page_h, fill=1, stroke=0)
    c.drawImage(cover_reader, ox, oy, width=dw, height=dh, mask="auto")
    if title and _ensure_font():
        if TITLE_UPPERCASE:
            title = title.upper()
        nominal = TITLE_FRAC * dh
        max_w = dw * (1 - 2 * TITLE_SIDE_MARGIN)
        lines, tpt = fit_title(title, nominal, max_w)
        c.setFont(_FONT_NAME, tpt)
        c.setFillColorRGB(*TITLE_RGB)
        y = (oy + dh) - (TITLE_TOP_FRAC * dh) - tpt * 0.35
        for line in lines:
            c.drawCentredString(page_w / 2, y, line)
            y -= tpt * 1.18
        if subtitle:
            spt = tpt * 0.42
            c.setFont(_FONT_NAME, spt)
            c.drawCentredString(page_w / 2, y - tpt * 0.10, subtitle)
    c.showPage()


def _art_ink(readers):
    """The brand's own ink colour, taken from its footer (else cover) art.

    Reading the colour off the kit means the stamps are on-brand for whichever
    kit is loaded, with nothing hardcoded. If that colour is too pale to read on
    the stamp band's light background, fall back to near-black.
    """
    for key in ("footer", "cover", "header", "back"):
        r = readers.get(key)
        if r is not None:
            rgb = r._corner_rgb
            lum = 0.2126 * rgb[0] + 0.7152 * rgb[1] + 0.0722 * rgb[2]
            return rgb if lum <= 0.6 else (0.10, 0.10, 0.10)
    return None


def _draw_overlay(c, readers, page_w, page_h, stamps=None):
    """Draw one content-page overlay: header strip on top, footer on bottom, optional
    stamp band just above the footer, watermark over the content band.

    Returns (header_h, footer_h, stamp_h). The middle stays transparent, and the
    document's own content is lifted clear of both the footer and the stamps."""
    header, footer, wm = readers.get("header"), readers.get("footer"), readers.get("watermark")
    hh = _strip_height(header, page_w) if header else 0.0
    fh = _strip_height(footer, page_w) if footer else 0.0
    sh = stamps.height(page_w, _FONT_NAME) if (stamps and _ensure_font()) else 0.0
    total = page_h + hh + fh + sh
    c.setPageSize((page_w, total))
    if header:
        c.drawImage(header, 0, page_h + fh + sh, width=page_w, height=hh, mask="auto")
    if footer:
        c.drawImage(footer, 0, 0, width=page_w, height=fh, mask="auto")
    if sh:
        stamps.draw(c, page_w, fh, _FONT_NAME)
    if wm:
        ww = WATERMARK_WIDTH_FRAC * page_w
        wpx, hpx = wm._px_size
        wh = ww * hpx / wpx
        c.drawImage(wm, (page_w - ww) / 2, fh + sh + (page_h - wh) / 2,
                    width=ww, height=wh, mask="auto")
    c.showPage()
    return hh, fh, sh


def _carry_over_navigation(writer, reader, offset=1):
    """Bring the source document's bookmarks and form registration across.

    Widget and link annotations survive the page merge on their own, but two
    document-level things do not, and both are content the reader paid for:

    * the **outline** (bookmarks) — a long installation manual arrives with an
      empty navigation pane without it;
    * the **/AcroForm** catalog entry — without it the widget annotations that
      *did* come across are orphaned and the fields stop working.

    Bookmarks point at page numbers, which all shift by `offset` because we add a
    front cover. Never raises: losing a bookmark must not lose the document.
    """
    try:
        acro = reader.trailer.get("/Root", {}).get("/AcroForm")
        if acro is not None:
            writer._root_object[NameObject("/AcroForm")] = acro.clone(writer).indirect_reference
    except Exception:
        pass

    def add(items, parent=None):
        for item in items:
            if isinstance(item, list):
                add(item, parent)            # children of the item just added
                continue
            try:
                page_no = reader.get_destination_page_number(item)
                title = str(item.get("/Title", "") or "").strip()
            except Exception:
                continue
            if page_no is None or not title:
                continue
            target = page_no + offset
            if not (0 <= target < len(writer.pages)):
                continue
            try:
                parent = writer.add_outline_item(title, target, parent=parent)
            except Exception:
                continue

    try:
        outline = reader.outline
    except Exception:
        return
    if not outline:
        return
    try:
        # Top level only carries a parent of None; nested lists inherit the item
        # added immediately before them, which is how pypdf models the tree.
        last = None
        for item in outline:
            if isinstance(item, list):
                add(item, last)
            else:
                try:
                    pn = reader.get_destination_page_number(item)
                    t = str(item.get("/Title", "") or "").strip()
                except Exception:
                    continue
                if pn is None or not t or not (0 <= pn + offset < len(writer.pages)):
                    continue
                try:
                    last = writer.add_outline_item(t, pn + offset)
                except Exception:
                    continue
    except Exception:
        pass


def rebrand_pdf(input_pdf, output_pdf, kit, title, subtitle=None, author=None, stamps=None):
    """Rebrand a single PDF. Returns a small dict of stats. Raises on hard failure.

    `stamps` is an optional docrefine.stamps.Stamps carrying the per-page small
    print (attribution, version/updated, tagline, disclaimer). With none, the
    output is exactly what earlier versions produced.
    """
    input_pdf, output_pdf = Path(input_pdf), Path(output_pdf)
    author = author or kit.brand_name
    reader = PdfReader(str(input_pdf))
    pages = reader.pages
    doc_or = _dominant_orientation(pages)
    if not kit.has(doc_or):
        gap = ", ".join(kit.missing(doc_or)) or "assets"
        raise ValueError(f"Brand kit is missing its '{doc_or}' {gap} image(s).")

    # Build ImageReaders fresh for THIS call (never shared across threads). Within
    # this one canvas the same reader objects are reused across pages, so each image
    # still embeds exactly once per document.
    cover_set = kit.live_readers(doc_or)
    orient_sets = {doc_or: cover_set}
    DW, DH = page_size(pages[0])
    if stamps and stamps.ink is None:
        stamps.ink = _art_ink(cover_set)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf)
    _draw_cover(c, cover_set["cover"], DW, DH, title, subtitle=subtitle)
    strip_dims = []
    for pg in pages:
        pw, ph = page_size(pg)
        po = _orientation(pw, ph)
        if po not in orient_sets:
            orient_sets[po] = kit.live_readers(po) if kit.has(po) else cover_set
        strip_dims.append(_draw_overlay(c, orient_sets[po], pw, ph, stamps=stamps))
    _draw_cover(c, cover_set["back"], DW, DH, None)
    c.save()
    buf.seek(0)

    overlays = PdfReader(buf)
    writer = PdfWriter()
    writer.append(overlays)  # preserves shared image XObjects

    # writer pages: [front cover, overlay_0..N-1, back cover]
    for i, (hh, fh, sh) in enumerate(strip_dims):
        target = writer.pages[1 + i]
        src = pages[i]
        try:
            src.transfer_rotation_to_content()
        except Exception:
            pass
        # place the original page UNDER the overlay, lifted clear of the footer
        # strip AND the stamp band so nothing is ever drawn over the document
        target.merge_transformed_page(src, Transformation().translate(0, fh + sh), over=False)

    # The document's own bookmarks and form registration, which the page merge
    # does not carry. Pages shift by one because of the front cover.
    _carry_over_navigation(writer, reader, offset=1)

    writer.add_metadata({
        "/Author": author,
        "/Title": title or input_pdf.stem,
        "/Producer": "DocRefine Pro",
    })
    # Write to a sibling temp file and swap it in atomically. A run that is killed
    # or loses power mid-write must never leave a half-written PDF behind: the
    # resume check treats any non-empty file as finished, so a partial one would
    # be shipped as the deliverable.
    tmp = None
    try:
        fd, tmp = tempfile.mkstemp(dir=str(output_pdf.parent), suffix=".part")
        with os.fdopen(fd, "wb") as f:
            writer.write(f)
        os.replace(tmp, output_pdf)
        tmp = None
    finally:
        if tmp and os.path.exists(tmp):
            try: os.remove(tmp)
            except OSError: pass

    size_mb = output_pdf.stat().st_size / 1e6
    return {"orientation": doc_or, "source_pages": len(pages),
            "output_pages": len(pages) + 2, "size_mb": round(size_mb, 2)}


# --- Naming helpers ---

# Cover titles the customer actually reads. The hand-made Batch 1 covers set the
# house style: short, plain, no model numbers — "PRODUCT CARE & CLEANING",
# "FIVE YEAR PRODUCT WARRANTY", "INSTALLATION MANUAL". A document's asset type is
# almost always the right title on its own, so that is what we render.
ASSET_TYPE_TITLES = {
    "installation-guide": "Installation Manual",
    "installation-manual": "Installation Manual",
    "installation-instructions": "Installation Instructions",
    "assembly-instructions": "Assembly Instructions",
    "mounting-instructions": "Mounting Instructions",
    "user-guide": "User Guide",
    "manual": "Product Manual",
    "maintenance-manual": "Maintenance Manual",
    "spec-sheet": "Specification Sheet",
    "specification-sheet": "Specification Sheet",
    "specifications": "Specification Sheet",
    "cut-sheet": "Specification Sheet",
    "product-cutsheet": "Specification Sheet",
    "drawing": "Technical Drawing",
    "cad-drawing": "Technical Drawing",
    "technical-drawing": "Technical Drawing",
    "warranty": "Product Warranty",
    "certification": "Product Certification",
    "catalog": "Product Catalog",
    "brochure": "Product Brochure",
    "care-cleaning": "Product Care & Cleaning",
    "document": "Product Documentation",
}

DEFAULT_TITLE = "Product Documentation"

# Tokens carrying a digit are model/part numbers (4C11D, 3635RL, 1570-12) and must
# keep their own casing — .title() would render them "4C11D" -> "4C11d".
_HAS_DIGIT = re.compile(r"\d")

MAX_DELIVERY_NAME = 60   # hard cap from the delivery brief


def title_from_filename(stem):
    """A tidy, human-readable cover title derived from a filename."""
    s = re.sub(r"[_\-]+", " ", stem)
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return DEFAULT_TITLE
    return " ".join(w if _HAS_DIGIT.search(w) else w.title() for w in s.split())


def title_for(asset_type, fallback_stem=None, doc_type=None):
    """The cover title for a document: plain, short, no model numbers.

    Prefers the canonical name for the asset type (what the customer needs to
    know the document *is*), then the model's doc_type, then the filename.
    """
    key = slugify(asset_type or "")
    if key in ASSET_TYPE_TITLES:
        return ASSET_TYPE_TITLES[key]
    dkey = slugify(doc_type or "")
    if dkey in ASSET_TYPE_TITLES:
        return ASSET_TYPE_TITLES[dkey]
    if key and key != "document":
        return key.replace("-", " ").title()
    if fallback_stem:
        return title_from_filename(fallback_stem)
    return DEFAULT_TITLE


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "-", (s or "").lower()).strip("-")


def _clip_words(slug, limit):
    """Shorten a slug to `limit` characters on a word boundary.

    Cutting mid-word produces names like 'wall-mount-mailbo'; dropping the whole
    word reads as a deliberate abbreviation instead.
    """
    if len(slug) <= limit:
        return slug
    cut = slug[:limit]
    if "-" in cut:
        cut = cut[:cut.rindex("-")]
    return cut.strip("-")


def output_filename(stem, brand_slug="budget-mailboxes", max_len=60):
    """Slugged delivery filename: <cleaned-name>-<brand>.pdf, lowercase, <= max_len."""
    tail = f"-{brand_slug}.pdf"
    keep = max(1, max_len - len(tail))
    base = slugify(stem)[:keep].strip("-") or "document"
    return f"{base}{tail}"


def output_filename_from_fields(product, asset_type, brand_slug="budget-mailboxes", max_len=60):
    """Brief's delivery pattern: <product>-<asset-type>-budget-mailboxes.pdf, <= max_len.

    The product is truncated in preference to the asset type, so a long product
    name never squeezes out the part that says what the document *is*.
    """
    p, a = slugify(product), slugify(asset_type)
    tail = f"-{brand_slug}.pdf"
    keep = max(1, max_len - len(tail))
    if p and a and len(f"{p}-{a}") > keep:
        p = _clip_words(p, max(1, keep - len(a) - 1))
    core = "-".join(x for x in (p, a) if x) or "document"
    return f"{_clip_words(core, keep) or 'document'}{tail}"


def numbered_filename(name, n, max_len=MAX_DELIVERY_NAME):
    """`name` with a `-n` disambiguator, still inside the delivery length cap.

    Appending the suffix naively pushed an already-capped name to 61 characters.
    """
    p = Path(name)
    stem, ext = p.stem, p.suffix
    suffix = f"-{n}"
    over = len(stem) + len(suffix) + len(ext) - max_len
    if over > 0:
        stem = stem[:max(1, len(stem) - over)].rstrip("-")
    return f"{stem}{suffix}{ext}"


def delivery_filename(stem, product="", asset_type="", brand_slug="budget-mailboxes", max_len=60):
    """The delivery filename for one document.

    Falls back to the source filename when no product was identified. The model
    leaves `product` blank on a large share of documents, and without this every
    unidentified installation guide in a batch collapses onto
    `installation-guide-budget-mailboxes.pdf` and gets a meaningless numbered
    suffix — the source name is where the model or part number actually lives.
    """
    product = (product or "").strip() or (stem or "")
    return output_filename_from_fields(product, asset_type,
                                       brand_slug=brand_slug, max_len=max_len)
