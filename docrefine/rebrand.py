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
import re
from pathlib import Path
from PIL import Image
from pypdf import PdfReader, PdfWriter, Transformation
from reportlab.pdfgen import canvas as rl_canvas
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

from .config import SystemUtils

# --- Layout constants (tuned against the Canva reference) ---
TITLE_FRAC = 0.027      # cover title size as a fraction of the cover-art height
TITLE_TOP_FRAC = 0.11   # title vertical position from the top of the cover art
TITLE_RGB = (1, 1, 1)   # white
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

    def __init__(self, root):
        self.root = Path(root)
        self._paths = {}
        self._readers = {}
        for orient, names in (("portrait", self._PORTRAIT_DIRS), ("landscape", self._LANDSCAPE_DIRS)):
            d = self._match_dir(names)
            if d:
                self._paths[orient] = self._resolve_assets(d)

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
        return orientation in self._paths

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
        counts[_orientation(float(pg.mediabox.width), float(pg.mediabox.height))] += 1
    if counts["landscape"] > counts["portrait"]:
        return "landscape"
    if counts["portrait"] > counts["landscape"]:
        return "portrait"
    return _orientation(float(pages[0].mediabox.width), float(pages[0].mediabox.height))


def _strip_height(reader, page_w):
    pw, ph = reader._px_size
    return page_w * ph / pw


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
        tpt = TITLE_FRAC * dh
        c.setFont(_FONT_NAME, tpt)
        c.setFillColorRGB(*TITLE_RGB)
        y = (oy + dh) - (TITLE_TOP_FRAC * dh) - tpt * 0.35
        c.drawCentredString(page_w / 2, y, title)
        if subtitle:
            spt = tpt * 0.42
            c.setFont(_FONT_NAME, spt)
            c.drawCentredString(page_w / 2, y - tpt * 1.05, subtitle)
    c.showPage()


def _draw_overlay(c, readers, page_w, page_h):
    """Draw one content-page overlay: header strip on top, footer on bottom, watermark
    over the content band. Returns (header_h, footer_h). Middle stays transparent."""
    header, footer, wm = readers.get("header"), readers.get("footer"), readers.get("watermark")
    hh = _strip_height(header, page_w) if header else 0.0
    fh = _strip_height(footer, page_w) if footer else 0.0
    total = page_h + hh + fh
    c.setPageSize((page_w, total))
    if header:
        c.drawImage(header, 0, page_h + fh, width=page_w, height=hh, mask="auto")
    if footer:
        c.drawImage(footer, 0, 0, width=page_w, height=fh, mask="auto")
    if wm:
        ww = WATERMARK_WIDTH_FRAC * page_w
        wpx, hpx = wm._px_size
        wh = ww * hpx / wpx
        c.drawImage(wm, (page_w - ww) / 2, fh + (page_h - wh) / 2, width=ww, height=wh, mask="auto")
    c.showPage()
    return hh, fh


def rebrand_pdf(input_pdf, output_pdf, kit, title, subtitle=None, author="Budget Mailboxes"):
    """Rebrand a single PDF. Returns a small dict of stats. Raises on hard failure."""
    input_pdf, output_pdf = Path(input_pdf), Path(output_pdf)
    reader = PdfReader(str(input_pdf))
    pages = reader.pages
    doc_or = _dominant_orientation(pages)
    if not kit.has(doc_or):
        raise ValueError(f"Brand kit has no '{doc_or}' asset set.")

    # Build ImageReaders fresh for THIS call (never shared across threads). Within
    # this one canvas the same reader objects are reused across pages, so each image
    # still embeds exactly once per document.
    cover_set = kit.live_readers(doc_or)
    orient_sets = {doc_or: cover_set}
    DW = float(pages[0].mediabox.width)
    DH = float(pages[0].mediabox.height)

    buf = io.BytesIO()
    c = rl_canvas.Canvas(buf)
    _draw_cover(c, cover_set["cover"], DW, DH, title, subtitle=subtitle)
    strip_dims = []
    for pg in pages:
        pw, ph = float(pg.mediabox.width), float(pg.mediabox.height)
        po = _orientation(pw, ph)
        if po not in orient_sets:
            orient_sets[po] = kit.live_readers(po) if kit.has(po) else cover_set
        hh, fh = _draw_overlay(c, orient_sets[po], pw, ph)
        strip_dims.append((hh, fh))
    _draw_cover(c, cover_set["back"], DW, DH, None)
    c.save()
    buf.seek(0)

    overlays = PdfReader(buf)
    writer = PdfWriter()
    writer.append(overlays)  # preserves shared image XObjects

    # writer pages: [front cover, overlay_0..N-1, back cover]
    for i, (hh, fh) in enumerate(strip_dims):
        target = writer.pages[1 + i]
        src = pages[i]
        try:
            src.transfer_rotation_to_content()
        except Exception:
            pass
        # place the original page UNDER the overlay, lifted above the footer strip
        target.merge_transformed_page(src, Transformation().translate(0, fh), over=False)

    writer.add_metadata({
        "/Author": author,
        "/Title": title or input_pdf.stem,
        "/Producer": "DocRefine Pro",
    })
    with open(output_pdf, "wb") as f:
        writer.write(f)

    size_mb = output_pdf.stat().st_size / 1e6
    return {"orientation": doc_or, "source_pages": len(pages),
            "output_pages": len(pages) + 2, "size_mb": round(size_mb, 2)}


# --- Naming helpers (placeholder logic for Phase 1; the LLM refines these in Phase 2) ---

def title_from_filename(stem):
    """A tidy, human-readable cover title derived from a filename."""
    s = re.sub(r"[_\-]+", " ", stem)
    s = re.sub(r"\s+", " ", s).strip()
    return s.title() if s else "Product Documentation"


def slugify(s):
    return re.sub(r"[^a-z0-9]+", "-", (s or "").lower()).strip("-")


def output_filename(stem, brand_slug="budget-mailboxes", max_len=60):
    """Slugged delivery filename: <cleaned-name>-<brand>.pdf, lowercase, <= max_len."""
    tail = f"-{brand_slug}.pdf"
    keep = max(1, max_len - len(tail))
    base = slugify(stem)[:keep].strip("-") or "document"
    return f"{base}{tail}"


def output_filename_from_fields(product, asset_type, brand_slug="budget-mailboxes", max_len=60):
    """Brief's delivery pattern: <product>-<asset-type>-budget-mailboxes.pdf, <= max_len."""
    core = "-".join(x for x in (slugify(product), slugify(asset_type)) if x) or "document"
    tail = f"-{brand_slug}.pdf"
    keep = max(1, max_len - len(tail))
    return f"{core[:keep].strip('-') or 'document'}{tail}"
