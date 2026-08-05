# SAVE AS: docrefine/stamps.py
"""
Per-page stamps — the small print the rebranding SOP asks for on every page.

A branded page already carries the kit's own art: a navy header bar, and a navy
footer bar holding the URL and phone number. What the SOP asks for *beyond* the
art is text — a tagline, the manufacturer attribution ("Manufactured by [X] |
Sold by [brand]"), a version and last-updated line, and a standard disclaimer.

Two rules shape this module:

* **Nothing is ever drawn over the document.** The stamps get their own light
  band, added between the document content and the footer art, exactly as the
  header and footer strips extend the page. The original page is never
  overlapped or cropped.
* **The wording is data, not code.** Both task briefs say "tagline" and
  "standard disclaimer" without giving the words — those live in the SOP. So
  every string is read from the brand kit's `brand.json`, and anything absent is
  simply not drawn. No wording is invented here, and a kit for a different brand
  cannot inherit another brand's tagline.

The manufacturer name is cleaned before it is ever printed: on the real Budget
Mailboxes batch that column contained websites ("Florencemailboxes.com"), the
seller ("WebLife Stores LLC") and the brand itself — which read as "Manufactured
by Budget Mailboxes | Sold by Budget Mailboxes". A value we cannot vouch for
means the attribution is left off that document rather than printed wrong.
"""
import re

from reportlab.pdfbase import pdfmetrics

# --- Layout, all proportional to page width so it holds on a 44x34in drawing ---
STAMP_PT_FRAC = 0.0108     # base text size as a fraction of page width (~6.6pt on Letter)
STAMP_MIN_PT = 5.0         # never shrink past legibility on a small page
STAMP_SIDE_FRAC = 0.045    # clear margin either side
STAMP_PAD_FRAC = 0.90      # vertical padding above/below, in multiples of text size
                           # (measured on a real page: less than this and the last
                           # line of the disclaimer crowds the footer bar)
STAMP_LEADING = 1.34       # line spacing
TAGLINE_SCALE = 1.18       # the tagline reads as a line of brand, not small print
DISCLAIMER_SCALE = 0.92
DISCLAIMER_MAX_LINES = 4   # a disclaimer longer than this is truncated, not allowed to
                           # swallow the page

DEFAULT_INK = (0.12, 0.05, 0.40)   # only used if neither the kit nor its art supplies one
DEFAULT_ATTRIBUTION = "Manufactured by {manufacturer} | Sold by {brand}"
DEFAULT_VERSION_LABEL = "Version 1.0"
UPDATED_PREFIX = "Last Updated"
SEPARATOR = "  ·  "

# Names that identify the seller or a storefront rather than who made the
# product. Any of these as a "manufacturer" means the value is unusable.
SELLER_HINTS = ("weblife", "budget mailboxes", "budgetmailboxes", "mailboxworks",
                "mailbox works")
_DOMAIN = re.compile(r"\.(com|net|org|co|us|biz|info)\b", re.I)
_WS = re.compile(r"\s+")


def _hex_rgb(value, fallback):
    """A '#1f2a5a'-style colour as a 0-1 RGB triple, or the fallback."""
    s = str(value or "").strip().lstrip("#")
    if len(s) == 6:
        try:
            return tuple(int(s[i:i + 2], 16) / 255 for i in (0, 2, 4))
        except ValueError:
            pass
    return fallback


def clean_manufacturer(value, brand_name="", aliases=None):
    """The manufacturer name as it can safely be printed, or "" if it cannot.

    Returning "" is a deliberate outcome, not a failure: an attribution line is
    a factual claim about who made the product, and a wrong one is worse than an
    absent one. Callers count the omissions so the gap is visible.
    """
    raw = _WS.sub(" ", str(value or "").strip())
    if not raw:
        return ""
    key = raw.lower()
    for k, v in (aliases or {}).items():           # kit-supplied corrections win
        if str(k).strip().lower() == key:
            return _WS.sub(" ", str(v).strip())
    if _DOMAIN.search(raw):                        # a website is not a manufacturer
        return ""
    brand = _WS.sub(" ", str(brand_name or "").strip()).lower()
    if brand and (key == brand or brand in key):   # "Manufactured by <us>" says nothing
        return ""
    for hint in SELLER_HINTS:
        if hint in key:
            return ""
    if len(raw) < 3:
        return ""
    return raw


def wrap(text, font, size, max_w):
    """Greedily wrap text to max_w. Always returns at least one line for non-empty text."""
    words = str(text or "").split()
    if not words:
        return []
    lines, cur = [], ""
    for w in words:
        trial = f"{cur} {w}".strip()
        if not cur or pdfmetrics.stringWidth(trial, font, size) <= max_w:
            cur = trial
        else:
            lines.append(cur)
            cur = w
    if cur:
        lines.append(cur)
    return lines


def _fill(template, manufacturer, brand):
    """Substitute the attribution placeholders without str.format.

    The template comes from a user-edited JSON file, where a stray brace would
    otherwise raise instead of just rendering.
    """
    out = str(template or "")
    out = out.replace("{manufacturer}", manufacturer).replace("{brand}", brand)
    return _WS.sub(" ", out).strip()


class Stamps:
    """The stamp text resolved for one document, ready to measure and draw.

    Build with `for_document`; a set with nothing to say is falsy, and callers
    skip the band entirely (so output is byte-for-byte as before when every
    stamp is switched off).
    """

    def __init__(self, attribution="", version_line="", tagline="", disclaimer="",
                 ink=None, bg=(1, 1, 1)):
        self.attribution = attribution
        self.version_line = version_line
        self.tagline = tagline
        self.disclaimer = disclaimer
        self.ink = ink
        self.bg = bg
        self._cache = {}

    def __bool__(self):
        return bool(self.attribution or self.version_line or self.tagline or self.disclaimer)

    # -- layout ------------------------------------------------------------
    def layout(self, page_w, font):
        """Rows to draw and the band height, for a page of this width.

        Returns (rows, height). Each row is (kind, size, payload):
        'pair' -> (left, right), 'center' -> text, 'left' -> text.
        """
        key = (round(page_w, 3), font)
        if key in self._cache:
            return self._cache[key]

        base = max(STAMP_MIN_PT, STAMP_PT_FRAC * page_w)
        margin = STAMP_SIDE_FRAC * page_w
        text_w = max(1.0, page_w - 2 * margin)
        rows = []

        if self.attribution or self.version_line:
            # One line carries both: who made it (left) and how current it is (right).
            # If they cannot share a line, the version drops to its own row so the
            # attribution is never clipped.
            left, right = self.attribution, self.version_line
            if left and right:
                together = (pdfmetrics.stringWidth(left, font, base)
                            + pdfmetrics.stringWidth(right, font, base)
                            + pdfmetrics.stringWidth(SEPARATOR, font, base))
                if together > text_w:
                    rows.append(("left", base, left))
                    rows.append(("left", base, right))
                else:
                    rows.append(("pair", base, (left, right)))
            else:
                rows.append(("left", base, left or right))

        if self.tagline:
            size = base * TAGLINE_SCALE
            for line in wrap(self.tagline, font, size, text_w):
                rows.append(("center", size, line))

        if self.disclaimer:
            size = base * DISCLAIMER_SCALE
            lines = wrap(self.disclaimer, font, size, text_w)
            if len(lines) > DISCLAIMER_MAX_LINES:
                lines = lines[:DISCLAIMER_MAX_LINES]
                lines[-1] = lines[-1].rstrip(" ,;") + "…"
            for line in lines:
                rows.append(("left", size, line))

        if not rows:
            out = ([], 0.0)
        else:
            body = sum(size * STAMP_LEADING for _, size, _ in rows)
            out = (rows, body + 2 * STAMP_PAD_FRAC * base)
        self._cache[key] = out
        return out

    def height(self, page_w, font):
        return self.layout(page_w, font)[1]

    # -- drawing -----------------------------------------------------------
    def draw(self, c, page_w, y0, font):
        """Draw the band with its baseline at y0 (its bottom edge). Returns the height."""
        rows, h = self.layout(page_w, font)
        if not rows:
            return 0.0
        base = max(STAMP_MIN_PT, STAMP_PT_FRAC * page_w)
        margin = STAMP_SIDE_FRAC * page_w
        c.setFillColorRGB(*self.bg)
        c.rect(0, y0, page_w, h, fill=1, stroke=0)
        c.setFillColorRGB(*(self.ink or DEFAULT_INK))
        y = y0 + h - STAMP_PAD_FRAC * base
        for kind, size, payload in rows:
            y -= size * STAMP_LEADING
            baseline = y + size * 0.28
            c.setFont(font, size)
            if kind == "pair":
                left, right = payload
                c.drawString(margin, baseline, left)
                c.drawRightString(page_w - margin, baseline, right)
            elif kind == "center":
                c.drawCentredString(page_w / 2, baseline, payload)
            else:
                c.drawString(margin, baseline, payload)
        return h


def updated_line(brand, today=None):
    """"Version 1.0 · Last Updated August 2026" from the kit's wording plus the run date.

    The date is taken once per run and passed in, so every document in a batch
    carries the same "Last Updated" month even if the run crosses midnight.
    """
    label = str(brand.get("version_label") or DEFAULT_VERSION_LABEL).strip()
    updated = str(brand.get("last_updated") or "").strip()
    if not updated and today is not None:
        updated = f"{UPDATED_PREFIX} {today.strftime('%B %Y')}"
    return SEPARATOR.join(p for p in (label, updated) if p)


def for_document(brand, manufacturer="", *, attribution=False, version=False,
                 tagline=False, disclaimer=False, today=None, ink=None):
    """Resolve the stamps for one document from the kit's wording and the toggles.

    A toggle that is on but has no wording in the kit yields nothing — callers
    check `missing_wording` once per run and warn, rather than failing per file.

    `ink` left as None means "take the colour from the kit's own art", which the
    renderer fills in; a kit can override it outright with `stamp_ink`.
    """
    brand = brand or {}
    brand_name = str(brand.get("name") or "").strip()
    s = Stamps(
        ink=_hex_rgb(brand.get("stamp_ink"), ink),
        bg=_hex_rgb(brand.get("stamp_bg"), (1, 1, 1)),
    )
    if attribution:
        who = clean_manufacturer(manufacturer, brand_name,
                                 brand.get("manufacturer_aliases"))
        if who:
            s.attribution = _fill(brand.get("attribution") or DEFAULT_ATTRIBUTION,
                                  who, brand_name)
    if version:
        s.version_line = updated_line(brand, today)
    if tagline:
        s.tagline = str(brand.get("tagline") or "").strip()
    if disclaimer:
        s.disclaimer = str(brand.get("disclaimer") or "").strip()
    return s


def missing_wording(brand, *, tagline=False, disclaimer=False):
    """Stamps that are switched on but have no text in the kit, for one clear warning."""
    brand = brand or {}
    gaps = []
    if tagline and not str(brand.get("tagline") or "").strip():
        gaps.append("tagline")
    if disclaimer and not str(brand.get("disclaimer") or "").strip():
        gaps.append("disclaimer")
    return gaps
