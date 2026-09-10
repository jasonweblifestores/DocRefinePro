from pypdf import PdfReader

FILES = {
 "SOURCE": r"C:\Users\WORK\Documents\DUMP\Imperial-Street-Sign-Brochure.pdf",
 "SHIPPED (v144 run)": r"C:\Users\WORK\Documents\Batch 4\_unique-to-rebrand_rebranded\imperial-street-specification-sheet-budget-mailboxes.pdf",
 "FIXED (v146)": r"C:\Users\WORK\Documents\DUMP_rebranded\imperial-street-signs-spec-sheet-budget-mailboxes.pdf",
}
for label, path in FILES.items():
    rd = PdfReader(path)
    print(f"--- {label}  ({len(rd.pages)} pages)")
    print("    Author:", (rd.metadata or {}).get("/Author"))
    for i, p in enumerate(rd.pages):
        mb = p.mediabox
        w = float(mb.right) - float(mb.left); h = float(mb.top) - float(mb.bottom)
        txt = p.extract_text() or ""
        print(f"    p{i+1}: box=[{float(mb.left):.0f},{float(mb.bottom):.0f},{float(mb.right):.0f},{float(mb.top):.0f}] "
              f"w={w:.0f} h={h:.0f} rot={p.get('/Rotate')} chars={len(txt)}")
        try:
            raw = p.get_contents().get_data().decode("latin-1")
        except Exception:
            raw = ""
        cms = [ln for ln in raw.splitlines() if ln.strip().endswith(" cm")][:4]
        for c in cms:
            print("        cm:", c.strip())
