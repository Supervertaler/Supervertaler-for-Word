"""Proof of concept: anchor, translate as tracked changes, save, reopen, verify.

Run:  python spike/roundtrip_demo.py [out_dir]
Uses a hidden Word instance; nothing touches your open documents.
"""
from __future__ import annotations

import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w                     # noqa: E402
from svword.project_xml import ProjectRecord, SegmentRecord, load, save  # noqa: E402

PARAGRAPHS = [
    "A mashup is a Web application that combines data from one or more sources. "
    "The term implies easy, fast integration.",
    "An example of a mashup is the use of cartographic data from a mapping program. "
    "This creates a new and distinct Web service.",
]

# Stub translator – any callable (str) -> str will do here; the real one is the
# Workbench/Otto pipeline.
STUB = {
    "A mashup is a Web application that combines data from one or more sources.":
        "Een mashup is een webtoepassing die gegevens uit één of meer bronnen combineert.",
    "The term implies easy, fast integration.":
        "De term impliceert een eenvoudige, snelle integratie.",
    "An example of a mashup is the use of cartographic data from a mapping program.":
        "Een voorbeeld van een mashup is het gebruik van cartografische gegevens uit een kaarttoepassing.",
    "This creates a new and distinct Web service.":
        "Dit creëert een nieuwe en onderscheidende webdienst.",
}


def check(label, got, want):
    ok = got == want
    print(("  ok   " if ok else "  FAIL ") + label)
    if not ok:
        print("       got : %r" % got)
        print("       want: %r" % want)
    return ok


UNIT = "paragraph" if "--paragraph" in sys.argv else "sentence"


def main(out_dir: str) -> int:
    os.makedirs(out_dir, exist_ok=True)
    out_dir = os.path.abspath(out_dir)
    app = w.word_app(visible=False)
    failures = 0
    doc = app.Documents.Add()
    try:
        # ---- build the source document -------------------------------------
        for p in PARAGRAPHS:
            doc.Content.InsertAfter(p + "\r")
        original_text = doc.Content.Text

        # ---- anchor every sentence ----------------------------------------
        record = ProjectRecord("en", "nl")
        ccs = w.anchor_document(doc, UNIT)
        for cc in ccs:
            record.segments.append(SegmentRecord(w.segment_id(cc), cc.Range.Text))
        print("anchored %d segments" % len(ccs))

        # ---- the Felix moment: sentence under the cursor ------------------
        doc.Range(5, 5).Select()
        got = w.sentence_at_cursor(doc).Text.strip()
        failures += not check("sentence_at_cursor", record.segments[0].source.startswith(got), True)

        # ---- translate as tracked changes ---------------------------------
        def stub(src):
            return " ".join(STUB[k] for k in STUB if k in src) or src
        for cc, seg in zip(ccs, record.segments):
            seg.target = stub(seg.source)
            seg.origin = "ai"
            w.set_target(app, doc, cc, seg.target)
            w.annotate(doc, cc, "AI draft · no TM match")
        save(doc, record)
        failures += not check("user name restored", app.UserName != w.REVISION_AUTHOR, True)

        path = os.path.join(out_dir, "roundtrip.docx")
        doc.SaveAs2(path, 16)                 # 16 = wdFormatXMLDocument
    except BaseException:
        doc.Close(0)
        app.Quit()
        raise
    else:
        doc.Close(0)                          # 0 = wdDoNotSaveChanges

    # ---- reopen and verify every claim from the document alone -------------
    doc = app.Documents.Open(path)
    try:
        rec = load(doc)
        n = len(STUB) if UNIT == "sentence" else len(PARAGRAPHS)
        failures += not check("xml part reloaded", rec is not None and len(rec.segments), n)
        ccs = list(w.anchors(doc))
        failures += not check("anchors survived save", len(ccs), n)
        by_id = {s.id: s for s in rec.segments}
        for cc in ccs:
            seg = by_id[w.segment_id(cc)]
            failures += not check("source of %s" % seg.id, w.source_of(cc), seg.source)
            failures += not check("target of %s" % seg.id, w.target_of(cc), seg.target)
        failures += not check("revision author", doc.Revisions(1).Author, w.REVISION_AUTHOR)
        failures += not check("comments", doc.Comments.Count, n)

        # Reject All -> source document back; Accept All -> target document.
        doc.Revisions.RejectAll()
        failures += not check("Reject All == original", doc.Content.Text, original_text)
        doc.Undo()
        doc.Revisions.AcceptAll()
        target_text = doc.Content.Text
        ok = all(t in target_text for t in STUB.values()) and not any(s in target_text for s in STUB)
        failures += not check("Accept All == target only", ok, True)
        failures += not check("anchors survive Accept All", sum(1 for _ in w.anchors(doc)), n)
    finally:
        doc.Close(0)
        app.Quit()

    print("\n%s  (%s)" % ("ALL PASSED" if not failures else "%d FAILED" % failures, path))
    return 1 if failures else 0


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    out = args[0] if args else tempfile.mkdtemp(prefix="svword_")
    sys.exit(main(out))
