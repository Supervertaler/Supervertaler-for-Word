"""Anchor a real source document and fill in targets from a list of pairs.

    python spike/build_example.py <source.docx> <out.docx> [--pairs pairs.json] [--langs nl en] [--open]

pairs.json is [[source, target], ...]. Sentences whose source matches a pair
get the target written as a tracked change (origin "tm"); the rest stay
untranslated. Without --pairs every sentence is anchored untranslated.
--open then shows the result in a visible Word of its own.
"""
from __future__ import annotations
import json, os, re, shutil, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w
from svword.project_xml import ProjectRecord, SegmentRecord, save

norm = lambda s: re.sub(r"\s+", " ", s).strip().lower()


def main():
    argv = sys.argv[1:]
    def opt(name, n):
        if name in argv:
            i = argv.index(name); vals = argv[i + 1:i + 1 + n]; del argv[i:i + 1 + n]; return vals
        return None
    pairs_json = (opt("--pairs", 1) or [None])[0]
    langs = opt("--langs", 2) or ["en", "nl"]
    args = [a for a in argv if not a.startswith("--")]
    src_docx, out_docx = args
    pairs = {norm(s): t for s, t in json.load(open(pairs_json, encoding="utf-8")) if t.strip()} if pairs_json else {}
    os.makedirs(os.path.dirname(os.path.abspath(out_docx)), exist_ok=True)
    shutil.copy(src_docx, out_docx)

    app = w.word_app(visible=False)
    try:
        doc = app.Documents.Open(os.path.abspath(out_docx))
        record = ProjectRecord(langs[0], langs[1])
        ccs = w.anchor_document(doc, "sentence")
        hit = 0
        for cc in ccs:
            seg = SegmentRecord(w.segment_id(cc), cc.Range.Text)
            t = pairs.get(norm(seg.source))
            if t:
                seg.target, seg.origin, seg.status = t, "tm", "draft"
                w.set_target(app, doc, cc, t, "draft")
                hit += 1
            record.segments.append(seg)
        save(doc, record)
        doc.Save()
        doc.Close(0)
        print(f"anchored {len(ccs)} sentences, {hit} with a target, -> {out_docx}")
    finally:
        app.Quit()

    if "--open" in sys.argv:
        import win32com.client as com
        vis = com.DispatchEx("Word.Application")
        vis.Visible = True
        d = vis.Documents.Open(os.path.abspath(out_docx))
        d.TrackRevisions = True
        d.ActiveWindow.View.RevisionsFilter.Markup = 0
        vis.Activate()


if __name__ == "__main__":
    main()
