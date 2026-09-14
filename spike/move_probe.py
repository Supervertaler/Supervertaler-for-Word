"""What survives when the translator rearranges text on the page?

Each scenario runs on a fresh copy of roundtrip.docx in a hidden Word and
reports per anchor: still present, in what document order, source and target.
Findings are summarised in CLAUDE.md. Locking anchors was tried and rejected:
it blocks cutting any paragraph that contains one.

  A  cut a sentence's anchor and paste it at the end of the other paragraph, tracking ON
  B  the same with tracking OFF
  C  move the whole second paragraph above the first, tracking ON
  D  merge sentences 1 and 2 by deleting across the anchor boundary, tracking ON
  E  delete a whole sentence (its anchor range), locked, tracking OFF
  F  delete a whole sentence, UNLOCKED, tracking OFF   (the accident case)
"""
from __future__ import annotations
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w
from svword.project_xml import load

SRC = os.path.join(os.path.dirname(__file__), "out", "roundtrip.docx")
WD_MOVED_FROM, WD_MOVED_TO = 14, 15


def state(doc, ids):
    """[(id, source, target)] in document order, plus which ids are missing."""
    rec = {s.id: s for s in load(doc).segments}
    present = []
    for cc in w.anchors(doc):                      # document order
        sid = w.segment_id(cc)
        present.append((sid, w.source_of(cc), w.target_of(cc, rec[sid].source)))
    missing = [i for i in ids if i not in {p[0] for p in present}]
    return present, missing


def report(name, doc, ids, expect_order=None):
    present, missing = state(doc, ids)
    order = [p[0] for p in present]
    print(f"\n[{name}]  anchors {len(present)}/{len(ids)}  missing={missing}")
    if expect_order:
        print("  order as expected:", order == expect_order)
    rec = {s.id: s for s in load(doc).segments}
    for sid, src, tgt in present:
        where = "in doc     " if src == rec[sid].source else "record only"
        print(f"  {sid} source {where} | tgt={tgt!r}")
    revs = {}
    for r in doc.Revisions:
        revs[r.Type] = revs.get(r.Type, 0) + 1
    print("  revision types:", revs, "(14/15 = moved from/to)")


def fresh(app, name, lock=True):
    path = os.path.join(tempfile.mkdtemp(prefix="svword_"), name + ".docx")
    shutil.copy(SRC, path)
    doc = app.Documents.Open(path)
    for cc in w.anchors(doc):
        cc.LockContentControl = lock
    ids = [w.segment_id(cc) for cc in w.anchors(doc)]
    return doc, ids


def cc_by_index(doc, i):
    return list(w.anchors(doc))[i]


def scenario_cut_paste(app, name, tracking, whole_control):
    doc, ids = fresh(app, name, lock=False)
    try:
        doc.TrackRevisions = tracking
        src = cc_by_index(doc, 1)                 # sentence 2, paragraph 1
        if whole_control:
            src.Cut()                              # what Word does when the whole sentence is selected
        else:
            src.Range.Cut()                        # a partial selection: contents only
        dest = doc.Paragraphs(2).Range
        dest.Collapse(0)                           # end of paragraph 2 (before the mark)
        dest.Move(1, -1)
        dest.Paste()
        report(name, doc, ids, expect_order=[ids[0], ids[2], ids[3], ids[1]])
    finally:
        doc.Close(0)


def scenario_move_paragraph(app):
    doc, ids = fresh(app, "C_move_paragraph", lock=False)
    try:
        doc.TrackRevisions = True
        p2 = doc.Paragraphs(2).Range
        p2.Cut()
        top = doc.Range(0, 0)
        top.Paste()
        report("C_move_paragraph", doc, ids, expect_order=[ids[2], ids[3], ids[0], ids[1]])
    finally:
        doc.Close(0)


def scenario_merge(app):
    doc, ids = fresh(app, "D_merge_across_boundary", lock=False)
    try:
        doc.TrackRevisions = True
        a, b = cc_by_index(doc, 0), cc_by_index(doc, 1)
        # delete from the last word of A's target to the first word of B's target
        start = a.Range.Words(a.Range.Words.Count).Start
        end = b.Range.Words(1).End
        doc.Range(start, end).Delete()
        report("D_merge_across_boundary", doc, ids)
    finally:
        doc.Close(0)


def scenario_delete_sentence(app, name, lock):
    doc, ids = fresh(app, name, lock=lock)
    try:
        doc.TrackRevisions = False
        cc = cc_by_index(doc, 1)
        try:
            cc.Range.Delete()
            how = "Range.Delete ran"
        except Exception as e:                      # noqa: BLE001
            how = "Range.Delete refused: " + str(e).splitlines()[0][:60]
        print("\n  ", how)
        report(name, doc, ids)
    finally:
        doc.Close(0)


def main():
    app = w.word_app(visible=False)
    try:
        scenario_cut_paste(app, "A1_contents_cut_tracking_on", True, False)
        scenario_cut_paste(app, "A2_control_cut_tracking_on", True, True)
        scenario_cut_paste(app, "B1_contents_cut_tracking_off", False, False)
        scenario_cut_paste(app, "B2_control_cut_tracking_off", False, True)
        scenario_move_paragraph(app)
        scenario_merge(app)
        scenario_delete_sentence(app, "E_delete_locked", lock=True)
        scenario_delete_sentence(app, "F_delete_unlocked", lock=False)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
