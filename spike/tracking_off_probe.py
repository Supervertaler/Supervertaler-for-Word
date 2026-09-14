"""Can the source be lost by editing the target?  Three scenarios on a copy of
roundtrip.docx, each run on its own fresh copy, hidden Word:

  A  tracking OFF, user types extra words at the end of a target
  B  tracking OFF, user deletes the first word of a target
  C  tracking ON (as the user, not Supervertaler), same deletion as B

For each: is the source still recoverable, and what does target_of() report?
"""
from __future__ import annotations
import os, shutil, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w

SRC = os.path.join(os.path.dirname(__file__), "out", "roundtrip.docx")


def scenario(app, name, tracking, edit):
    path = os.path.join(tempfile.mkdtemp(prefix="svword_"), name + ".docx")
    shutil.copy(SRC, path)
    doc = app.Documents.Open(path)
    try:
        cc = next(w.anchors(doc))
        src0, tgt0 = w.source_of(cc), w.target_of(cc)
        doc.TrackRevisions = tracking
        edit(doc, cc)
        cc = next(w.anchors(doc))            # re-fetch: COM proxies can go stale
        print(f"\n[{name}] tracking={'on' if tracking else 'off'}")
        print(f"  source intact : {w.source_of(cc) == src0}   ({w.source_of(cc)!r})")
        print(f"  target before : {tgt0!r}")
        print(f"  target after  : {w.target_of(cc)!r}")
        print(f"  live text     : {cc.Range.Text!r}")
        print(f"  revisions     : {[(r.Type, r.Author) for r in cc.Range.Revisions]}")
    finally:
        doc.Close(0)


def type_at_end(doc, cc):
    r = cc.Range
    r.Collapse(0)                      # 0 = wdCollapseEnd
    r.InsertAfter(" (extra)")


def delete_first_word(doc, cc):
    cc.Range.Words(1).Delete()


def main():
    app = w.word_app(visible=False)
    try:
        scenario(app, "A_off_insert", False, type_at_end)
        scenario(app, "B_off_delete", False, delete_first_word)
        scenario(app, "C_on_delete", True, delete_first_word)
    finally:
        app.Quit()


if __name__ == "__main__":
    main()
