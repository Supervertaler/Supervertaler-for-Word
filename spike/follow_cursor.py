"""Follow the Word cursor: a small window that shows the source of whatever
sentence the cursor is in, lets you edit the target, and writes it back as a
tracked change. Word stays in No Markup, so the page is clean target text.

    python spike/follow_cursor.py [path.docx]        (default: spike/out/roundtrip.docx)
    python spike/follow_cursor.py --selftest         (hidden Word, no window)

Rules proved here:
  * the document is the truth: leaving an anchor in Word re-reads it and
    refreshes the window and the XML record
  * the window writes back only on Ctrl+Enter, then moves the Word cursor on
"""
from __future__ import annotations

import os
import sys
import time

import pythoncom
import win32com.client as com

sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w                                  # noqa: E402
from svword.project_xml import load, save                          # noqa: E402

WD_MARKUP_NONE = 0     # wdRevisionsMarkupNone
DEFAULT = os.path.join(os.path.dirname(__file__), "out", "roundtrip.docx")


def anchor_at_selection(doc):
    """The Supervertaler anchor containing the selection, or None."""
    try:
        cc = doc.ActiveWindow.Selection.Range.ParentContentControl
    except pythoncom.com_error:
        return None
    if cc is None or not str(cc.Tag).startswith(w.TAG_PREFIX):
        return None
    return cc


def find_anchor(doc, seg_id):
    for cc in w.anchors(doc):
        if w.segment_id(cc) == seg_id:
            return cc
    return None


def next_anchor(doc, cc):
    """The first Supervertaler anchor that starts after `cc`, or None."""
    end = cc.Range.End
    for other in w.anchors(doc):
        if other.Range.Start >= end:
            return other
    return None


class Follower:
    """Holds the one piece of state that must exist: which anchor the cursor
    was last in, so that leaving it can be detected."""

    def __init__(self, app, doc, on_change):
        self.app, self.doc, self.on_change = app, doc, on_change
        self.current_id = None
        self.record = load(doc)
        self.by_id = {s.id: s for s in self.record.segments} if self.record else {}

    def poll(self):
        """Called on every selection change (event) or tick (fallback)."""
        cc = anchor_at_selection(self.doc)
        seg_id = w.segment_id(cc) if cc is not None else None
        if seg_id == self.current_id:
            return
        if self.current_id is not None:
            self._left(self.current_id)
        self.current_id = seg_id
        self.on_change(cc, self.by_id.get(seg_id))

    def _left(self, seg_id):
        """Word is the truth: re-read the anchor we just left into the record."""
        cc = find_anchor(self.doc, seg_id)
        seg = self.by_id.get(seg_id)
        if cc is not None and seg is not None:
            seg.target = w.target_of(cc)
            save(self.doc, self.record)

    def write_back(self, target: str, status: str = "confirmed"):
        cc = anchor_at_selection(self.doc)
        if cc is None:
            return False
        w.set_target(self.app, self.doc, cc, target, status)
        seg = self.by_id.get(w.segment_id(cc))
        if seg is not None:
            seg.target, seg.status, seg.origin = target, status, "human"
            save(self.doc, self.record)
        nxt = next_anchor(self.doc, cc)
        if nxt is not None:
            nxt.Range.Select()
        return True


class AppEvents:
    """pywin32 event sink for Word.Application."""
    follower: Follower | None = None

    def OnWindowSelectionChange(self, sel):
        if self.follower:
            self.follower.poll()


def open_doc(app, path):
    doc = app.Documents.Open(os.path.abspath(path))
    doc.TrackRevisions = True
    try:
        doc.ActiveWindow.View.RevisionsFilter.Markup = WD_MARKUP_NONE
    except pythoncom.com_error:
        pass
    return doc


def pump(n=20):
    for _ in range(n):
        pythoncom.PumpWaitingMessages()
        time.sleep(0.01)


# ------------------------------------------------------------------ self-test

def selftest():
    import shutil
    import tempfile
    path = os.path.join(tempfile.mkdtemp(prefix="svword_"), "follow.docx")
    shutil.copy(DEFAULT, path)

    app = w.word_app(visible=False)
    events = com.WithEvents(app, AppEvents)
    seen = []
    doc = open_doc(app, path)
    failures = 0
    try:
        f = Follower(app, doc, lambda cc, seg: seen.append(seg.id if seg else None))
        events.follower = f
        ids = [w.segment_id(c) for c in w.anchors(doc)]
        for cc in list(w.anchors(doc)):              # click into each anchor
            doc.Range(cc.Range.Start + 1, cc.Range.Start + 1).Select()
            pump()
        events_fired = list(seen)
        doc.Range(0, 0).Select(); pump(); f.poll()
        ok = events_fired == ids
        failures += not ok
        print(("ok   " if ok else "FAIL ") + "selection events fired for each anchor in order")
        if not ok:
            print("     events:", events_fired, "anchors:", ids)

        # edit in Word with tracking on, leave the anchor: record must follow
        first = find_anchor(doc, ids[0])
        doc.Range(first.Range.Start + 1, first.Range.Start + 1).Select(); pump(); f.poll()
        first.Range.Words(1).Delete()
        doc.Range(0, 0).Select(); pump(); f.poll()
        got = load(doc).segments[0].target
        live = w.target_of(find_anchor(doc, ids[0]))
        ok2 = got == live and not got.startswith("Een ") and got.startswith("mashup")
        failures += not ok2
        print(("ok   " if ok2 else "FAIL ") + "leaving an edited anchor refreshed the record: %r" % got)

        # write back from the window
        second = find_anchor(doc, ids[1])
        doc.Range(second.Range.Start + 1, second.Range.Start + 1).Select(); pump(); f.poll()
        new = "De term impliceert een makkelijke, snelle integratie."
        f.write_back(new)
        cc1 = find_anchor(doc, ids[1])
        ok3 = (w.target_of(cc1) == new
               and w.source_of(cc1) == "The term implies easy, fast integration."
               and cc1.Title == "confirmed"
               and load(doc).segments[1].status == "confirmed")
        failures += not ok3
        print(("ok   " if ok3 else "FAIL ") + "write_back kept source, set target, status and record")
        if not ok3:
            print("     target:", repr(w.target_of(cc1)), "source:", repr(w.source_of(cc1)), cc1.Title)

        # after write_back the cursor should be in the next anchor
        pump(); f.poll()
        ok4 = f.current_id == ids[2]
        failures += not ok4
        print(("ok   " if ok4 else "FAIL ") + "cursor moved on to the next segment (%s)" % f.current_id)
    finally:
        doc.Close(0)
        app.Quit()
    print("\n" + ("ALL PASSED" if not failures else "%d FAILED" % failures))
    return 1 if failures else 0


# --------------------------------------------------------------------- window

def window(path):
    import tkinter as tk
    from tkinter import scrolledtext

    app = w.word_app(visible=True)
    events = com.WithEvents(app, AppEvents)
    doc = open_doc(app, path)

    root = tk.Tk()
    root.title("Supervertaler for Word – spike")
    root.attributes("-topmost", True)
    root.geometry("560x380+40+40")
    head = tk.Label(root, text="Click into a sentence in Word", anchor="w",
                    font=("Segoe UI", 10, "bold"))
    head.pack(fill="x", padx=8, pady=(8, 2))
    src = scrolledtext.ScrolledText(root, height=5, wrap="word", font=("Segoe UI", 11),
                                    state="disabled")
    src.pack(fill="both", expand=True, padx=8)
    tk.Label(root, text="Target  (Ctrl+Enter writes it back as a tracked change and moves on)",
             anchor="w").pack(fill="x", padx=8, pady=(6, 0))
    tgt = scrolledtext.ScrolledText(root, height=5, wrap="word", font=("Segoe UI", 11))
    tgt.pack(fill="both", expand=True, padx=8, pady=(0, 8))

    def show(cc, seg):
        src.configure(state="normal")
        src.delete("1.0", "end")
        tgt.delete("1.0", "end")
        if cc is None:
            head.configure(text="Not in a segment")
        else:
            head.configure(text="Segment %s · %s" % (w.segment_id(cc), cc.Title))
            src.insert("1.0", seg.source if seg else w.source_of(cc))
            tgt.insert("1.0", w.target_of(cc))
        src.configure(state="disabled")

    f = Follower(app, doc, show)
    events.follower = f

    def write_back(_=None):
        f.write_back(tgt.get("1.0", "end").strip())
        return "break"
    tgt.bind("<Control-Return>", write_back)

    def tick():
        pythoncom.PumpWaitingMessages()
        f.poll()                       # fallback for any click the event missed
        root.after(150, tick)

    def close():
        try:
            doc.Save()
        finally:
            root.destroy()
    root.protocol("WM_DELETE_WINDOW", close)
    root.after(150, tick)
    root.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv:
        sys.exit(selftest())
    window(sys.argv[1] if len(sys.argv) > 1 else DEFAULT)
