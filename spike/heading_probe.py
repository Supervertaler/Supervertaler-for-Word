"""Headings without a full stop must get exactly one anchor each."""
import os, sys, tempfile
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from svword import word_com as w

app = w.word_app(False)
doc = app.Documents.Add()
fails = 0
try:
    for p in ["TITLE IN CAPS", "Heading two", "Sentence one. Sentence two.", "Figure 1", "Last one"]:
        doc.Content.InsertAfter(p + "\r")
    doc.Paragraphs(1).Range.Font.AllCaps = True
    ccs = w.anchor_document(doc, "sentence")
    texts = [cc.Range.Text for cc in w.anchors(doc)]
    want = ["TITLE IN CAPS", "Heading two", "Sentence one.", "Sentence two.", "Figure 1", "Last one"]
    ok = texts == want and doc.ContentControls.Count == len(want)
    fails += not ok
    print(("ok   " if ok else "FAIL ") + "one anchor per unit, no paragraph marks: %s (controls=%d)" % (texts, doc.ContentControls.Count))
    marks = [cc for cc in doc.ContentControls if "\r" in cc.Range.Text]
    fails += bool(marks)
    print(("ok   " if not marks else "FAIL ") + "no anchor contains a paragraph mark")
    paras = [cc.Range.ParentContentControl for cc in doc.ContentControls]
    nested = sum(1 for cc in doc.ContentControls if cc.ParentContentControl is not None)
    fails += bool(nested)
    print(("ok   " if not nested else "FAIL ") + "no nested anchors (%d)" % nested)
finally:
    doc.Close(0); app.Quit()
print("ALL PASSED" if not fails else "%d FAILED" % fails)
sys.exit(1 if fails else 0)
