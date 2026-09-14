"""Thin COM layer over Word. Everything here is a plain function that takes the
Word Document object; no state is kept outside the document itself.

Word constants are written as ints so nothing needs makepy.
"""
from __future__ import annotations

import contextlib
import uuid

import win32com.client as com

WD_CC_RICH_TEXT = 0          # wdContentControlRichText
WD_SENTENCE = 3              # wdSentence (Range.Expand unit)
WD_REV_INSERT = 1            # wdRevisionInsert
WD_REV_DELETE = 2            # wdRevisionDelete

TAG_PREFIX = "sv:seg:"
REVISION_AUTHOR = "Supervertaler"


# --------------------------------------------------------------------- Word app

def word_app(visible: bool = False):
    """Attach to a running Word or start one."""
    try:
        app = com.GetActiveObject("Word.Application")
    except Exception:  # noqa: BLE001 – no running instance
        app = com.Dispatch("Word.Application")
    app.Visible = visible
    return app


@contextlib.contextmanager
def revision_author(app, name: str = REVISION_AUTHOR):
    """Revisions made inside the block are attributed to `name`.
    The user's own name is restored even when the block raises."""
    old = app.UserName
    app.UserName = name
    try:
        yield
    finally:
        app.UserName = old


# ------------------------------------------------------------------- sentences

def sentence_at_cursor(doc):
    """The Felix moment: expand the selection to the sentence under the cursor.
    Returns a Range. Uses Word's own sentence detection for now; the Workbench
    segmenter should replace it (Word splits on 'Fig.' and 'e.g.')."""
    rng = doc.ActiveWindow.Selection.Range
    rng.Expand(WD_SENTENCE)
    return rng


def sentences_in_paragraph(paragraph):
    """Yield sentence Ranges of one paragraph, skipping the empty tail."""
    for s in paragraph.Range.Sentences:
        if s.Text.strip():
            yield s


# ---------------------------------------------------------------------- anchors

def new_segment_id() -> str:
    return uuid.uuid4().hex[:12]


def anchor(doc, rng, seg_id: str | None = None):
    """Wrap a Range in a rich-text content control tagged with the segment id.
    Trailing whitespace is left outside the control so paragraphs keep their
    spacing when the text inside is replaced."""
    seg_id = seg_id or new_segment_id()
    text = rng.Text
    trailing = len(text) - len(text.rstrip())
    if trailing:
        rng.MoveEnd(1, -trailing)     # 1 = wdCharacter
    cc = doc.ContentControls.Add(WD_CC_RICH_TEXT, rng)
    cc.Tag = TAG_PREFIX + seg_id
    cc.Title = "draft"
    return cc


def anchors(doc):
    """All Supervertaler content controls in document order."""
    for cc in doc.ContentControls:
        if str(cc.Tag).startswith(TAG_PREFIX):
            yield cc


def segment_id(cc) -> str:
    return str(cc.Tag)[len(TAG_PREFIX):]


def set_target(app, doc, cc, target: str, status: str = "draft"):
    """Replace the anchored source with the target as a tracked change.
    In Word: Original view = source, No Markup = target, All Markup = bilingual."""
    was_tracking = doc.TrackRevisions
    doc.TrackRevisions = True
    try:
        with revision_author(app):
            cc.Range.Text = target
    finally:
        doc.TrackRevisions = was_tracking
    cc.Title = status


def _supervertaler_deletions(cc):
    """Deleted revisions made by Supervertaler inside an anchor, in document
    order, deduplicated by span (Word's Revisions collection can list a nested
    revision twice)."""
    seen, out = set(), []
    for r in cc.Range.Revisions:
        if r.Type != WD_REV_DELETE or r.Author != REVISION_AUTHOR:
            continue
        key = (r.Range.Start, r.Range.End)
        if key not in seen:
            seen.add(key)
            out.append((key, r.Range.Text))
    return [t for _, t in sorted(out)]


def source_of(cc) -> str:
    """Source text = what Supervertaler deleted when it wrote the target. The
    XML project record is the authority; this reads it from the document
    alone. An untranslated anchor returns its live text."""
    deleted = _supervertaler_deletions(cc)
    return "".join(deleted) if deleted else cc.Range.Text


def target_of(cc) -> str:
    """Target text = the anchor's live text. Range.Text already excludes every
    deleted revision, whoever made it, and includes text typed with tracking
    off. An untranslated anchor has no target yet."""
    return cc.Range.Text if _supervertaler_deletions(cc) else ""


# --------------------------------------------------------------------- comments

def annotate(doc, cc, text: str):
    """Attach a Word comment to an anchor (match info, agent questions)."""
    return doc.Comments.Add(cc.Range, text)
