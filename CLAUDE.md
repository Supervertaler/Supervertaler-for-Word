# Supervertaler for Word – Claude Code reference

Experimental CAT tool living inside Word. Read README.md first; it is short.

## Principles

- **The document is the project.** All state lives in the .docx: content controls,
  tracked changes, a custom XML part, comments. The hover tool is stateless.
- **Plain functions over the Word Document object.** No wrapper classes, no
  dependency injection. See `svword/word_com.py`.
- **Every spike verifies itself.** A spike prints ok/FAIL per claim and exits
  non-zero on failure. Run against a hidden Word instance; never touch the
  user's open documents.
- **Failure paths clean up.** Word instances are quit and the user's revision
  author name is restored in `finally` blocks. A crashed spike must not leave a
  hidden WINWORD.EXE behind.

## Word COM notes

- Constants are ints (no makepy). wdContentControlRichText=0, wdSentence=3,
  wdRevisionInsert=1, wdRevisionDelete=2, wdFormatXMLDocument=16.
- `Range.Text` EXCLUDES deleted-revision text (in every markup view), but
  `Range.Start/End` still count those characters. So target = `Range.Text`;
  source = the Supervertaler-authored deletions in `Range.Revisions`. Never map
  revision offsets onto `Range.Text` indices.
- The Revisions collection can list a nested revision twice; dedupe by span.
- A user deleting inside the AI insertion with tracking on creates a deletion
  in the user's name; filter by author or it pollutes the source.
- Setting `cc.Range.Text` while `doc.TrackRevisions` is on records delete+insert
  in one go. The content control survives Accept All and Reject All.
- `CustomXMLParts.SelectByNamespace(ns)` finds the project record; it is
  replaced wholesale on save.
- Word's `Sentences` collection splits on abbreviations ("Fig.", "e.g."). Replace
  with the Workbench segmenter from `Supervertaler-Python-Core` when wiring up.

## Related repos

- `Supervertaler-Workbench` – desktop CAT tool; source of the segmenter, TM, termbase code
- `Supervertaler-Python-Core` – shared Qt-free pipeline (private)
- `otto` – server-side one-click translation; natural producer of these files
- `Supervertaler-Sidekick` – floating window and global hotkeys; natural home of the hover tool

## Style

British English in user-facing text. En dashes, never em dashes.
