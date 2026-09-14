# Supervertaler for Word

**A CAT tool that lives inside Microsoft Word. The document is the project.**

<img width="2560" height="1440" alt="image" src="https://github.com/user-attachments/assets/f2e67426-8da1-4360-b258-441716007b92" />

*Word shows the clean target text. The pane shows the sentence under the cursor
with its source, TM matches with differences marked, termbase hits, and a grid
that scrolls with the document. Test data, not a real job.*

Experimental. Started 14 September 2026. Not a product yet.

## The idea

Earlier Word-based CAT tools (Trados Workbench, Wordfast Classic, MetaTexis, Felix)
kept their state in fragile places: hidden text, purple styles, or a popup window
that forgot everything when it closed. Word has had four durable structures since
2007 that none of them used. This project uses all four:

| Structure | Role |
|---|---|
| **Content controls** | Anchor each segment. The tag `sv:seg:<id>` travels with the text through cut, paste and reflow. |
| **Tracked changes** | Source is the deletion, target the insertion. Word's *Original* view is the source document, *No Markup* is the translation, *All Markup* is a bilingual review file. Accept All delivers; Reject All restores. |
| **Custom XML part** | The project record (alignment map, statuses, origin, match %) rides inside the .docx. No sidecar files. |
| **Comments** | Match info and, later, the agent's questions to the translator. |

The editing surface is a Word task pane, docked by Word itself. The translator
works in No Markup, so the page is clean target text, and the pane follows the
cursor. The pane holds no state of its own: close Word, reopen the file
tomorrow, and everything resumes.

## The preview is the document

Every CAT tool has a preview pane, and every one of them is an approximation:
a rendering generated from an export, refreshed when the tool gets round to it,
never quite the file the client will open. Here there is no preview. The page
in Word *is* the document, live, and the grid in the pane is a second view of
the same anchors. Click a row and Word jumps to that sentence. Confirm a
segment and the page changes in front of you. Move a paragraph in Word and the
grid still knows which segment is which, because the anchors moved with the
text. What you are looking at while you translate is the deliverable, in its
final format, at every moment.

## There is no project file

The project file is the Word document. Literally.

Few people know it, but a .docx is a zip archive. Rename one to .zip and open
it: the text is in one XML file, the styles in another, the comments in
another, and so on. Word reads the parts it knows and carries the rest around
untouched. Our project record is one more part inside that zip. When you email
the document, copy it to another machine, or open it in six months, the project
travels with it, because it *is* the document.

| Inside the .docx | What it holds |
|---|---|
| The main document text | The target text, live, exactly as Word shows it |
| Tracked changes | The source, as a deletion under each target |
| Content controls | The segment anchors and their status |
| Comments | Match info now; the agent's questions later |
| One custom XML part | The record: language pair, and every segment's id, source, target, status and origin |

The record is small, one element per segment, so a 300-page document grows by
a few hundred kilobytes. It is also redundant on purpose: source and target can
be rebuilt from the tracked changes alone, which is what the round-trip tests
prove. Lose the record and the document still knows everything.

What is *not* in the file, and never will be: the translation memory and the
termbases. Those span many jobs and stay in the local engine. The document
only ever holds its own sentences, so when it leaves your machine nothing of
your TM goes with it beyond the sentences the client is paying for.

Two consequences. Delivery is a save: Accept All Changes and the file is the
deliverable, or send it as it is and let the client see the bilingual view.
And cleaning is optional: a strip command can remove the anchors and the
record for clients who want a plain file.

## Why tags mostly disappear

In Trados or memoQ, tags exist because the editor is not the document.
Formatting has to be smuggled through the grid as placeholders and put back on
export. Here the document is right there, live, and the target text lives
inside it. So the rule can be: the pane handles the text, and Word handles the
formatting. If a sentence needs a word in italics, you click into it in Word
and press Ctrl+I, and nothing in the pane or the record needs to know. The
record stores the plain target, the TM gets a plain sentence, and the
formatting lives where it belongs, in the file.

Most inline formatting a translator ever touches is bold, italics, underline,
super and subscript, and hyperlinks. None of that needs a tag.

What is left is a short list of things that are not text: footnote
references, cross-reference fields, index entries, inline images, bookmarks.
Those will appear in the pane as small opaque chips you can move but not edit,
and the pane will refuse to confirm a segment if one is missing. On most
documents that list is empty.

## Terminology in the source, not beside it

The source box in the pane is a TermLens: termbase targets sit under the words
they belong to, numbered, and Alt+1 to Alt+9 insert them at the caret in the
target box. A term with several candidates unfolds into a small picker on
click. No separate terminology window, no chips to scan.

One file, three consumers: Otto can emit it, the pane edits it in place, and
Supervertaler Workbench can open it as a project.

## What works today

- **Storage model**, proven over COM by `spike/roundtrip_demo.py`: anchor every
  sentence, translate as tracked changes, save, reopen, and verify that source
  and target are recoverable from the document alone, that Reject All yields
  the original and Accept All the target, and that anchors and the XML record
  survive both. The source cannot be lost by editing the target with tracking
  on or off (`spike/tracking_off_probe.py`).
- **Cursor following**, proven by `spike/follow_cursor.py`: Word's selection
  events reach Python, leaving an anchor refreshes the record, writing back
  moves the cursor on.
- **The task pane**, `spike/taskpane/`: an Office.js add-in served from
  localhost that reads the anchors and record, follows the cursor, writes
  targets back as tracked changes without disturbing the source deletion, and
  shows fuzzy TM matches, termbase hits and a grid. Loads in desktop Word 365
  over plain `http://localhost`, no certificate needed.

TM and termbase matches in the pane are a mock inside the page. The engine
will run locally so client data never leaves the machine.

## Try the pane

```powershell
cd "spike\taskpane"
powershell -NoProfile -File register.ps1      # once; -Remove to undo
python -m http.server 3000                    # leave running
```

Open a document prepared by `spike/roundtrip_demo.py`, then in Word:
Home > Add-ins > More Add-ins > Developer Add-ins > Supervertaler for Word.
After the first time there is a Supervertaler button on the Home tab.
Set Review > No Markup.

## What is installed, and what will be

Today nothing is installed; the spike is three loose pieces. The pane is a few
static files served from localhost by a Python web server. The manifest is an
XML file registered in the user's registry as a developer add-in, which is
what puts the Supervertaler button on Word's ribbon. And the Python side
anchors documents and runs the tests. Another person would need the repo,
Python, the server and the registration script. That is a developer setup.

The product is one local application, the local engine:

1. **The engine** runs in the background. It serves the pane's files on a
   localhost port and answers the pane's requests: TM matches, termbase hits,
   LLM drafts, Prepare document. Python, packaged the way the Workbench is,
   with a Mac build the same way.
2. **The installer** copies the engine, writes a manifest pointing at the
   engine's port, and registers it: a registry key on Windows, a manifest
   folder on Mac. Word shows the button on next start.
3. **Later, for reach**: the pane's static files hosted on a Supervertaler
   subdomain and the manifest listed in Microsoft's add-in store, so Word users
   find it under Add-ins and only the engine needs a download. A pane loaded
   over https can talk to the local engine because Word's browser exempts
   localhost from mixed-content blocking. To be verified before relying on it.

## Licensing

The pane is JavaScript served to a browser and cannot be protected, and does
not need to be. The engine is where the value is, so the engine is the gate,
using the same Lemon Squeezy flow as Supervertaler for Trados:

- activate a key against Lemon Squeezy with a machine instance name, validate
  on start-up and periodically, deactivate when moving machines. Their
  per-key activation limit gives per-seat control.
- cache the last successful validation and allow a set number of days offline,
  then fall back to a "please reconnect" state. No home-grown cryptography,
  the engine is not worth attacking.
- the pane asks the engine for licence status on start and shows a key entry
  screen if needed. The document itself keeps working; what stops without a
  licence is what the engine provides.
- no key ever lives in the pane.

A document anchored by a licensed user opens in Word for anyone, bilingual
view included. Only translating needs the product.

## Roadmap

1. **Prepare document** in the pane: segment and anchor a fresh Word file.
2. **Local engine** on localhost in front of the Workbench TM, termbase and LLM code.
3. **Formatting-safe write-back**: bold, italics, links and fields inside a sentence survive confirm.
4. **The Workbench segmenter** instead of Word's sentence detection.
5. **Packaging**: the engine serves the pane, an installer registers the add-in,
   licensing through Lemon Squeezy in the engine.

Then: agent questions as comments, a rulebook learned from what the translator
changes, and Draft All backed by the Otto pipeline running locally.

## Layout

```
svword/word_com.py     COM layer: anchors, tracked-change targets, sentence at cursor
svword/project_xml.py  the custom XML project record
spike/                 proofs of concept, each self-verifying
spike/taskpane/        the Office.js task pane, manifest and registration script
```

Windows only for now. The task pane itself is cross-platform; the COM spikes and
the local engine are not yet.

## Requirements

Python 3.10+, `pywin32`, Microsoft Word 365.

## Licence

MIT, like the rest of the Supervertaler family.
