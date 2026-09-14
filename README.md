# Supervertaler for Word

**A CAT tool that lives inside Microsoft Word. The document is the project.**

<img width="2560" height="1440" alt="image" src="https://github.com/user-attachments/assets/179cfea8-6efc-455b-a0b0-ae549bad4242" />

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

## Roadmap

1. **Prepare document** in the pane: segment and anchor a fresh Word file.
2. **Local engine** on localhost in front of the Workbench TM, termbase and LLM code.
3. **Formatting-safe write-back**: bold, italics, links and fields inside a sentence survive confirm.
4. **The Workbench segmenter** instead of Word's sentence detection.
5. **Packaging**: the engine serves the pane, an installer registers the add-in.

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
