# Supervertaler for Word

**A CAT tool that lives inside Microsoft Word. The document is the project.**

Experimental. Started 14 September 2026. Nothing here is a product yet.

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

The editing surface is a hover tool in the Felix style: click a sentence in Word,
a floating window is the CAT tool. The hover tool holds no document state, so
closing it and reopening Word tomorrow resumes exactly where you were.

One file, three consumers: Otto can emit it, the hover tool edits it in place,
and Supervertaler Workbench can open it as a project.

## Status

`spike/roundtrip_demo.py` proves the storage model end to end against a real
Word instance over COM: anchor every sentence, translate as tracked changes,
save, reopen, and verify that source and target are both recoverable from the
document alone, that Reject All yields the original and Accept All the target,
and that anchors and the XML part survive both. All checks pass on Word 16.

```
python spike/roundtrip_demo.py spike\out
```

Not yet built: the hover window, TM and termbase lookup, a real translator
behind the stub, the Workbench segmenter in place of Word's sentence detection,
and any handling of documents that arrive with their own tracked changes.

## Layout

```
svword/word_com.py     thin COM layer: anchors, tracked-change targets, sentence at cursor
svword/project_xml.py  the custom XML project record
spike/                 proofs of concept, each self-verifying
```

Windows only for now (COM). An Office.js add-in would extend this to Mac and
Word for the web; the storage model is the same.

## Requirements

Python 3.10+, `pywin32`, Microsoft Word.

## Licence

MIT, like the rest of the Supervertaler family.
