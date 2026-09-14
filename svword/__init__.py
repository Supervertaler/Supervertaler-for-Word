"""Supervertaler for Word – the document is the project.

Source and target live inside the .docx itself:
  * content controls anchor each segment (tag = sv:seg:<id>)
  * tracked changes present source (deletion) and target (insertion)
  * a custom XML part holds the alignment map and statuses
  * comments carry match info and the agent's questions
"""
