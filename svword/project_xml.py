"""The project record stored inside the .docx as a custom XML part.

One part per document, found by namespace. Replaced wholesale on save: the
record is small (one element per segment) and Word's node-level API is not
worth its complexity at this size. Designed for ~10,000 segments; a 300-page
document is ~5,000.
"""
from __future__ import annotations

import xml.etree.ElementTree as ET
from dataclasses import dataclass, field

NS = "urn:supervertaler:word:1"


@dataclass
class SegmentRecord:
    id: str
    source: str
    target: str = ""
    status: str = "draft"      # draft | confirmed | locked
    origin: str = ""           # tm | ai | human
    match: int = 0


@dataclass
class ProjectRecord:
    source_lang: str
    target_lang: str
    segments: list[SegmentRecord] = field(default_factory=list)
    resources: dict = field(default_factory=dict)   # termbases, tms, project_termbase: the pane's choices

    def to_xml(self) -> str:
        ET.register_namespace("sv", NS)
        root = ET.Element(f"{{{NS}}}project",
                          {"source_lang": self.source_lang, "target_lang": self.target_lang})
        if self.resources:
            ET.SubElement(root, f"{{{NS}}}resources", {k: str(v) for k, v in self.resources.items()})
        for s in self.segments:
            el = ET.SubElement(root, f"{{{NS}}}segment",
                               {"id": s.id, "status": s.status, "origin": s.origin,
                                "match": str(s.match)})
            ET.SubElement(el, f"{{{NS}}}source").text = s.source
            ET.SubElement(el, f"{{{NS}}}target").text = s.target
        return ET.tostring(root, encoding="unicode")

    @classmethod
    def from_xml(cls, xml: str) -> "ProjectRecord":
        root = ET.fromstring(xml)
        rec = cls(root.get("source_lang", ""), root.get("target_lang", ""))
        res = root.find(f"{{{NS}}}resources")
        if res is not None:
            rec.resources = dict(res.attrib)
        for el in root.findall(f"{{{NS}}}segment"):
            rec.segments.append(SegmentRecord(
                id=el.get("id"),
                source=el.findtext(f"{{{NS}}}source", ""),
                target=el.findtext(f"{{{NS}}}target", ""),
                status=el.get("status", "draft"),
                origin=el.get("origin", ""),
                match=int(el.get("match", "0")),
            ))
        return rec


def save(doc, record: ProjectRecord) -> None:
    for part in doc.CustomXMLParts.SelectByNamespace(NS):
        part.Delete()
    doc.CustomXMLParts.Add(record.to_xml())


def load(doc) -> ProjectRecord | None:
    parts = doc.CustomXMLParts.SelectByNamespace(NS)
    if parts.Count == 0:
        return None
    return ProjectRecord.from_xml(parts.Item(1).XML)
