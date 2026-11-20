# Known gaps (pptx)

The steps oxoffice blocks or gets wrong in this release, each with its workaround. "Fallback" means: do that
step only with the default `pptx` skill (python-pptx), and keep oxoffice for the rest and for the verification.

Before reaching for lxml, try the raw XML of the element (`shape.element.xml` / `element.replace_xml(xml)`, also on
text frames, paragraphs, runs, tables, cells, slides, layouts and masters: `references/python-api.md`, "Raw XML").

No gap is open in this release. A part holding a control character that XML 1.0 forbids is refused at open
(`references/cli.md`, validate): delete the character from that part's bytes, then open the file.

Every public mutating call costs the same at 5 slides and at 20.

A step that oxoffice blocks, or gets wrong, and that this page does not list is a new gap: report it as `SKILL.md`
describes (Reporting a bug or a missing feature).

Also: `shape.text = ...`, `text_frame.text = ...` and `paragraph.text = ...` drop run formatting, as in
python-pptx; edit `runs[k].text` to keep it.
