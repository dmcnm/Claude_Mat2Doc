---
title: "Mat2Doc manual"
---

# Mat2Doc manual

**Mat2Doc** is an exact functional MATLAB replica of
[python-docx](https://github.com/python-openxml/python-docx) v1.2.0 — a
base-MATLAB (R2024b, no toolboxes) library for creating and manipulating Word
`.docx` files. Every symbol is a faithful, byte-equivalence-validated port of
its Python original: the OPC package layer, the order-preserving OOXML element
tree, the WordprocessingML document/paragraph/run/table object model, and the
shared value-type and simple-type machinery beneath them.

This manual is the API reference. Pages are organised by architectural layer,
from the foundation value types up through the high-level document, paragraph,
run, and table APIs. Each page documents one module's ported symbols — syntax,
behaviour, examples, and the `Ported from python-docx v1.2.0` provenance line
that ties every MATLAB symbol back to its source.

:::{note}
API pages are generated from the MATLAB help headers. Until the generator
lands they are maintained by hand to the same shape (one section per symbol:
syntax, description, example, ported-from).
:::

## Sections

- **Foundation & utilities** — the shared length-unit value types (the
  `Length` family and its EMU-based conversions) and the `RGBColor` value
  object.
- **XML layer (oxml)** — the byte-fidelity OOXML foundation: the namespace
  machinery, the order-preserving parser and `XmlElement` tree, the part-XML
  serializer, and the xmlchemy element-class / attribute-descriptor / XPath
  engine.
- **OPC layer (opc)** — the two-tier serializer that regenerates
  `[Content_Types].xml` and the `.rels` parts (`CT_Types` / `CT_Default` /
  `CT_Override` / `CT_Relationships` / `CT_Relationship`), plus the OPC
  constants, the `default_content_types` spec pair-list, and the shared helpers;
  and the **packaging tier** above it — the `PackURI` partname algebra, the
  `phys_pkg` physical zip reader/writer factories (with the reproducible
  `1980-01-01` zip writer), and the `Relationships` collection that regenerates
  each `.rels` part in insertion order; the **package assembly tier** above it
  — `PackageReader` (the serialized read path) and `PackageWriter` (whose
  zip-entry traversal is the byte-critical M1 order), with the serialized
  part/relationship value objects and the reader/writer content-type maps; and
  the **package/part object model** on top — `OpcPackage` (open / save /
  graph traversal), the `Part` / `XmlPart` base parts, and the `PartFactory`
  content-type→class registry whose XmlPart-vs-Part split decides M1
  whitespace-collapse (the full 17-part open→save round-trip is byte-proven here).
- **Python → MATLAB mapping** — the living module→package and dunder/idiom
  reference tables.

## Relationship to Mat2Ppt

Mat2Doc is a **standalone** toolbox — it shares no code with
[Mat2Ppt](https://github.com/python-openxml/python-docx) (the python-pptx
port). The two projects share only the binding translation rules
(`design.md`) and the cross-language validation harness, which live in the
management repository and ship in neither toolbox. Where a symbol has a direct
Mat2Ppt analogue (as most of the `Length` family does), the design is
**re-implemented** in the `mat2doc:` namespace against the **python-docx
v1.2.0 source of truth**, never copied — so, for example, docx's `Twips`
class and its half-to-even `twips` conversion are ported, while pptx's
`Centipoints` (absent from docx) is not.

## Deviation posture

Divergences from the Python original are all on **dead paths** (unreachable
through any ported call site, API-invisible). Mat2Doc pre-adopts the Mat2Ppt
deviation rulings, carried `mat2doc:`-namespaced (see
`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`):
**D-STYPE-1..4** (int/float indistinguishability), **D-002** (exotic
string-input dead paths), **D-003** (multiplier-constructor error-message
wording; exception class faithful), and **D-004** (non-finite / wrong-type
error identifiers). No new D-number has been opened.
