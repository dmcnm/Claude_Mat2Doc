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
  `Length` family and its EMU-based conversions), the `RGBColor` value object,
  and `pyStr`, the mandated Python-`str()` numeric→XML formatting helper (H14).
- **Proxy, exception & protocol tier** — the base tier that **every later API
  proxy extends** (P3–P6): `ElementProxy` (H5 element-identity equality + the
  `part` None-guard), the standalone `Parented` / `StoryChild` part-provider
  bases, `TextAccumulator`, the `lazyproperty` / `write_only_property` idiom
  guides, the `mat2doc.exc` exception raisers (flat identifiers), and the
  `mat2doc.types` structural protocols — plus the byte-neutral, Python-faithful
  `Document → ElementProxy` retrofit.
- **XML layer (oxml)** — the byte-fidelity OOXML foundation: the namespace
  machinery, the order-preserving parser and `XmlElement` tree, the part-XML
  serializer, and the xmlchemy element-class / attribute-descriptor / XPath
  engine — plus the **simple-type validator tier** (`+oxml\+simpletypes`: the 4
  Base / 10 Xsd / 21 `ST_*` classes that validate and convert every typed `w:`
  attribute value the P4–P6 element classes read and write; the D-STYPE-2 float
  re-home, the `ST_HexColor`/`ST_HexColorAuto` split, `ST_OnOff`, and the
  `ST_UniversalMeasure` unit grammar live here); and the **text/font element
  tier** (`+oxml\+text` + `+oxml\+shared`: the **first real `w:`-content element
  classes** — `CT_RPr`, the M2 byte-critical `<w:rPr>` run-properties container
  whose 27 `ZeroOrOne` descriptors ride the non-contiguous H11 successor slices
  of the 39-entry `_tag_seq`, its 6 child `CT_*` classes, the `_new_color`
  black-seed override, and the shared `CT_OnOff` (D-delta-1 tri-state) /
  `CT_String` leaves — registered byte-neutrally with M1 preserved).
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
  whitespace-collapse (the full 17-part open→save round-trip is byte-proven here);
  and the **core-properties tier** on top — `CT_CoreProperties` (the 15
  Dublin-Core descriptors, the W3CDTF date grammar, and the byte-critical `xsi`
  namespace hoist), the `CoreProperties` API wrapper, and the
  `CorePropertiesPart` part class (`/docProps/core.xml`; the 721 B re-serialize
  and 681 B fresh build are byte-proven).
- **WordprocessingML document tier (docx) — M1 → P2 complete** — the public
  `mat2doc.Document()` entry function and the object graph beneath it
  (`package.Package` / `parts.DocumentPart` / `document.Document`) that satisfies
  the **M1 milestone** (`mat2doc.Document().save()` produces `default.docx`
  byte-identical to python-docx — 17/17 L1 three-way — and opens clean in real
  Word), **completed by P2-3** with the `CT_Document`/`CT_Body` element classes
  (registered for `w:document`/`w:body` — byte-neutral, the last P2 step to touch
  the `document.xml` parse path), the `_Body` proxy and the `BlockItemContainer`
  (`blkcntnr`) base, and the `Document._body`/`_block_width` accessors. Documents
  the byte-neutral registration, the H11 successor ordering, the
  object-graph-complete/features-stubbed-at-P4+ state, and the **Phase 2 complete**
  milestone.
- **Story-part tier (parts)** — the part classes an opened `.docx` graph is
  built from: `StoryPart` (the real base inserted above `DocumentPart`, with
  `next_id` and the `_document_part` lazyproperty), the reparented
  `DocumentPart` (whose `_styles_part` / `_settings_part` now return real
  sibling parts), and the thin `StylesPart` / `SettingsPart` / `NumberingPart`
  `XmlPart` shells. Documents the **object-graph-vs-feature-surface**
  distinction P2-2 draws (part plumbing live, feature accessors stubbed at their
  real-phase owner), the three byte-neutral PartFactory flips, and **task #60**
  (the `xpath` hoist onto `XmlElement`, `Relationships.delitem`, and the live
  `Part.drop_rel` with its `< 2` refcount threshold).
- **Enumerations (enum)** — two pages. The **enumeration base tier**: `BaseEnum`
  (MS-API-value enums) and `BaseXmlEnum` (XML-attribute-mapping enums), the base
  machinery **every docx enum extends** (the concrete `WD_*` enums and the P4–P6
  element classes that (de)serialize enumerated attributes). Documents the
  MATLAB enum-with-associated-data idiom (value classdef + `enumeration` block),
  the four **docx-vs-pptx semantic deltas** (chiefly `from_xml`'s absent
  None-guard — load-bearing for the `INHERITED` members — and `to_xml`'s
  Python-falsy reject), and the H3 None/`""`/`<missing>` tri-state. Then the
  **concrete enum tier** (P3-3, `+enum\{+text,+section,+dml}`): the 12
  `WD_*` / `MSO_*` enums (108 members) that stand on that base, the
  `WD_BREAK_TYPE.TEXT_WRAPPING` member-alias idiom, the `None`-valued
  `INHERITED` members and their `from_xml(None)` linkage, the
  **`MSO_THEME_COLOR_INDEX` docx-vs-pptx delta** (full-word tokens,
  `NOT_THEME_COLOR="UNMAPPED"`, no `MIXED`), and the 7 module-level class
  aliases.
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
