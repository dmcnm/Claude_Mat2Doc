---
title: "Python → MATLAB mapping"
---

# Python → MATLAB mapping

Two living reference tables for the Mat2Doc port: **module → package** (where
each python-docx source module's symbols land in the `+mat2doc` namespace) and
**dunder / idiom** (how recurring Python constructs are rendered in base
MATLAB R2024b). Rows are appended per work package; this page starts with
**P1-1** (`docx/shared.py`).

## Module → package

| python-docx module `src/docx/...` | Mat2Doc package / files | WP | notes |
|---|---|---|---|
| `shared.py` — `Length`, `Inches`, `Cm`, `Emu`, `Mm`, `Pt`, `Twips`, `RGBColor` | `+mat2doc\+shared\` (one file per class) + `+shared\private\pyIntArg.m` | P1-1 | `Length < double` (int-subclass degradation); `RGBColor` is a value class (docx home package, not `dml/color.py`). `Twips` present / `Centipoints` absent per docx v1.2.0. `pyIntArg` is project infrastructure (no Python counterpart). |
| `oxml/ns.py` — `nsmap`, `pfxmap`, `nspfxmap`, `nsdecls`, `qn`, `NamespacePrefixedTag` | `+mat2doc\+oxml\` (`nsmap.m`, `pfxmap.m`, `nspfxmap.m`, `nsdecls.m`, `qn.m`, `NamespacePrefixedTag.m`) | P1-2 | `nsmap` is **public** in docx (16 entries, **no** `_nsmap`→`nsmap_` rotation); adds `dgm`/`w14`, rebinds `sl`, drops the pptx-only prefixes. `nspfxmap` is docx's subset function (pptx's `namespaces`/`nsuri` absent). OPC prefixes `ct`/`pr`/`r` live in a separate `opc/oxml.py` map (OPC-layer WP). |
| `oxml/parser.py` — `parse_xml`, `OxmlElement`, `register_element_cls` + `element_class_lookup`, `oxml_parser.makeelement` | `+mat2doc\+oxml\` (`parse_xml.m`, `OxmlElement.m`, `registry.m`, `createElement.m`) | P1-2 | `OxmlElement` is **3-arg** `(nsptag_str, attrs, nsdecls)` (docx delta vs pptx 2-arg); `attrs` is an Nx2 string `[name value]` (Clark keys aren't struct fields). `registry` EMPTY at P1-2 (120-row target lands with the `CT_*` WPs). `createElement` = the makeelement-equivalent (design-realization). |
| lxml `_Element` surface + `etree.fromstring(xml, oxml_parser)` read side (design-realization, **D-001**) | `+mat2doc\+oxml\XmlElement.m`, `XmlParser.m` | P1-2 | Own order-preserving OOXML-subset parser + `_Element` tree node (Heterogeneous root, Sealed eq/ne/deepcopy). `matlab.io.xml.dom` sorts attributes → harness-only, never on the toolbox path. Rejects `<!DOCTYPE>` (**D-006**), ASCII digit char-ref grammar (**D-002**), FIX-1 ws-tail-drop rule. |
| `opc/oxml.py::serialize_part_xml` (byte-matched to lxml `etree.tostring`) | `+mat2doc\+oxml\serialize_part_xml.m` | P1-2 | Placed in `+oxml` though the Python symbol lives in `opc/oxml.py` (XML-layer infra, design.md §3). lxml-matched: single-quote declaration, insertion-order attrs, H7 escaping, verbatim-until-moved nsdecls (**D-serializer-nsdecl**); residual foreign-file prefix-rewrite L2 (**D-nsprefix-rewrite**, case `a03`). |

**Not ported from `shared.py` in P1-1** (proxy tier, deferred to P2-1):
`lazyproperty`, `write_only_property`, `ElementProxy`, `Parented`,
`StoryChild`, `TextAccumulator`.

## Dunder / idiom

| Python idiom | MATLAB rendering | first WP |
|---|---|---|
| `class Length(int)` (int subclass; arithmetic degrades to a plain `int`) | `classdef Length < double` — defining properties forces every op to return a plain `double`; `isa(x,'double')` true; EMU held exactly (magnitudes ≪ 2^53, design.md §8) | P1-1 |
| `int.__new__(cls, emu)` | `obj@double(fix(emu)+0)` after `pyIntArg(emu,"parse")`; the `+0` normalizes IEEE `-0.0 → +0.0` | P1-1 |
| `int(x)` on a constructor argument | `fix(x)` (truncate toward zero — H6) after domain coercion by `pyIntArg` | P1-1 |
| `self / float(_EMUS_PER_*)` (true division) | `double(obj) / <Const>` (dependent-property getter) | P1-1 |
| `int(round(self / float(635)))` (Python-3 `round()` = half-to-even, returns int) | static `pyRoundHalfToEven_` replicating CPython `float.__round__` (`r=round(x); if abs(x-r)==0.5, r=2*round(x/2)`) + `+0` signed-zero normalize — H6/H14 | P1-1 |
| `@property` `emu` `return self` | `value = obj` (the getter returns the `Length` instance itself) | P1-1 |
| private class attrs `_EMUS_PER_*` | `Constant` properties `EMUS_PER_*_` (leading underscore rotated to trailing, design.md §2) | P1-1 |
| `class RGBColor(Tuple[int,int,int])` (immutable tuple, equality-by-value) | value `classdef` with immutable `r/g/b` properties + by-value `eq`/`ne`; comparison to a non-`RGBColor` is `false` | P1-1 |
| `__str__` / `__repr__` = `"%02X..."` / `"RGBColor(0x%02x...)"` | `str_()` / `repr_()` returning a `string` via `sprintf` (dunders rotate to trailing-underscore methods) — direct `%X`/`%x` hex, **not** routed through `pyStr` (H14) | P1-1 |
| `int(rgb_hex_str[:2], 16)` etc. | `parseHexByte_(s(1:2))` → validated `hex2dec` (base MATLAB, case-insensitive) | P1-1 |
| `raise TypeError(msg)` / `raise ValueError(msg)` | `error("mat2doc:TypeError", "%s", msg)` / `error("mat2doc:ValueError", "%s", msg)` — one identifier per Python exception type (D-004 namespace) | P1-1 |
| `isinstance(val, int)` where a MATLAB `double` cannot distinguish int from float | integer-**valued** real numeric scalar accepted as the Python int (D-STYPE-1); non-integral → `TypeError` | P1-1 |
| ordered `dict` prefix→URI (`nsmap`, `nspfxmap`) | scalar `struct`, insertion-ordered fields (H11); iterate `fieldnames`; never `containers.Map` | P1-2 |
| `dict` keyed by a URI or Clark name (not a valid struct field name) — `pfxmap`, `OxmlElement` `attrs` | **Nx2 `string` array** `[key, value]`, document order (H11) | P1-2 |
| `nsmap[pfx]` / dict `KeyError` | `isfield` guard → `error("mat2doc:KeyError", "%s", …)` | P1-2 |
| `etree.fromstring(xml, oxml_parser)` | `mat2doc.oxml.parse_xml(xml)` driving `XmlParser.parse` — own OOXML-subset reader (D-001), order-preserving | P1-2 |
| `oxml_parser.makeelement(clark, attrib, nsmap)` | `createElement` (+ `.set` per attr) with registry-then-`XmlElement` fallback | P1-2 |
| lxml `_Element` identity / `el1 is el2` (H5) | `handle` identity; `Sealed` `==`/`~=` on the Heterogeneous `XmlElement` root | P1-2 |
| `el.insert(i, c)` / `parent.index(child)` (0-based) | 1-based on the MATLAB surface (H1): `insert(i+1, c)`, `index(child)-1` for data | P1-2 |
| `copy.deepcopy(el)` / `__deepcopy__` | `XmlElement.deepcopy` (Sealed, structural copy of own decls + verbatim tag URI) | P1-2 |
| `etree.tostring(el, encoding="UTF-8", standalone=True)` | `mat2doc.oxml.serialize_part_xml(el)` — own writer, byte-matched to lxml (single `unicode2native` at the boundary, H2/H7) | P1-2 |
| bytes ↔ text at the XML boundary | `native2unicode` / `unicode2native('UTF-8')` only (H2); astral chars via surrogate pairs → 4-byte UTF-8 | P1-2 |

## Namespace policy

Mat2Doc uses the `mat2doc:` error-identifier namespace and the `mat2doc.*`
package namespace throughout — **no shared code with Mat2Ppt**. Designs common
to both toolboxes (the `Length` family, `pyIntArg`) are re-implemented, not
copied. The P0 same-package name-collision scan found **zero** collisions, so
**no symbol was renamed** (FLAG-3-docx resolved by convention).
