---
title: "mat2doc.opc — OPC two-tier serializer + constants / spec / shared"
---

# `mat2doc.opc` — OPC two-tier serializer + constants / spec / shared

Ported from python-docx v1.2.0 modules `src/docx/opc/oxml.py`,
`src/docx/opc/constants.py`, `src/docx/opc/spec.py`, and `src/docx/opc/shared.py`
(packages `+mat2doc/+opc/` and `+mat2doc/+opc/+oxml/`). This is the **OPC
(Open Packaging Conventions) serializer layer** — the machinery that regenerates
the two byte-critical package-map parts every `.docx` carries:
**`[Content_Types].xml`** (the `CT_Types` / `CT_Default` / `CT_Override` tree)
and every **`.rels`** part (the `CT_Relationships` / `CT_Relationship` tree),
plus the OPC constants (`CONTENT_TYPE`, `RELATIONSHIP_TYPE`, `NAMESPACE`,
`RELATIONSHIP_TARGET_MODE`), the `default_content_types` spec pair-list, and the
`CaseInsensitiveDict` / `cls_method_fn` shared helpers.

It is a **two-tier serializer** in the OOXML sense: the *content-types* tier
(`ct:` namespace) declares the format of every part, and the *relationships*
tier (`pr:` / `r:` namespaces) wires parts together. Both tiers sit directly on
the P1-2 `+oxml` byte-fidelity foundation (`serialize_part_xml`, the
order-preserving parser, `XmlElement`).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## The M1 byte-critical role — what P1-4 PROVES and DEFERS

This layer is the byte source for the two package-map parts, so it carries the
core of the **M1 milestone** (a byte-identical `.docx` regenerated from a
round-trip). P1-4 proves that source **in isolation**, and defers the
surrounding assembly to later WPs. The Gate-3 bar (adopted verbatim):

- **PROVES — isolated-serializer L1.** Hand-fed `CT_Types` (3 `<Default>` + 11
  `<Override>` rows, in the frozen document order) serializes **byte-identical**
  to the frozen 1738-byte `[Content_Types].xml`, and hand-fed
  `CT_Relationships` serialize **byte-identical** to each frozen `.rels`
  (734 B `_rels/.rels`, 1227 B `word/_rels/document.xml.rels`, 295 B
  `customXml/_rels/item1.xml.rels`) — each **three ways**: MATLAB output ==
  python-docx `docx.opc.oxml` oracle == frozen `references\s0001\parts`.
- **DEFERS** (not proven here): default-vs-override **classification** logic,
  the sort / row-order derivation, `.rels` ordering-from-graph,
  partname→content-type resolution, and zip-entry DFS order → **P1-5 / P1-6a**.
  The integrated full-package M1 L1 proof is **P1-8** (the M1 gate).

So these element classes **append verbatim, in caller order, and never
reorder** — the sort and the `<Default>`-before-`<Override>` grouping are the
caller's job (`pkgwriter._ContentTypesItem`, P1-6a), by design. This WP builds
the *serializer*, proven byte-exact when fed the rows in the right order; the
*ordering* is a separate, later concern.

## The `<Default>`-before-`<Override>` emission and append-verbatim contract

`CT_Types.add_default` / `add_override` and `CT_Relationships.add_rel` each build
a child (`CT_Default` / `CT_Override` / `CT_Relationship`) and **append** it —
plain lxml `append`, never a `ZeroOrMore` schema-sequence inserter. The child
lands at the end, in call order. The frozen `[Content_Types].xml` has its three
`<Default>` rows before its eleven `<Override>` rows because the P1-6a caller
adds them in that order, **not** because these classes sort them. Matching
Python, the `add_*` methods return **nothing**.

## The raw-etree shape (docx is deliberately rawer than pptx)

The docx OPC `CT_*` classes subclass bare `etree.ElementBase` and read/write
their attributes with plain `get` / `set` / `find` / `findall` / `append` — **no
`RequiredAttribute` / `OptionalAttribute` descriptors, no simple-type
conversion, no attribute validation**. The port follows that exactly: the
getters are one-line `obj.get(...)` and the `.new` factories parse a standalone
element then `set` attributes. In particular `CT_Default.content_type` /
`extension` return `[]` (Python `None`, H3) when the attribute is absent — they
do **not** raise. This is intentionally different from the Mat2Ppt pptx `CT_*`
(xmlchemy descriptor classes that raise `InvalidXmlError`); adding that
validation would be **behavior-adding**, so it is not added. Consequently the
OPC base class `mat2doc.opc.oxml.BaseOxmlElement` extends
`mat2doc.oxml.XmlElement` (the `ElementBase` analogue), **not**
`mat2doc.oxml.BaseOxmlElement` — the OPC elements deliberately do not inherit
`xpath` / `getChild` / `getAttrTyped`.

## TargetMode is emitted only when External (H3 tri-state)

`CT_Relationship.new` writes the `TargetMode` attribute **only** inside
`if target_mode == "External"` (string equality, not a boolean). An **Internal**
relationship **omits the attribute entirely** — it is not written as
`TargetMode="Internal"`. The `target_mode` getter uses the two-argument
`get("TargetMode", RELATIONSHIP_TARGET_MODE.INTERNAL)`, so an absent attribute
reads back as the literal `"Internal"` (never `[]`). This tri-state is exactly
what the frozen `.rels` bytes require: `_rels/.rels` and
`word/_rels/document.xml.rels` carry only internal rels (no `TargetMode`), and an
external rel (e.g. a hyperlink target) emits `TargetMode="External"`.

## The `default_content_types` pair-list vs dict trap (CRITICAL)

`mat2doc.opc.default_content_types` returns an **Nx2 `string` pair-list**, not a
dict / `containers.Map`. `spec.py` is a **duplicate-key tuple list**: several
extensions map to more than one content type — the extension `bin` appears
**three** times (printer-settings for pptx / xlsx / docx). Python consumes it as
a *membership test on the whole pair*:

```python
(ext.lower(), content_type) in default_content_types
```

so **both** fields must match for a part to be classified as a `<Default>`
rather than an `<Override>` (`pkgwriter._ContentTypesItem`, P1-6a). Porting it as
a dictionary keyed by extension would **collapse** the three `bin` rows to one
and silently break default-vs-override precedence. It is therefore an Nx2 row
list, and membership is a **ROW test** over both columns:

```matlab
any(pairs(:,1) == ext & pairs(:,2) == ct)
```

never a key lookup. The membership test is order-insensitive (a set test in
Python), so the row order carries no behavior — it is preserved only for
readability / audit fidelity to `spec.py`.

## `CaseInsensitiveDict` — the lowercasing scope (H15)

`mat2doc.opc.CaseInsensitiveDict` is a `handle` class (Python subclasses `dict`)
whose keys match **without respect to case**. The fold is confined to exactly the
three overridden dunders — `set` (`__setitem__`), `get` (`__getitem__`), and
`isKey` (`__contains__`): each lowercases its key argument. The **stored** key is
the lowercased form, so `keys()` returns lowercased keys (which is what the
P1-6a consumer's `sorted(keys())` sees, and the lowercased extension is what
`add_default` emits as the `Extension` attribute). **Values are stored
verbatim** — never case-folded. It is created **empty** (constructor keys are
deliberately unhandled, per the docstring), matching the `dict` subclass that is
populated only via `__setitem__`.

## Registry — 5 OPC classes by raw Clark name

The 5 OPC element classes are merged into the single
`mat2doc.oxml.registry` **by raw Clark name** (via a local `registerClark_`,
plan-audit condition B2 option A), sourcing their namespace URIs from
`mat2doc.opc.NAMESPACE` so the Clark key equals the `xmlns` the `.new` factories
emit and the parser resolves:

| Clark tag key | MATLAB class |
|---|---|
| `{…/package/2006/content-types}Default` | `mat2doc.opc.oxml.CT_Default` |
| `{…/package/2006/content-types}Override` | `mat2doc.opc.oxml.CT_Override` |
| `{…/package/2006/content-types}Types` | `mat2doc.opc.oxml.CT_Types` |
| `{…/package/2006/relationships}Relationship` | `mat2doc.opc.oxml.CT_Relationship` |
| `{…/package/2006/relationships}Relationships` | `mat2doc.opc.oxml.CT_Relationships` |

In docx there are **two** parsers with two `element_class_lookup`s — the main
WordprocessingML one (the 120 `w:*` classes) and this OPC one (the 5 `CT_*`
classes). The port keeps **one** parser (`mat2doc.oxml.parse_xml`) and **one**
registry, which is behavior-preserving because the key spaces are **disjoint
namespaces**: an OPC part contains only `ct:` / `pr:` tags and a
WordprocessingML part only the main-map tags, so no tag resolves differently
under the merged table than under docx's two separate lookups. The OPC-local
`parse_xml` / `serialize_part_xml` / `serialize_for_reading` / `nsmap` / `qn`
files are thin delegators kept for **citation fidelity** to the
`docx/opc/oxml.py` module boundary.

## Deviation posture (adopt-and-verify, ZERO new D-numbers)

Every divergence exercised in this layer is a **recurrence of an already-adopted
Mat2Ppt deviation**, carried in the `mat2doc:` namespace per
`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`.
Gate-3 confirmed **0 new D-numbers**:

- **D-001** (own OOXML-subset parser) — the `CT_*.new` string literals transit
  `mat2doc.opc.oxml.parse_xml` → `mat2doc.oxml.parse_xml`; re-verified to parse
  and round-trip the frozen OPC parts byte-identically.
- **D-serializer-nsdecl** (verbatim-until-moved) — the load-bearing ruling that
  produces the single-root-`xmlns` byte form. Each `CT_*.new` parses a standalone
  `<X xmlns="…"/>` (the parser marks `nsVerbatim`); `add_default` / `add_override`
  / `add_rel` call `append`, which **clears** `nsVerbatim` → the MOVED path
  suppresses the child's stored default-URI declaration against the root's
  identical binding → a single `xmlns` on the root, exactly as in the frozen
  bytes. This mechanism *produces* the L1 bytes; it is not an L2 divergence.
- **Declaration single-quote + LF, insertion-order attributes, escaping table** —
  the P1-2 L1 conventions (goldens `declaration_docx.bin`, `attr_order_docx.bin`),
  exercised here, not deviations.
- **D-006** (DOCTYPE rejection) — inherited via the shared parser.

One **port-shape** note (not output-visible, not a deviation): Python
`CT_Relationships` overrides the base `.xml` property to return the `.rels`
**file bytes**; MATLAB cannot override an inherited property, so that byte member
is rotated to `xml_file_bytes` (plan-audit condition B4, the same rotation used
in Mat2Ppt). `CT_Relationships` inherits the base pretty `.xml` unchanged. The
sole byte consumer is `Relationships.xml` (rel.py, **P1-5**), which must call the
rotated name `xml_file_bytes`.

---

## `CT_Types`

**Syntax**

```matlab
t = mat2doc.opc.oxml.CT_Types.new()
    t.add_default(ext, content_type)
    t.add_override(partname, content_type)
d = t.defaults      % list of <ct:Default> children (document order)
o = t.overrides     % list of <ct:Override> children (document order)
```

**Description**

The `<Types>` element, the container / root of `[Content_Types].xml`.
`add_default` builds a `CT_Default` and appends it; `add_override` builds a
`CT_Override` and appends it — **verbatim, in caller order, never reordered**
(the sort and the `<Default>`-before-`<Override>` grouping are the P1-6a
caller's job). `defaults` / `overrides` return the child lists via the OPC-local
`qn`. Registered for `ct:Types`.

**Example**

```matlab
t = mat2doc.opc.oxml.CT_Types.new();
t.add_default("xml", mat2doc.opc.CONTENT_TYPE.XML);
t.add_override("/word/document.xml", mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN);
disp(numel(t.defaults))    % 1
disp(numel(t.overrides))   % 1
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::CT_Types` (lines 209–237)*

---

## `CT_Default`

**Syntax**

```matlab
d = mat2doc.opc.oxml.CT_Default.new(ext, content_type)
ct = d.content_type      % ContentType attribute, or [] (None, H3) if absent
ex = d.extension         % Extension attribute, or [] (None, H3) if absent
```

**Description**

A `<Default>` element in `[Content_Types].xml` — the default content type applied
to any part with the given extension. A **raw etree element**: the two attributes
are read via plain `get` and written by the `new` factory via plain `set`, with
no validation and no simple-type conversion. A missing attribute returns `[]`
(Python `None`), it does **not** raise. `new` sets `Extension` then `ContentType`
(that insertion order is the serialized order, H11). Registered for `ct:Default`.

**Example**

```matlab
d = mat2doc.opc.oxml.CT_Default.new("png", mat2doc.opc.CONTENT_TYPE.PNG);
disp(d.extension)      % "png"
disp(d.content_type)   % "image/png"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::CT_Default` (lines 88–112)*

---

## `CT_Override`

**Syntax**

```matlab
o = mat2doc.opc.oxml.CT_Override.new(partname, content_type)
ct = o.content_type      % ContentType attribute, or [] (None, H3) if absent
pn = o.partname          % PartName attribute, or [] (None, H3) if absent
```

**Description**

An `<Override>` element in `[Content_Types].xml` — the content type applied to
the single part with the given partname. Raw etree element (see `CT_Default`):
`new` sets `PartName` then `ContentType`. Registered for `ct:Override`.

**Example**

```matlab
o = mat2doc.opc.oxml.CT_Override.new("/word/document.xml", ...
    mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN);
disp(o.partname)       % "/word/document.xml"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::CT_Override` (lines 115–138)*

---

## `CT_Relationships`

**Syntax**

```matlab
rels = mat2doc.opc.oxml.CT_Relationships.new()
       rels.add_rel(rId, reltype, target, is_external)
bytes = rels.xml_file_bytes      % .rels file bytes (declaration + no whitespace)
list  = rels.Relationship_lst    % list of <pr:Relationship> children (document order)
```

**Description**

The `<Relationships>` element, the root of a `.rels` part. `add_rel` builds a
`CT_Relationship` and **appends** it (verbatim, caller order; returns nothing).
`xml_file_bytes` (the rotated Python `.xml` byte-override, condition B4) returns
the `.rels` file bytes via `serialize_part_xml`; `.xml` (inherited from the base)
remains the pretty-print test helper. Registered for `pr:Relationships`.

**Example**

```matlab
rels = mat2doc.opc.oxml.CT_Relationships.new();
rels.add_rel("rId1", mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, ...
    "word/document.xml", false);
disp(numel(rels.Relationship_lst))       % 1
b = rels.xml_file_bytes;                  % uint8 .rels bytes with declaration
disp(class(b))                            % "uint8"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::CT_Relationships` (lines 181–206)*

---

## `CT_Relationship`

**Syntax**

```matlab
r = mat2doc.opc.oxml.CT_Relationship.new(rId, reltype, target, target_mode)
id  = r.rId            % Id attribute
ty  = r.reltype        % Type attribute
tg  = r.target_ref     % Target attribute
tm  = r.target_mode    % TargetMode attribute, or default "Internal" if absent
```

**Description**

A single `<Relationship>` in a `.rels` part. Raw etree element with four
attributes; `new` sets `Id` → `Type` → `Target` → (optionally) `TargetMode`. The
**`TargetMode` attribute is written only when `target_mode == "External"`** — an
internal relationship omits it entirely, and the `target_mode` getter defaults to
`"Internal"` (two-argument `get`). `target_mode` has default
`RELATIONSHIP_TARGET_MODE.INTERNAL`. Registered for `pr:Relationship`.

**Example**

```matlab
r = mat2doc.opc.oxml.CT_Relationship.new("rId1", ...
    mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, "word/document.xml");
disp(r.rId)           % "rId1"
disp(r.target_mode)   % "Internal"  (no TargetMode attribute written)
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::CT_Relationship` (lines 140–178)*

---

## `nsmap` (OPC-local)

**Syntax**

```matlab
map = mat2doc.opc.oxml.nsmap()
```

**Description**

The OPC-local prefix→URI map, a scalar struct with fields `ct`, `pr`, `r` (in
that Python source order, H11). This is a **separate** map from the
WordprocessingML `mat2doc.oxml.nsmap`, which has no `ct` / `pr` bindings — the
OPC `CT_*` classes and their `qn` lookups resolve `ct:` / `pr:` **only** through
this map (H8).

**Example**

```matlab
m = mat2doc.opc.oxml.nsmap();
disp(m.ct)   % "http://schemas.openxmlformats.org/package/2006/content-types"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::nsmap` (lines 24–28)*

---

## `qn` (OPC-local)

**Syntax**

```matlab
clark = mat2doc.opc.oxml.qn(tag)
```

**Description**

Turns a prefixed OPC tag (`"pr:Relationship"`, `"ct:Default"`) into a
Clark-notation qualified name, resolving the prefix against the **OPC-local**
`nsmap` — not `mat2doc.oxml.qn`, which would `KeyError` on `ct` / `pr`. It is the
OPC-layer twin of the main `qn`.

**Example**

```matlab
disp(mat2doc.opc.oxml.qn("pr:Relationship"))
% "{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::qn` (lines 41–50)*

---

## `parse_xml` (OPC-local)

**Syntax**

```matlab
root = mat2doc.opc.oxml.parse_xml(xml)
```

**Description**

The `etree.fromstring()` replacement using the OPC parser — parses `xml` (UTF-8
bytes or a decoded string) and returns the root `XmlElement`, applying the
registry so a parsed `<Types>` / `<Relationships>` / `<Default>` / `<Override>` /
`<Relationship>` instantiates the corresponding `CT_*` subclass. A thin delegator
to `mat2doc.oxml.parse_xml` (one unified parser + registry, condition B2); kept as
a distinct file for citation fidelity.

**Example**

```matlab
xml = "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships""/>";
root = mat2doc.opc.oxml.parse_xml(xml);
disp(class(root))   % "mat2doc.opc.oxml.CT_Relationships"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::parse_xml` (lines 36–38)*

---

## `serialize_part_xml` (OPC-local)

**Syntax**

```matlab
xml_bytes = mat2doc.opc.oxml.serialize_part_xml(part_elm)
```

**Description**

Serializes `part_elm` to UTF-8 part-file bytes with the XML declaration — the
`etree.tostring(part_elm, encoding="UTF-8", standalone=True)` replacement. A thin
delegator to the P1-2 byte-critical `mat2doc.oxml.serialize_part_xml` (docx
*defines* the symbol in `opc/oxml.py`, but the implementation was pre-ported at
P1-2); all byte conventions — single-quote + LF declaration, insertion-order
attributes, H7 escaping, verbatim-until-moved nsdecls — are that engine's.

**Example**

```matlab
t = mat2doc.opc.oxml.CT_Types.new();
t.add_default("xml", mat2doc.opc.CONTENT_TYPE.XML);
b = mat2doc.opc.oxml.serialize_part_xml(t);
disp(class(b))   % "uint8"
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::serialize_part_xml` (lines 53–59)*

---

## `serialize_for_reading` (OPC-local + core)

**Syntax**

```matlab
xml = mat2doc.opc.oxml.serialize_for_reading(element)   % OPC-local delegator
xml = mat2doc.oxml.serialize_for_reading(element)       % the engine
```

**Description**

Pretty-printed XML (no declaration) — the
`etree.tostring(element, encoding="unicode", pretty_print=True)` replacement,
backing the `.xml` test-helper on both `BaseOxmlElement` bases. docx defines a
byte-identical `serialize_for_reading` in **both** `oxml/xmlchemy.py` and
`opc/oxml.py`; the port has **one engine** (`mat2doc.oxml.serialize_for_reading`,
new at P1-4) and the OPC-local file is a thin delegator.

This is a **test-only** surface — no docx production path reads it; the byte-exact
part path is `serialize_part_xml`. The libxml2 pretty-print algorithm is
reproduced faithfully (2-space indent per level; formatting turns **off** for a
mixed-content element and **propagates** off through its whole subtree; one
trailing LF after the root, including the root's tail), verified live against
lxml 5.3.0 at Gate-2 / Gate-3 (13/13 vectors).

**Example**

```matlab
t = mat2doc.opc.oxml.CT_Types.new();
t.add_default("xml", mat2doc.opc.CONTENT_TYPE.XML);
disp(t.xml)   % pretty-printed <Types>…</Types>\n (no declaration)
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::serialize_for_reading` (lines 62–67; engine `src/docx/oxml/xmlchemy.py::serialize_for_reading`, lines 22–27)*

---

## `BaseOxmlElement` (OPC-local)

**Syntax**

```matlab
% Base class for the OPC CT_* elements; not usually constructed directly.
classdef CT_Default < mat2doc.opc.oxml.BaseOxmlElement
```

**Description**

The OPC-local base class for the `CT_*` elements — `class
BaseOxmlElement(etree.ElementBase)` in docx, a **separate, minimal** base
**distinct** from the xmlchemy `mat2doc.oxml.BaseOxmlElement`. It adds only the
`.xml` pretty-print test-helper property; the `CT_*` subclasses are raw etree
elements using plain `get` / `set` / `find` / `findall` / `append`. The port
extends `mat2doc.oxml.XmlElement` (the `ElementBase` analogue), so the OPC
elements deliberately do **not** inherit `xpath` / `getChild` / `getAttrTyped`.
The constructor is a transparent pass-through (design.md §2 INT-1) so the parser
can instantiate registered `CT_*` classes via `feval(cls, name, ownDecls)`.

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::BaseOxmlElement` (lines 75–85)*

---

## `CONTENT_TYPE`

**Syntax**

```matlab
ct = mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN
```

**Description**

Content-type URIs (like MIME types) specifying a part's format — **87** class
constants transcribed verbatim from the docx source. Python implicit-string-
concatenation entries (a value split across two source lines for line length) are
joined into the single string they denote.

**Example**

```matlab
disp(mat2doc.opc.CONTENT_TYPE.PNG)   % "image/png"
```

*Ported from python-docx v1.2.0: `src/docx/opc/constants.py::CONTENT_TYPE` (lines 7–156)*

---

## `RELATIONSHIP_TYPE`

**Syntax**

```matlab
rt = mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT
```

**Description**

Open XML relationship-type URIs — **73** class constants transcribed verbatim.
Note the anomalous `PIVOT_CACHE_RECORDS`, whose value interleaves
`…/relationships` + `/spreadsheetml/pivotCacheRecords` (an infix differing from
its siblings), joined exactly.

**Example**

```matlab
disp(mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT)
% "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
```

*Ported from python-docx v1.2.0: `src/docx/opc/constants.py::RELATIONSHIP_TYPE` (lines 178–306)*

---

## `NAMESPACE`

**Syntax**

```matlab
ns = mat2doc.opc.NAMESPACE.OPC_CONTENT_TYPES
```

**Description**

The OPC-layer namespace URIs used by `[Content_Types].xml` and `.rels`; the
OPC-local prefix map `{ct, pr, r}` binds prefixes to three of them, and the
registry keys the 5 OPC classes on Clark names built from them.

**Example**

```matlab
disp(mat2doc.opc.NAMESPACE.OPC_CONTENT_TYPES)
% "http://schemas.openxmlformats.org/package/2006/content-types"
```

*Ported from python-docx v1.2.0: `src/docx/opc/constants.py::NAMESPACE` (lines 159–168)*

---

## `RELATIONSHIP_TARGET_MODE`

**Syntax**

```matlab
mode = mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL   % "External"
mode = mat2doc.opc.RELATIONSHIP_TARGET_MODE.INTERNAL   % "Internal"
```

**Description**

The Open XML relationship target modes. Hard dependency of `CT_Relationship.new`
(its `target_mode` default is `INTERNAL`) and `CT_Relationships.add_rel` (chooses
`EXTERNAL` vs `INTERNAL` from `is_external`).

**Example**

```matlab
disp(mat2doc.opc.RELATIONSHIP_TARGET_MODE.EXTERNAL)   % "External"
```

*Ported from python-docx v1.2.0: `src/docx/opc/constants.py::RELATIONSHIP_TARGET_MODE` (lines 171–175)*

---

## `default_content_types`

**Syntax**

```matlab
pairs = mat2doc.opc.default_content_types()   % Nx2 string [extension, content_type]
```

**Description**

The `(extension, content-type)` pairs eligible for a `<Default>` — an **Nx2
string pair-list** in Python source order, **not a dict** (see the pair-list-vs-
dict trap above). Duplicate keys are preserved (`bin` × 3); membership is a
**row** test over both columns, order-insensitive.

**Example**

```matlab
pairs = mat2doc.opc.default_content_types();
CT = mat2doc.opc.CONTENT_TYPE;
is_default = any(pairs(:,1) == "png" & pairs(:,2) == CT.PNG);   % true
disp(is_default)   % 1
```

*Ported from python-docx v1.2.0: `src/docx/opc/spec.py::default_content_types` (lines 5–24)*

---

## `CaseInsensitiveDict`

**Syntax**

```matlab
d = mat2doc.opc.CaseInsensitiveDict()
    d.set(key, value)
v = d.get(key)
tf = d.isKey(key)
k = d.keys()
```

**Description**

A `handle` mapping that matches keys without respect to case (H15). The
lowercasing is confined to `set` / `get` / `isKey`; stored keys are the
lowercased form (so `keys()` returns them lowercased), values are stored
verbatim. Created empty (constructor keys unhandled, per the docstring).

**Example**

```matlab
d = mat2doc.opc.CaseInsensitiveDict();
d.set("PNG", "image/png");
disp(d.isKey("png"))   % 1     (key folded to lowercase, H15)
disp(d.get("Png"))     % "image/png"
disp(d.keys())         % "png" (stored lowercased)
```

*Ported from python-docx v1.2.0: `src/docx/opc/shared.py::CaseInsensitiveDict` (lines 10–26)*

---

## `cls_method_fn`

**Syntax**

```matlab
fn = mat2doc.opc.cls_method_fn(cls, method_name)
```

**Description**

The `getattr(cls, method_name)` replacement — returns a handle to the named
static / class method of `cls`, so a caller can invoke it later without
re-deriving the class. `cls` is a fully-qualified class-name **string** and the
result is `str2func(cls + "." + method_name)`. The sole docx consumer is
`PartFactory` (`opc/part.py:192`), ported in a later WP (**P1-6b**); `str2func`
does not validate existence until the handle is invoked (VERIFY carried to P1-6b).

**Example**

```matlab
% (illustrative — the referenced class is ported in P1-6b)
fn = mat2doc.opc.cls_method_fn("mat2doc.opc.CONTENT_TYPE", "empty");
disp(class(fn))   % "function_handle"
```

*Ported from python-docx v1.2.0: `src/docx/opc/shared.py::cls_method_fn` (lines 29–31)*
