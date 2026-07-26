---
title: "mat2doc core-properties tier — CT_CoreProperties + CoreProperties + CorePropertiesPart"
---

# core-properties tier — `CT_CoreProperties` + `CoreProperties` + `CorePropertiesPart`

Ported from python-docx v1.2.0 modules `src/docx/oxml/coreprops.py`,
`src/docx/opc/coreprops.py`, and `src/docx/opc/parts/coreprops.py`, plus the
shared numeric-formatting helper `src/docx/shared.py`-adjacent `pyStr` (see
`shared.md`). This is the **document-metadata tier** — the layer that reads and
writes the Dublin-Core properties stored in the OPC part `/docProps/core.xml`.
It spans three packages:

- **`+mat2doc/+oxml/+coreprops/CT_CoreProperties.m`** — the custom oxml element
  class for the `<cp:coreProperties>` root: the 15 Dublin-Core child
  descriptors, the W3CDTF date parse/format machinery, and the byte-critical
  `xsi` namespace hoist for the two date-with-type children.
- **`+mat2doc/+opc/CoreProperties.m`** — the API wrapper: 15 read/write
  properties that delegate to the wrapped `CT_CoreProperties` element.
- **`+mat2doc/+opc/+parts/CorePropertiesPart.m`** — the `XmlPart` subclass for
  `/docProps/core.xml`, plus the `default()` factory that builds a fresh
  core-properties part for a core-less package.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## Byte-fidelity is proven at P1-7 (headline)

The M1-relevant `core.xml` bytes are byte-identical to python-docx on both the
parse→re-serialize path and the fresh-build path (Gate-3
`validate_P1-7_coreprops.md`):

- **721 B `core.xml` re-serialize — three-way byte-identical.** Parsing the
  frozen template `core.xml` yields a `CT_CoreProperties` root (registry-routed,
  the one new `registry.m` row), and re-serializing it reproduces the frozen
  721 B **byte-for-byte** — MATLAB == python-docx `parse_xml→serialize_part_xml`
  == frozen bytes (sha256 `d14be828…`).
- **681 B fresh-build (xsi hoist) — byte-identical.** A fresh
  `CT_CoreProperties.new()` with a
  title/author/keywords/created/modified/lastPrinted/revision battery serializes
  to the **same 681 B** as python-docx (sha256 `3be041bb…`): `xmlns:xsi` appears
  exactly once on the root, `xsi:type="dcterms:W3CDTF"` on exactly the two date
  children, nsdecl order `cp/dc/dcterms/xsi`. This path also re-proves H7
  escaping, H2 non-ASCII, and the H14 `pyStr` numeric route on **generated**
  content.

The wall-clock `modified` stamp of `default()` (D-coreprops-time) is the only
non-byte-reproducible surface, and it is **M1-unreachable** — see the
`CorePropertiesPart` section.

## Deviation posture (all adopt-only, ZERO new D-numbers)

Every divergence in this tier is a **recurrence of an already-signed deviation**,
carried `mat2doc:`-namespaced. Gate-3 confirmed **0 new D-numbers**
(`validation\summary\deviation_ledger.md`):

| D-number | Site | What it is / reachability |
|---|---|---|
| **D-001** | own parser/serializer (inherited) | `parse_xml` / `serialize_part_xml` re-exercised, byte-proven by the 721 B + 681 B matches. |
| **D-002** | `parse_int_` (revision) + `parse_W3CDTF_to_datetime_` (dates) | ASCII `[0-9]` int/date grammar under-accepts Unicode digits / `_` grouping / `>2^53` vs CPython `int()` / `_strptime`. The **safe** under-accept direction; divergent inputs fall back to `0` / `[]` (or the ledgered nearest-double); dead on real `core.xml`. Never emits `NaN`/`Inf`. The P1-7 docx site citation was added to the existing D-002 ledger row. |
| **D-serializer-nsdecl** | `set_element_datetime_` xsi hoist | places `xmlns:xsi` once on the root via `set_nsdecl_`; byte-verified once-on-root by the 681 B build. |
| **D-005** | `set_element_datetime_` non-datetime error | message reports the MATLAB class token where CPython reports `type(value)`; dead path (values arrive as `datetime`). |
| **D-coreprops-time** | `CorePropertiesPart.default()` | wall-clock `modified` stamp; **M1-unreachable** (see below). Same family as D-zip-time. |

## The docx tier is line-for-line the pptx tier

python-docx `coreprops.py` is structurally identical to python-pptx
`coreprops.py` (both inherit the shared OPC core-properties lineage), so the
Mat2Ppt solutions to the hard parts — the W3CDTF round-trip, the `xsi` nsdecl
hoist, the revision int grammar — carry over verbatim (each re-verified
line-by-line against docx v1.2.0). The only docx-vs-pptx **value** deltas are the
`default()` literals: title `"Word Document"` (pptx `"PowerPoint Presentation"`)
and last_modified_by `"python-docx"` (pptx `"python-pptx"`), both faithful to
docx v1.2.0 `parts/coreprops.py`.

---

## `CT_CoreProperties`

**Syntax**

```matlab
element = mat2doc.oxml.coreprops.CT_CoreProperties.new()
```

**Description**

Custom element class for the `<cp:coreProperties>` element — the root of the
Core Properties part stored as `/docProps/core.xml`. It implements the
Dublin-Core document-metadata elements and is registered for
`{...core-properties}coreProperties` in `registry.m` (the one new main-map row,
`oxml/__init__.py:96`), so parsing a real `core.xml` yields a
`CT_CoreProperties` root.

**The 15 Dublin-Core child descriptors.** Each is a `ZeroOrOne` child with
`successors=()` (so insertion always appends at the end of the element), ported
per design.md §2 as a Constant `TAG` schema table plus one-line delegating
members over the `BaseOxmlElement` child engine
(`getChild`/`getOrAddChild`/`insertChildInSequence`/`addChild`/`removeChild`).
The Dublin-Core accessor properties expose them:

| Accessor | Element | Type |
|---|---|---|
| `author_text` | `<dc:creator>` | string |
| `category_text` | `<cp:category>` | string |
| `comments_text` | `<dc:description>` | string |
| `contentStatus_text` | `<cp:contentStatus>` | string |
| `identifier_text` | `<dc:identifier>` | string |
| `keywords_text` | `<cp:keywords>` | string |
| `language_text` | `<dc:language>` | string |
| `lastModifiedBy_text` | `<cp:lastModifiedBy>` | string |
| `subject_text` | `<dc:subject>` | string |
| `title_text` | `<dc:title>` | string |
| `version_text` | `<cp:version>` | string |
| `created_datetime` | `<dcterms:created>` | datetime |
| `lastPrinted_datetime` | `<cp:lastPrinted>` | datetime |
| `modified_datetime` | `<dcterms:modified>` | datetime |
| `revision_number` | `<cp:revision>` | int |

**String tri-state (H3).** A string accessor returns `""` when the descriptor
child is absent **or** its text is `None` (`[]`) — `""` is a real empty string,
never a None test. Setting a string routes through `pyStr` (H14, F2) so a
non-string value serializes exactly as Python `str(value)` would, then is
length-limited to 255 Unicode **code points** (counting code points, not UTF-16
units, so an astral char counts 1 — matching Python `len(str)`).

**W3CDTF dates.** `parse_W3CDTF_to_datetime_` mirrors the CPython `strptime`
loop: four anchored regex templates replicating `_strptime` field widths (`%Y`
exactly four digits; `%m`/`%d`/`%H`/`%M`/`%S` one-or-two digits), **last**
successful match wins, the whole parseable part must match, then
`datetime(...)` component round-trip rejects rolled-over fields (out-of-range →
`strptime` `ValueError`) plus a `year >= 1` floor (CPython `MINYEAR = 1`). The
numeric-offset handling preserves the **inverted-sign** convention verbatim
(`sign_factor = -1` for `+`), so `2003-12-31T10:14:55-08:00` reads as
`18:14:55`. `set_element_datetime_` serializes with
`strftime("%Y-%m-%dT%H:%M:%SZ")` (whole-second; MATLAB `ss` truncates like
`%S`).

:::{note}
**Getters return a naive datetime (API-value only, VERIFY-tz).** The date
getters return a MATLAB datetime with `TimeZone=''`, mirroring the validated
Mat2Ppt precedent, where python-docx returns a **tz-aware UTC** datetime. This
does **not** affect XML bytes — the format path emits the same wall-clock fields
plus a literal `Z` (proven by the 721 B / 681 B byte-identity) — but an API
probe comparing the returned object sees naive-vs-UTC. Gate-3 compares
**wall-clock fields, never tzinfo**.
:::

**The `xsi` namespace hoist (byte-critical, D-serializer-nsdecl).** The
`created`/`modified` setters must place `xmlns:xsi` on the **root once** (not
per child) and stamp `xsi:type="dcterms:W3CDTF"` on the child. python-docx does
this with an opaque lxml reconciliation trick (set a throwaway `xsi:foo` on the
root, set `xsi:type` on the child, delete `xsi:foo`) whose documented intent is
exactly "add the `xsi` namespace to the root element rather than each child
element." The port translates the trick to its stated intent —
`set_nsdecl_("xsi", nsmap().xsi)` on the root — after which the serializer
renders the child attribute as `xsi:type` with no redundant per-child `xmlns`.
The result is byte-identical to lxml: trailing decl placement
`cp,dc,dcterms,xsi`, `xmlns:xsi` exactly once.

**The revision int grammar (D-002).** `revision_number` reads `<cp:revision>`
through `parse_int_`, an ASCII `[0-9]` int-literal grammar: absent child → `0`;
present-but-empty `<cp:revision/>` (text `None`) replicates `int(str(None))`
failing → `0`; `"1"` → 1; `"abc"`/`"-5"`/`"2.5"` → 0. The setter guards for a
positive int and, on rejection, reports the offending value through `pyStr` (so
`2.5` renders `'2.5'`, not `'2'`, in the `mat2doc:ValueError`). Unicode-digit /
underscore-grouped / `>2^53` inputs are the signed D-002 divergences (fall back
to `0` or the nearest-double); dead on real `core.xml`.

**Example**

```matlab
e = mat2doc.oxml.coreprops.CT_CoreProperties.new();
e.title_text = "Report";         % creates <dc:title>Report</dc:title>
disp(e.title_text)               % "Report"
disp(e.revision_number)          % 0 (no <cp:revision> yet)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/coreprops.py::CT_CoreProperties`
(registered for `<cp:coreProperties>`, `oxml/__init__.py:96`).*

---

## `CoreProperties`

**Syntax**

```matlab
cp = mat2doc.opc.CoreProperties(element)
```

**Description**

Dublin-Core document properties for the package, corresponding to the part named
`/docProps/core.xml`. Provides broadly-standardized document metadata as 15
read/write properties, each delegating to the wrapped `CT_CoreProperties`
element's `*_text` / `*_datetime` / `revision_number` accessor:

`author`, `category`, `comments`, `content_status`, `created`, `identifier`,
`keywords`, `language`, `last_modified_by`, `last_printed`, `modified`,
`revision`, `subject`, `title`, `version`.

It is a `handle` class (design.md §2): a proxy that wraps a shared element tree,
so two `CoreProperties` over the same element are views of one object — matching
Python reference semantics. Because every accessor delegates to the element, the
element's deviations apply transitively: `revision` reads through `parse_int_`
and `created`/`modified`/`last_printed` through the W3CDTF grammar (both D-002),
and the `created`/`modified` setters carry the `xsi` nsdecl hoist
(D-serializer-nsdecl).

**Example**

```matlab
e  = mat2doc.oxml.coreprops.CT_CoreProperties.new();
cp = mat2doc.opc.CoreProperties(e);
cp.author = "Ada";                % <dc:creator>Ada</dc:creator>
disp(cp.author)                   % "Ada"
disp(cp.revision)                 % 0
```

*Ported from python-docx v1.2.0: `src/docx/opc/coreprops.py::CoreProperties`.*

---

## `CorePropertiesPart`

**Syntax**

```matlab
part = mat2doc.opc.parts.CorePropertiesPart.default(package)
props = part.core_properties
```

**Description**

The `XmlPart` subclass corresponding to `/docProps/core.xml` ("core" is short
for Dublin Core). It parses on load and re-serializes on save through
`serialize_part_xml` (the byte-matched path), wraps a `CT_CoreProperties`
element, and exposes it through a `CoreProperties` proxy (`core_properties`). Its
constructor takes the **docx** argument order
`CorePropertiesPart(partname, content_type, element, package)` — element third,
package last. MATLAB does not inherit constructors or dispatch inherited static
methods to a subclass, so this class declares its own pass-through constructor
and its own `load` (the `PartFactory` entry point) — the faithful realization of
Python's inherited-but-`cls`-bound `XmlPart.load`.

**`default()` builds a fresh core-properties part** for a package that lacks
one, stamping the docx literals: title `"Word Document"`, last_modified_by
`"python-docx"`, revision `1`, and `modified` = the current wall-clock time
(D-coreprops-time).

:::{note}
**Both `default()` and this subclass are M1-UNREACHABLE.** (1) The bundled
`default.docx` **has** a `/docProps/core.xml`, so python-docx never builds one
from scratch — `default()` only fires for a core-less package via the (P2)
`Package.core_properties` fallback. (2) At M1 `PartFactory.part_type_for_` maps
`OPC_CORE_PROPERTIES` → base `mat2doc.opc.XmlPart` (the XmlPart-vs-Part split;
the row comment reads "P2: CorePropertiesPart"), so `core.xml` loads as a base
`XmlPart` and this subclass is not instantiated. It is ported now for P2 wiring;
the **element** registry (`cp:coreProperties` → `CT_CoreProperties`) is wired at
P1-7, the **part-class** registry is a byte-neutral P2 refinement — either
dispatch parses the same `CT_CoreProperties` root and re-serializes identical
bytes. The wall-clock `modified` (D-coreprops-time) is therefore dead for any
round-trip of a package that has a `core.xml`.
:::

**Example**

```matlab
% At P2, once OpcPackage.core_properties + the PartFactory row land:
% part = mat2doc.opc.parts.CorePropertiesPart.default(package);
% part.core_properties.title   % "Word Document"
```

*Ported from python-docx v1.2.0:
`src/docx/opc/parts/coreprops.py::CorePropertiesPart`.*

---

## `pyStr` (H14, `+mat2doc/+shared/`)

The revision setter and the string-setter path route every numeric→text
conversion through `mat2doc.shared.pyStr` — **the only permitted
numeric-serialization idiom** at an XML site (raw `num2str` / `sprintf('%g')`
there is a defect per H14 / design.md §8). `CT_CoreProperties` is the **first
numeric-serialization site in the docx port**, so P1-7 establishes the helper:
it is a faithful re-port of the Mat2Ppt `+util\pyStr.m` (identifiers rebound to
`mat2doc`, verified code-identical modulo namespaces, 20/20 vs CPython `str()`).
It matches Python's shortest-round-trip float repr (`2.0`→`"2.0"`,
`1/3`→`"0.3333333333333333"`, `1e16`→`"1e+16"`, `inf`/`-inf`/`nan`), so the
revision `ValueError` shows `got '2.5'`, not `'2'`. Full API in
[`mat2doc.shared`](shared.md).

*See `mat2doc.shared.pyStr` — Mat2Doc infrastructure (no python-docx
counterpart), mandated by design.md §8 and hazard H14.*
