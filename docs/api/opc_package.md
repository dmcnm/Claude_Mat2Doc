---
title: "mat2doc.opc — package/part tier: OpcPackage + Part + PartFactory"
---

# `mat2doc.opc` — package/part tier: OpcPackage + Part + PartFactory

Ported from python-docx v1.2.0 modules `src/docx/opc/package.py` and
`src/docx/opc/part.py` (plus the registration block in `src/docx/__init__.py`,
lines 37-51); package `+mat2doc/+opc/`. This is the **package/part object model
tier** that sits *above* the P1-6a package-assembly tier (`opc_pkgrw.md` —
`PackageReader` / `PackageWriter` and the serialized value objects) and *below*
the P1-8 WordprocessingML `Document`: the live logical object graph an
`.docx` is read into and written back out of, and the content-type→class
registry that decides — part by part — whether a part's bytes are re-serialized
or kept verbatim.

Five classes across two source modules:

- **`OpcPackage`** — the main API class: `open` / `save` round-trip, the
  `iter_parts` / `iter_rels` depth-first graph traversals (returning a
  heterogeneous `Part` **object array**, not a cell), `part_related_by` /
  `relate_to` / `main_document_part` / `next_partname`, and the lazy package
  `rels`.
- **`Unmarshaller`** — the load-time graph builder: constructs every part (via
  the `PartFactory` handle), realizes the relationship graph (resolving each
  serialized target to its live `Part`), then fires `after_unmarshal`.
- **`Part`** — the base part (and the default, **verbatim-blob** part class): a
  partname, content type, blob, package back-reference, and lazy `rels`.
- **`XmlPart`** — the base part for an XML payload: parses its blob to an element
  tree on load and **re-serializes** it on `blob` (the whitespace-collapse path).
- **`PartFactory`** — the content-type→class registry: the **XmlPart-vs-Part
  split** that decides M1 whitespace-collapse.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## The 17-part M1 open→save sweep is GREEN at P1-6b (headline)

This WP joins the P1-6a reader to the P1-6a writer, so
`mat2doc.opc.OpcPackage.open(default.docx).save()` is the **first** time the
full 17-part round-trip runs end-to-end through the real logical model. Gate-3
proved it:

- **`pkgcompare` OVERALL PASS** — all 17 parts byte-identical (16 XML parts
  **L1**, the thumbnail **bin**), in the **exact frozen M1 zip-entry order**
  (L0 PASS: inventory + order + content-types + all three `.rels`), verified
  **three-way**: **MATLAB ≡ python-docx `OpcPackage.open→save` (s0008) ≡ the
  frozen `s0001` `Document().save()` reference.**
- The **Finding-1 whitespace split reproduces exactly** (§ PartFactory below):
  the 4 reserialized XML parts (`document` / `styles` 349458 B / `settings` /
  `numbering`) plus the regenerated `core.xml` collapse their template
  pretty-print whitespace to L1, while the 8 passthrough parts
  (`stylesWithEffects` 438131 B / `webSettings` / `fontTable` / `theme1` /
  `app` / `customXml/item1` / `itemProps1` / `thumbnail`) stay byte-verbatim.
- **Second save == first save** and **reopen→save == candidate**, whole-zip
  byte-identical (MATLAB↔MATLAB byte-stable under D-zip-time; idempotent).

This is the strongest possible pre-M1 evidence — the OPC open→save round-trip is
byte-proven ahead of the formal gate — **but the official M1 gate stays P1-8**:
the from-scratch `Document()` build wiring (which produces the graph rather than
reading it) is P1-8's job. Every byte of the M1 acceptance set is already
emitted correctly by the P1-6b read+write path.

## `iter_parts` returns a heterogeneous `Part` OBJECT ARRAY, never a cell (VERIFY-1b)

`OpcPackage.iter_parts` depth-first walks the relationship graph and returns
**exactly one reference to each part**, in traversal order, as a
`mat2doc.opc.Part` **object array** — a mix of `Part` and `XmlPart` elements in
a single vector, `iscell == false`. This is the load-bearing hand-off the P1-6a
`PackageWriter.write_parts_` requires: the writer indexes `parts(k)` and reads
`part.partname` / `part.blob` / `part.rels` / `part.partname.rels_uri` — none of
which work on a cell.

The heterogeneous array is possible because
`Part < handle & matlab.mixin.Heterogeneous` and `XmlPart < Part`, so a mixed
`[XmlPart, Part]` vector forms one `mat2doc.opc.Part` array (design.md §2). The
identity dedup `any(visited == part)` requires `eq` to be **Sealed** in the root
(MATLAB refuses to dispatch an unsealed method on a heterogeneous array), so
`Part` seals `eq` / `ne`, forwarding to `eq@handle` / `ne@handle`. **VERIFY-1b is
DISCHARGED** — Gate-3 confirmed a non-cell heterogeneous array of 13 reachable
parts that round-trips through the writer byte-stably.

`iter_parts` and `iter_rels` are the ported generators (Python `walk_parts` /
`walk_rels` with a shared mutable `visited=[]`), realized as **precomputed handle
arrays** with the `visited` list a **nested-function closure variable** (fresh
per top-level call, H9). No mutation-during-iteration in the Python original, so
eager materialization is unobservable.

## H5 identity — the diamond dedup (yield vs recursion)

The two traversals treat the shared `visited` set differently, faithful to
Python:

- **`iter_parts`** yields each **part** once — the `visited` set gates both the
  yield and the recursion.
- **`iter_rels`** yields **every relationship** (external ones too), and the
  `visited` set gates **recursion only** — a part reached by several rels still
  has each of those rels yielded once.

Gate-3 pinned this with a diamond graph (pkg→A, pkg→B, A→C, B→C, C shared):
`iter_parts` yields **3 parts** (C exactly once) while `iter_rels` yields **4
rels** (both A→C and B→C). And `A.rels.values(1).target_part == C` **and**
`B.rels.values(1).target_part == C` — the **same handle** from both reads (H5
element identity), with the Sealed `eq` dispatching on the heterogeneous array.

## The PartFactory XmlPart-vs-Part registry split (M1 whitespace-collapse)

`PartFactory.create(partname, content_type, reltype, blob, package)` is the
ported factory `__new__`: it consults the **part-class selector** (reltype
`IMAGE` → the image part), then the **content-type registry**, then the default
base `Part`, and returns `PartClass.load(...)`. **Why it matters for M1** (frozen
`references\m1_skeleton_target.md` **Finding 1**): on a plain open→save,
python-docx re-serializes exactly the XML content types it registers to an
`XmlPart` subclass, and keeps every other part's bytes **verbatim** through the
generic base `Part`. Because `XmlPart` parses with `remove_blank_text=True` and
re-serializes, the template's pretty-print indentation **collapses** — that is
the entire mechanism behind the M1 finding. Mat2Doc's registry must reproduce
the split so the same parts collapse and the other XML siblings stay verbatim.

| bucket | at M1 | content types | `blob` behavior |
|---|---|---|---|
| **XmlPart** (reserialize) | all 8 registered → `mat2doc.opc.XmlPart` | `OPC_CORE_PROPERTIES`, `WML_COMMENTS`, `WML_DOCUMENT_MAIN`, `WML_FOOTER`, `WML_HEADER`, `WML_NUMBERING`, `WML_SETTINGS`, `WML_STYLES` | parse + re-serialize (whitespace collapse → L1) |
| **base Part** (verbatim) | selector `IMAGE` → `Part`; everything unregistered → `Part` (default) | `stylesWithEffects`, `webSettings`, `fontTable`, `theme`, `app`, `application/xml`, `customXmlProperties`, `image/jpeg`, … | blob kept byte-verbatim (passthrough) |

**M1 stand-in.** The eight registered part classes are all `XmlPart` subclasses
in python-docx (`CorePropertiesPart`, `CommentsPart`, `DocumentPart`,
`FooterPart`, `HeaderPart`, `NumberingPart`, `SettingsPart`, `StylesPart` — each
extends `XmlPart` or `StoryPart < XmlPart`), and `ImagePart` (via the
`part_class_selector` for reltype `IMAGE`) is a plain `Part`. Those feature
subclasses are not ported until P2, so every registered content type here maps to
the **base** `mat2doc.opc.XmlPart` and the `IMAGE` selector to the **base**
`mat2doc.opc.Part`. This yields the correct M1 byte behavior — an `XmlPart`
subclass and the base `XmlPart` share `XmlPart.blob` (parse + re-serialize), and
`ImagePart` and the base `Part` share the verbatim `Part.blob`. **P2 refines**
each row to its specific subclass; because those subclasses inherit the same
`blob` unchanged, the emitted bytes are identical — only the reloaded part's
**type** changes.

The registration set (`part_class_selector` + `part_type_for` +
`default_part_type`) lives in `docx/__init__.py` 37-51, **not** in `opc/part.py`;
it is baked here as static tables mirroring those lines exactly (H10 — an
explicit switch/registry, design.md §2). The selector is looked up through
`cls_method_fn` (the `getattr(cls, "part_class_selector")` analogue), which
resolves the P1-5 VERIFY: the `cls` currency is a fully-qualified class-name
**string** (`"mat2doc.opc.PartFactory"`).

## `XmlPart` lazy-parse / on-demand reserialize (H9)

An `XmlPart` parses its blob into an `XmlElement` tree **once at `load` time**
and holds it; `element` returns that live tree, and `blob` **re-serializes it on
each access** via `mat2doc.opc.oxml.serialize_part_xml` (the P1-4 byte-critical
writer). This matches python-docx exactly: `XmlPart.load` parses eagerly
(part.py 229-232), `element` is a stored attribute, and `blob` re-serializes on
demand (part.py 220-222). Gate-3 confirmed the currency: `element()` twice
returns the same handle (parsed once, no re-parse); `blob()` is `uint8`, stable
across calls, starts with the single-quote lxml `<?xml` declaration, and reflects
a live element mutation (re-serializes, not cached). The base `Part.blob` returns
its stored bytes **verbatim** (Python `self._blob or b""`; `blob` None → empty
`uint8`, H4).

`XmlPart.blob` is a genuine **override** of `Part.blob`. Because MATLAB cannot
override a superclass **property** accessor, every Python `@property` on `Part`
(`blob`, `content_type`, `package`, `partname`, `related_parts`, `rels`) is
ported as a **zero-argument method** so `XmlPart` can override `blob` (a method).
`part.blob`, `part.partname`, `part.rels` still read naturally with no
parentheses (MATLAB invokes a zero-arg method on field access) — exactly how the
P1-6a writer consumes them.

## `save` REGENERATES — it does not copy the original zip

`OpcPackage.save` fires `before_marshal` on every part, then calls
`PackageWriter.write(pkg_file, self.rels, self.parts)`. It re-emits the package
from the **live logical model**, so `XmlPart` bodies are re-serialized and
base-`Part` bodies are kept verbatim — the same split as above. (Python `save`
evaluates `self.parts` twice; MATLAB captures the DFS result once, proven
equivalent because docx v1.2.0 never overrides `before_marshal` — only the two
no-op base defs — so the graph cannot change between the two evaluations.)

## The docx-vs-pptx `Part.__init__` arg order (blob before package)

python-docx `Part.__init__` is `(partname, content_type, blob=None,
package=None)` — **blob before package**, both defaulting to `None`.
(python-pptx used `(partname, content_type, package, blob)`.) The docx order is
preserved here, and in `Part.load`, `XmlPart`, and `PartFactory.create`.
`XmlPart.__init__(partname, content_type, element, package)` places the element
in the blob slot and calls the base ctor with **no** blob (the element is stored
separately). docx `Part.blob` is a getter-only `@property` (no setter, unlike
pptx), so there is deliberately no `set_blob`.

## Deviation posture (adopt-only, ZERO new D-numbers)

Gate-2 (Opus auditor) **APPROVE** and Gate-3 **PASS** both confirmed **0 new
D-numbers**. The 17/17 L1 result proves **zero output-visible divergence**
anywhere in this WP. Every divergence is a recurrence of an already-adopted
ruling:

- **D-001** — the own OOXML parser, inherited via `mat2doc.oxml.parse_xml` in
  `XmlPart.load`. Re-exercised (byte-proven) by the 4 reserialized-part L1
  matches.
- **D-serializer-nsdecl** — the lxml-convention serializer, inherited via
  `mat2doc.opc.oxml.serialize_part_xml` in `XmlPart.blob`. Re-exercised by the
  same L1 matches plus the regenerated `[Content_Types].xml` / `.rels`.
- **D-zip-time** — the MATLAB↔MATLAB whole-zip byte-stability of `save` (fixed
  1980 entry timestamps, the P1-5 `ZipPkgWriter_`). `pkgcompare` never compares
  container bytes (part-level after unzip), so this is envelope-only; the part
  bytes are L1. Observed in the second-save / reopen→save identities.

### Finding 1 — deferred/tracked (NOT triggered by the sweep), tracked as #60

`XmlPart.rel_ref_count_` faithfully ports part.py:246
`self._element.xpath("//@r:id")`, but the P1-3 merged `mat2doc.oxml.XmlElement`
(the parser's fallback class for unregistered tags — all tags at M1) has no
`xpath` method; only `BaseOxmlElement` does. So `Part.drop_rel` on an `XmlPart`
would raise `MATLAB:noSuchMethodOrField` before reaching its intended
`mat2doc:notYetPorted` stub. **M1-UNREACHABLE:** `drop_rel`'s only callers
(`DocumentPart.drop_header_part`, `Section` footer handling) are P2, and
`default.docx` open→save never invokes it — the 17-part sweep is clean. Routed to
**task #60** (P1-3 follow-up), with the P2 DocumentPart WP obligated to (a) add
`Relationships` `delitem`, (b) remove the stub, (c) clear the `xpath`
reachability — one P2 gate item. Not fixed here (outside the five WP files; their
translation is correct as written).

---

## `OpcPackage`

**Syntax**

```matlab
pkg   = mat2doc.opc.OpcPackage.open(pkg_file)   % (Static) path OR uint8 zip bytes
        pkg.save(out_file)                       % regenerate + write the package
parts = pkg.iter_parts()   % heterogeneous mat2doc.opc.Part object array (DFS, once each)
rels  = pkg.iter_rels()    % Relationship_ handle array (every rel, DFS)
part  = pkg.main_document_part()          % part_related_by(OFFICE_DOCUMENT)
part  = pkg.part_related_by(reltype)      % KeyError none / ValueError >1
rId   = pkg.relate_to(part, reltype)      % existing-or-new rId
uri   = pkg.next_partname("/word/header%d.xml")   % next free numeric suffix
r     = pkg.rels()                        % the package Relationships (lazy, base "/")
```

**Description**

Main API class for an OPC package. `open` reads the file, builds the instance,
and drives `Unmarshaller.unmarshal` to construct every part (via `PartFactory`)
and realize the relationship graph. A **handle** class — the package and its
parts are a live object graph shared by reference (H5), exactly like the Python
proxy model; a part reached through two relationships is the **same** handle, and
the traversals dedup on that identity. `iter_parts` returns a **heterogeneous
`Part` object array** (never a cell, VERIFY-1b); `iter_rels` yields every
relationship (external too), the `visited` set gating recursion only. `save`
**regenerates** the package from the live model (it does not copy the original
zip): it fires `before_marshal` on each part, then `PackageWriter.write` — so
`XmlPart` bodies re-serialize and base-`Part` bodies stay verbatim.

**Example**

```matlab
tpl = fullfile(fileparts(fileparts(which("mat2doc.opc.OpcPackage"))), ...
    "templates", "default.docx");
pkg = mat2doc.opc.OpcPackage.open(tpl);
parts = pkg.iter_parts();
disp(iscell(parts))                      % 0    (object array, not a cell)
disp(class(parts))                       % "mat2doc.opc.Part"  (heterogeneous root)
disp(numel(parts))                       % 13   (parts reachable at M1)
out = [tempname '.docx'];
pkg.save(out);                           % regenerate + write (17-part M1 order)
disp(isfile(out))                        % 1
```

*Ported from python-docx v1.2.0: `src/docx/opc/package.py::OpcPackage` (lines 24-179)*

---

## `Unmarshaller`

**Syntax**

```matlab
mat2doc.opc.Unmarshaller.unmarshal(pkg_reader, package, part_factory)   % (Static)
% pkg_reader  : a mat2doc.opc.PackageReader
% package     : the OpcPackage being populated
% part_factory: a function handle, e.g. @mat2doc.opc.PartFactory.create
```

**Description**

Hosts the static methods that build a package's part graph and relationships from
a `PackageReader`. `unmarshal` constructs every part (delegating to the
`part_factory` handle — the MATLAB stand-in for the Python `PartFactory` class
object), realizes the relationship graph (resolving each serialized target to its
live `Part`), then fires `after_unmarshal` on every part and on the package.
Consumes the P1-6a reader currency (`iter_sparts` / `iter_srels` struct arrays,
H9). The internal `{partname → Part}` map is an **insertion-ordered** parallel
`{keys, vals}` struct — never `containers.Map` (which sorts, H11); `parts.vals`
is a heterogeneous `Part` array iterated in insertion order.

**Example**

```matlab
% Unmarshaller is what OpcPackage.open drives internally; the public entry is
% OpcPackage.open (which wires reader + package + @PartFactory.create together).
tpl = fullfile(fileparts(fileparts(which("mat2doc.opc.OpcPackage"))), ...
    "templates", "default.docx");
reader  = mat2doc.opc.PackageReader.from_file(tpl);
package = mat2doc.opc.OpcPackage();
mat2doc.opc.Unmarshaller.unmarshal(reader, package, @mat2doc.opc.PartFactory.create);
disp(numel(package.iter_parts()))        % 13   (graph built + related)
```

*Ported from python-docx v1.2.0: `src/docx/opc/package.py::Unmarshaller` (lines 182-219)*

---

## `Part`

**Syntax**

```matlab
p  = mat2doc.opc.Part(partname, content_type, blob, package)   % blob BEFORE package (docx)
p  = mat2doc.opc.Part.load(partname, content_type, blob, package)   % (Static)
b  = p.blob()            % stored bytes VERBATIM (Python self._blob or b"")
ct = p.content_type()    % string
pn = p.partname()        % PackURI
rp = p.related_parts()   % {rId -> Part} internal targets
r  = p.rels()            % this part's Relationships (lazy, over partname.baseURI)
      p.set_partname(PackURI(...))          % TypeError if not a PackURI
rId = p.relate_to(target, reltype, is_external)   % existing-or-new rId
```

**Description**

Base class for package parts — **and** the default, **verbatim-blob** part class.
A part has a partname (`PackURI`), a content type, an optional blob (its
serialized bytes) and a back-reference to its package. The base `blob` returns the
bytes **verbatim** (the path that keeps images, thumbnails, and any unregistered
XML part byte-stable); `XmlPart` overrides it to re-serialize. A
**heterogeneous root** (`handle & matlab.mixin.Heterogeneous`) with **Sealed**
`eq` / `ne` (forwarding to `handle`) so `iter_parts` can return a mixed
`Part` / `XmlPart` object array and dedup by identity (VERIFY-1b, H5). Every
Python `@property` is a **zero-argument method** so `XmlPart` can override `blob`.
Ctor arg order is the docx `(partname, content_type, blob=None, package=None)` —
blob before package.

**Example**

```matlab
% A base Part returns its blob VERBATIM -- no re-serialization (the path that
% keeps thumbnails and passthrough XML parts byte-stable).
p = mat2doc.opc.Part( ...
    mat2doc.opc.PackURI("/word/theme/theme1.xml"), ...
    mat2doc.opc.CONTENT_TYPE.OFC_THEME, uint8('<a:theme/>'), []);
disp(char(p.blob()))         % "<a:theme/>"   (identical to the stored bytes)
disp(string(p.partname()))   % "/word/theme/theme1.xml"
```

*Ported from python-docx v1.2.0: `src/docx/opc/part.py::Part` (lines 21-162)*

---

## `XmlPart`

**Syntax**

```matlab
xp = mat2doc.opc.XmlPart(partname, content_type, element, package)   % element in blob slot
xp = mat2doc.opc.XmlPart.load(partname, content_type, blob, package)   % (Static) parses blob
b  = xp.blob()      % OVERRIDE: serialize_part_xml(element) -- re-serialized on demand
el = xp.element()   % the live XmlElement tree (parsed once at load)
```

**Description**

Base class for package parts carrying an XML payload (most of them). `load`
parses the blob into an `XmlElement` tree **once** and stores it; `element`
returns that live tree, and `blob` **re-serializes it on each access** via the
P1-4 `serialize_part_xml` (byte-matched to lxml). This is **why** open→save
regenerates content parts rather than copying their bytes: parsing with
`remove_blank_text` strips the template's pretty-print indentation, so
re-serializing collapses the whitespace — the mechanism behind the M1 finding
(see `PartFactory` and `m1_skeleton_target.md` Finding 1). `XmlPart.blob` is a
genuine method override of `Part.blob` (H9 lazy/eager currency: parsed once,
serialized on demand).

**Example**

```matlab
% An XmlPart RE-SERIALIZES its parsed element on blob (contrast the verbatim
% base Part): loading parses the XML; blob emits it afresh with the declaration.
xml = "<w:p xmlns:w='" + ...
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>";
xp  = mat2doc.opc.XmlPart.load( ...
    mat2doc.opc.PackURI("/word/document.xml"), ...
    mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, ...
    uint8(unicode2native(xml, "UTF-8")), []);
disp(class(xp))                              % "mat2doc.opc.XmlPart"
disp(startsWith(char(xp.blob()), "<?xml"))   % 1   (re-serialized, declaration prepended)
disp(isa(xp.element(), "mat2doc.oxml.XmlElement"))   % 1   (live parsed tree)
```

*Ported from python-docx v1.2.0: `src/docx/opc/part.py::XmlPart` (lines 207-247)*

---

## `PartFactory`

**Syntax**

```matlab
part = mat2doc.opc.PartFactory.create(partname, content_type, reltype, blob, package)   % (Static)
cls  = mat2doc.opc.PartFactory.part_cls_for_(content_type)   % class-name string for a CT
```

**Description**

Constructs the registered `Part` subtype for a content type / reltype. `create`
consults the **part-class selector** (reltype `IMAGE` → the image part), then the
**content-type registry** (`part_type_for_`, an ordered Nx2 table mirroring
`docx/__init__.py` 44-51), then the default base `Part`, and returns
`PartClass.load(...)`. The **M1 whitespace-collapse decider** (Finding 1): the 8
registered XML content types → `mat2doc.opc.XmlPart` (parse + re-serialize →
collapse → L1), everything else → base `mat2doc.opc.Part` (bytes verbatim). An
explicit switch/registry (H10 — no MATLAB `__new__`), dispatched through
`cls_method_fn` (the `getattr(cls, name)` analogue; `cls` is a class-name
string). Arg order is the docx `(partname, content_type, reltype, blob,
package)` — blob before package. **P2** refines each registered row to its
specific subclass; the emitted bytes are unchanged (same inherited `blob`), only
the reloaded part's type changes.

**Example**

```matlab
CT = mat2doc.opc.CONTENT_TYPE;
f  = @mat2doc.opc.PartFactory.part_cls_for_;
disp(f(CT.WML_STYLES))            % "mat2doc.opc.XmlPart"   (registered -> reserialize)
disp(f(CT.WML_DOCUMENT_MAIN))     % "mat2doc.opc.XmlPart"
disp(f(CT.WML_WEB_SETTINGS))      % "mat2doc.opc.Part"      (unregistered -> passthrough)
disp(f(CT.OFC_THEME))             % "mat2doc.opc.Part"
% create() end-to-end: a document part loads as an XmlPart, webSettings as a base Part.
xml = uint8('<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>');
d = mat2doc.opc.PartFactory.create(mat2doc.opc.PackURI("/word/document.xml"), ...
    CT.WML_DOCUMENT_MAIN, mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, xml, []);
disp(class(d))                    % "mat2doc.opc.XmlPart"
```

*Ported from python-docx v1.2.0: `src/docx/opc/part.py::PartFactory` (lines 165-204) + `src/docx/__init__.py` (lines 37-51)*
