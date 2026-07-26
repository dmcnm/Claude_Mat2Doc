---
title: "mat2doc.opc — package assembly tier: PackageReader + PackageWriter"
---

# `mat2doc.opc` — package assembly tier: PackageReader + PackageWriter

Ported from python-docx v1.2.0 modules `src/docx/opc/pkgreader.py` and
`src/docx/opc/pkgwriter.py` (package `+mat2doc/+opc/`). This is the **package
assembly tier** that sits *above* the P1-5 packaging tier (`opc_packaging.md` —
`PackURI` / `phys_pkg` / `Relationships`) and *below* the P1-6+ package/part
object model: the read path that walks a serialized `.docx` into a graph of
serialized parts and relationships, and the write path whose zip-entry traversal
**is** the byte-critical M1 order.

Two directions, seven classes:

- **Read path** — `PackageReader` (the load entry point + the `iter_sparts` /
  `iter_srels` iterators the unmarshaller consumes), `SerializedPart_`,
  `SerializedRelationship_` (the **V-3** `target_partname` lazy-cache lives here),
  `SerializedRelationships_` (`load_from_xml`), and `ContentTypeMap_` (the
  *reader-side* content-type resolver — both its maps case-insensitive).
- **Write path** — `PackageWriter` (the **M1 zip-DFS traversal order**) and
  `ContentTypesItem_` (the *writer-side* content-types composer — overrides
  case-**preserving**).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## Where this tier sits in the M1 byte path — PROVES and DEFERS

This is the **second** M1 byte-critical WP (P1-4 was the isolated serializer;
P1-5 the packaging tier). `PackageWriter`'s traversal *is* the frozen M1
zip-entry order, and `PackageReader` is the load path a round-trip re-reads. What
Gate-3 proves here, and what defers to P1-8:

- **PROVES — the writer/reader MECHANISM at full-ladder strength on a synthetic
  package.** A 5-part → 8-entry stub written by *both* languages is
  `pkgcompare` **OVERALL PASS**: L0 zip-entry order + content-types + all `.rels`
  identical, and **8/8 parts byte-identical** (`[Content_Types].xml` **773 B L1**,
  both `.rels` L1). The reader is proven at round-trip strength on the **real
  17-entry `default.docx`** — `iter_sparts` (13 parts) and `iter_srels` (13 rows)
  are `probe_diff`-exact against python-docx's own `PackageReader`, byte-for-byte
  on every blob, in DFS order.
- **DEFERS** — the authoritative **full 17-part L0 order proof** requires
  `OpcPackage.iter_parts()` + real `Part` objects, which do **not** exist yet:
  it is the **P1-8 M1 gate**. `PackageWriter` preserves whatever `parts` order it
  is handed (no sort, no reorder — proven equal to python-docx's own
  `PackageWriter` on the stub); the *derivation* of that order (the DFS of the
  relationship graph) is `iter_parts`, the caller's job, at **P1-6b / P1-8**.

## The M1 zip-DFS traversal order (`PackageWriter`)

`PackageWriter.write(pkg_file, pkg_rels, parts)` emits zip entries in this exact
order — the frozen M1 target shape:

0. **`[Content_Types].xml`** — `write_content_types_stream_` (composed from
   `parts` by `ContentTypesItem_.from_parts`).
1. **`/_rels/.rels`** — `write_pkg_rels_` (the package relationships).
2. …**N.** each part in the **given `parts` order**, and immediately after each
   part **its `.rels` item — iff the part has ≥1 relationship** (`write_parts_`).

Three rules make this byte-exact:

- **Part-then-its-own-`.rels` interleave.** A part's `.rels` is written
  *immediately after* the part, not batched at the end.
  `word/document.xml` → `word/_rels/document.xml.rels` → `word/styles.xml` …
- **No-rels parts emit no `.rels`** (`if part.rels.len > 0`, the H4 port of
  Python `if len(part.rels):`). In the stub, `styles` / `jpeg` / `png` / `core`
  produce no `.rels` item.
- **`write_parts_` never reorders or de-duplicates.** It faithfully preserves the
  caller's `parts` sequence; the DFS that *produces* that sequence — including its
  visited-dedup (a part reachable by two rel paths appears once) — is
  `iter_parts` (P1-8), not this writer.

The Gate-3 stub proved the full interleave byte-for-byte against python-docx's
own `PackageWriter`:
`[Content_Types].xml | _rels/.rels | word/document.xml |
word/_rels/document.xml.rels | word/styles.xml | docProps/thumbnail.jpeg |
word/media/IMG.PNG | docProps/Core.XML`. The integrated 17-part `default.docx`
order proof is **P1-8**.

## The reader/writer `_ContentTypeMap` ↔ `_ContentTypesItem` case asymmetry (H15)

`[Content_Types].xml` is handled by **two** classes with a *deliberate* case
asymmetry — faithful to docx, not an accident, and verified not-unified at Gate-2:

| side | class | overrides map | defaults map |
|---|---|---|---|
| **reader** | `ContentTypeMap_` (pkgreader) | `CaseInsensitiveDict` (partname CI) | `CaseInsensitiveDict` (extension CI) |
| **writer** | `ContentTypesItem_` (pkgwriter) | plain `dictionary` (partname case-**preserved**) | `CaseInsensitiveDict` (extension folded lower) |

So on the **read** side an override partname lookup is case-**insensitive**
(`/WORD/DOCUMENT.XML` still resolves to the `/word/document.xml` override), while
on the **write** side an override `PartName` attribute is written with the raw
part's case **preserved** (`/docProps/Core.XML` emits `PartName="/docProps/Core.XML"`).
Both sides fold the *extension* of a `<Default>` to lowercase (`IMG.PNG` →
`Extension="png"`). The 773 B `[Content_Types].xml` L1 match pins the writer side;
the s0007 precedence probe pins the reader side.

`ContentTypesItem_` also applies **sort-at-emit**, not sort-at-insertion:
`element_` builds a fresh `<Types>` with `<Default>` children sorted by extension
**then** `<Override>` children sorted by partname (MATLAB `sort` == Python
`sorted()` for the all-ASCII OOXML partname corpus, H11).

## Override-beats-Default precedence (`ContentTypeMap_`, the reader)

`ContentTypeMap_.getitem(partname)` resolves a content type with **Override
(exact partname) beating Default (by extension)**, exactly as pkgreader.py 95-105:

1. an exact override partname match wins (case-insensitive);
2. else the partname's extension matches a default (case-insensitive);
3. else raise `mat2doc:KeyError` — message verbatim to Python
   (`no content type for partname '…' in [Content_Types].xml`).

The key **must** be a `PackURI` (Python `isinstance` guard); a non-`PackURI` key
raises `mat2doc:KeyError` — the identifier is exact and the *kind* matches Python,
but the message spells the type as MATLAB `string` where Python interpolates
`<class 'str'>`. This is **VERIFY-3**, an accepted **unreachable-diagnostic**
divergence (library callers always pass a `PackURI`) — **no D-number**.

## The V-3 `target_partname` lazy-cache lives HERE (not in `rel.py`)

`SerializedRelationship_.target_partname` is the **only** relationship
`target_partname` cache in python-docx v1.2.0 — `pkgreader.py` 211-227, guarded by
`hasattr(self, "_target_partname")`. The P1-5 `Relationship_` (`rel.py`) is
**non-caching** and has **no** `target_partname` at all; the P1-5 brief's "relcache"
note pointed one layer too high. This WP places the cache **here**, faithfully:

- a private cache property `target_partname_` plus a **logical**
  `has_target_partname_` flag as the `hasattr` sentinel — **never** `isempty`
  (design.md §2 lazyproperty pattern; an empty value could be a legitimate cache
  hit);
- for an **internal** relationship, `PackURI.from_rel_ref(baseURI, target_ref)`
  computed once (posix `..`-traversal: `/word` + `../docProps/core.xml` →
  `/docProps/core.xml`), then returned unchanged on repeat access;
- for an **external** relationship, `target_partname` raises `mat2doc:ValueError`
  (message verbatim: `… undefined where TargetMode == "External"`).

The full ruling is
`validation\summary\decision_2026-07-25_mat2doc_relcache_not_in_docx_rel.md`
(V-3 discharged: cache placed in the pkgreader port, `rel.py` untouched).

## Iterator currency — generators become struct arrays (H9)

The three Python generators (`iter_sparts`, `iter_srels`, `_walk_phys_parts`)
become **precomputed struct arrays** (design.md §2 — a generator with no
mutation-during-iteration materializes to an array):

- `iter_sparts()` → `1×N` struct, fields `partname` (PackURI), `content_type`
  (string), `reltype` (string), `blob` (uint8).
- `iter_srels()` → `1×M` struct, fields `source_uri` (PackURI), `srel`
  (`SerializedRelationship_`) — package rels first (`source_uri` = `PACKAGE_URI`),
  then each part's rels.

A consumer mirrors Python tuple-unpacking with
`for s = reader.iter_sparts()` and reads the named fields. The DFS's shared
mutable `visited_partnames` list is replicated by **threading** `visited` as an
in/out argument through the recursion, so a part reachable by two paths is emitted
exactly once, globally. **VERIFY-1b (P1-6b hand-off):** `write_parts_` and
`from_parts` index `parts(k)` and read `.partname` / `.content_type` / `.blob` /
`.rels`, so the P1-6b `iter_parts` result **must be an object array over a common
`Part` base class**, not a cell array.

## Deviation posture (adopt-only, ZERO new D-numbers)

Gate-2 (Fable) APPROVE — zero defects, zero inline fixes — and Gate-3 PASS both
confirmed **0 new D-numbers**. Every divergence is a recurrence of an
already-adopted ruling:

- **D-zip-time** — the s0005 package is written by the P1-5 `ZipPkgWriter_`
  (1980-01-01 `setTime` + the Info-ZIP `0x5455` UT extra field, same mechanism).
  `pkgcompare` never compares container/zip bytes (part-level after unzip), so this
  is envelope-only; the part bytes are L1. See `proofs\D-zip-time_explained.md`.
- **D-001 / D-serializer-nsdecl** — the `[Content_Types].xml` and `.rels` bytes
  transit the P1-4 `serialize_part_xml` / `parse_xml` path; re-exercised
  incidentally by the 773 / 298 / 283 B L1 matches. This WP adds **no**
  serialization or parsing logic of its own.
- **VERIFY-3** — the non-`PackURI` `KeyError` message spelling (above); accepted
  unreachable-diagnostic, no D-number.

---

## `PackageReader`

**Syntax**

```matlab
reader = mat2doc.opc.PackageReader.from_file(pkg_file)   % (Static) path OR uint8 bytes
sp     = reader.iter_sparts()   % 1xN struct {partname, content_type, reltype, blob}
sr     = reader.iter_srels()    % 1xM struct {source_uri, srel}
% low-level constructor (used by from_file; content_types is accepted-and-ignored):
reader = mat2doc.opc.PackageReader(content_types, pkg_srels, sparts)
```

**Description**

Low-level, read-only access to a serialized OPC (`.docx`) package. `from_file`
opens the physical package, resolves the content-type map, DFS-walks the
relationship graph from the package rels, and collects a `SerializedPart_` per
reached part. `iter_sparts` / `iter_srels` expose the results as **struct arrays**
(H9) — package rels come first in `iter_srels`. The DFS emits each reachable part
exactly once (visited-dedup, threaded across the recursion). Faithful to Python,
the constructor **accepts but does not store** `content_types` (used only during
`from_file`).

**Example**

```matlab
tpl = fullfile(fileparts(which("mat2doc.opc.PackageReader")), "..", ...
    "templates", "default.docx");
reader = mat2doc.opc.PackageReader.from_file(tpl);
sp = reader.iter_sparts();               % 1xN struct
disp(numel(sp))                          % 13   (parts reachable at M1)
disp(sp(1).content_type)                 % content type of the first DFS-reached part
sr = reader.iter_srels();                % 1xM struct
disp(string(sr(1).source_uri))           % "/"  (package rels emitted first)
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgreader.py::PackageReader` (lines 10-83)*

---

## `PackageWriter`

**Syntax**

```matlab
mat2doc.opc.PackageWriter.write(pkg_file, pkg_rels, parts)   % (Static)
% pkg_file : path (string/char), or [] for the in-memory stream currency
% pkg_rels : a mat2doc.opc.Relationships (the package-level relationships)
% parts    : an OBJECT array of parts, each with .partname/.content_type/.blob/.rels
```

**Description**

Writes a zip-format OPC package. `write` is **static** (the class is never
instantiated). It emits, in the byte-critical M1 order: (0) `[Content_Types].xml`
composed from `parts`, (1) `/_rels/.rels` from `pkg_rels`, then (2..N) each part
in the **given order** followed immediately by its `.rels` item **iff the part has
≥1 relationship** (`if part.rels.len > 0`). `write_parts_` **preserves** the
caller's order — it never sorts, reorders, or de-duplicates (that DFS is
`iter_parts`, P1-8). Every value handed to the physical writer is `uint8` bytes,
so no char/string ever reaches the zip layer (H2). The `parts` sequence must be an
object array over a common `Part` base class (VERIFY-1b).

**Example**

```matlab
% Skeleton package: no parts yet (real Part objects arrive at P1-6b). This writes
% the two package-map entries -- [Content_Types].xml then /_rels/.rels -- in order.
tmp = [tempname '.docx'];
pkg_rels = mat2doc.opc.Relationships("/");                 % empty package rels
mat2doc.opc.PackageWriter.write(tmp, pkg_rels, ...
    mat2doc.opc.SerializedPart_.empty(1, 0));              % no parts
r = mat2doc.opc.PhysPkgReader.factory(tmp);
disp(startsWith(char(r.content_types_xml), '<?xml'))       % 1  ([Content_Types].xml emitted)
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgwriter.py::PackageWriter` (lines 22-59)*

---

## `SerializedPart_`

**Syntax**

```matlab
sp = mat2doc.opc.SerializedPart_(partname, content_type, reltype, blob, srels)
pn = sp.partname       % PackURI
ct = sp.content_type   % string
b  = sp.blob           % uint8 part bytes
rt = sp.reltype        % string (the referring relationship type)
sr = sp.srels          % SerializedRelationships_
```

**Description**

Value object for one package part in **serialized** form — targets are referred to
by partname, not by an in-memory `Part`. Read-only access to the partname, content
type, referring relationship type, blob, and serialized relationships. A `handle`
class (the Python original is a plain object with identity; `PackageReader`
collects and iterates them by reference). Underscore rotation
`_SerializedPart` → `SerializedPart_` (design.md §2).

**Example**

```matlab
sp = mat2doc.opc.SerializedPart_( ...
    mat2doc.opc.PackURI("/word/document.xml"), ...
    "application/xml", ...
    mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, ...
    uint8('<w:document/>'), ...
    mat2doc.opc.SerializedRelationships_());
disp(string(sp.partname))    % "/word/document.xml"
disp(sp.content_type)        % "application/xml"
disp(char(sp.blob))          % "<w:document/>"
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgreader.py::_SerializedPart` (lines 130-164)*

---

## `SerializedRelationship_`

**Syntax**

```matlab
srel = mat2doc.opc.SerializedRelationship_(baseURI, rel_elm)   % rel_elm = a CT_Relationship
x    = srel.is_external
rt   = srel.reltype
id   = srel.rId
tm   = srel.target_mode
ref  = srel.target_ref       % the raw Target attribute (relative for internal)
pn   = srel.target_partname   % PackURI (internal, lazy-cached); ValueError if external
```

**Description**

Value object for a serialized relationship — the target is a **partname**, not an
in-memory `Part`. Built from the source part's `baseURI` and a parsed
`<Relationship>` element. `target_partname` is the **V-3 lazy-cache**: for an
internal rel it computes `PackURI.from_rel_ref(baseURI, target_ref)` once
(`hasattr`-guarded by a logical flag, never `isempty`) and returns the same
`PackURI` on repeat access; for an external rel it raises `mat2doc:ValueError`
(message verbatim). This is the **only** relationship class that caches its target
partname in docx v1.2.0 — the `rel.py` `Relationship_` does not (see the relcache
decision doc). A `handle` class (the cache needs reference semantics).

**Example**

```matlab
rel_elm = mat2doc.opc.oxml.CT_Relationship.new("rId1", ...
    mat2doc.opc.RELATIONSHIP_TYPE.OFFICE_DOCUMENT, "document.xml");
srel = mat2doc.opc.SerializedRelationship_("/", rel_elm);
disp(srel.rId)                       % "rId1"
disp(srel.is_external)               % 0
disp(string(srel.target_partname))   % "/document.xml"  (baseURI "/" + Target, cached)
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgreader.py::_SerializedRelationship` (lines 167-227)*

---

## `SerializedRelationships_`

**Syntax**

```matlab
srels = mat2doc.opc.SerializedRelationships_.load_from_xml(baseURI, rels_item_xml)  % (Static)
arr   = srels.to_array()   % 1xN SerializedRelationship_, document (append) order
```

**Description**

Read-only sequence of `SerializedRelationship_` objects, corresponding to one
`.rels` part. `load_from_xml` parses the `.rels` bytes and appends a
`SerializedRelationship_` per `<Relationship>`; when `rels_item_xml` is `[]`
(Python `None` — the `.rels` item is absent) it returns an **empty** collection,
not an error (H3). `to_array` is the `__iter__` mapping (design.md §2:
`for x in srels` → `for x = srels.to_array()`), returning the array in document
order. **V-3:** this loader lives here (`pkgreader.py`), not in `rel.py`. A
`handle` class.

**Example**

```matlab
xml = "<Relationships xmlns=""http://schemas.openxmlformats.org/package/2006/relationships"">" + ...
      "<Relationship Id=""rId1"" Type=""http://x/styles"" Target=""styles.xml""/>" + ...
      "<Relationship Id=""rId2"" Type=""http://x/img"" Target=""http://ext"" TargetMode=""External""/>" + ...
      "</Relationships>";
srels = mat2doc.opc.SerializedRelationships_.load_from_xml("/word", uint8(char(xml)));
arr = srels.to_array();
disp(numel(arr))                        % 2
disp(string(arr(1).target_partname))    % "/word/styles.xml"  (internal, joined + cached)
disp(arr(2).is_external)                % 1
disp(numel(mat2doc.opc.SerializedRelationships_.load_from_xml("/word", []).to_array()))   % 0
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgreader.py::_SerializedRelationships` (lines 230-254)*

---

## `ContentTypeMap_`

**Syntax**

```matlab
ctmap = mat2doc.opc.ContentTypeMap_.from_xml(content_types_xml)   % (Static) uint8 bytes
ct    = ctmap.getitem(partname)   % content type for a PackURI; mat2doc:KeyError on a miss
```

**Description**

The **reader-side** dictionary of content type by partname, built from
`[Content_Types].xml`. `getitem` resolves with **Override (exact partname) beats
Default (by extension)**; both maps are `CaseInsensitiveDict`, so an override
partname match *and* a default extension match are case-**insensitive**. A miss on
both raises `mat2doc:KeyError` (message verbatim to Python). The key must be a
`PackURI` (a non-`PackURI` key raises `mat2doc:KeyError` — VERIFY-3, accepted
unreachable-diagnostic). Contrast the writer-side `ContentTypesItem_`, whose
overrides map is case-**preserving**.

**Example**

```matlab
t = mat2doc.opc.oxml.CT_Types.new();
t.add_default("xml", mat2doc.opc.CONTENT_TYPE.XML);
t.add_override("/word/document.xml", mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN);
ctmap = mat2doc.opc.ContentTypeMap_.from_xml(mat2doc.opc.oxml.serialize_part_xml(t));
disp(ctmap.getitem(mat2doc.opc.PackURI("/word/document.xml")))   % …document.main+xml (Override)
disp(ctmap.getitem(mat2doc.opc.PackURI("/WORD/DOCUMENT.XML")))   % same (Override, case-insensitive)
disp(ctmap.getitem(mat2doc.opc.PackURI("/a/b.xml")))             % "application/xml" (xml Default)
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgreader.py::_ContentTypeMap` (lines 86-127)*

---

## `ContentTypesItem_`

**Syntax**

```matlab
cti = mat2doc.opc.ContentTypesItem_.from_parts(parts)   % (Static) object array of parts
b   = cti.blob                                          % uint8 [Content_Types].xml bytes
```

**Description**

The **writer-side** composer of `[Content_Types].xml` from a list of parts. Its
single public entry is the static `from_parts`, which pre-seeds the `rels` and
`xml` defaults, then classifies each part: a part is a `<Default>` **iff**
`(ext.lower(), content_type)` is a **row** of `default_content_types` (the
duplicate-key pair-list — both fields must match), otherwise an `<Override>`.
`_defaults` is a `CaseInsensitiveDict` (extension folded lower); `_overrides` is a
plain `dictionary` keyed by the **case-preserved** partname (the reader/writer
asymmetry). `blob` serializes a fresh `<Types>` with `<Default>` children **sorted
by extension** then `<Override>` children **sorted by partname** (sort-at-emit,
H11). A `handle` class.

**Example**

```matlab
% A part exposing .partname (PackURI) + .content_type is all from_parts reads.
sp = mat2doc.opc.SerializedPart_( ...
    mat2doc.opc.PackURI("/word/document.xml"), ...
    mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, "", uint8([]), ...
    mat2doc.opc.SerializedRelationships_());
cti = mat2doc.opc.ContentTypesItem_.from_parts(sp);
b = cti.blob;
disp(class(b))                                  % "uint8"
disp(contains(char(b), "document.main+xml"))    % 1  (document.xml -> <Override>)
disp(contains(char(b), "Extension=""rels"""))   % 1  (rels default pre-seeded)
```

*Ported from python-docx v1.2.0: `src/docx/opc/pkgwriter.py::_ContentTypesItem` (lines 62-115)*
