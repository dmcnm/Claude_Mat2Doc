---
title: "mat2doc.opc — packaging tier: PackURI + phys_pkg + relationships"
---

# `mat2doc.opc` — packaging tier: PackURI + phys_pkg + relationships

Ported from python-docx v1.2.0 modules `src/docx/opc/packuri.py`,
`src/docx/opc/phys_pkg.py`, and `src/docx/opc/rel.py` (package
`+mat2doc/+opc/`). This is the **packaging tier** that sits *above* the P1-4 OPC
serializer (`opc_oxml.md`) and *below* the P1-6+ package/part object model: the
pack-URI (partname) algebra, the physical read/write of the zip container, and
the relationship collection that regenerates each `.rels` part from a graph of
relationships.

Three concerns:

- **`PackURI`** — a string-like value object for a pack URI (partname such as
  `/word/document.xml`), with the `baseURI` / `ext` / `filename` / `idx` /
  `membername` / `rels_uri` algebra and the `from_rel_ref` / `relative_ref`
  helpers.
- **`phys_pkg`** — the physical package layer: the `PhysPkgReader` /
  `PhysPkgWriter` factories that dispatch to a directory or zip reader / the zip
  writer, the concrete `_DirPkgReader` / `_ZipPkgReader` / `_ZipPkgWriter`, and
  the audited `java.util.zip` byte boundary (`bytesToJava` / `bytesFromJava`).
- **`rel`** — `Relationships` (a `Dict[str, _Relationship]` collection keyed by
  rId) and `Relationship_`, which together add, look up, and serialize
  relationships into `.rels` bytes.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## Where this tier sits in the M1 byte path

P1-4 (`opc_oxml.md`) proved the OPC *serializer* in isolation — hand-fed rows in
frozen order emit byte-identical `[Content_Types].xml` / `.rels`. This tier is
what will *drive* that serializer at assembly time:

- `Relationships.xml` is the **sole consumer** of the P1-4 `xml_file_bytes`
  rotation — it builds a `CT_Relationships`, appends each held relationship in
  insertion order, and returns the `.rels` file bytes.
- `PhysPkgWriter` / `_ZipPkgWriter` is the byte sink P1-6a's `PackageWriter` will
  write those `.rels` bytes (and every part blob) into.
- `PackURI` is the partname currency threaded through both — `membername` is the
  zip entry key, `rels_uri` locates a part's `.rels` companion.

The integrated full-package M1 L1 proof remains **P1-8**. P1-5 proves this tier
at the component level: `PackURI` values `probe_diff`-match the python-docx
oracle, the three `.rels` are **L1 byte-identical**, the zip writer is
byte-reproducible and byte-identical to the shipped Mat2Ppt writer, and the Java
boundary round-trips 0x00–0xFF cleanly.

## `PackURI` is a value class, not a `str` subclass (H2)

python-docx `PackURI` subclasses `str`. MATLAB `classdef` cannot subclass
`string`, so `PackURI` is a **value class** wrapping the URI text in a private
property, with explicit `string()` / `char()` conversions and `eq` / `ne` so
partnames compare and serve as lookup keys like the Python `str`. Any Python call
site that uses a `PackURI` directly *as* a `str` (dict key, formatting, path
join) maps to `string(obj)` here. Because Python `str` is immutable, the value
class is the faithful realization — copies are indistinguishable.

The constructor enforces the leading slash (`PackURI("/word/document.xml")`); a
URI that does not begin with `/` raises `mat2doc:ValueError`, mirroring the
Python `ValueError`. The internal posix-path helpers (`split` / `splitext` /
`join` / `abspath` / `relpath`) are pure-string ports of `posixpath`, matching it
exactly on the absolute `/`-separated URIs docx uses.

## The docx `idx` regex is `[1-9][0-9]*` — a leading-zero suffix is NOT an index (H12)

`PackURI.idx` returns the integer "tuple" partname index (`slide1.xml` → `1`) or
`[]` (Python `None`). The regex is docx-specific: `packuri.py:18` uses
`([a-zA-Z]+)([1-9][0-9]*)?` — the trailing numeric group's first digit is **1–9,
no leading zero**. This differs from python-pptx's `([0-9][0-9]*)?` (which
Mat2Ppt ported). The docx form was ported exactly:

| filename | `idx` |
|---|---|
| `image1.jpeg` | `1` |
| `image10.png` | `10` |
| `image01.png` | `[]` (None — leading-zero suffix does not match) |
| `media.png` | `[]` (None — no trailing integer) |

Three distinct Python `None` returns collapse to `[]` (H3 tri-state): no filename,
no regex match, and no numeric group. The value is returned as **data** — the
OOXML part number — never index-shifted (H1). The `image01.png` → `[]` case is
exactly what distinguishes the docx regex from pptx's and is pinned in the Gate-3
probe.

## `PhysPkgReader` / `PhysPkgWriter` are PUBLIC factories (H10)

In python-docx the factory bases are **public** (`PhysPkgReader`, no leading
underscore — contrast the private pptx `_PhysPkgReader`). Their Python `__new__`
dispatches on the argument and returns a *subtype instance*. A MATLAB constructor
cannot return a different class, so the dispatch is a static `factory` method (the
Mat2Ppt precedent) and the concrete readers subclass the base:

- **`PhysPkgReader.factory(pkg_file)`** (`phys_pkg.py:13-25`, faithful to docx
  which tests `str` **first**):
  - a path that is a **directory** → `DirPkgReader_`
  - a path that is a **zip file** → `ZipPkgReader_`
  - a path that is neither → raise `mat2doc:PackageNotFoundError`
  - a **stream** (our currency: `uint8` whole-zip bytes) → `ZipPkgReader_`
- **`PhysPkgWriter.factory(pkg_file)`** (`phys_pkg.py:31-32`) — **always**
  `ZipPkgWriter_`. There is no directory writer: a directory package can be read
  but not written.

The concrete readers rotate their leading underscore to trailing per design.md §2
(`_DirPkgReader` → `DirPkgReader_`, `_ZipPkgReader` → `ZipPkgReader_`,
`_ZipPkgWriter` → `ZipPkgWriter_`); the public factory bases keep the docx name.

### Per-reader error-id fidelity (H3, subtle)

docx `_DirPkgReader.blob_for` opens the member with `open(...,"rb")` — a missing
member raises **IOError**; `_ZipPkgReader.blob_for` does `zipf.read(name)` — a
missing member raises **KeyError**. These are *different* except clauses, so the
port keeps them distinct: `DirPkgReader_.blob_for` raises `mat2doc:IOError`,
`ZipPkgReader_.blob_for` raises `mat2doc:KeyError`, and each `rels_xml_for`
catches **its own** id (returning `[]` = None) and rethrows anything else.
`content_types_xml` does **not** catch — a missing content-types stream
propagates, matching Python.

## The stateful zip writer and D-zip-time

docx `_ZipPkgWriter` is **stateful**, not a batch writer: the constructor opens
the archive, `write(pack_uri, blob)` appends one DEFLATED entry per call
(membername = `pack_uri.membername`), and `close()` finalizes it (flushing the
bytes to the `pkg_file` path, and retaining them for `to_bytes()` in the stream
case). The port mirrors that interface exactly, realized over an in-memory
`java.util.zip` pass (`ByteArrayOutputStream` → `ZipOutputStream`, DEFLATED).

**D-zip-time (ADOPTED deviation, ZERO new D-number).** Python
`ZipFile.writestr` stamps each entry with the wall-clock time, so whole-file zip
bytes are not reproducible. Every entry time here is pinned to `1980-01-01 00:00`
(`GregorianCalendar(1980,0,1,0,0,0).setTime`, the DOS-time floor) so
MATLAB↔MATLAB output is byte-reproducible. This is envelope metadata discarded on
extraction — whole-file zip bytes are **out of scope** for equivalence
(part-level after unzip). The mechanism, its signed scope, and the Info-ZIP
`0x5455` extended-timestamp extra field the JVM emits alongside the DOS stamp are
described in
`validation\summary\proofs\D-zip-time_explained.md` and pre-adopted in
`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`. The
`0x5455` field is part of the *same* `setTime` mechanism (not a separate
divergence); the P1-5 writer was proven **byte-identical to Mat2Ppt's shipped,
M3-validated writer** for the same entries/order, so both emit it.

The signed `byte[]` ↔ `uint8` crossing goes **only** through `bytesToJava` /
`bytesFromJava` (`typecast`, not `cast`, so `128 → -128` bit patterns are
preserved) — the sole audited Java boundary of the OPC layer (design.md §4). They
are **re-ported** into `+mat2doc\+opc` (design.md §7 forbids sharing code between
the two toolboxes; the Mat2Ppt twin is not referenced), logically identical to
their Mat2Ppt counterparts.

## `Relationship_` does NOT cache (faithful to docx v1.2.0)

python-docx v1.2.0 `opc/rel.py::_Relationship` declares **every** accessor
(`is_external`, `reltype`, `rId`, `target_part`, `target_ref`) as a **plain
`@property`** — no `@lazyproperty`, no cache member, and **no `target_partname`**
on this class at all. `target_ref` is **recomputed on each access**. The port
mirrors that: `Relationship_` is non-caching.

This deliberately does **not** follow the Mat2Ppt #11 "relcache" precedent, which
was for pptx `package.py`'s `@lazyproperty`-based `_Relationship` — a *different*
(refactored) class in a different module. Adding a cache here would be
behavior-adding (design.md §7). Serialized bytes are identical either way
(`target_ref` is computed once per emit). The caching the P1-5 brief cited is real
— it lives one layer down in `opc/pkgreader.py::_SerializedRelationship` (P1-6a
scope), and must be placed **there**, not re-injected into `rel.py`. Full ruling:
`validation\summary\decision_2026-07-25_mat2doc_relcache_not_in_docx_rel.md`.

For an **internal** relationship `target_ref` is the target partname made relative
to the `baseURI` (the serialization form); for an **external** relationship it is
the URI verbatim, and `target_part` raises `mat2doc:ValueError`.

## `Relationships.xml` emits in INSERTION order — no rId sort (H11, docx-vs-pptx)

The `Relationships` collection subclasses `Dict[str, _Relationship]` (rel.py:13):
relationships are keyed by rId and iterated in dict **insertion order**. Its `xml`
accessor iterates `self.values()` in that order and **does not sort by rId** —
contrast the Mat2Ppt pptx `_Relationships.xml`, which sorts `<Relationship>`
elements numerically. The docx behavior was ported exactly: relationships live in
insertion-ordered parallel arrays, and the emitted `.rels` element order is the
order they were added.

The Gate-3 `rels2` fixture pins this: adding `rId3` **before** `rId1` emits
`rId3, rId1, rId2` (insertion, not numeric) — byte-identical to python-docx. Three
ordering sites all follow docx: (a) `xml` emits in insertion order; (b)
`add_relationship` preserves dict position on a key-replace; (c) `_next_rId` scans
`1..len+1` ascending, **reusing the lowest gap** (so `{rId3, rId1}` → next
`rId2`). No `containers.Map` anywhere; `related_parts` uses a MATLAB `dictionary`
(insertion-preserving) with cell-wrapped Part values.

The `xml` accessor is the sole consumer of the P1-4 `xml_file_bytes` rotation
(condition B4): Python `Relationships.xml` returns the `.rels` **file bytes** via
`CT_Relationships.xml`, but MATLAB cannot override the inherited pretty `.xml`, so
that byte member was rotated to `CT_Relationships.xml_file_bytes`. P1-6a's
`PackageWriter` calls `part.rels.xml` (this accessor), which routes through
`xml_file_bytes` → `ZipPkgWriter_.write`.

## `related_parts` currency (H3/H11, carry-forward V-1)

`related_parts` returns the `{rId → target Part}` map for **internal**
relationships only (external rels are omitted, matching `if not is_external`), as a
MATLAB `dictionary(string → cell)` — a consumer reads a part with `rp(rId){1}`
(the cell-wrapped Part handle, the established Mat2Doc map currency). Because
MATLAB `dictionary` is a **value** type, the getter returns a snapshot copy where
Python returns the live dict. docx reads it transiently (safe); a P1-6+ consumer
that *caches* it across `add_relationship` calls would not see later additions —
carried as VERIFY item V-1.

## Deviation posture (adopt-only, ZERO new D-numbers)

Gate-2 (Opus) APPROVE and Gate-3 PASS both confirmed **0 new D-numbers**. Every
divergence exercised is a recurrence of an already-adopted ruling:

- **D-zip-time** — the `1980-01-01` `setTime` stamp (and the `0x5455` UT extra
  field, same mechanism). Re-verified byte-identical to the Mat2Ppt writer;
  envelope-only, within the signed same-machine determinism scope. See
  `proofs\D-zip-time_explained.md`.
- **D-001** — the own OOXML-subset serializer: the `.rels` bytes transit
  `CT_Relationships` → the P1-4 `serialize_part_xml` path; re-exercised
  incidentally by the 474 / 629 / 579 B byte matches.
- Declaration single-quote + LF, insertion-order attributes, escaping — the
  P1-2/P1-4 L1 conventions, inherited via `xml_file_bytes`.

---

## `PackURI`

**Syntax**

```matlab
uri = mat2doc.opc.PackURI(pack_uri_str)   % pack_uri_str must begin with "/"
b   = uri.baseURI        % directory portion ("/word" for "/word/document.xml")
e   = uri.ext            % extension without the leading dot ("" if none)
f   = uri.filename       % final component ("" for "/")
i   = uri.idx            % int partname index, or [] (None) — docx [1-9][0-9]*
m   = uri.membername     % URI without the leading slash (the zip key)
r   = uri.rels_uri       % PackURI of this partname's .rels companion
ref = uri.relative_ref(baseURI)                 % relative reference from baseURI
u2  = mat2doc.opc.PackURI.from_rel_ref(baseURI, relative_ref)   % (Static)
```

**Description**

A string-like value object for a pack URI (partname). Wraps the URI text with
explicit `string()` / `char()` and `eq` / `ne` so it compares and keys like the
Python `str` it subclasses. The constructor enforces the leading slash
(`mat2doc:ValueError` otherwise). `idx` uses the docx-specific `[1-9][0-9]*`
regex — a leading-zero numeric suffix yields `[]` (None), not an index. The posix
helpers are pure-string ports of `posixpath`.

**Example**

```matlab
uri = mat2doc.opc.PackURI("/word/slides/slide1.xml");
disp(uri.baseURI)              % "/word/slides"
disp(uri.membername)           % "word/slides/slide1.xml"  (zip key)
disp(uri.idx)                  % 1
disp(string(uri.rels_uri))     % "/word/slides/_rels/slide1.xml.rels"
disp(mat2doc.opc.PackURI("/word/image01.png").idx)   % []  (leading-zero suffix)
```

*Ported from python-docx v1.2.0: `src/docx/opc/packuri.py::PackURI` (lines 15-105)*

---

## `PACKAGE_URI`

**Syntax**

```matlab
uri = mat2doc.opc.PACKAGE_URI()   % PackURI("/")
```

**Description**

The pack URI for the package pseudo-partname `/`. Mirrors the module constant
`PACKAGE_URI = PackURI("/")`. Returned fresh each call; `PackURI` is an immutable
value type, so this is indistinguishable from a shared constant.

**Example**

```matlab
disp(string(mat2doc.opc.PACKAGE_URI()))            % "/"
disp(string(mat2doc.opc.PACKAGE_URI().rels_uri))   % "/_rels/.rels"
```

*Ported from python-docx v1.2.0: `src/docx/opc/packuri.py::PACKAGE_URI` (line 108)*

---

## `CONTENT_TYPES_URI`

**Syntax**

```matlab
uri = mat2doc.opc.CONTENT_TYPES_URI()   % PackURI("/[Content_Types].xml")
```

**Description**

The pack URI of the content-types stream. Mirrors the module constant
`CONTENT_TYPES_URI = PackURI("/[Content_Types].xml")`. It is what the readers'
`content_types_xml` accessor resolves to fetch `[Content_Types].xml`.

**Example**

```matlab
disp(string(mat2doc.opc.CONTENT_TYPES_URI()))   % "/[Content_Types].xml"
```

*Ported from python-docx v1.2.0: `src/docx/opc/packuri.py::CONTENT_TYPES_URI` (line 109)*

---

## `PhysPkgReader`

**Syntax**

```matlab
reader = mat2doc.opc.PhysPkgReader.factory(pkg_file)   % (Static)
```

**Description**

Public factory base for physical package readers. `factory` dispatches
(`phys_pkg.py:13-25`, `str` tested first): a directory path → `DirPkgReader_`, a
zip path → `ZipPkgReader_`, neither → `mat2doc:PackageNotFoundError`; a stream
(`uint8` whole-zip bytes) → `ZipPkgReader_`. The concrete readers subclass this
base; the base keeps the docx public name (no underscore rotation).

**Example**

```matlab
% Build a tiny in-memory package, then dispatch a reader over its bytes.
w = mat2doc.opc.PhysPkgWriter.factory([]);   % [] -> stream (bytes only)
w.write(mat2doc.opc.CONTENT_TYPES_URI(), uint8('<Types/>'));
w.close();
reader = mat2doc.opc.PhysPkgReader.factory(w.to_bytes());
disp(class(reader))                          % "mat2doc.opc.ZipPkgReader_"
disp(char(reader.content_types_xml))         % "<Types/>"
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py::PhysPkgReader` (lines 10-25)*

---

## `PhysPkgWriter`

**Syntax**

```matlab
writer = mat2doc.opc.PhysPkgWriter.factory(pkg_file)   % (Static) -> ZipPkgWriter_
```

**Description**

Public factory base for physical package writers. `factory` **always** returns a
`ZipPkgWriter_` (`phys_pkg.py:31-32`) — there is no directory writer. `pkg_file`
is a path string (written to disk on `close()`) or `[]` for the bytes-only stream
case (retrieve via `to_bytes()`).

**Example**

```matlab
w = mat2doc.opc.PhysPkgWriter.factory([]);
disp(class(w))     % "mat2doc.opc.ZipPkgWriter_"
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py::PhysPkgWriter` (lines 28-32)*

---

## `DirPkgReader_`

**Syntax**

```matlab
reader = mat2doc.opc.DirPkgReader_(path)   % path = an expanded-package directory
blob   = reader.blob_for(pack_uri)
cts    = reader.content_types_xml
rels   = reader.rels_xml_for(source_uri)   % [] (None) if no rels item
reader.close()                              % no-op (a directory needs no closing)
```

**Description**

`PhysPkgReader` for an OPC package extracted into a directory tree. `blob_for`
resolves a `PackURI` to `<path>/<membername>` and reads it; a missing member
raises `mat2doc:IOError`, so `rels_xml_for` (which catches IOError → `[]`) matches
the Python except clause. Read-only — docx has no `_DirPkgWriter`.

**Example**

```matlab
d = fullfile(tempname);  mkdir(d);
fid = fopen(fullfile(d, "[Content_Types].xml"), "wb");
fwrite(fid, uint8('<Types/>'), "uint8");  fclose(fid);
reader = mat2doc.opc.DirPkgReader_(d);
disp(char(reader.content_types_xml))                          % "<Types/>"
disp(isempty(reader.rels_xml_for(mat2doc.opc.PACKAGE_URI()))) % 1  (no .rels -> [])
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py::_DirPkgReader` (lines 35-68)*

---

## `ZipPkgReader_`

**Syntax**

```matlab
reader = mat2doc.opc.ZipPkgReader_(pkg_file)   % path OR uint8 whole-zip bytes
blob   = reader.blob_for(pack_uri)             % mat2doc:KeyError if absent
cts    = reader.content_types_xml
rels   = reader.rels_xml_for(source_uri)       % [] (None) if no rels item
reader.close()                                  % no-op (in-memory member map)
```

**Description**

`PhysPkgReader` for a zip-file (`.docx`) package, or an in-memory `uint8`
whole-zip byte vector. On construction the archive is enumerated once into an
ordered `{membername → uint8 blob}` map via an in-memory `java.util.zip` pass
(design.md §4 realization of docx's lazy `ZipFile.read`). `blob_for` is then an
observable-equivalent lookup, raising `mat2doc:KeyError` on a missing member
(matching `zipfile`'s `read` KeyError).

**Example**

```matlab
w = mat2doc.opc.ZipPkgWriter_([]);
w.write(mat2doc.opc.PackURI("/word/document.xml"), uint8('<w:document/>'));
w.close();
z = mat2doc.opc.ZipPkgReader_(w.to_bytes());
disp(char(z.blob_for(mat2doc.opc.PackURI("/word/document.xml"))))   % "<w:document/>"
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py::_ZipPkgReader` (lines 71-101)*

---

## `ZipPkgWriter_`

**Syntax**

```matlab
w = mat2doc.opc.ZipPkgWriter_(pkg_file)   % path string, or [] (bytes only)
w.write(pack_uri, blob)                   % append one DEFLATED entry (stateful)
w.close()                                 % finalize; flush to path if a path
bytes = w.to_bytes()                      % whole-zip uint8 bytes (after close)
```

**Description**

Stateful `PhysPkgWriter` mirroring docx `_ZipPkgWriter`: the constructor opens the
archive, `write` appends one entry per call, `close` finalizes. Every entry time
is pinned to `1980-01-01 00:00` (**D-zip-time**, adopted, no new D-number) for
byte-reproducible output. Byte-parity with the shipped Mat2Ppt writer is proven
for the same entries/order. The `byte[]` ↔ `uint8` boundary goes only through
`bytesToJava`.

**Example**

```matlab
w = mat2doc.opc.ZipPkgWriter_([]);
w.write(mat2doc.opc.PackURI("/hello.txt"), uint8('hi'));
w.close();
b = w.to_bytes();
disp(class(b))                       % "uint8"
disp(char(b(1:2)))                   % "PK"  (local-file-header signature)
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py::_ZipPkgWriter` (lines 104-119)*

---

## `bytesToJava` / `bytesFromJava`

**Syntax**

```matlab
jbytes = mat2doc.opc.bytesToJava(blob)     % uint8 -> int8 (Java byte[])
blob   = mat2doc.opc.bytesFromJava(jbytes) % int8 -> uint8
```

**Description**

The sole audited Java byte boundary of the OPC layer (design.md §4). `bytesToJava`
is `typecast(blob, 'int8')` — the signed representation MATLAB hands to
`java.io` / `java.util.zip` (a direct `uint8`→java conversion **saturates** values
> 127 and corrupts binary parts). `bytesFromJava` is the inverse. `typecast` (not
`cast`) preserves the bit pattern (`128 ↔ -128`, `255 ↔ -1`). Re-ported into
`+mat2doc\+opc` (no cross-toolbox sharing, design.md §7); logically identical to
the Mat2Ppt twins.

**Example**

```matlab
j = mat2doc.opc.bytesToJava(uint8([0 127 128 255]));
disp(double(j'))                                       % 0   127  -128  -1
u = mat2doc.opc.bytesFromJava(j);
disp(isequal(u, uint8([0 127 128 255])))               % 1  (round-trip identity)
```

*Ported from python-docx v1.2.0: `src/docx/opc/phys_pkg.py` (java.util.zip boundary, design-realization per design.md §4 / Spike S3)*

---

## `Relationships`

**Syntax**

```matlab
rels = mat2doc.opc.Relationships(baseURI)
rel  = rels.add_relationship(reltype, target, rId, is_external)
rel  = rels.get_or_add(reltype, target_part)
rId  = rels.get_or_add_ext_rel(reltype, target_ref)
part = rels.part_with_reltype(reltype)   % KeyError none / ValueError >1
rp   = rels.related_parts                 % dictionary(rId -> {Part}) internal only
b    = rels.xml                           % uint8 .rels bytes (insertion order)
% dict surface: rels.contains(rId), rels.getitem(rId), rels.keys, rels.values, rels.len
```

**Description**

Collection of `Relationship_` objects with `Dict[str, _Relationship]` semantics,
keyed by rId in insertion order. Adds relationships, looks up a part by
relationship type, exposes the internal-target map, and regenerates the `.rels`
bytes. The `xml` accessor emits `<Relationship>` elements in **insertion order
(no rId sort, H11)** — the docx-vs-pptx divergence — and routes through the P1-4
`xml_file_bytes` rotation. `_next_rId` reuses the lowest gap.

**Example**

```matlab
r  = mat2doc.opc.Relationships("/");
RT = mat2doc.opc.RELATIONSHIP_TYPE;
id1 = r.get_or_add_ext_rel(RT.HYPERLINK, "http://a.example");
id2 = r.get_or_add_ext_rel(RT.HYPERLINK, "http://b.example");
id3 = r.get_or_add_ext_rel(RT.HYPERLINK, "http://a.example");   % dedup -> rId1
disp(id1 + " " + id2 + " " + id3)                 % "rId1 rId2 rId1"
disp(r.len)                                        % 2
disp(startsWith(char(r.xml), '<?xml'))             % 1  (.rels bytes)
```

*Ported from python-docx v1.2.0: `src/docx/opc/rel.py::Relationships` (lines 13-111)*

---

## `Relationship_`

**Syntax**

```matlab
rel = mat2doc.opc.Relationship_(rId, reltype, target, baseURI, external)
x   = rel.is_external
ty  = rel.reltype
id  = rel.rId
ref = rel.target_ref     % external: URI verbatim; internal: partname rel. to baseURI
p   = rel.target_part    % raises mat2doc:ValueError when external
```

**Description**

Value object for a single relationship. `target` is a Part (internal) or a string
URI (external). **NON-caching** — every accessor is a plain property and
`target_ref` is recomputed each access, faithful to docx v1.2.0 `rel.py`
(`@lazyproperty`/`target_partname` are pptx-only, and belong in `pkgreader.py` at
P1-6a; see the relcache decision doc). `handle` class for reference semantics,
never mutated after construction.

**Example**

```matlab
rel = mat2doc.opc.Relationship_("rId1", ...
    mat2doc.opc.RELATIONSHIP_TYPE.HYPERLINK, "http://example", "/", true);
disp(rel.is_external)    % 1
disp(rel.rId)            % "rId1"
disp(rel.target_ref)     % "http://example"  (URI verbatim; external)
```

*Ported from python-docx v1.2.0: `src/docx/opc/rel.py::_Relationship` (lines 114-153)*
