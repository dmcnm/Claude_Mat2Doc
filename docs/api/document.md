---
title: "mat2doc — the WordprocessingML entry: Document + Package + DocumentPart"
---

# WordprocessingML entry — `mat2doc.Document` + `Package` + `DocumentPart` + `document.Document`

Ported from python-docx v1.2.0 modules `src/docx/api.py`, `src/docx/package.py`,
`src/docx/parts/document.py`, and `src/docx/document.py` (plus two byte-neutral
row flips in the `src/docx/__init__.py` PartFactory registration block); packages
`+mat2doc/`, `+mat2doc/+package/`, `+mat2doc/+parts/`, `+mat2doc/+document/`.
This is the **docx-level entry tier** — the public `mat2doc.Document()` function
and the thin walking-skeleton object graph beneath it — that sits *above* the
OPC package/part tier (`opc_package.md`): the layer a user actually calls to open
a `.docx` and save it back out.

The tier is the **M1 milestone WP** (P1-8). It wires four new symbols over
machinery that already emits every M1 part byte-identical (proven at P1-6b), so
the byte stream does not change — what P1-8 adds is the *public entry* and the
*live object graph* the entry traverses:

- **`mat2doc.Document`** (`+mat2doc/Document.m`) — the public entry **function**
  (`api.py::Document`): default-template resolution, `Package.open`, the
  content-type guard, and `return document_part.document`. Ported **in full**;
  api.py is finished at P1-8.
- **`mat2doc.package.Package`** (`+mat2doc/+package/Package.m`) — the docx
  `OpcPackage` subclass: its **own static `open`**, live `after_unmarshal` →
  image-part gathering, and the (P7) image-part stubs.
- **`mat2doc.parts.DocumentPart`** (`+mat2doc/+parts/DocumentPart.m`) — the
  minimal main-document part: **own static `load`**, live `document` / `save` /
  `core_properties`, every feature accessor stubbed until P2.
- **`mat2doc.document.Document`** (`+mat2doc/+document/Document.m`) — the minimal
  document **proxy** class: live `save` / `core_properties` / `part`, every
  content member stubbed until P2.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## ★ M1 MILESTONE — ACHIEVED (headline)

**M1 is the first end-to-end Mat2Doc acceptance milestone: `mat2doc.Document()`
opens the bundled default template and `.save()` writes it back byte-identical to
python-docx AND the result opens clean in real Microsoft Word.** At P1-8 both
legs are GREEN — **M1 is ACHIEVED.**

**The acceptance bar (both legs required):**

| Leg | Bar | Result |
|---|---|---|
| **Byte leg** (Gate-3) | `mat2doc.Document().save(tmp)` (from a foreign cwd) → `pkgcompare` vs the frozen 17-part acceptance set `references\s0001`: **L0** (17-part inventory in the exact zip-DFS order, 3 Default + 11 Override, all three `.rels` order-sensitive) + **16 XML parts L1 byte-identical** + `docProps/thumbnail.jpeg` **bin byte-identical** = **17/17**, verified **three-way** `MATLAB ≡ python-docx Document().save ≡ frozen s0001` | **PASS — 17/17 L1 three-way** |
| **Word COM leg** (mso-office-verifier) | The M1 round-trip file opens in real Word with **no error, no repair/recovery prompt, zero dropped content**, and survives an **open-edit-save round-trip** (Word re-emits its own `.docx` and reopens it clean); a repair prompt = FAIL regardless of bytes | **PASS — open-clean + edit-save round-trip, no repair** |

Both flipped parts (`docProps/core.xml`, `word/document.xml`) are byte-identical
to the P1-6b **pre-flip** baseline (`s0008`): the row flips changed the reloaded
part **type**, not one emitted byte. Zero `mat2doc:notYetPorted` stub fires on a
clean `Document().save`. **Deviations: ZERO new D-numbers** — the round-trip
re-exercises only already-signed adopt-only rulings (D-001 own parser, byte-proven
17/17; D-serializer-nsdecl / D-coreprops-time both unreachable at M1 because
`core.xml` is present and read+reserialized, not created; D-zip-time is
container-only and never compared part-level).

### The walking skeleton → P2 roadmap

P1-8 is deliberately a **walking skeleton** — the thinnest object graph that
satisfies the M1 public entry. It ports the open→save spine and stubs every
feature member (each `mat2doc:notYetPorted` stub names its un-stub WP). Nothing
P1-8 writes is thrown away; every P2 change is a **refinement on top**
(superclass insertion, further row flips, stub removal), each byte-neutral by
construction:

- **P2-1 — `ElementProxy` / `Parented` base.** Supplies the proxy base classes
  and retrofits `mat2doc.document.Document` onto the real `ElementProxy` base
  (**VERIFY-M1-DOC-BASE**). Mat2Ppt precedent: WP8 reparented `Presentation`
  onto `PartElementProxy` with byte-neutral effect. Until then `Document` derives
  `handle` only and stores `element_`/`part_` directly.
- **P2-2 — `StoryPart` + real sibling parts.** Inserts `BaseStoryPart` /
  `StoryPart` **above** `DocumentPart` (**VERIFY-M1-DOCPART-BASE**), ports the
  real `StylesPart` / `SettingsPart` / `NumberingPart` (and flips their
  PartFactory rows — byte-identical output, only the reloaded type changes),
  un-stubs the DocumentPart sibling accessors, discharges **task #60** (the
  `drop_rel` → `rel_ref_count_` `xpath` reachability gap), and adds
  `Relationships` `delitem` (`drop_rel`). Because those bases add methods only
  (blob/load/element inherit from `XmlPart` unchanged), the reparented classes
  emit identical bytes.
- **P2-3 — `Document` feature shell.** Fills the content members
  (`add_paragraph` / `add_heading` / `add_page_break` / `paragraphs` /
  `iter_inner_content`, the `_Body` block-container, and stream-save hardening).

## The walking skeleton — what is LIVE vs STUBBED at M1

**LIVE (the open→save spine):** `mat2doc.Document` (entry) → `Package.open`
(own static) → `main_document_part` (base `OpcPackage`) → content-type guard →
`DocumentPart.document` → `Document.save` → `DocumentPart.save` → `Package.save`
(base `OpcPackage`) → `PackageWriter` (P1-6a). `after_unmarshal` →
`gather_image_parts_` runs **live on every open** (a no-op on `default.docx`,
which has no internal `RT.IMAGE` relationship). `core_properties` is live on both
`DocumentPart` and `Document` (P1-7 made the core-properties path real).

**STUBBED (`mat2doc:notYetPorted`, feature-only — never on open/save):** every
other public member. Each stub message names its un-stub owner WP:

| Class | Stubbed members | Un-stub WP |
|---|---|---|
| `Package` | `get_or_add_image_part`, `image_parts` | **P7** image tier |
| `DocumentPart` | `styles` / `settings` | **P2-2** |
| `DocumentPart` | `numbering_part` | **P2-2** |
| `DocumentPart` | `comments`, header/footer part accessors + `drop_header_part`, `inline_shapes`, `get_style` / `get_style_id` | **P2** feature tiers |
| `Document` | `add_paragraph` / `add_heading` / `add_page_break` / `iter_inner_content` / `paragraphs` | **P2-3** |
| `Document` | `add_table` / `tables`, `add_section` / `sections`, `add_comment` / `comments`, `inline_shapes`, `settings` / `styles` | **P2** feature tiers |
| `Document` | `add_picture` | **P7** image tier |

**Zero-stubs-on-save is the mechanical proof.** A clean `mat2doc.Document().save`
completes with no exception; had any stub been on the open/save path, MATLAB would
have raised `mat2doc:notYetPorted`. (Opening an *image-bearing* docx would reach
`image_parts` via `gather_image_parts_` and raise the P7 stub — that is correct
not-yet-ported behavior; it never happens for any M1 input.)

## The two byte-neutral PartFactory row flips

P1-8 flips exactly two rows of `PartFactory.part_type_for_` (the ordered Nx2
table mirroring `docx/__init__.py` 44-51), from the P1-6b base-`XmlPart` M1
stand-in to their real subclasses:

| content type | before (M1 stand-in) | after (P1-8) | why byte-neutral |
|---|---|---|---|
| `WML_DOCUMENT_MAIN` | `mat2doc.opc.XmlPart` | `mat2doc.parts.DocumentPart` | `DocumentPart` inherits `XmlPart.blob` (parse + `serialize_part_xml`) unchanged; its own static `load` constructs the subclass. **Required for M1** — `mat2doc.Document` needs `main_document_part.document`. |
| `OPC_CORE_PROPERTIES` | `mat2doc.opc.XmlPart` | `mat2doc.opc.parts.CorePropertiesPart` | `CorePropertiesPart` (P1-7) inherits `XmlPart.blob` unchanged; its own static `load`. Closes the P1-6b **VERIFY-core-props** obligation at zero byte-risk. |

Because both subclasses IS-A `XmlPart` and inherit `blob` unchanged, the emitted
bytes are identical — only the reloaded part's **type** changes. Gate-3 proved it:
the saved `core.xml` / `document.xml` are byte-identical to the pre-flip `s0008`
baseline. The other six rows stay base `XmlPart` until **P2-2** refines them the
same way.

## Deviation posture (adopt-only, ZERO new D-numbers)

Gate-2 (Fable, cross-model) **APPROVE** and Gate-3 **PASS** both confirmed **0 new
D-numbers**, and the Word COM leg confirms the file for the first time in real
Office. The 17/17 L1 result proves zero output-visible divergence. Every
divergence is a recurrence of an already-signed ruling:

- **D-001** — the own OOXML parser, inherited via `mat2doc.oxml.parse_xml` in
  `DocumentPart.load`. Re-exercised (byte-proven) by the 17/17 L1 sweep.
- **D-serializer-nsdecl** — **unreachable at M1**: no element is created or
  mutated on a plain open→save; every part is read and re-serialized verbatim.
- **D-zip-time** — container-level 1980 entry timestamps only; `pkgcompare` never
  compares whole-zip bytes (part-level after unzip), and the Word COM leg is the
  gate that accepts the envelope.
- **D-coreprops-time** — **unreachable**: `default.docx` HAS a `core.xml`, so
  `CorePropertiesPart.default()`'s wall-clock stamp never runs.

## Gate-2 fix folded in (F1 — H3 `''`/`""` is NOT `None`)

The public entry's None test uses the **strict** double-`0×0` sentinel
(`strcmp(class(docx),'double') && isequal(size(docx),[0,0])`), not a bare
`isequal(docx,[])`. A bare test is TRUE for the `0×0` char `''`, so
`mat2doc.Document('')` would silently open the default template — whereas
python-docx (`docx='' is not None`) treats `''` as a real path and raises
`PackageNotFoundError`. This is the ratified None-idiom for public entries
(`decision_2026-07-26_mat2doc_none_idiom.md`: inline `isequal(x,[])` internally,
**strict form at any public entry that accepts a user path**). Post-fix,
`Document('')` and `Document("")` raise `mat2doc:PackageNotFoundError` (matching
Python's exception class, message byte-verbatim), while `Document([])` and the
no-arg form still open the template.

---

## `mat2doc.Document` (entry function)

**Syntax**

```matlab
d = mat2doc.Document()        % open the bundled default template
d = mat2doc.Document(docx)    % open the .docx at path `docx` (string)
```

**Description**

The public entry point: opens a `.docx` and returns a `mat2doc.document.Document`
proxy. With no argument (or the `[]` None sentinel) it opens the bundled
`+mat2doc\templates\default.docx`, resolved **package-relative** via
`fileparts(mfilename('fullpath'))` so it works from any working directory (the
`os.path.split(__file__)[0]` analogue). It opens through the docx-level
`mat2doc.package.Package` subclass so every part receives a `Package`
back-reference (and `document_part.core_properties` resolves), takes the
`main_document_part`, **rejects a non-Word main part** with a verbatim
`mat2doc:ValueError` (`file '%s' is not a Word file, content type is '%s'`), and
returns `document_part.document`. The None test is the **strict** public-entry
form (F1): an empty char/string `''`/`""` is a real path, not None, and raises
`mat2doc:PackageNotFoundError` — matching python-docx (H3). This file
`+mat2doc\Document.m` is the entry **function**; the document **proxy class** is
the fully-qualified `mat2doc.document.Document` (the two names do not collide,
mirroring Python `docx.api.Document` vs `docx.document.Document`).

**Example**

```matlab
d = mat2doc.Document();          % open the default template
out = [tempname '.docx'];
d.save(out);                     % write it back (17-part M1 round-trip)
disp(isfile(out))                % 1
disp(class(d))                   % "mat2doc.document.Document"
% d = mat2doc.Document("report.docx");   % open an existing file
```

*Ported from python-docx v1.2.0: `src/docx/api.py::Document` (lines 19-37)*

---

## `mat2doc.package.Package`

**Syntax**

```matlab
pkg = mat2doc.package.Package.open(pkg_file)   % (Static) own override; builds a Package
        pkg.after_unmarshal()                  % LIVE: gathers image parts (no-op on default.docx)
ip  = pkg.image_parts()                        % (STUB, P7) ImageParts collection
part = pkg.get_or_add_image_part(descriptor)   % (STUB, P7)
```

**Description**

The WordprocessingML customization of `mat2doc.opc.OpcPackage`. A **handle**
class (one live object graph shared by reference). Its **own static `open`**
overrides the inherited `OpcPackage.open` — MATLAB statics are not polymorphic in
`cls`, so without the override the base would construct a base `OpcPackage` and
parts would receive the wrong package back-reference; the override constructs a
`mat2doc.package.Package` (the faithful realization of Python `Package.open` =
inherited `OpcPackage.open` with `cls` bound to `docx.package.Package`).
`after_unmarshal` overrides the base no-op (without forwarding to super, matching
Python) to run `gather_image_parts_` — **live on every open**, a no-op on
`default.docx` (no internal `RT.IMAGE` relationship; the reltype guard
short-circuits before `image_parts` is evaluated). `get_or_add_image_part` and
`image_parts` require the **P7** image tier (`docx.image.image.Image`,
`docx.parts.image.ImagePart`, the `ImageParts` collection) and are stubbed;
neither is reachable on a plain open→save of `default.docx`.

**Example**

```matlab
tpl = fullfile(fileparts(fileparts(which("mat2doc.package.Package"))), ...
    "templates", "default.docx");
pkg = mat2doc.package.Package.open(tpl);
disp(class(pkg))                        % "mat2doc.package.Package"
disp(class(pkg.main_document_part()))   % "mat2doc.parts.DocumentPart"
out = [tempname '.docx'];
pkg.save(out);                          % regenerate + write (17-part M1 order)
disp(isfile(out))                       % 1
```

*Ported from python-docx v1.2.0: `src/docx/package.py::Package` (lines 15-47)*

---

## `mat2doc.parts.DocumentPart`

**Syntax**

```matlab
dp = mat2doc.parts.DocumentPart(partname, content_type, element, package)   % element BEFORE package (docx)
dp = mat2doc.parts.DocumentPart.load(partname, content_type, blob, package) % (Static) own override; parses blob
d  = dp.document()          % FRESH mat2doc.document.Document each access (non-cached)
     dp.save(path)          % -> package.save(path)
cp = dp.core_properties()   % -> package.core_properties() (LIVE, P1-7)
% styles/settings/numbering_part/comments/inline_shapes, header & footer parts,
% get_style/get_style_id, drop_header_part -> (STUB, P2/P2-2)
```

**Description**

The main-document part (`/word/document.xml`) — the walking-skeleton entry part.
It subclasses `mat2doc.opc.XmlPart`, so it **parses on load and re-serializes on
save** through `serialize_part_xml` (byte-matched to lxml), inheriting `blob`
unchanged — which is why the PartFactory flip `WML_DOCUMENT_MAIN` →
`DocumentPart` is byte-neutral. It declares its **own constructor** and **own
static `load`** (the inherited-static trap: MATLAB does not inherit constructors
or dispatch inherited statics to the subclass — without its own `load`, the
factory would silently build a base `XmlPart` and `.document` would be missing).
Arg order is the docx `(partname, content_type, element, package)` — element
third. Only `document`, `save`, and `core_properties` are live; every feature
accessor is a `mat2doc:notYetPorted` stub (P2/P2-2). **`document` is a plain
`@property`** (parts/document.py 58-61) — it constructs a **fresh**
`mat2doc.document.Document` on each access (not cached), so two reads are distinct
proxies, exactly as in python-docx. **VERIFY-M1-DOCPART-BASE:** in python-docx
`DocumentPart` extends `StoryPart < BaseStoryPart < XmlPart`; that tier is not
ported at M1 — **P2-2** inserts those bases above this class (byte-neutral,
methods-only).

**Example**

```matlab
CT  = mat2doc.opc.CONTENT_TYPE;
xml = "<w:document xmlns:w='" + ...
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + ...
    "<w:body/></w:document>";
dp  = mat2doc.parts.DocumentPart.load( ...
    mat2doc.opc.PackURI("/word/document.xml"), ...
    CT.WML_DOCUMENT_MAIN, uint8(unicode2native(xml, "UTF-8")), []);
disp(class(dp))              % "mat2doc.parts.DocumentPart"
d = dp.document();           % a mat2doc.document.Document (fresh each access)
disp(class(d))               % "mat2doc.document.Document"
```

*Ported from python-docx v1.2.0: `src/docx/parts/document.py::DocumentPart` (the M1 slice: `core_properties` 52-56, `document` 58-61, `save` 111-114)*

---

## `mat2doc.document.Document`

**Syntax**

```matlab
d = mat2doc.document.Document(element, part)   % not called directly; use mat2doc.Document(...)
    d.save(path)             % LIVE: -> part.save(path) (the M1 headline path)
cp = d.core_properties()     % LIVE: -> part.core_properties() (P1-7 real)
p  = d.part()                % LIVE: the owning DocumentPart
% add_paragraph/add_heading/add_page_break/paragraphs/iter_inner_content -> (STUB, P2-3)
% add_table/tables, add_section/sections, add_comment/comments, inline_shapes,
% settings/styles, add_picture -> (STUB, P2/P7)
```

**Description**

The top API proxy object for a WordprocessingML document — **not** intended to be
constructed directly; use the `mat2doc.Document(...)` entry function. A **handle**
class wrapping a shared element tree and its part. At M1 only `save`,
`core_properties`, and the trivial `part` accessor are live; every content member
is a `mat2doc:notYetPorted` stub (P2-3 for document content, P2/P7 for the other
feature tiers), and none is on the open→save path. **VERIFY-M1-DOC-BASE:** in
python-docx `Document` extends `ElementProxy` (which supplies element identity and
the `element`/`part` accessors); `ElementProxy`/`Parented` are not ported at M1,
so this class derives `handle` only and stores `element_`/`part_` directly.
**P2-1** supplies the proxy bases and retrofits this class onto the real base (the
Mat2Ppt VERIFY-M1-C precedent) — byte-neutral (the base adds identity/accessor
behavior, not serialized output). Save accepts a path today; python-docx also
accepts a file-like object, and stream currency is inherited unchanged from the
`OpcPackage.save` contract (any stream-specific hardening is a P2-3 concern,
deliberately not stubbed — a stub would falsely block a path save that works).

**Example**

```matlab
d = mat2doc.Document();          % opens the bundled default template
out = [tempname '.docx'];
d.save(out);                     % LIVE save (M1 round-trip)
disp(isfile(out))                % 1
disp(class(d.part()))            % "mat2doc.parts.DocumentPart"
disp(class(d.core_properties())) % "mat2doc.opc.CoreProperties"
```

*Ported from python-docx v1.2.0: `src/docx/document.py::Document` (the M1 slice: `__init__` 35-39, `core_properties` 165-168, `part` 193-196, `save` 198-204)*
