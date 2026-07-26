---
title: "mat2doc — the WordprocessingML document tier: Document + Package + the CT_Document/CT_Body object graph"
---

# WordprocessingML document tier — `mat2doc.Document` + `Package` + `DocumentPart` + `document.Document` + `CT_Document`/`CT_Body`

Ported from python-docx v1.2.0 modules `src/docx/api.py`, `src/docx/package.py`,
`src/docx/parts/document.py`, `src/docx/document.py`, `src/docx/oxml/document.py`
(`CT_Document`/`CT_Body`, **P2-3**), and `src/docx/blkcntnr.py`
(`BlockItemContainer`, **P2-3**) — plus the byte-neutral row flips and the two
`w:document`/`w:body` registrations in the `src/docx/__init__.py` blocks; packages
`+mat2doc/`, `+mat2doc/+package/`, `+mat2doc/+parts/`, `+mat2doc/+document/`,
`+mat2doc/+oxml/+document/`, and the package-root `+mat2doc/BlockItemContainer.m`.
This is the **docx-level document tier** — the public `mat2doc.Document()` function
and the object graph beneath it — that sits *above* the OPC package/part tier
(`opc_package.md`): the layer a user actually calls to open a `.docx` and save it
back out. It began as the M1 walking skeleton (P1-8) and is **completed by P2-3**
(the `CT_Document`/`CT_Body` element classes and the `_Body`/`blkcntnr` object
graph — **the last P2 work package**).

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

- **P2-1 — `ElementProxy` / `Parented` base (DONE — see `proxy.md`).** Supplied
  the proxy base classes and retrofitted `mat2doc.document.Document` onto the real
  `ElementProxy` base (**VERIFY-M1-DOC-BASE discharged**, byte-neutral; the
  Mat2Ppt WP8 `Presentation`→`PartElementProxy` precedent). `Document` now
  inherits `element()`/`eq`/`ne`; the wrapped root lives on the base's protected
  `element_`.
- **P2-2 — `StoryPart` + real sibling parts (DONE — see `parts.md`).** Inserts
  `StoryPart` **above** `DocumentPart` (**VERIFY-M1-DOCPART-BASE** discharged;
  note the real class is `StoryPart`, not the `BaseStoryPart` the M1 header
  wrongly claimed — v1.2.0 has no such class), ports the thin `StylesPart` /
  `SettingsPart` / `NumberingPart` `XmlPart` shells (and flips their PartFactory
  rows — byte-identical output, only the reloaded type changes), un-stubs the
  `DocumentPart._styles_part` / `_settings_part` object-graph accessors (feature
  accessors stay stubbed at their real-phase owner), discharges **task #60** (the
  `xpath` hoist onto `XmlElement` that closes the `drop_rel` → `rel_ref_count_`
  reachability gap), and adds `Relationships` `delitem` + the live `Part.drop_rel`
  (`< 2` refcount threshold). Because those bases add methods only
  (blob/load/element inherit from `XmlPart` unchanged), the reparented classes
  emit identical bytes (17/17 M1 sweep unchanged). The story-part tier is
  documented in **`parts.md`**.
- **P2-3 — document shell + `blkcntnr` + `CT_Document`/`CT_Body` (DONE — this
  WP; COMPLETES PHASE 2).** Registers the two document-root element classes, wires
  the `_Body`/`BlockItemContainer` object graph and the `Document._body` /
  `_block_width` accessors, and re-proves `document.xml` L1. The content adders
  (`add_paragraph`/`add_heading`/`paragraphs`/`add_table`/…) stay stubbed at their
  P4/P5/P6/P7/P8 owners — **the object graph, not the features** (below).

## ★ P2-3 — the document object graph, completed (COMPLETES PHASE 2)

P2-3 is the **final P2 work package**; on its merge **Phase 2 is complete**. It
closes the last gap in the walking skeleton by registering the two document-root
element classes and wiring the block-container object graph the entry traverses —
while every *content-authoring* member stays a `mat2doc:notYetPorted` stub at its
real-phase (P4/P5/P6/P7/P8) owner. The distinction is the same one P2-2 drew for
the part tier (`parts.md`): **P2 un-stubs the OBJECT GRAPH, not the FEATURES.**

**What went LIVE at P2-3:**

- **`mat2doc.oxml.document.CT_Document` / `CT_Body`** — the custom element classes
  registered for `w:document` / `w:body` (the ONE P2 step that touches the
  `document.xml` parse path). `CT_Document.body` (`ZeroOrOne`) + the `sectPr_lst`
  union property; `CT_Body`'s `ZeroOrMore(p)` / `ZeroOrMore(tbl)` /
  `ZeroOrOne(sectPr)` member families (incl. the public `add_p`/`add_tbl` docx
  adders), `clear_content`, and `inner_content_elements`.
- **`Document._body` → `mat2doc.document.Body_`** — the `_Body` block-item
  container proxy over the `CT_Body` element, cached through the manual
  `__body is None` idiom (`body__`); repeated `body_()` returns the **same** handle.
- **`mat2doc.BlockItemContainer`** — the `blkcntnr` base (`< StoryChild`) that
  `_Body` extends: the container wiring (`_element` + the inherited parent→part
  chain).
- **`Document._block_width`** — the section-width helper, ported faithfully
  (`page_width or Inches(8.5)` truthiness, `sections[-1]` → `section(end)`) but
  transitively blocked on the P5 `sections` stub — dead until P5.

**Still STUBBED at their real-phase owner** (each stub names its WP): every
content adder/reader — `Document.add_paragraph`/`add_heading`/`add_page_break`/
`paragraphs`/`iter_inner_content` (**P4**), `add_table`/`tables` (**P6**),
`add_section`/`sections` (**P5**), `add_picture`/`inline_shapes` (**P7**),
`add_comment`/`comments` (**P8-2**), `settings` (**P5-1**), `styles` (**P4-7**);
the mirrored `BlockItemContainer`/`Body_` adders (**P4**/**P6**); and
`CT_Body.add_section_break` (**P5**). Each stubs exactly at the item-construction
boundary — the container structure is real, only the returned `Paragraph`/`Table`
proxy is not yet ported.

### The byte-risk — registered, and proven byte-neutral

Registering `w:document`→`CT_Document` and `w:body`→`CT_Body` (both
`BaseOxmlElement` subclasses) means `word/document.xml` now **parses** to
`CT_Document`/`CT_Body` with the full descriptor machinery instead of a generic
`XmlElement`, and **re-serializes** through the `BaseOxmlElement` path on save.
This was the boundary-audit escalation flag for all of P2 — and it moved **zero
bytes**. `mat2doc.Document().save` from a foreign cwd → `pkgcompare` vs the frozen
17-part `references\s0001` set is **OVERALL PASS — L0 + 16 XML L1 +
`thumbnail.jpeg` bin = 17/17**, three-way `MATLAB ≡ python-docx Document().save ≡
s0001`, with **`word/document.xml` 1548 B byte-identical**. Neither CT class
overrides any serialization member; both `XmlElement` and `BaseOxmlElement` exit
through the identical `+oxml\serialize_part_xml` walk, so the class flip changes
only the parsed node's **type**, never its bytes. The deep-re-audit trigger did
not fire; no Fable escalation.

### H11 child ordering — the WP hazard

`CT_Body`'s `p` and `tbl` descriptors carry `successors=("w:sectPr",)`; `sectPr`
carries `successors=()`. A newly added `w:p` / `w:tbl` therefore inserts **before**
any existing `w:sectPr` (the section-properties sentinel that must stay last),
while a fresh `sectPr` appends at the end. Gate-3 pinned both orderings against the
python-docx oracle: `get_or_add_sectPr → add_p → add_tbl → add_p` yields
`["p","tbl","p","sectPr"]`, and `add_p → add_tbl → get_or_add_sectPr` yields
`["p","tbl","sectPr"]`. The successor wiring (`PTBL_SUCCESSORS = "w:sectPr"` for
p/tbl; `NO_SUCCESSORS` for sectPr) is proven at the oxml layer today; it lights up
through the real `Paragraph`/`Table` path when P4/P6 land (Gate-2/3 carry-forward).

### Forward dependencies resolve generic (scope-driven, byte-neutral)

`CT_P` (P4), `CT_Tbl` (P6) and `CT_SectPr` (P5) are **not** registered yet, so the
`w:p`/`w:tbl`/`w:sectPr` that `add_p`/`add_tbl`/`get_or_add_sectPr` create fall
back to a generic `mat2doc.oxml.XmlElement`. They still serialize correctly and
insert in the right sequence — only the element **class** is generic. In
python-docx these are real `CT_P`/`CT_Tbl`/`CT_SectPr`; the class difference is
intentional at P2-3 and byte-neutral (identical serialization + identical
insertion order, proven by the 17/17 sweep and the H11 pins), so Gate-3 keeps it
out of the shared probe and checks it MATLAB-side only.

### ★ PHASE 2 COMPLETE

With P2-3 the **document object graph is fully wired end-to-end**:
`mat2doc.Document()` → `Package` → `DocumentPart` (real `StoryPart` tier, P2-2) →
`document.Document` (on the real `ElementProxy` base, P2-1) → `CT_Document` →
`_Body` / `CT_Body` (P2-3). Phase 2 delivered, in order, the **proxy base tier**
(P2-1 — `proxy.md`), the **object-graph un-stub** (P2-2 — `parts.md`), and this
**document shell** (P2-3). Throughout, **M1 was preserved** — every P2 WP re-ran
the 17/17 byte-neutrality sweep three-way and moved zero bytes; **zero new
D-numbers across all of P2**. What remains is the **feature surface**: content
authoring lights up at **P3** (enums + simpletypes) → **P4** (paragraph/run) and
onward toward **M2**.

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
e  = d.element()             % LIVE (P2-1): the CT_Document root (inherited ElementProxy)
b  = d.body_()               % LIVE (P2-3): the _Body proxy (cached same handle)
% block_width_                 -> LIVE (P2-3) but transitively raises at the P5 sections stub
% add_paragraph/add_heading/add_page_break/paragraphs/iter_inner_content -> (STUB, P4)
% add_table/tables -> (STUB, P6); add_section/sections -> (STUB, P5)
% add_comment/comments -> (STUB, P8-2); inline_shapes/add_picture -> (STUB, P7)
% settings -> (STUB, P5-1); styles -> (STUB, P4-7)
```

**Description**

The top API proxy object for a WordprocessingML document — **not** intended to be
constructed directly; use the `mat2doc.Document(...)` entry function. A **handle**
class on the real **`mat2doc.shared.ElementProxy`** base (P2-1). `save`,
`core_properties`, `part`, and the inherited `element`/`eq`/`ne` are live; **P2-3**
made the private object-graph accessors `body_` (the cached `_body`) and
`block_width_` (`_block_width`) live too. Every *content-authoring* member remains
a `mat2doc:notYetPorted` stub naming its real-phase owner (P4 paragraph/heading,
P5 section, P6 table, P7 picture/shape, P8-2 comment, P4-7 styles, P5-1 settings),
and none is on the open→save path. **`body_`** returns a `mat2doc.document.Body_`
over `element.body`, cached via the manual `__body is None` idiom (`body__`), so
repeated reads return the **same** proxy handle (faithful to python-docx's cached
`@property`, document.py 241-246). **`block_width_`** is ported faithfully
(document.py 232-239 — `page_width or Inches(8.5)` H4 truthiness, `sections[-1]` →
`section(end)`) but calls `sections()` first, so it raises the **P5** stub before
any `Section` member resolves; it is unreachable by a live path at P2-3 (its only
caller, `add_table`, is itself a P6 stub). **VERIFY-M1-DOC-BASE (RESOLVED, P2-1):**
`Document` now extends `ElementProxy` (element identity + the `element` accessor);
the retrofit was byte-neutral (17/17 M1 sweep) and made `Document ==`
Python-faithful. `part` stays overridden to return this document's own `_part`.
Save accepts a path today; python-docx also accepts a file-like object, and stream
currency is inherited unchanged from the `OpcPackage.save` contract.

**Example**

```matlab
d = mat2doc.Document();          % opens the bundled default template
out = [tempname '.docx'];
d.save(out);                     % LIVE save (M1 round-trip)
disp(isfile(out))                % 1
disp(class(d.part()))            % "mat2doc.parts.DocumentPart"
disp(class(d.core_properties())) % "mat2doc.opc.CoreProperties"
b1 = d.body_(); b2 = d.body_();  % P2-3: the _Body proxy
disp(class(b1))                  % "mat2doc.document.Body_"
disp(b1 == b2)                   % 1   (manual __body cache -> same handle)
```

*Ported from python-docx v1.2.0: `src/docx/document.py::Document` (live members: `__init__` 35-39, `core_properties` 165-168, `part` 193-196, `save` 198-204, `_block_width` 232-239, `_body` 241-246; base `ElementProxy` retrofitted P2-1, `_body`/`_block_width` added P2-3)*

---

## `mat2doc.oxml.document.CT_Document` (P2-3)

**Syntax**

```matlab
d    = mat2doc.oxml.OxmlElement("w:document")   % registry -> CT_Document
body = d.get_or_add_body()      % <w:body> child, created (append) if absent
b    = d.body                   % the <w:body> child, or [] (None) if absent
secs = d.sectPr_lst             % all directly-accessible w:sectPr (document order)
```

**Description**

The custom element class for the `<w:document>` root, registered for `w:document`
(`oxml/__init__.py:101`). `CT_Document < mat2doc.oxml.BaseOxmlElement`, so an
opened `word/document.xml` now parses to a `CT_Document` with the full descriptor
machinery instead of a generic `XmlElement`. It carries one `ZeroOrOne("w:body")`
child descriptor (`body` / `get_or_add_body` / the generated `new_body_` /
`insert_body_` / `add_body_` / `remove_body_` engine members) and the read-only
`sectPr_lst` `@property` — the union
`./w:body/w:p/w:pPr/w:sectPr | ./w:body/w:sectPr`, returned in document order with
the body-level `sectPr` last. **Byte-neutral registration:** `BaseOxmlElement`
adds no serialization override, so the frozen M1 `document.xml` (1548 B) does not
move one byte (17/17 sweep). The `sectPr_lst` xpath descends into
`w:p/w:pPr/w:sectPr` — `CT_P` (P4) / `CT_SectPr` (P5) are unregistered, so those
matched nodes are generic `XmlElement`; the returned node-set is identical either
way (only the match's element **class** changes, not its identity/order).

**Example**

```matlab
d    = mat2doc.oxml.OxmlElement("w:document");  % a CT_Document (registered)
body = d.get_or_add_body();                     % creates <w:body/> if absent
disp(class(d))                                  % "mat2doc.oxml.document.CT_Document"
disp(class(body))                               % "mat2doc.oxml.document.CT_Body"
disp(numel(d.sectPr_lst))                       % 0   (none yet)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/document.py::CT_Document` (registered for `<w:document>`, `oxml/__init__.py:101`)*

---

## `mat2doc.oxml.document.CT_Body` (P2-3)

**Syntax**

```matlab
body = d.get_or_add_body()      % a CT_Body (from a CT_Document)
p    = body.add_p()             % new <w:p>, inserted BEFORE any <w:sectPr> (H11)
tbl  = body.add_tbl()           % new <w:tbl>, inserted BEFORE any <w:sectPr> (H11)
sp   = body.get_or_add_sectPr() % the <w:sectPr>, appended LAST (successors=())
ps   = body.p_lst               % list of <w:p> children (document order, 1x0 if none)
els  = body.inner_content_elements  % all <w:p> and <w:tbl>, document order
       body.clear_content()     % remove all content, preserving <w:sectPr> if present
```

**Description**

The custom element class for `<w:body>`, registered for `w:body`
(`oxml/__init__.py:100`); `CT_Body < mat2doc.oxml.BaseOxmlElement`. Its three
descriptors are `p = ZeroOrMore("w:p", successors=("w:sectPr",))`,
`tbl = ZeroOrMore("w:tbl", successors=("w:sectPr",))`, and
`sectPr = ZeroOrOne("w:sectPr", successors=())`. Per the docx `ZeroOrMore` form
(D-delta-4), each list descriptor generates a **public `add_p`/`add_tbl`** adder
(plus the `x_lst` getter and the `new_`/`insert_`/`add_` engine members) — there
is **no** bare single getter and **no** remover. **H11 child ordering (the WP
hazard):** `p`/`tbl` carry `successors=("w:sectPr",)`, so a new `w:p`/`w:tbl`
inserts before any existing `w:sectPr`; `sectPr` carries `successors=()` and
appends last (the sentinel). `clear_content` (xpath `./*[not(self::w:sectPr)]`)
materializes the non-sectPr children and removes them, preserving the section
properties; `inner_content_elements` (xpath `./w:p | ./w:tbl`) is a fixed
two-branch child-axis union — a `w:p` nested inside a `w:ins` is **excluded**.
`add_section_break` is a **P5** stub (it needs `CT_SectPr.clone` + `CT_P.set_sectPr`;
`add_p` itself is live, only the section-break orchestration stubs).

**Example**

```matlab
d    = mat2doc.oxml.OxmlElement("w:document");
body = d.get_or_add_body();
sp   = body.get_or_add_sectPr();   % sectPr present
body.add_p(); body.add_tbl();      % each inserts BEFORE the sectPr (H11)
kids = arrayfun(@(e) string(e.tag), body.inner_content_elements);
disp(numel(kids))                  % 2   (the p and the tbl; sectPr excluded)
body.clear_content();              % removes the p and tbl, keeps the sectPr
disp(numel(body.p_lst))            % 0
```

*Ported from python-docx v1.2.0: `src/docx/oxml/document.py::CT_Body` (registered for `<w:body>`, `oxml/__init__.py:100`)*

---

## `mat2doc.document.Body_` (`_Body`, P2-3)

**Syntax**

```matlab
b = d.body_()                   % the _Body proxy for a Document (cached same handle)
b = mat2doc.document.Body_(body_elm, parent)   % ctor (not called directly)
    b.clear_content()           % clear content, preserving section properties
% add_paragraph/add_table/paragraphs/tables/iter_inner_content -> inherited STUB (P4/P6)
```

**Description**

The proxy for the `<w:body>` element — the block-item container for the main
document story. `Body_ < mat2doc.BlockItemContainer < mat2doc.shared.StoryChild`
(the `_Body` → `Body_` underscore rotation, design.md §2). It inherits the
block-item add/read surface from `BlockItemContainer` (all P4/P6 stubs at this WP)
and the parent→part chain from `StoryChild`. Beyond the container wiring it stores
the `CT_Body` element again in its own `_body` field (rotated `body_`, faithful to
python-docx's redundant store, document.py 255-257) and overrides `clear_content`
to delegate to the **live** `CT_Body.clear_content`. `Document.body_` caches the
`Body_` handle through the manual `__body is None` idiom, so repeated access
returns the same proxy.

**Example**

```matlab
d = mat2doc.Document();          % opens the bundled default template
b = d.body_();                   % a mat2doc.document.Body_ (the _Body proxy)
disp(class(b))                   % "mat2doc.document.Body_"
b.clear_content();               % delegates to CT_Body.clear_content
disp(b == d.body_())             % 1   (cached same handle)
```

*Ported from python-docx v1.2.0: `src/docx/document.py::_Body`*

---

## `mat2doc.BlockItemContainer` (`blkcntnr`, P2-3)

**Syntax**

```matlab
c = mat2doc.BlockItemContainer(element, parent)   % a block element + a story parent
p = c.part()                    % delegated up the StoryChild parent chain -> a StoryPart
% add_paragraph/add_paragraph_/add_table/paragraphs/tables/iter_inner_content -> STUB (P4/P6)
```

**Description**

Base class for proxy objects that can contain block-level items (paragraphs and
tables): `_Body`, `_Cell`, header/footer, footnote/endnote, comment, and text-box
objects all derive it. `BlockItemContainer < mat2doc.shared.StoryChild` — because
`docx.blkcntnr` is a **top-level** module, the class lands at the `+mat2doc`
package **root** (`mat2doc.BlockItemContainer`), mirroring
`from docx.blkcntnr import BlockItemContainer`; it does **not** collide with the
`+mat2doc\Document.m` entry function (distinct names). At P2-3 only the container
wiring is live — the constructor stores the wrapped block element (rotated
`element_`) and delegates the parent to `StoryChild`. Every content member
(`add_paragraph` / `add_paragraph_` / `add_table` / `paragraphs` / `tables` /
`iter_inner_content`) is a `mat2doc:notYetPorted` stub that raises **exactly at the
item-construction boundary**: `blkcntnr.py:99-101` reduces `_add_paragraph` to
`Paragraph(self._element.add_p(), self)` — the `add_p()` is live (P2-3), but the
`Paragraph` proxy needs `CT_P` + `Paragraph` (**P4**) and `Table` needs `CT_Tbl` +
`Table` (**P6**), so the whole method must raise up front (calling `add_p()` first
would leave a stray `<w:p>` in the tree). This is the **canonical** stub location
for the block-item adders; the `Document` feature adders stub above it as faithful
placeholders (both layers stub, so no live path reaches an unported dependency).

**Example**

```matlab
d        = mat2doc.Document();          % a ProvidesStoryPart
body_elm = d.element().body;            % the CT_Body element
c        = mat2doc.BlockItemContainer(body_elm, d);
disp(class(c))                          % "mat2doc.BlockItemContainer"
disp(class(c.part()))                   % "mat2doc.parts.DocumentPart" (up the parent chain)
```

*Ported from python-docx v1.2.0: `src/docx/blkcntnr.py::BlockItemContainer`*
