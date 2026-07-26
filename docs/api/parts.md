---
title: "mat2doc.parts — the story-part tier: StoryPart + DocumentPart + the thin sibling parts"
---

# `mat2doc.parts` — the story-part tier: `StoryPart` + `DocumentPart` + `StylesPart` / `SettingsPart` / `NumberingPart`

Ported from python-docx v1.2.0 modules `src/docx/parts/story.py`,
`src/docx/parts/document.py`, `src/docx/parts/styles.py`,
`src/docx/parts/settings.py`, and `src/docx/parts/numbering.py` (plus three
byte-neutral row flips in the `src/docx/__init__.py` PartFactory registration
block and the **task #60** oxml/opc changes described below); package
`+mat2doc/+parts/`. This is the **story-part tier** that sits between the OPC
`XmlPart` base (`opc_package.md`) and the WordprocessingML `Document` entry
(`document.md`): the part classes an opened `.docx` graph is actually built
from.

At **M1 (P1-8)** `DocumentPart` was a walking-skeleton stand-in
(`DocumentPart < XmlPart` directly, every sibling accessor stubbed). **P2-2
un-stubs the OBJECT GRAPH — not the FEATURES** (see the distinction below):

- **`StoryPart`** (`+mat2doc/+parts/StoryPart.m`) — the real base class
  (`class StoryPart(XmlPart)`) inserted **above** `DocumentPart`, supplying
  `next_id` (H1), the `_document_part` lazyproperty, and the live
  `get_style` / `get_style_id` delegation. (Discharges the M1
  **VERIFY-M1-DOCPART-BASE** hand-off.)
- **`DocumentPart`** (`+mat2doc/+parts/DocumentPart.m`) — reparented to
  `StoryPart` (`DocumentPart < StoryPart < XmlPart`), with `_styles_part` /
  `_settings_part` now returning **real** sibling parts.
- **`StylesPart`** / **`SettingsPart`** / **`NumberingPart`**
  (`+mat2doc/+parts/*.m`) — pure `XmlPart` shells (`load` + `default()` +
  their templates), the targets of the three byte-neutral PartFactory flips.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## ★ Un-stub the OBJECT GRAPH, not the FEATURES (the P2-2 thesis)

P2-2 makes the **part-level plumbing** live — the object graph the entry
traverses — while every **feature surface** still raises
`mat2doc:notYetPorted` at its real-phase owner. The line runs exactly where a
part hands off to a *proxy* or a *feature* it cannot yet build:

| tier | at P2-2 | why |
|---|---|---|
| **Object graph (part plumbing)** — LIVE | `StoryPart` base + `next_id` + `_document_part`; `DocumentPart._styles_part`→`StylesPart`, `_settings_part`→`SettingsPart`; the `StylesPart`/`SettingsPart`/`NumberingPart` shells (`load`/`default`) | these are pure part-graph plumbing (`part_related_by`, `XmlPart.blob`), reachable and byte-neutral |
| **Feature surface** — STILL STUB | `DocumentPart.styles`/`settings`/`numbering_part`/`comments`/`inline_shapes`/`get_style`/`get_style_id`/header-footer parts; `StoryPart.get_or_add_image`/`new_pic_inline`; `StylesPart.styles`; `SettingsPart.settings`; `NumberingPart.numbering_definitions` | each needs a proxy / element class not ported until its real phase |

Every still-stub names its **correct real-phase owner** (the FLAG-B relabel
sweep folded in at P2-2 replaced the old catch-all "P2 tier" labels):

| still-stubbed member | real-phase owner |
|---|---|
| `DocumentPart.styles`, `get_style`, `get_style_id`; `StylesPart.styles` | **P4-7** (styles resolution) |
| `DocumentPart.settings`; `SettingsPart.settings` | **P5-1** (Settings proxy) |
| `DocumentPart.numbering_part`; `NumberingPart.numbering_definitions` | **P8-1** (numbering definitions) |
| `DocumentPart.comments` / `_comments_part` | **P8-2** (CommentsPart + Comments) |
| `DocumentPart` header / footer parts | **P5-3b** (Header/FooterPart) |
| `DocumentPart.inline_shapes`; `StoryPart.get_or_add_image` / `new_pic_inline` | **P7** (image / shape tier) |

**Zero-stubs-on-save remains the mechanical proof:** a clean
`mat2doc.Document().save` never touches any of these members, so the 17/17 M1
byte sweep stays green (below).

## The three byte-neutral PartFactory row flips

P2-2 flips three rows of `PartFactory.part_type_for_` (the ordered Nx2 table
mirroring `docx/__init__.py` 49-51), from the M1 base-`XmlPart` stand-in to the
new shells:

| content type | before (M1 stand-in) | after (P2-2) | why byte-neutral |
|---|---|---|---|
| `WML_STYLES` | `mat2doc.opc.XmlPart` | `mat2doc.parts.StylesPart` | shell IS-A `XmlPart`, inherits `blob` (parse + `serialize_part_xml`) unchanged; its **own static `load`** constructs the subclass |
| `WML_SETTINGS` | `mat2doc.opc.XmlPart` | `mat2doc.parts.SettingsPart` | same; the custom ctor stores `self._settings = element` (dead until P5-1) but emits no bytes |
| `WML_NUMBERING` | `mat2doc.opc.XmlPart` | `mat2doc.parts.NumberingPart` | same |

Because each shell IS-A `XmlPart` and inherits `blob` unchanged, the reloaded
part's **type** changes but not one emitted byte. Gate-3 re-proved it: a
foreign-cwd `mat2doc.Document().save` → `pkgcompare` vs the frozen 17-part
`references\s0001` set is **OVERALL PASS — L0 + 16 XML L1 + `thumbnail.jpeg`
bin = 17/17**, three-way `MATLAB ≡ python-docx Document().save ≡ s0001`. The
`styles.xml` / `settings.xml` / `numbering.xml` parts (now loaded as the shells)
and the reparented `document.xml` are all byte-identical.

**The inherited-static trap** is what makes each flip actually land: MATLAB does
not dispatch an inherited static method to the calling subclass, so a shell that
did *not* declare its own `load` would let the factory build a base `XmlPart`,
making the flip inert. All three shells (and `DocumentPart`) declare their own
static `load` — the faithful realization of Python's `cls`-bound
`XmlPart.load` (opc/part.py 229-232).

## Task #60 — `xpath` hoisted onto `XmlElement` (the reachability fix)

M1 tracked **task #60** (`opc_package.md` Finding 1): `XmlPart.rel_ref_count_`
and `StoryPart.next_id` call `element.xpath(...)` on a plain parsed root, but at
M1 `xpath` lived only on `BaseOxmlElement`, not on the parser-fallback
`XmlElement` — so those consumers would raise `MATLAB:noSuchMethodOrField`
before reaching their intended behavior. P2-2 closes it:

- **`xpath` is hoisted to `mat2doc.oxml.XmlElement`** (the lxml `_Element`
  analogue), a **pure relocation** — the P1-3 method body moved verbatim, the
  only change being the `arguments` context type widened from
  `BaseOxmlElement` to `XmlElement` (a widening every `BaseOxmlElement` still
  satisfies). `BaseOxmlElement` (which `< XmlElement`) **inherits it unchanged**
  — the metaclass `DefiningClass` of `BaseOxmlElement.xpath` is now
  `mat2doc.oxml.XmlElement` (a single definition, confirmed). python-docx's
  `BaseOxmlElement.xpath` override (xmlchemy.py 687-692) merely injects the
  fixed public `nsmap`; the underlying `lxml._Element.xpath` is universal, so
  the hoist is source-faithful.
- A plain **unregistered** `w:document` root (`class ==
  mat2doc.oxml.XmlElement`, before CT_Document registration at P2-3) now
  evaluates `//@id` and `//@r:id` — the exact production consumers
  `StoryPart.next_id` (`//@id`, story.py 84) and `XmlPart.rel_ref_count_`
  (`//@r:id`, part.py 246).
- **Byte-neutral:** adds a method, changes no serialization; the P1-3 xpath
  regression (Test_p1_3a/3b/3x, 158 methods) stays green.

## `drop_rel` / `delitem` — the `< 2` refcount threshold

The other two #60 items make the relationship-drop path live:

- **`Relationships.delitem`** (`del rels[rId]`) is the inherited dict
  `__delitem__` (`Relationships(Dict[str,_Relationship])` does not override it):
  it prunes the `rId → _Relationship` mapping only (`rIds_`/`rels_`) and
  **KeyErrors if absent**. **Faithful quirk (H3/H11):** it does **not** touch
  the parallel `target_parts_by_rId_` map, so a stale `related_parts` entry
  survives a drop exactly as plain `dict.__delitem__` leaves
  `_target_parts_by_rId` untouched upstream.
- **`Part.drop_rel`** is made LIVE (was a P1-6b `notYetPorted` stub):
  `if rel_ref_count_(rId) < 2: rels().delitem(rId)`, where `rel_ref_count_` is
  `XmlPart`'s `sum(element.xpath("//@r:id") == rId)`.

Gate-3 pinned the truth table against the python-docx oracle — the threshold is
`< 2` (part.py:81), so a rel referenced **exactly once IS dropped**:

| call | refcount | verdict |
|---|---|---|
| `drop_rel(rId)` | 2 | not `< 2` → **KEPT** |
| `drop_rel(rId)` | 1 | `1 < 2` → **dropped** |
| `drop_rel(rId)` | 0 (present rel) | `0 < 2` → **dropped** (implicit rel) |
| `drop_rel(rId)` | absent | `del rels[..]` → `mat2doc:KeyError` |

(This **corrects the P2-2 brief's hypothesis** that a rel referenced once is
not dropped — the port implements `< 2` faithfully and matches the oracle.)

## Deviation posture (adopt-only, ZERO new D-numbers)

Gate-2 (Opus auditor) **APPROVE** and Gate-3 **PASS** both confirmed **0 new
D-numbers**; the 17/17 L1 sweep proves zero output-visible divergence. Every
divergence is a recurrence of an already-signed ruling:

- **D-001** (own OOXML parser), **D-serializer-nsdecl** (unreachable at P2-2 —
  no created/mutated elements on the open→save path), **D-zip-time**
  (container-only) — adopted transitively via the unchanged `XmlPart.blob`
  path; re-proven by the 17/17 sweep.
- The `next_id` ASCII-`isdigit` filter (Python `str.isdigit()` realized as
  non-empty `[0-9]`) and the `delitem`-leaves-`related_parts`-stale quirk are
  **faithful reproductions / unexercised-grammar notes in the already-signed
  D-002 family** — Word and python-docx only ever emit ASCII-integer ids — so
  **no new D-number and no new ledger row**.

---

## `mat2doc.parts.StoryPart`

**Syntax**

```matlab
sp = mat2doc.parts.StoryPart(partname, content_type, element, package)   % element BEFORE package (docx)
sp = mat2doc.parts.StoryPart.load(partname, content_type, blob, package) % (Static, inherited)
n  = sp.next_id()               % LIVE (H1): max existing //@id value + 1, or 1
% get_style/get_style_id  -> LIVE delegation to _document_part (P4-7 feature stub)
% get_or_add_image/new_pic_inline -> (STUB, P7 image tier)
```

**Description**

The base class for story parts — parts that can carry textual content (the
document part and the header / footer / comments parts), sharing content
behaviors like `paragraphs` / `add_paragraph` / `add_table`. It subclasses
`mat2doc.opc.XmlPart` (parses on load, re-serializes on save), adding **methods
only**, so inserting it above `DocumentPart` is **byte-neutral**. `next_id` is
**LIVE** (H1): it collects `element.xpath("//@id")` — reachable on a plain
parsed root thanks to the #60 hoist — filters to ASCII-digit values, and
returns `max + 1` (gaps not filled) or `1` if none; this is **data arithmetic on
id VALUES, not an index shift**. `_document_part` is a `@lazyproperty`
(logical-flag cached) resolving the package's `main_document_part` — the same
`DocumentPart` handle on repeat. `get_style` / `get_style_id` are **live
delegations** to `_document_part`, whose implementations are P4-7 feature stubs,
so the `notYetPorted` surfaces at the P4-7 owner. `get_or_add_image` /
`new_pic_inline` are **P7** feature stubs (never on the open/save path).

**NAME (ratified):** python-docx v1.2.0 defines `class StoryPart(XmlPart)` —
there is **no `BaseStoryPart`** anywhere in the clone (the M1 DocumentPart header
claimed one; the P2-2 brief inherited the error). Per design.md §1 (exact Python
spelling) the class is `StoryPart`. Confirmed at Gate-2 (RATIFIED) and Gate-3.

**Example**

```matlab
% next_id: the next available //@id VALUE (max + 1), gaps not filled.
w   = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
xml = "<w:document xmlns:w='" + w + "'><w:body>" + ...
    "<w:p id='7'/><w:p id='12'/><w:p id='notnum'/></w:body></w:document>";
dp  = mat2doc.parts.DocumentPart.load( ...        % DocumentPart IS-A StoryPart
    mat2doc.opc.PackURI("/word/document.xml"), ...
    mat2doc.opc.CONTENT_TYPE.WML_DOCUMENT_MAIN, ...
    uint8(unicode2native(xml, "UTF-8")), []);
disp(isa(dp, "mat2doc.parts.StoryPart"))   % 1
disp(dp.next_id())                          % 13   (max{7,12} + 1; "notnum" skipped)
```

*Ported from python-docx v1.2.0: `src/docx/parts/story.py::StoryPart` (P2-2 thin slice: `next_id` 76-88 + `_document_part` 90-95 LIVE; `get_style`/`get_style_id` 41-58 live delegation to the P4-7 stub; `get_or_add_image` 27-39 / `new_pic_inline` 60-74 P7 stubs)*

---

## `mat2doc.parts.DocumentPart`

**Syntax**

```matlab
dp = mat2doc.parts.DocumentPart(partname, content_type, element, package)   % element BEFORE package
dp = mat2doc.parts.DocumentPart.load(partname, content_type, blob, package) % (Static) own override
d  = dp.document()          % FRESH mat2doc.document.Document each access (non-cached)
     dp.save(path)          % -> package.save(path)
cp = dp.core_properties()   % -> package.core_properties() (LIVE, P1-7)
sp = dp.styles_part_()      % -> real mat2doc.parts.StylesPart   (P2-2 LIVE)
st = dp.settings_part_()    % -> real mat2doc.parts.SettingsPart (P2-2 LIVE)
% styles/settings/numbering_part/comments/inline_shapes, header & footer parts,
% get_style/get_style_id -> (STUB; owners P4-7/P5-1/P8-1/P8-2/P7/P5-3b)
```

**Description**

The main-document part (`/word/document.xml`). At **P2-2** it is reparented to
`StoryPart` (`DocumentPart < StoryPart < XmlPart`, `class
DocumentPart(StoryPart)`), a **byte-neutral** superclass insertion (StoryPart
adds methods only), and its **object-graph accessors go live**: `styles_part_`
(`_styles_part`, document.py 156-169) returns a real
`mat2doc.parts.StylesPart`, `settings_part_` (`_settings_part`, 142-154) a real
`mat2doc.parts.SettingsPart`. Both are plain `@property` (NOT lazyproperty) —
each call runs `part_related_by`; on a package that HAS the part (default.docx
does) it returns the **same** loaded handle every time (H5/H9 identity via the
live rels), and only when the part is ABSENT do they create a default and relate
it (materialize-once). The **KeyError-only** catch is faithful: a
`part_related_by` `ValueError` (duplicate rel) propagates, exactly as Python's
`except KeyError:` lets it through. `document` / `save` / `core_properties`
remain M1-live; `document` is a plain `@property` constructing a **fresh**
`mat2doc.document.Document` each access. Every remaining feature accessor stays a
`mat2doc:notYetPorted` stub naming its real-phase owner (table above). Own
constructor + own static `load` handle the inherited-static trap.

**Example**

```matlab
d  = mat2doc.Document();          % open the bundled default template
dp = d.part();                    % the DocumentPart
disp(class(dp))                              % "mat2doc.parts.DocumentPart"
disp([isa(dp,"mat2doc.parts.StoryPart") isa(dp,"mat2doc.opc.XmlPart")])   % 1  1
sp1 = dp.styles_part_();  sp2 = dp.styles_part_();
disp(class(sp1))                             % "mat2doc.parts.StylesPart"
disp(sp1 == sp2)                             % 1   (same handle on repeat)
disp(class(dp.settings_part_()))             % "mat2doc.parts.SettingsPart"
```

*Ported from python-docx v1.2.0: `src/docx/parts/document.py::DocumentPart` (P2-2: reparented to `StoryPart`; `_styles_part` 156-169 / `_settings_part` 142-154 LIVE; feature accessors stubbed at their real-phase owners)*

---

## `mat2doc.parts.StylesPart`

**Syntax**

```matlab
sp = mat2doc.parts.StylesPart(partname, content_type, element, package)
sp = mat2doc.parts.StylesPart.load(partname, content_type, blob, package)   % (Static) own override
sp = mat2doc.parts.StylesPart.default(package)                              % (Static) fresh from template
b  = sp.blob()      % inherited XmlPart.blob: serialize_part_xml(element)
% styles -> (STUB, P4-7)
```

**Description**

The `styles.xml` part — a pure `XmlPart` shell at P2-2. It parses on load and
re-serializes on save via `serialize_part_xml` (byte-matched to lxml), inheriting
`blob` / `element` unchanged; that is why the `WML_STYLES → StylesPart` flip is
byte-neutral. `default(package)` builds a fresh part from the byte-identical
`+mat2doc/templates/default-styles.xml` (**NOT** on the M1 open/save path —
default.docx ships its own `styles.xml`, so it loads via `load`; `default()`
runs only for a styles-less package, validated at P4-7). The `styles` proxy
accessor is a **P4-7** feature stub (the `Styles` proxy / `StyleFactory` land
there). Its own static `load` makes the flip land on the subclass
(inherited-static trap).

**Example**

```matlab
d  = mat2doc.Document();
sp = d.part().styles_part_();
disp(class(sp))                              % "mat2doc.parts.StylesPart"
disp(isa(sp, "mat2doc.opc.XmlPart"))         % 1
disp(startsWith(char(sp.blob()), "<?xml"))   % 1   (re-serialized on demand)
```

*Ported from python-docx v1.2.0: `src/docx/parts/styles.py::StylesPart` (`default` 22-28 + `_default_styles_xml` 36-42 LIVE; `styles` 30-34 → P4-7 stub)*

---

## `mat2doc.parts.SettingsPart`

**Syntax**

```matlab
sp = mat2doc.parts.SettingsPart(partname, content_type, element, package)   % ctor stores self._settings
sp = mat2doc.parts.SettingsPart.load(partname, content_type, blob, package) % (Static) own override
sp = mat2doc.parts.SettingsPart.default(package)                            % (Static) fresh from template
% settings -> (STUB, P5-1)
```

**Description**

The `settings.xml` part — a pure `XmlPart` shell with one faithful addition:
unlike the plain `XmlPart`, docx `SettingsPart.__init__` also stores
`self._settings = element` (settings.py 26), ported as the private `settings_`
field (the cache the P5-1 `settings` proxy will read). That state is **dead until
P5-1** but the constructor is faithful. `default(package)` builds from the
byte-identical `default-settings.xml` template (not on the M1 path). The
`settings` proxy accessor is a **P5-1** feature stub. Own static `load` for the
flip; byte-neutral (`WML_SETTINGS → SettingsPart`).

**Example**

```matlab
d  = mat2doc.Document();
sp = d.part().settings_part_();
disp(class(sp))                              % "mat2doc.parts.SettingsPart"
disp(isa(sp, "mat2doc.opc.XmlPart"))         % 1
```

*Ported from python-docx v1.2.0: `src/docx/parts/settings.py::SettingsPart` (`__init__` 22-26, `default` 28-34 + `_default_settings_xml` 44-50 LIVE; `settings` 36-42 → P5-1 stub)*

---

## `mat2doc.parts.NumberingPart`

**Syntax**

```matlab
np = mat2doc.parts.NumberingPart(partname, content_type, element, package)
np = mat2doc.parts.NumberingPart.load(partname, content_type, blob, package)   % (Static) own override
     mat2doc.parts.NumberingPart.new()      % (Static) -> mat2doc:NotImplementedError (FAITHFUL, upstream)
% numbering_definitions -> (STUB, P8-1)
```

**Description**

The `numbering.xml` part — a pure `XmlPart` shell. Own static `load` for the
byte-neutral `WML_NUMBERING → NumberingPart` flip. Two members are notable:

- **`new()`** raises `mat2doc:NotImplementedError` — this is **faithful, NOT a
  port stub**: python-docx v1.2.0 itself declares `def new(cls): raise
  NotImplementedError` (numbering.py 11-14). Its only caller is
  `DocumentPart.numbering_part`, itself a P8-1 feature stub, so it is unreached
  at P2-2. (Upstream raises a bare `NotImplementedError`; MATLAB `error(id,'')`
  does not throw, so the port supplies explanatory text — the identifier is the
  faithful element.)
- **`numbering_definitions`** is a **P8-1** feature stub — its faithful body
  builds a `_NumberingDefinitions` over the element, whose `__len__` reads the
  `CT_Numbering` `num_lst` descriptor (not ported until P8-1). Never on the
  open/save path.

**Example**

```matlab
d  = mat2doc.Document();
np = mat2doc.opc.PartFactory.part_cls_for_(mat2doc.opc.CONTENT_TYPE.WML_NUMBERING);
disp(np)                                     % "mat2doc.parts.NumberingPart"
% NumberingPart.new is upstream-NotImplementedError (faithful), unreached at P2-2:
try
    mat2doc.parts.NumberingPart.new();
catch ME
    disp(ME.identifier)                      % "mat2doc:NotImplementedError"
end
```

*Ported from python-docx v1.2.0: `src/docx/parts/numbering.py::NumberingPart` (`new` 11-14 faithful `NotImplementedError`; `numbering_definitions` 16-20 → P8-1 stub)*
