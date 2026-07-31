---
title: "mat2doc.oxml.section — the section oxml layer (CT_SectPr · CT_PageSz · CT_PageMar · CT_SectType · CT_HdrFtrRef · CT_HdrFtr · the section block-element iterator)"
---

# `mat2doc.oxml.section` — the section oxml layer (`CT_SectPr` + page geometry + `CT_HdrFtr` + the section iterator)

Ported from python-docx v1.2.0 `src/docx/oxml/section.py` (in the NEW package
`+mat2doc/+oxml/+section/`) — the `<w:sectPr>` section-properties element and the
four child element classes its page-geometry / start-type / header-footer
accessors read and write, plus the **seven `register_element_cls` rows** they
require (`src/docx/oxml/__init__.py` :84, :123, :126–:130).

:::{note}
**★ Sections tier opens — section oxml core done; header/footer bodies + the
section iterator next.** This is the **second work package of Phase 5** (the
sections tier). It ports the section-properties *element* core — `CT_SectPr` and
its page-size / margin / section-type / header-footer-reference children — and
registers `w:sectPr` so the shipped `word/document.xml` (the **CENTRAL M1 part**,
which carries a `<w:sectPr>`) now parses through `CT_SectPr`, **byte-neutrally**
(M1 stays 17/17). The `<w:hdr>`/`<w:ftr>` body roots (`CT_HdrFtr`) and the block
walk (`_SectBlockElementIterator`) follow at **P5-2b**; the `Section`/`Sections`
proxy surface at **P5-3a**; the separate-part header/footer rels at **P5-3b**.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The one-package layout

`oxml/section.py` (a `docx/oxml/` module) lands in the single new package
`+mat2doc\+oxml\+section` — following the established module→subpackage
convention. It collides with no existing package.

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `oxml/section.py` | `+mat2doc\+oxml\+section\` | `CT_SectPr`, `CT_PageSz`, `CT_PageMar`, `CT_SectType`, `CT_HdrFtrRef` |

**Deferred to P5-2b (cleanly, not stubbed here):** `CT_HdrFtr` (`w:hdr`/`w:ftr`
roots) and `_SectBlockElementIterator`. The single `CT_SectPr` member that needs
the iterator, `iter_inner_content`, is a clean `mat2doc:notYetPorted` stub naming
`_SectBlockElementIterator` and owner P5-2b — nothing M1- or geometry-critical is
stubbed.

(id-section-registry-parse-path)=
## The registry-adding, M1-byte-clean parse path — the headline

This WP registers **seven** tags, exactly as python-docx registers them, in
`docx/oxml/__init__.py` source order:

| tag | class | `oxml/__init__.py` | role |
|---|---|---|---|
| `w:titlePg` | `shared.CT_OnOff` | :84 | the title-page toggle child (**P5-1 deferral CLOSED**) |
| `w:footerReference` | `section.CT_HdrFtrRef` | :123 | a footer reference child of `w:sectPr` |
| `w:headerReference` | `section.CT_HdrFtrRef` | :126 | a header reference child of `w:sectPr` |
| `w:pgMar` | `section.CT_PageMar` | :127 | the page-margins child |
| `w:pgSz` | `section.CT_PageSz` | :128 | the page-size child |
| `w:sectPr` | `section.CT_SectPr` | :129 | the **section-properties root** |
| `w:type` | `section.CT_SectType` | :130 | the section-start-type child |

`default.docx`'s `word/document.xml` carries
`<w:sectPr><w:pgSz .../><w:pgMar .../><w:cols/><w:docGrid/></w:sectPr>`, so
registering these tags makes that subtree transit the new `CT_*` classes on
**every load** of the central document part. Because each new class extends
`BaseOxmlElement` and reserializes through the same `serialize_part_xml` walk with
**no serialization override** (they only *add* descriptors and `@property`
accessors), registering a tag changes only the parsed node's **class**, never its
**bytes**.

The **M1 17/17 byte-neutrality sweep holds**: `mat2doc.Document().save()` → unzip
→ all 17 parts byte-identical to the frozen `references\s0001` reference, with
`word/document.xml` (1548 B, SHA-256 `0e4dd503…7836327`, independently re-derived
through the new `CT_SectPr` parse path) L1 byte-identical. This is the standing
obligation every registry-adding `CT_*` WP inherits (P4-6 `styles.xml` / P5-1
`settings.xml` precedent), and here it lands on the **M1-central** part.

:::{note}
**`w:type` is an ELEMENT registration, disjoint from every `@w:type` attribute.**
`register_element_cls("w:type", CT_SectType)` registers the *element* `w:type`.
Every other `w:type` in the tree (`CT_Br`, `CT_Style`, table `tblLayout`/`tblW`)
is the `@w:type` *attribute* — a disjoint namespace from element-tag
registration. No prior registry row registers the element `w:type`; the
registry's duplicate-key guard did not fire. The `w:hdr`/`w:ftr` rows
(`__init__.py` :124–:125) are intentionally NOT registered here — they belong to
`CT_HdrFtr` (P5-2b).
:::

---

(id-ct_sectpr)=
## `CT_SectPr`

**Syntax**

```matlab
sp = mat2doc.oxml.OxmlElement("w:sectPr");      % a CT_SectPr (registered)
sp.page_width  = mat2doc.shared.Twips(12240);   % -> <w:pgSz w:w="12240"/>
sp.page_height = mat2doc.shared.Twips(15840);
sp.orientation                                   % PORTRAIT (no @w:orient)
sp.start_type                                    % NEW_PAGE (no w:type child)
```

**Description**

The `<w:sectPr>` element — the section-properties container. It carries the
page-geometry, section-start, title-page, and header/footer-reference surface for
one section. `CT_SectPr` extends `BaseOxmlElement`; its accessors read and write
the `w:pgSz` / `w:pgMar` / `w:type` / `w:titlePg` children and the
`w:headerReference` / `w:footerReference` reference lists.

(id-ct_sectpr-tagseq)=
**The H11 child successor sequence — the correctness crux.** The `_tag_seq`
(`section.py` 112–133, ported **verbatim** as the Constant `TAG_SEQ`) is a
**20-tag** tuple giving the OOXML schema order of every `<w:sectPr>` child.
Each `ZeroOrOne` / `ZeroOrMore` descriptor declares `successors=_tag_seq[N:]`
(everything schema-*after* its own tag), and the Python 0-based slice start `N`
maps to the MATLAB 1-based slice `TAG_SEQ(N+1:end)` — the base shift applied once
at the slice (the P5-1 lesson). A wrong slice inserts a child in the wrong
position → Word repair / byte divergence. Per descriptor:

| descriptor | kind | Python `successors` | own-tag 1-based idx | MATLAB slice | first successor |
|---|---|---|---|---|---|
| `headerReference` | `ZeroOrMore` | `_tag_seq` (full) | not in seq | `TAG_SEQ` | `w:footnotePr` |
| `footerReference` | `ZeroOrMore` | `_tag_seq` (full) | not in seq | `TAG_SEQ` | `w:footnotePr` |
| `type` | `ZeroOrOne` | `_tag_seq[3:]` | 3 (`w:type`) | `TAG_SEQ(4:end)` | `w:pgSz` |
| `pgSz` | `ZeroOrOne` | `_tag_seq[4:]` | 4 (`w:pgSz`) | `TAG_SEQ(5:end)` | `w:pgMar` |
| `pgMar` | `ZeroOrOne` | `_tag_seq[5:]` | 5 (`w:pgMar`) | `TAG_SEQ(6:end)` | `w:paperSrc` |
| `titlePg` | `ZeroOrOne` | `_tag_seq[14:]` | 14 (`w:titlePg`) | `TAG_SEQ(15:end)` | `w:textDirection` |

`headerReference` / `footerReference` are **not members** of `_tag_seq` (it begins
at `w:footnotePr`), so their successors span the **whole** sequence — a new
reference inserts after any existing references but before `type`/`pgSz`/`pgMar`/…
Gate-2 and Gate-3 proved the placement byte-for-byte: seeded a scrambled sectPr,
`get_or_add` in reverse order `titlePg`→`pgMar`→`pgSz`→`type`, then added
header + footer references — result `headerReference, footerReference,
headerReference, type, pgSz, pgMar, cols, titlePg, docGrid` (references first,
titlePg between cols and docGrid), serialized bytes identical to the lxml mirror.

(id-ct_sectpr-geometry)=
**Page geometry — `page_width` / `page_height` / the seven margins.** The page-size
accessors read/write `@w:w` / `@w:h` of the `w:pgSz` child; the margin accessors
read/write the seven `@w:*` of the `w:pgMar` child. All are `Length`-typed EMU
held exactly in doubles (H6), and all are H3 tri-state — `[]` (None) round-trips,
and a `[]` assignment removes the attribute (or leaves the child absent). The
port transcribes python-docx's **faithful setter ASYMMETRY** verbatim, **not**
normalized:

- **WRAP** `int | Length | None` via `Length(value)` — `bottom_margin`, `footer`,
  `gutter`, `header`, `left_margin` — realized by the private `asLengthOrNone_`
  static (`[]` and any `Length` subclass pass through; a bare number is wrapped
  `Length(value)`);
- **NO WRAP** `Length | None`, direct assign — `top_margin`, `right_margin`,
  `page_height`, `page_width`.

(id-ct_sectpr-orientation)=
**Orientation — the NO-w/h-swap semantics (H4, H10).** The getter returns
`WD_ORIENTATION.PORTRAIT` when `w:pgSz` is absent, else `pgSz.orient`. The setter
is Python `pgSz.orient = value if value else WD_ORIENTATION.PORTRAIT` — and
because `BaseXmlEnum` subclasses `int` with `PORTRAIT`'s int value `0`, `if value`
is **falsy for BOTH `None` AND `PORTRAIT`**. The port expands that faithfully
(`isFalsy = isequal(value,[]) || double(value.value) == 0`), observably identical
either way. The crucial fact proven at Gate-2/Gate-3: setting `LANDSCAPE`
**appends `@w:orient="landscape"` and does NOT swap `@w:w`/`@w:h`** —
`<w:pgSz w:w="12240" w:h="15840" w:orient="landscape"/>` keeps the portrait
dimensions (python-docx's `Section.page_width`/`page_height` are straight
pass-throughs with no orientation arithmetic; the port adds none). A port that
swapped dimensions on landscape would diverge — the LANDSCAPE G-scenario
(`references\s0041`, 1569 B, SHA `d165a628…`) is byte-identical and pins that it
does not.

(id-ct_sectpr-starttype)=
**`start_type` — identity semantics (H4).** The getter returns
`WD_SECTION_START.NEW_PAGE` when the `w:type` child or its `@w:val` is absent, else
`type.val`. The setter is Python `if value is None or value is NEW_PAGE` —
**identity**, so `[]` (None) or `NEW_PAGE` remove the `w:type` child; any other
member writes `<w:type w:val="…"/>` at the H11 position. Contrast the titlePg
setter below, which is `==`-membership.

(id-ct_sectpr-titlepg)=
**`titlePg_val` — the `[None, False]` breadth (H3, D-delta-1).** The getter returns
`false` when `w:titlePg` is absent, else `titlePg.val`. The setter is Python
`if value in [None, False]` — a **membership** test that uses `==`, so `None`,
`False`, **and any `x == False`** (e.g. numeric `0` / `0.0`) all remove the child.
The port reproduces the full breadth as
`isequal(value,[]) || isequal(value,false)` — and because `isequal(0,false)` is
true, the numeric-0 case is covered. Otherwise `get_or_add_titlePg().val = true`;
since `CT_OnOff.val = true` equals its default, `@w:val` is dropped and an **empty
`<w:titlePg/>`** is emitted (D-delta-1). This is deliberately different from
`CT_Settings.evenAndOddHeaders_val`, which uses **identity** (`is None or is
False`) — the two are genuinely different in the python-docx source, and both are
ported faithfully.

(id-ct_sectpr-hdrftr)=
**The header/footer reference surface.** Six methods + two list getters:
`add_headerReference(type_, rId)` / `add_footerReference(type_, rId)` (the PUBLIC
geometry adders that shadow the generated `add_x` — xmlchemy skips a generated
name when the class already defines it, so only the private `add_headerReference_`
is generated), `get_headerReference(type_)` / `get_footerReference(type_)` (xpath
by `@w:type`, returning the **live** child handle — H5 — or `[]`),
`remove_headerReference(type_)` / `remove_footerReference(type_)` (return the rId
and detach by handle; a verbatim `mat2doc:ValueError` when absent — the
"should-never-happen" guard), and the `headerReference_lst` / `footerReference_lst`
document-order list getters.

`clone()` deep-copies the tree and clears **all** root attributes
(`deepcopy` + loop `remove_attrib` over `attrib_names()`) — the section-break
duplicate. `preceding_sectPr` returns the nearest `w:sectPr` on the reverse
`preceding` axis (`./preceding::w:sectPr[1]`, already 1-based, never shifted) or
`[]`.

**Example**

```matlab
sp = mat2doc.oxml.OxmlElement("w:sectPr");        % a CT_SectPr
sp.page_width  = mat2doc.shared.Twips(12240);     % <w:pgSz w:w="12240"/>
sp.page_height = mat2doc.shared.Twips(15840);
disp(sp.orientation == mat2doc.enum.section.WD_ORIENTATION.PORTRAIT);  % 1 (no @w:orient)
sp.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;        % appends @w:orient, NO w/h swap
disp(sp.start_type == mat2doc.enum.section.WD_SECTION_START.NEW_PAGE); % 1 (no w:type child)
sp.titlePg_val = true;   % -> <w:titlePg/> (empty, D-delta-1)
sp.titlePg_val = 0;      % 0 == false -> removes the child (the [None,False] breadth)
disp(sp.titlePg_val);    % 0
```

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_SectPr`*

---

(id-ct_pagesz)=
## `CT_PageSz`

**Syntax**

```matlab
ps = mat2doc.oxml.OxmlElement("w:pgSz");
ps.orient                             % PORTRAIT (@w:orient absent — the non-None default)
ps.w = mat2doc.shared.Twips(12240);   % <w:pgSz w:w="12240"/>
```

**Description**

The `<w:pgSz>` element — page dimensions and orientation. Three
`OptionalAttribute`s: `w` (`@w:w`, `ST_TwipsMeasure` → `Length`, default None),
`h` (`@w:h`, `ST_TwipsMeasure` → `Length`, default None), and `orient` (`@w:orient`,
`WD_ORIENTATION`, **default `PORTRAIT`**). `orient` is the only attribute here with
a **non-None** default: the getter returns `PORTRAIT` when `@w:orient` is absent,
and the setter **removes `@w:orient`** when the value equals `PORTRAIT` (the
`value == default` delta). `w`/`h` are the standard H3 tri-state (`[]` when absent;
`[]`-assign removes).

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_PageSz`*

---

(id-ct_pagemar)=
## `CT_PageMar`

**Syntax**

```matlab
pm = mat2doc.oxml.OxmlElement("w:pgMar");
pm.top = mat2doc.shared.Twips(1440);   % <w:pgMar w:top="1440"/>
```

**Description**

The `<w:pgMar>` element — the seven page margins, all `OptionalAttribute`,
all default None: `top` / `bottom` (`ST_SignedTwipsMeasure` — **signed** 32-bit)
and `right` / `left` / `header` / `footer` / `gutter` (`ST_TwipsMeasure` —
**unsigned**). The signed-vs-unsigned split is transcribed **exactly** against
`section.py` 63–83 — a wrong simple-type would change the accepted range and the
parse path (signed negatives such as `w:top="-720"` are proven to serialize
exactly). All seven are H3 tri-state: `[]` when absent; `[]`-assign removes the
attribute.

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_PageMar`*

---

(id-ct_secttype)=
## `CT_SectType`

**Syntax**

```matlab
t = mat2doc.oxml.OxmlElement("w:type");
t.val = mat2doc.enum.section.WD_SECTION_START.EVEN_PAGE;  % <w:type w:val="evenPage"/>
```

**Description**

The `<w:type>` element — the section start type. One `OptionalAttribute`, `val`
(`@w:val`, `WD_SECTION_START`, default None). (The python-docx docstring reads
`<w:sectType>`, but the element is registered and used as `w:type` —
`__init__.py:130`; the tag ported here is `w:type`.) H3 tri-state: `[]` when
`@w:val` is absent, `[]`-assign removes. `CT_SectPr.start_type` reads `.val` on
this child, returning `NEW_PAGE` when the child or its `@w:val` is absent.

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_SectType`*

---

(id-ct_hdrftrref)=
## `CT_HdrFtrRef`

**Syntax**

```matlab
ref = mat2doc.oxml.OxmlElement("w:headerReference");   % a CT_HdrFtrRef
ref.type_ = mat2doc.enum.section.WD_HEADER_FOOTER.PRIMARY;
ref.rId   = "rId7";   % <w:headerReference w:type="default" r:id="rId7"/>
```

**Description**

The `<w:headerReference>` / `<w:footerReference>` elements. Two
`RequiredAttribute`s: `type_` (`@w:type`, `WD_HEADER_FOOTER` → a member) and `rId`
(`@r:id`, `XsdString` → string). python-docx names the first member `type_` with a
trailing underscore already (to dodge the `type` builtin) — **no** underscore
rotation applies (rotation is for LEADING underscores only), so the MATLAB member
is `type_` verbatim. Both attributes are REQUIRED: a missing `@w:type` or `@r:id`
on load raises `mat2doc:InvalidXmlError` (verbatim message), never a default.
`WD_HEADER_FOOTER` is python-docx's alias of `WD_HEADER_FOOTER_INDEX`; the MATLAB
alias class forwards `from_xml`/`to_xml` to the canonical enum. **`rId` is a plain
`r:id` string attribute only** — the separate-part header/footer relationship
wiring (`add_header_part`, `relate_to`) is P5-3b; this class does not resolve or
create relationships.

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_HdrFtrRef`*

---

(id-write-paths)=
## The geometry-write paths — byte-proven round-trips

The parse path is byte-neutral; the **novel** paths are the geometry / start-type /
titlePg writes on a real body `w:sectPr`. All are frozen as permanent references
(Gate-3, co-located `.gitattributes` `* binary` pins) and are **byte-identical to
python-docx**:

| scenario | ops | `word/document.xml` | reference |
|---|---|---|---|
| geometry write | `page_width=Twips(15840)`, `page_height=Twips(12240)`, `orientation=LANDSCAPE`, `top_margin=Twips(360)` | 1568 B, SHA `698367cd…` | `references\s0040` |
| LANDSCAPE only (w/h-swap guard) | `orientation=LANDSCAPE` | 1569 B, SHA `d165a628…` (w/h NOT swapped) | `references\s0041` |
| margin removed | `top_margin = []` (None) | 1535 B, SHA `b45b2c6b…` | (pinned) |
| start type | `start_type = EVEN_PAGE` | 1574 B, SHA `23f89f6b…` | (pinned) |
| titlePg | `titlePg_val = true` (`<w:titlePg/>`) | 1560 B, SHA `7b41c7a5…` | (pinned) |

Each is a full 17-part package byte-identical to the python-docx mirror (only
`word/document.xml` differs from the M1 default); the mutated subtree
`<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>` shows `@w:orient`
appended after `w`,`h` with no swap, and `cols`/`docGrid` untouched. **Zero new
D-numbers** — every equivalence leg is L1 byte-identical or value-exact.

(id-d-nsprefix-rewrite)=
## D-nsprefix-rewrite recurrence — dead-on-generation

During probe design a serialization divergence surfaced on **loose (un-rooted)**
sectPr elements carrying MULTIPLE header/footer references: lxml auto-numbers the
`r:` namespace prefix with a GLOBAL counter (`ns0`, `ns1`, `ns2`, …) while the
port reuses `ns0`. This is the pre-existing, **SIGNED (accepted)** permanent-L2
deviation **D-nsprefix-rewrite** (ledger: "dead-on-generation,
canonically-equivalent, foreign-file-only") — no new D-number. It is
**dead-on-generation** and reaches ZERO package bytes, proven three ways:

1. In every real package `r:` is declared at the `w:document` root, so
   header/footer references render with the in-scope `r:` prefix (`r:id="rId1"`).
2. **Rooted** (r: declared) multi-reference serialization is **byte-identical** on
   both sides —
   `<w:headerReference w:type="default" r:id="rId1"/><w:footerReference … r:id="rId2"/><w:headerReference w:type="first" r:id="rId3"/>`.
3. M1 17/17 + the geometry/LANDSCAPE/extended 17/17 + the parse-path round-trip
   are all L1 — no `ns0`/`nsN` ever appears in generated output.

The three reference-bearing probe legs are therefore built r:-ROOTED (faithful to
real API usage), and `probe_diff` is a clean MATCH.

---

(id-p5-2b)=
## P5-2b — `CT_HdrFtr`, the section block-element iterator, and the `iter_inner_content` un-stub

:::{note}
**★ Section oxml layer COMPLETE (core + hdr/ftr bodies + the section iterator).**
P5-2b adds the header/footer **PART root** `CT_HdrFtr` and the
`_SectBlockElementIterator` that partitions a body's block elements into sections,
and un-stubs `CT_SectPr.iter_inner_content` so it is now **live**. It registers
`w:hdr`/`w:ftr` — **M1-neutral** (separate-part roots absent from `default.docx`) —
so M1 stays 17/17 with **zero re-pins** and **zero new D-numbers**. What remains in
**Phase 5** is the API surface: the `Section`/`Sections` proxies (**P5-3a**) and the
`_Header`/`_Footer` separate-part rels (**P5-3b**).
:::

(id-ct_hdrftr)=
### `CT_HdrFtr` — the header/footer part root

**Syntax**

```matlab
h = mat2doc.oxml.OxmlElement("w:hdr");   % a CT_HdrFtr (registered for w:hdr AND w:ftr)
p = h.add_p();                            % new <w:p>, appended at end
elms = h.inner_content_elements;          % 1xN [CT_P | XmlElement(tbl)], document order
```

**Description**

`<w:hdr>` / `<w:ftr>` is the **root of a SEPARATE package part**
(`word/header1.xml`, `word/footer1.xml`, …) — never a child of `document.xml`. It
holds block content (`w:p` / `w:tbl` children) exactly as `CT_Body` holds the main
story, and a single class serves **both** the header and the footer part (registered
for both tags in `oxml/__init__.py` :124–:125).

**The `ZeroOrMore` descriptor family (tag-based, append-at-end).** `p` and `tbl`
are both `ZeroOrMore` with `successors=()`, generating the docx-form member set
`p_lst` / `new_p_` / `insert_p_` / `add_p_` / `add_p` (and the `tbl` analogues) — a
public `add_x` adder, **no** bare getter, **no** get-or-add, **no** remover
(D-delta-4). `successors=()` maps to `NO_SUCCESSORS` (append at end), mirroring the
`CT_Body` sentinel pattern.

(id-ct_hdrftr-ice)=
**`inner_content_elements` — the tag-based child union (CT_Tbl included as generic).**
The accessor is `xpath("./w:p | ./w:tbl")` — a **tag-based** child union, not
`isinstance` dispatch. `CT_Tbl` is not registered until **P6**, so a `w:tbl` child
resolves to a **generic `XmlElement`** (the `CT_Body.tbl_lst` precedent, live since
P2-3) and is **INCLUDED in document order**, never dropped, never crashed — only its
element *class* is generic. This class therefore ports **COMPLETE with zero stubs**:
no `notYetPorted` table branch belongs at the oxml tier (the only table-branch
decision is `Section.iter_inner_content`'s `isinstance` dispatch at P5-3a, whose
else-branch is owner P6-4a). A `w:p` **nested inside `w:ins`** (or any other wrapper)
is **NOT** included — the xpath is a fixed two-branch **child** union, not a
descendant scan. The return is the materialized `1×N` heterogeneous `XmlElement`
array (H9 — the Python `@property` returns a list); a **typed empty array** when no
p/tbl are present (never `[]` / None). H5: two reads return the same live child
handles.

**Gate-3 header-part byte proof.** A real python-docx v1.2.0 header part (a
paragraph + a 1×2 table + a paragraph, non-ASCII in both a paragraph and a cell) is
frozen as `references\s0043\header1.xml` (1842 B, SHA `e6abe568…b397fcb2`, a
co-located `.gitattributes` `* binary` pin). MATLAB `parse_xml` → `CT_HdrFtr` →
`inner_content_elements` iterated → `serialize_part_xml` reproduces the input
**byte-identical**, and that SHA equals **python-docx's own** parse→serialize
round-trip — extending the D-001 own-parser round-trip proof to a real header PART (a
fresh input class).

**Example**

```matlab
h = mat2doc.oxml.OxmlElement("w:hdr");   % a CT_HdrFtr
h.add_p();                                % append <w:p>
h.add_tbl();                              % append a generic <w:tbl>
elms = h.inner_content_elements;          % 1x2 [CT_P, XmlElement]
disp(numel(elms));                        % 2 (the tbl is INCLUDED, as a generic element)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::CT_HdrFtr`*

---

(id-sect-iterator)=
### `SectBlockElementIterator_` — partitioning a body into sections

**Description**

`_SectBlockElementIterator` (FLAG-3 → `SectBlockElementIterator_`) walks the
block-item elements (`w:p` / `w:tbl`) that **belong to one section**: the block
elements after the previous section's terminus up to and including this section's
terminal element. The static entry
`iter_sect_block_elements(sectPr)` returns them as a **materialized `1×N`
heterogeneous `XmlElement` array** in document order (`CT_P` for `w:p`, generic
`XmlElement` for `w:tbl`) — the Python `_iter_sect_block_elements` **generator**
collapses to a precomputed array (H9: consumers iterate with no tree mutation, so
laziness is unobservable). This is the surface `CT_SectPr.iter_inner_content` and
(at P5-3a) `Section.iter_inner_content` consume.

(id-sect-iterator-strategy)=
**The boundary strategy (section.py 454–537).** Get all block elements from the
start of the document to **and including** this section, then compute the count of
those that came from **prior** sections and skip that many, leaving only this
section's. `sectPrs` is every `w:sectPr` in document order; `sectPr_idx` is this
sectPr's position; `n_blks_to_skip` is `0` for the first section, else the block
count of the previous section's "in-and-above" set; the result is
`blocks-in-and-above(this)[n_blks_to_skip:]`. Three xpath shapes assemble the
"in-and-above" node-set, all built in `blocks_in_and_above_section_xpath_` verbatim
from section.py 502–518:

| selector | XPath | selects |
|---|---|---|
| `p_sect_term_block` | `./parent::w:pPr/parent::w:p` | the `w:p` a `sectPr` sits in |
| `body_sect_term` | `self::w:sectPr[parent::w:body]` | the final (body) `sectPr` |
| `pred_ps_and_tbls` | `preceding-sibling::*[self::w:p \| self::w:tbl]` | the p/tbl before the context node |

`p_sect_term_block` and `body_sect_term` are **mutually exclusive** (a `sectPr`
lives in a `w:pPr` OR directly in `w:body`), so exactly one shape contributes per
`sectPr`. The engine returns the union in **document order** (`docSortDedupe`,
matching lxml), which is the precondition the skip-count arithmetic depends on.

(id-sect-iterator-hazards)=
**Index and identity hazards.** H1 — `sectPrs.index(sectPr)` is Python 0-based; the
MATLAB `find` result is 1-based, so the first-section test `sectPr_idx == 0` maps to
`idx == 1`, the previous-section access `sectPrs[sectPr_idx - 1]` to
`sectPrs(idx - 1)` (base-invariant predecessor arithmetic), and the slice
`[n_blks_to_skip:]` to `(n_blks_to_skip + 1 : end)` (the 0→1 slice-start
conversion; an empty `1×0` result when `n_blks_to_skip == numel(blocks)`). H5 —
`sectPrs.index` is an element **identity** search (`find(sectPrs == sectPr, 1)` over
the Sealed `==`); a not-found search raises `mat2doc:ValueError` "sectPr is not in
list" (faithful to Python `list.index()`'s implicit contract, unreachable in
practice). Two Python perf optimizations are dropped as **unobservable** (H9): the
two compiled `etree.XPath` class attributes (the engine re-parses each call) and the
separate `count(...)` xpath (`numel` of the same node-set is value-identical).

**Gate-3 partition corpus.** The adversarial 4-document / 11-section corpus
(`references\s0044\probe.json`, frozen) matches the python-docx
`_SectBlockElementIterator` oracle **exactly**, and per document the concatenation of
all partitions equals the body's block children in order with **no drop and no
double-count of any boundary paragraph** — across every edge: an empty first section
(bare break-p → `[p|]`), a break paragraph carrying content, a table first / last in
a section, a body section ending in a table, and an empty body section (a typed
`1×0` array, no error at the slice boundary).

*Ported from python-docx v1.2.0: `src/docx/oxml/section.py::_SectBlockElementIterator`*

---

(id-iter-inner-content-unstub)=
### `CT_SectPr.iter_inner_content` — now live

The P5-2a stub (a clean `mat2doc:notYetPorted` naming `_SectBlockElementIterator`)
is **un-stubbed COMPLETELY** at P5-2b: it delegates one line to
`SectBlockElementIterator_.iter_sect_block_elements(obj)`. Tag-based throughout, so
there is no CT_Tbl issue at this tier — the un-stub is full, not partial.
`probe_diff` (s0042) confirms it resolves (`resolves = true`, no `notYetPorted`) and
yields the correct per-section partitions on a realistic 3-section body.

---

(id-hdrftr-registry)=
### The `w:hdr` / `w:ftr` registry rows — M1-neutral

Two rows are added in `oxml/__init__.py` source order, between the existing
`w:footerReference` (:123) and `w:headerReference` (:126):

| tag | class | `oxml/__init__.py` | role |
|---|---|---|---|
| `w:ftr` | `section.CT_HdrFtr` | :124 | the footer-part root |
| `w:hdr` | `section.CT_HdrFtr` | :125 | the header-part root |

`default.docx` has **17 parts and no header/footer part**; the strings
"header"/"footer" do not occur in its `[Content_Types].xml`, and no M1 part carries a
`<w:hdr>`/`<w:ftr>` tag. So these rows light up **only** when a header/footer part is
actually loaded (first materialized at P5-3b) — they touch no M1 parse path. The WP
is **byte-neutral AND flip-neutral**: no existing exact-class `XmlElement` pin can see
a `w:hdr`/`w:ftr` class (contrast the P5-2a `sectPr` flip, which re-pinned one test).
M1 stays **17/17** (`word/document.xml` re-derived `0e4dd503…`, unchanged) with
**zero re-pins**.

---

## Section oxml layer COMPLETE — Section/Sections API (P5-3a) next

The section oxml layer is now whole: the section-properties **element core**
(`CT_SectPr` + its 20-tag `TAG_SEQ`, the page-geometry / start-type / titlePg
accessors, the header/footer-reference surface, and the four children `CT_PageSz` /
`CT_PageMar` / `CT_SectType` / `CT_HdrFtrRef`), the header/footer **part root**
`CT_HdrFtr`, and the **section block-element iterator** — with
`CT_SectPr.iter_inner_content` live. Registering `w:sectPr`/… (P5-2a) is byte-neutral
on the M1-central `word/document.xml`; registering `w:hdr`/`w:ftr` (P5-2b) is
M1-neutral; every geometry-write path is byte-proven, the header-part round-trip is
byte-identical, and every iterator partition equals the python-docx oracle — with
**zero new D-numbers** across both WPs.

What remains in **Phase 5** is the **API surface**: **P5-3a** (`Section`/`Sections`
reading this `CT_SectPr` surface, plus the `Document.sections` / `add_section` +
`CT_Body.add_section_break` un-stubs — the section-authoring path; its
`Section.iter_inner_content` table branch carries a P6-4a debt) and **P5-3b**
(`_Header`/`_Footer` + the separate-part header/footer rels, which consume the
`Settings.odd_and_even_pages_header_footer` toggle from P5-1, and which pre-decide
the `BlockItemContainer` element-accessor refactor).
