---
title: "mat2doc.section — the Section/Sections API tier (Section · Sections · the add_section authoring path)"
---

# `mat2doc.section` — the Section / Sections API tier (`Section` + `Sections` + the `add_section` authoring path)

Ported from python-docx v1.2.0 `src/docx/section.py::Section` and
`src/docx/section.py::Sections` (in the NEW package `+mat2doc/+section/`), plus
the **C1 authoring un-stubs** they light up — `document.py::Document.sections` /
`Document.add_section` (`+mat2doc/+document/Document.m`) and
`oxml/document.py::CT_Body.add_section_break`
(`+mat2doc/+oxml/+document/CT_Body.m`). This is the **API surface** over the
P5-2a/P5-2b section-oxml layer: `Section` is a transparent proxy over an
already-registered `CT_SectPr`, `Sections` is the section sequence, and
`add_section` is the section-authoring byte path.

:::{note}
**★ PHASE 5 COMPLETE — `Header_`/`Footer_` + the hdr/ftr separate parts landed
at P5-3b.** This is the **fourth work package of Phase 5** (the sections tier).
It ports the user-facing `Section` / `Sections` **proxies** that read and write
the `CT_SectPr` geometry/type surface built at P5-2a, and
**un-stubs the section-authoring path** (`Document.sections` / `Document.add_section`
+ `CT_Body.add_section_break`) so a document can now grow sections. It adds **no
`register_element_cls` row and no serialization code** — equivalence is
**behavioral** (proxy value parity) plus serialized-bytes parity on the novel
`add_section` and property-write paths — so it is **byte-neutral** (M1 stays
17/17, **zero new D-numbers**). The final Phase-5 tier — the
**headers/footers** (`Header_`/`Footer_` and the separate-part hdr/ftr
relationships, **P5-3b**, which consume the
`Settings.odd_and_even_pages_header_footer` toggle from P5-1) — is now landed on
the **[headers & footers page](headers_footers.md)**, completing the sections
work and **Phase 5**.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The one-package layout

`section.py` (a top-level `docx/` module) lands in the single new package
`+mat2doc\+section` — following the established module→subpackage convention. It
collides with no existing package. The oxml element classes it reads
(`CT_SectPr` + children) live separately in `+mat2doc\+oxml\+section` (P5-2a/P5-2b).

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `section.py` | `+mat2doc\+section\` | `Section`, `Sections` |

**The tier boundary.** `Section`/`Sections` add **NO** oxml logic, **NO** registry
rows and **NO** serialization code. Every `Section` geometry/type accessor
delegates **one-to-one** to the same-named `CT_SectPr` member (which owns all the
H6/H3/H4/H10 logic — see [`oxml_section.md`](oxml_section.md)); `Sections`
reads the live `CT_Document.sectPr_lst`. So equivalence is **behavioral** (proxy
return values) plus **serialized-bytes** on the novel authoring / property-write
paths — never a new byte-registry fixture.

(id-c1-unstub)=
## The C1 authoring un-stub — `Document.sections` / `add_section` now live

Since P1-8 `Document.sections` and `Document.add_section` were
`mat2doc:notYetPorted` stubs (no section proxy, no author path). P5-3a **un-stubs
all three** members of the section-authoring chain, whose deps are now all live:

| member | Python | now delegates to |
|---|---|---|
| `Document.sections` | `document.py:206-209` | `Sections(self._element, self._part)` |
| `Document.add_section(start_type)` | `document.py:140-148` | `body.add_section_break()` → `start_type` → `Section(...)` |
| `CT_Body.add_section_break` | `oxml/document.py:51-71` | `CT_SectPr.clone` (P5-2a) + `CT_P.set_sectPr` (P4) |

`Document.block_width_` (`document.py:232-239`) goes transitively live too — its
`self.sections[-1]` now reaches the real `Sections` collection (the stub-era
placeholder `sections()(end)`, which assumed a plain array, is corrected to
`s.getitem_(-1)`). It is reached only via `add_table` (a P6 stub), so it has no
live caller yet, but it now computes a real value rather than raising. **The C1
un-stub adds no registry rows and moves not a single byte of the default
`word/document.xml`** — a bare `Document().save()` fires ZERO stubs and stays M1
17/17 byte-identical (`0e4dd503…`/1548 B, independently re-derived).

---

(id-section)=
## `Section`

**Syntax**

```matlab
d   = mat2doc.Document();
sec = d.sections.getitem_(0);                  % the document's one default section
sec.top_margin                                  % 914400 EMU (Length) — <w:pgMar w:top="1440"/>
sec.top_margin  = mat2doc.shared.Inches(1.5);   % write-through to the CT_SectPr
sec.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;   % appends @w:orient, NO w/h swap
sec.start_type                                  % WD_SECTION_START.NEW_PAGE (no w:type child)
```

**Description**

`Section` provides access to the section- and page-setup properties of one
document section. It is a **pure API/proxy tier** over an already-registered
`CT_SectPr` (P5-2a): each geometry/type accessor delegates one-to-one to the
`CT_SectPr` `@property` of the same name. In python-docx it is a **plain object**
— `class Section:` (`section.py:24`), NOT an `ElementProxy`, NOT a `StoryChild`
— so it is `classdef Section < handle` with default handle identity (Python
default object identity) and defines **no `eq`/`ne`** (contrast an `ElementProxy`
subclass, whose equality is element identity). Its two attributes — `_sectPr`
and `_document_part` — rotate to `sectPr_` / `document_part_` (design.md §2).

(id-section-geometry)=
**The twelve geometry / type accessors (Dependent get/set).** All are `Dependent`
properties delegating to the same-named `CT_SectPr` member, **except** the three
aliases whose names differ from their backing member:

| `Section` property | delegates to `CT_SectPr` | type |
|---|---|---|
| `bottom_margin` / `top_margin` / `left_margin` / `right_margin` | same name | `Length` \| `[]` |
| `gutter` | `gutter` | `Length` \| `[]` |
| `page_width` / `page_height` | same name | `Length` \| `[]` |
| `footer_distance` | `sectPr.footer` (**alias**) | `Length` \| `[]` |
| `header_distance` | `sectPr.header` (**alias**) | `Length` \| `[]` |
| `orientation` | `orientation` | `WD_ORIENTATION` |
| `start_type` | `start_type` | `WD_SECTION_START` |
| `different_first_page_header_footer` | `sectPr.titlePg_val` (**alias**) | `bool` |

Because the tier is a transparent pass-through, every hazard lives in `CT_SectPr`
and is documented on [`oxml_section.md`](oxml_section.md): H6/H3 (`Length | None`
in EMU, `[]` round-trip and `[]`-assign removes), H10 (`orientation`
`WD_ORIENTATION` with the **NO-w/h-swap** semantics — setting `LANDSCAPE` appends
`@w:orient="landscape"` and does **not** swap `@w:w`/`@w:h`; `start_type`
`WD_SECTION_START` identity setter), and H4 (`different_first_page_header_footer`
is a bool via the `titlePg` `[None, False]` breadth). This tier adds none of it.

(id-section-part)=
**`part` and the story-parent contract.** `part` (`section.py:218-220`) is a
zero-argument accessor returning the owning `DocumentPart` (a `StoryPart`). It is
the `ProvidesStoryPart` hook the paragraphs minted by `iter_inner_content` reach
through: `Section.iter_inner_content` wraps each block element as
`Paragraph(element, self)`, i.e. the `Section` is the block item's parent, and
`Paragraph` (a `StoryChild`) delegates `part` up to `self._parent.part` — so
`Section.part` satisfies the contract exactly as in python-docx.

(id-section-iic)=
**`iter_inner_content` — the CT_P→Paragraph wrap and the P6-4a table debt (C2).**
`iter_inner_content` (`section.py:157-163`) yields each `Paragraph` **or** `Table`
belonging to this section, in document order, over
`CT_SectPr.iter_inner_content` (P5-2b, LIVE — a heterogeneous array: a `CT_P` for
each `<w:p>`, a generic `XmlElement` for each `<w:tbl>`). Because `Paragraph` and
`Table` are distinct proxy types sharing no `matlab.mixin.Heterogeneous` base,
the result is a **`1×N` cell** in document order (mirroring
`Paragraph.iter_inner_content`). The `CT_P` → `Paragraph(element, obj)` wrapping
is ported **FULLY**; the non-`CT_P` branch (a `<w:tbl>` → `Table`) **raises**
`mat2doc:notYetPorted` naming owner **P6-4a** — `Table` is a Phase-6 proxy, and a
`<w:tbl>` is **NEVER silently dropped** (C2). H9: the Python generator is
materialized (no tree mutation during iteration, so laziness is unobservable);
H10: `isinstance(element, CT_P)` → `isa(element, "mat2doc.oxml.text.CT_P")`.

(id-section-hdrftr-stubs)=
**The six header/footer properties — now LIVE (P5-3b).** Six members return
`Header_`/`Footer_` proxies over the separate hdr/ftr parts: `even_page_footer`
(`section.py:61-68`), `even_page_header` (`70-77`), `first_page_footer`
(`79-86`), `first_page_header` (`88-95`), `footer` (`97-104`, `@lazyproperty`),
`header` (`135-142`, `@lazyproperty`) — three `WD_HEADER_FOOTER` index kinds
(`PRIMARY`/`EVEN_PAGE`/`FIRST_PAGE`). The underlying `CT_SectPr`
hdr/ftr-reference accessors (`get`/`add`/`remove` `headerReference` /
`footerReference`) are **LIVE** (P5-2a); **P5-3b** added the proxy layer, the
`HeaderPart`/`FooterPart` part classes and the separate-part wiring — see the
**[headers & footers page](headers_footers.md)**.

**Example**

```matlab
d   = mat2doc.Document();
sec = d.add_section();                                  % new last section (WD_SECTION.NEW_PAGE)
sec.top_margin  = mat2doc.shared.Inches(1);
sec.orientation = mat2doc.enum.section.WD_ORIENTATION.LANDSCAPE;
disp(sec.start_type == mat2doc.enum.section.WD_SECTION_START.NEW_PAGE);   % 1
disp(sec.part == d.part);                               % 1 (H5 — same DocumentPart)
first = d.sections.getitem_(0);                         % the first section carries the body paragraph
for item = first.iter_inner_content()
    disp(class(item{1}));                               % mat2doc.text.Paragraph
end
```

*Ported from python-docx v1.2.0: `src/docx/section.py::Section`*

---

(id-sections)=
## `Sections`

**Syntax**

```matlab
d    = mat2doc.Document();
secs = d.sections;             % a mat2doc.section.Sections
secs.len_()                    % number of sections
last = secs.getitem_(-1);      % Python sections[-1] (negative wrap)
mid  = secs.getitem_(struct("start", 0, "stop", 2, "step", []));   % sections[0:2]
for s = secs.to_array(); disp(s.start_type); end
```

**Description**

`Sections` is the sequence of `Section` objects, one per section in the document,
supporting length, iteration and indexed access. In python-docx it is
`class Sections(Sequence[Section])` (`section.py:256`) — again a **plain object**
holding the document root element and the `DocumentPart`, not an `ElementProxy`
(it has no single wrapped element that is its identity) — so it is
`classdef Sections < handle` with default handle identity and **no `eq`/`ne`**.
Its section list is `self._document_elm.sectPr_lst` (`CT_Document.sectPr_lst`,
LIVE) — an xpath-ordered array of all directly-accessible `<w:sectPr>`.

(id-sections-sequence)=
**The Sequence surface — explicit methods, the 0-based `getitem_` convention.**
The shared 1-based `()` collection base (`RedefinesParen`) is a **future work
package** (the standing VERIFY-COLLECTION flag). Per the established Mat2Doc
precedent (`TabStops`, `Styles`), the Python `Sequence` dunders are ported as
**explicit methods** keeping line-for-line fidelity:

| Python | MATLAB | note |
|---|---|---|
| `sections[key]` | `sections.getitem_(key)` | key is the **Python 0-based** int, or a slice `struct` |
| `for s in sections` | `for s = sections.to_array()` | `__iter__` → materialized `1×N` |
| `len(sections)` | `sections.len_()` | `__len__` |

`getitem_` takes the **Python 0-based** key (the `TabStops` precedent — NOT the
Mat2Ppt 1-based `Slides` convention; the toolboxes share no code by rule), then
hits the 1-based MATLAB `sectPr_lst` with an explicit `+1` at the single indexing
site (H1). A **negative** key counts from the end (`getitem_(-1)` → the last
section, `i + n`); an **out-of-range** key raises `mat2doc:IndexError` with the
verbatim CPython message **`list index out of range`**.

(id-sections-slice)=
**The slice overload.** Python `sections[i:j:k]` returns a `List[Section]`. The
slice is represented as a **struct** with fields `start` / `stop` / `step` (each
a scalar double or `[]` for `None`) — the interim currency until the
`RedefinesParen` base lands (which will accept a native MATLAB range). `getitem_`
detects a slice via `isstruct(key)` (mirroring `isinstance(key, slice)`), computes
the selected positions with a **faithful port of CPython `slice.indices(n)` +
`range(...)`** (sign-dependent bounds, negative clamping, exclusive stop), and
returns a `1×N` `Section` object array (the Python list comprehension); an empty
slice yields a `1×0` `Section` array. A **zero step** raises `mat2doc:ValueError`
with the verbatim CPython message **`slice step cannot be zero`**. Gate-3
re-derived a 12-case slice battery (`[0:2]`, `[:]`, `[::2]`, `[-2:]`, empty
`[2:1]`, reverse `[::-1]`, `[-1::-1]`, `[1::-2]`, `[0:3:2]`, clamped `[5:]` /
`[-5:2]`, `[:1:-1]`) identical to python-docx `list[i:j:k]`.

**H5 identity.** Every `getitem_` / `to_array` element mints a **fresh**
`Section` view of its `<w:sectPr>` (python-docx does not cache `Section` objects);
the wrapped `CT_SectPr` is the shared identity.

**Example**

```matlab
d = mat2doc.Document();
d.add_section(mat2doc.enum.section.WD_SECTION.NEW_PAGE);
d.add_section(mat2doc.enum.section.WD_SECTION.ODD_PAGE);
secs = d.sections;
disp(secs.len_());                       % 3
disp(secs.getitem_(-1).start_type == mat2doc.enum.section.WD_SECTION_START.ODD_PAGE);  % 1
try
    secs.getitem_(3);                    % out of range
catch e
    disp(e.identifier);                  % mat2doc:IndexError
end
```

*Ported from python-docx v1.2.0: `src/docx/section.py::Sections`*

---

(id-add-section)=
## The `add_section` authoring path — byte-proven

`Document.add_section(start_type)` (`document.py:140-148`, default
`WD_SECTION.NEW_PAGE`, H13) is the section-authoring path. Its faithful three-line
body is:

```matlab
new_sectPr = obj.element_.body.add_section_break();   % clone the sentinel into a new trailing w:p
new_sectPr.start_type = start_type;                   % set the requested start type
section = mat2doc.section.Section(new_sectPr, obj.part_);
```

**The intermediate-sectPr nesting (`CT_Body.add_section_break`,
`oxml/document.py:51-71`).** The body's previously-last `<w:sectPr>` (the
sentinel) is cloned into a **new trailing paragraph** — `add_p()` inserts the new
`<w:p>` **before** the body sentinel (the `p` descriptor carries
`successors=("w:sectPr",)`, H11), and `CT_P.set_sectPr` nests the clone inside
its `<w:pPr>` — so the clone (carrying the prior section's geometry) becomes an
**intermediate** break paragraph and the sentinel stays last, governing the new
final section. The clone-before-removal ordering is load-bearing: the intermediate
break keeps the prior section's header/footer references, then the hdr/ftr
references are removed from the **sentinel only** (so the new final section
"inherits" them from the prior section). The sentinel handle is returned (H5), its
`start_type` set, and the `Section` proxy wraps it.

**Byte scenarios (part-level after unzip, vs python-docx v1.2.0, all 17/17
byte-identical, frozen `references\s0046`–`s0051`).** `NEW_PAGE` removes the
`w:type` child (identity default); the other four start types write
`<w:type w:val="…"/>`:

| scenario | call(s) | `word/document.xml` SHA-256 (PY == MAT) | bytes |
|---|---|---|---|
| `add_section(NEW_PAGE)` | 1 add | `4d49edc1…189d2bda` | 1792 |
| `add_section(CONTINUOUS)` | 1 add | `9d8f6d74…e1efbafe5` | 1820 |
| `add_section(NEW_COLUMN)` | 1 add | `1fec6403…8de2b5e50b1` | 1820 |
| `add_section(EVEN_PAGE)` | 1 add | `8cc56f42…b871f553a9a` | 1818 |
| `add_section(ODD_PAGE)` | 1 add | `531c2a63…f861f936f0c76` | 1817 |
| 3-section chain (`ODD_PAGE`, `EVEN_PAGE`) | 2 adds | `f6cc256d…6b3875f6eb` | 2087 |
| property-write on `sections[0]` (`top_margin=Inches(1.5)`, `orientation=LANDSCAPE`, `page_width=Inches(11)`) | writes | `5f3a0d58…185a94dc86` | 1569 |
| full property surface on `sections[0]` + a newly-added section | add + writes | `b6c6af5e…236ab715` | 1893 |

A **multi-section round-trip** is byte-faithful too: MATLAB opens a python-built
3-section `.docx`, probes every section, and saves unchanged — byte-identical to
python-docx's own open/save (`references\s0051`). **Zero new D-numbers** — the
`add_section` break construction (including the intermediate-sectPr nesting and
the hdr/ftr-reference-removal branch) and every property write (including
`[]`-removal on a newly-added section) are byte-identical to python-docx.

---

## Section API live — headers/footers (P5-3b) landed → PHASE 5 COMPLETE

P5-3a lands the section **API surface**: the `Section` geometry/type proxy (all
twelve accessors + `part` + `iter_inner_content` with the P6-4a table debt), the
`Sections` sequence (the 0-based `getitem_` int/negative/slice, `to_array`,
`len_`, and the verbatim `IndexError`/`ValueError`), and the **section-authoring
path** (`Document.sections` / `add_section` + `CT_Body.add_section_break`, byte-
proven). It is behavioral + un-stub with **no registry rows**, byte-neutral (M1
17/17, `document.xml` `0e4dd503…` unchanged), and carries **zero new D-numbers**.
The final Phase-5 tier — **P5-3b**, `Header_`/`Footer_` and the separate-part
header/footer relationships (`HeaderPart`/`FooterPart`), which un-stub the six
`Section` header/footer members above, consume the P5-1
`odd_and_even_pages_header_footer` toggle, and complete the sections work — is
now landed on the **[headers & footers page](headers_footers.md)** and is
**COM-verified in real Word** (all four header/footer packages open clean).
**Phase 5 is COMPLETE.**
