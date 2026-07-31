---
title: "Headers & footers — the separate-part hdr/ftr tier (Section hdr/ftr surface · _Header/_Footer · HeaderPart/FooterPart)"
---

# Headers & footers — the separate-part header/footer tier

Ported from python-docx v1.2.0 `src/docx/section.py::_Header` / `::_Footer` /
`::_BaseHeaderFooter` (the section-side proxies, in the package
`+mat2doc/+section/`) and `src/docx/parts/hdrftr.py::HeaderPart` / `::FooterPart`
(the part classes, in `+mat2doc/+parts/`), plus the six
`section.py::Section` header/footer accessors they light up and the five
`parts/document.py::DocumentPart` un-stubs that mint and drop the parts. This is
**P5-3b — the FINAL Phase-5 work package**, and it completes the sections tier.

:::{note}
**★ PHASE 5 COMPLETE — sections + settings + headers/footers, COM-verified.**
With P5-3b the whole of Phase 5 is done: the document **settings** proxy (P5-1),
the **section geometry** oxml core `CT_SectPr` (P5-2a), the header/footer PART
root `CT_HdrFtr` + the section content walk (P5-2b), the **Section / Sections
API** + the `add_section` authoring path (P5-3a), and now **headers/footers as
separate parts** (P5-3b). Every P5 WP is byte-identical to python-docx v1.2.0
with **zero new D-numbers**, and the four runtime-added header/footer packages
open **clean in real Microsoft Word** (see [the COM milestone](#id-hdrftr-com)).
The cross-part hazard (risk-register #4 — headers/footers are the **first
runtime-added parts** in Mat2Doc) is cleared.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## Why headers/footers are different — the separate-part model

Every earlier authoring surface (paragraphs, runs, styles, sections) writes into
`word/document.xml` — a part that **already exists** in the template. A header or
footer is different: it lives in **its own part** (`word/header1.xml`,
`word/footer1.xml`, `word/header2.xml`, …) that does **not** exist until content
is first added. Defining a header for the first time is therefore the **first
runtime-added part** in Mat2Doc, and it moves four coordinated pieces of the
package at once:

| Piece | What P5-3b adds | Byte proof |
|---|---|---|
| The new part `word/header1.xml` | a fresh `w:hdr` root parsed from `templates/default-header.xml`, re-serialized through the D-001 serializer | `header1.xml` 1329 B, SHA `3c75200fabed` |
| `[Content_Types].xml` | a new `<Override>` entry for the header part — the **first runtime-added Override** | 1866 B, SHA `4a3612de6729` |
| `word/_rels/document.xml.rels` | a new `<Relationship>` (`rId9`, reltype `.../header`, target `header1.xml`) appended in **insertion order — no rId sort** (H11) | 1355 B, SHA `29b4d0f49e57` |
| `word/document.xml` | a `<w:headerReference>` inside the section's `<w:sectPr>` | 1597 B, SHA `e3b5e38706b3` |

All four were re-derived independently on the python-docx side and matched
**byte-for-byte** by Mat2Doc (Gate-3 scenario `s0052`, frozen). The new part
lands in the zip **after `word/numbering.xml` and before
`docProps/thumbnail.jpeg`** — identical entry order on both sides.

(id-section-hdrftr)=
## `Section` — the six header/footer accessors (now live)

Since P5-3a the six `Section` header/footer members were clean
`mat2doc:notYetPorted` stubs. P5-3b **un-stubs all six**. Each returns a
`Header_` or `Footer_` proxy bound to one of the three `WD_HEADER_FOOTER` index
kinds:

| `Section` member | proxy | `WD_HEADER_FOOTER` index | Python (`section.py`) |
|---|---|---|---|
| `header` | `Header_` | `PRIMARY` | `135-142` (`@lazyproperty`) |
| `footer` | `Footer_` | `PRIMARY` | `97-104` (`@lazyproperty`) |
| `even_page_header` | `Header_` | `EVEN_PAGE` | `70-77` |
| `even_page_footer` | `Footer_` | `EVEN_PAGE` | `61-68` |
| `first_page_header` | `Header_` | `FIRST_PAGE` | `88-95` |
| `first_page_footer` | `Footer_` | `FIRST_PAGE` | `79-86` |

**Caching, H5.** `header` and `footer` are `@lazyproperty` in python-docx — they
cache their proxy, so two reads return the **same** handle. The four
`even_page_*` / `first_page_*` members are **plain** `@property` — a **fresh**
proxy each access. Mat2Doc reproduces both regimes exactly (a private cache +
computed-flag for `header`/`footer`; a new object each call for the other four).

**Even/first-page headers need the section toggles.** An even-page header is only
*active* when the document carries `w:evenAndOddHeaders` (the P5-1
`Settings.odd_and_even_pages_header_footer` toggle); a first-page header is only
active when the section carries `w:titlePg`
(`Section.different_first_page_header_footer`, P5-2a). P5-3b writes the parts and
references faithfully whether or not those toggles are set — matching python-docx
byte-for-byte — and real Word honors them at render time (see the COM record).

## `_Header` / `_Footer` — the header/footer proxy

`_Header` (`section.py:436-479`) and `_Footer` (`section.py:390-433`) are thin
subclasses of `_BaseHeaderFooter` (`section.py:289-387`). Under the FLAG-3
leading-underscore rotation they become `mat2doc.section.Header_`,
`mat2doc.section.Footer_` and `mat2doc.section.BaseHeaderFooter_`. Each is a
`BlockItemContainer` — so, like `_Body` and a table cell, it exposes the
block-item surface (`paragraphs`, `add_paragraph`) — but with a lazy element
(see [the C3 seam](#id-c3-seam)).

(id-is-linked)=
**`is_linked_to_previous` — the get/set that adds or drops the part.** A
header/footer is *linked to previous* when it has **no explicit definition** and
therefore inherits the prior section's header/footer. The getter is
`not self._has_definition`; the setter (`section.py:316-325`) is the authoring
trigger:

- setting it to `false` on a linked header **adds** an (empty) header part +
  reference — the header now has its own definition;
- setting it to `true` on an unlinked header **drops** the header part +
  reference — it goes back to inheriting;
- setting it to its current value is a **no-op** (the guard is ported exactly).

```matlab
d = mat2doc.Document();
h = d.sections.getitem_(0).header;      % a mat2doc.section.Header_
disp(h.is_linked_to_previous);          % 1 (true) — linked, no part yet
h.is_linked_to_previous = false;        % ADD an empty header definition -> word/header1.xml
disp(h.is_linked_to_previous);          % 0 (false) — now has its own part
h.is_linked_to_previous = true;         % DROP it -> back to the pristine package
```

**`paragraphs` / `add_paragraph`.** Inherited from `BlockItemContainer` — a
header, like a document body, must contain **at least one** paragraph, so a new
header holds a single empty paragraph. Add content by writing into
`paragraphs(1)`; calling `add_paragraph()` by itself leaves an empty paragraph
above the new one. `paragraphs` is a zero-argument method returning a `1×N`
`Paragraph` array, so assign it to a variable before indexing:

```matlab
d = mat2doc.Document();
h = d.sections.getitem_(0).header;
h.is_linked_to_previous = false;        % give the header its own part
ps = h.paragraphs;                      % 1×N Paragraph — assign, then index
ps(1).text = "Hello, Header!";          % writes into word/header1.xml
d.save(fullfile(tempdir, "hdr.docx"));  % -> the byte-proven s0052 package
```

**`part` — the story-parent hook (overridden).** `_BaseHeaderFooter.part`
(`section.py:327-336`) overrides `StoryChild.part` to return the header/footer's
own `HeaderPart` / `FooterPart` (created if necessary), so a `Paragraph` minted
as `Paragraph(element, self)` resolves image insertion etc. against the header
part — **not** the document part. H5 handle identity: `h.part == h.part`.

*Ported from python-docx v1.2.0: `src/docx/section.py::_Header`,
`src/docx/section.py::_Footer`, `src/docx/section.py::_BaseHeaderFooter`*

(id-inherit-walk)=
## The inherit-walk — how a header resolves through prior sections

A section that does not define its own header **inherits** the header of the
section before it, walking back until a definition is found or the first section
is reached. This is `_get_or_add_definition` (`section.py:356-374`), a plain
three-case recursion (H9 — no generator):

1. **case-1** — this header *has* a definition → return its own part.
2. **case-2** — inherited, and a **prior section exists** → recurse into the
   prior section's `_Header`/`_Footer` (`preceding_sectPr`) and return *its*
   definition — the two sections share **one** header part.
3. **case-3** — inherited, but this **is the first section** → add a new
   definition here (there is nothing to inherit from).

```matlab
d = mat2doc.Document();
d.add_section();                              % a second section (WD_SECTION.NEW_PAGE)
d.add_section();                              % a third section
secs = d.sections;
h2 = secs.getitem_(1).header;                 % section 2's header
disp(h2.is_linked_to_previous);               % 1 — inherits section 1
s1h = secs.getitem_(0).header;                % define section 1's header (case-3 add)
s1h.is_linked_to_previous = false;
h3 = secs.getitem_(2).header;                 % section 3 resolves via the walk...
disp(string(h3.part.partname));               % /word/header1.xml — shares section 1's part
```

Gate-3 proved the walk over a three-section document — `linked [T,T,T]`; after
section-1 defines its header, `linked [T,F,T]` and sections 2 and 3 resolve to
the **same** `header1.xml` (case-2 recursion via `preceding_sectPr`); after
section-3 defines its own, it gets a **distinct** `header2.xml` (case-3 add on
the first-of-a-run) — matching python-docx step-for-step.

(id-c3-seam)=
## The C3 element-accessor seam — a byte-neutral base-class refactor

python-docx's `_BaseHeaderFooter` **never stores** `self._element`; it overrides
`_element` as a **lazy `@property`** returning
`self._get_or_add_definition().element` (`section.py:351-354`), so the header
part is created only on **first content access**. MATLAB **cannot redefine a
stored property in a subclass**, so a naive port (a plain stored `element_`)
would give the header the wrong element.

**Resolution.** `BlockItemContainer.element_` was refactored from a stored
property into a zero-argument **protected method** (the house
property-as-method convention already used for `StoryChild.part`,
`Section.part`, …), backed by a renamed concrete store `element_store_`:

- the base seam `element_()` returns the stored handle — every internal
  `BlockItemContainer` element read routes through `obj.element_()`;
- `BaseHeaderFooter_.element_()` **overrides** the seam to
  `get_or_add_definition_().element()` (lazy part-creation on first content
  access), reproducing the python-docx semantics exactly.

Because a zero-argument method is read with the **same** `obj.element_`
dot-syntax a property would use, the refactor is **byte-neutral for every
existing subclass**: `Body_` passes a concrete element to the constructor, and
`Document` is not a `BlockItemContainer` (it extends `ElementProxy`). This was
the sharpest structural risk of the WP — a live base-class touch — so it was
re-proven at the byte level:

| Guard | Scenario | Result |
|---|---|---|
| **M1** default save | `mat2doc.Document().save()` vs the frozen `s0001` | **17/17 L1**, `document.xml` `0e4dd503bc09` unchanged |
| **M2** hello-world | `add_heading×3 + add_paragraph + save` vs `s0033` | **17/17 L1**, `document.xml` `a71e550253b8` unchanged |

The seam perturbed **zero bytes** on either the default or the content-bearing
path.

## `HeaderPart` / `FooterPart` — the part classes

`HeaderPart` (`parts/hdrftr.py:36-53`) and `FooterPart` (`:16-33`) are
`StoryPart` subclasses (`mat2doc.parts.HeaderPart`,
`mat2doc.parts.FooterPart`). They **parse on load** and **re-serialize on save**
through the D-001 serializer, inheriting `blob` / `element` / `paragraphs` /
`add_paragraph` from `StoryPart < XmlPart` unchanged — they add only the `new`
factory (which reads the byte-identical template and picks the next free
partname via `next_partname("/word/header%d.xml")`) and their own static `load`.

**PartFactory flip (byte-neutral).** The content-type registry now maps
`WML_HEADER → HeaderPart` and `WML_FOOTER → FooterPart` (previously base
`XmlPart`). Because `HeaderPart` inherits `XmlPart.blob` and declares its own
static `load`, a reloaded header part's **type** changes but the emitted bytes
are identical to the previous dispatch. No M1 fixture loads a header/footer part
(the default template has none), so the flip is inert on the M1 sweep.

**Inherited-static trap.** MATLAB does not dispatch an inherited static method to
the calling subclass, so `HeaderPart`/`FooterPart` each declare their **own**
`load` — the faithful realization of Python's `cls`-bound `XmlPart.load`
(`opc/part.py:229-232`). Without it, the PartFactory flip would silently build a
base `XmlPart` and the newly created header would round-trip through the wrong
class. Gate-3 byte-proved the **reload path** (`s0058`): a python-docx-generated
file containing `header1.xml`/`footer1.xml` opened in Mat2Doc and saved
unchanged is **19/19 L1**.

**Faithful header/footer asymmetry (preserved verbatim).** Dropping a header
calls `DocumentPart.drop_header_part(rId)`; dropping a footer calls
`drop_rel(rId)` **directly** — there is no `drop_footer_part` on `DocumentPart`.
python-docx is written that way (`section.py:414-417` vs `:460-463`); the port
does not tidy the asymmetry.

*Ported from python-docx v1.2.0: `src/docx/parts/hdrftr.py::HeaderPart`,
`src/docx/parts/hdrftr.py::FooterPart`*

(id-hdrftr-com)=
## The Word COM oracle — real Office accepts the runtime-added parts

The four P5-3b packages were opened in **real Microsoft Word (COM
`Word.Application`, version 16.0, build 16.0.20228)** with alerts suppressed —
every one **opened silently, with no repair/recovery prompt and no dropped
content**, and Word round-tripped each through its own save:

| Package | Content | Word verdict |
|---|---|---|
| `s0052` | primary header "Hello, Header!" | OPEN-clean; Header **Primary** = "Hello, Header!" |
| `s0054` | header + footer (two runtime parts, `rId9`/`rId10`) | OPEN-clean; both present and placed |
| `s0055` | primary + even + first-page headers + `titlePg` | OPEN-clean; **`DifferentFirstPage` honored**; first-page header renders |
| `s0058` | python-generated header+footer, reloaded + resaved | OPEN-clean; "Reloaded Header"/"Reloaded Footer" intact |

For `s0055`, page 1 was exported to PDF: the **first-page** header ("First")
renders on the title page — confirming the first-page header is wired to render
on the first page, not the primary header. Full record:
`validation\mat2doc\com_verify_P5_hdrftr.md` (verdict **PASS**).

## ★ Phase 5 COMPLETE

Phase 5 delivered the sections tier end-to-end, in five work packages, each
byte-identical to python-docx v1.2.0 with **zero new D-numbers**:

| WP | Delivered |
|---|---|
| **P5-1** | `Settings` proxy + `CT_Settings` (the `odd_and_even_pages_header_footer` toggle) — [settings page](settings.md) |
| **P5-2a** | `CT_SectPr` + page geometry (size/orientation/margins) — [oxml section page](oxml_section.md) |
| **P5-2b** | `CT_HdrFtr` (the `w:hdr`/`w:ftr` part root) + the section content walk |
| **P5-3a** | the `Section` / `Sections` API + the `add_section` authoring path — [section API page](section_api.md) |
| **P5-3b** | headers/footers as separate parts (this page) — the first runtime-added parts |

Throughout Phase 5 **M1 stayed 17/17** (every WP re-ran the byte-neutrality
sweep), and the cross-part hazard (risk-register #4) is **cleared and
COM-verified**. Next is **Phase 6 (tables)** — where `CT_Tc` (the table cell,
P6-3a) is the hardest WP and P5-3a's `Section.iter_inner_content` `w:tbl` debt
(owner P6-4a) is discharged.
