---
title: "mat2doc.enum.{text,section,dml,style,table,shape} — the concrete WD_*/MSO_* enums"
---

# `mat2doc.enum.{text,section,dml,style,table,shape}` — the concrete enums

Ported from python-docx v1.2.0 `src/docx/enum/text.py`, `section.py`, `dml.py`,
`style.py`, `table.py`, and `shape.py` — the **19 concrete `WD_*` / `MSO_*`
enumerations** built on the P3-1 `BaseEnum` / `BaseXmlEnum` machinery, plus the
`WD_BREAK_TYPE.TEXT_WRAPPING` member-alias and the **11 module-level class
aliases**. These are the value-layer constants the P4 (text/run/paragraph/style),
P5 (section), P6 (table), and P7 (inline-shape) element classes read and write —
alignment, break type, highlight color, underline style, line spacing, tab
alignment/leader, header-footer index, page orientation, section start, color
type, theme color, built-in style name, style type, cell vertical alignment, row
height rule, table alignment/direction, and inline-shape type. Like the base tier
they **emit no serialized output of their own**; each supplies a fixed member set
and (for the `BaseXmlEnum` subclasses) the `from_xml` / `to_xml` translation its
consumers call.

The tier lands across two work packages: **P3-3** ported the `text` / `section` /
`dml` enums (documented first below), and **P3-4** — the final P3 WP — added the
`style` / `table` / `shape` enums (documented under
[The style / table / shape enums](#style-table-shape-enums)). With P3-4
merged, **Phase 3 is complete** (see the [Phase 3 complete](#phase-3-complete)
note).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. This page documents the *concrete* enum
tier; the `BaseEnum` / `BaseXmlEnum` base machinery it stands on is on the
[enumeration base tier](enums.md) page.
:::

## Package layout — per-module subpackages

The docx `enum/` module structure is mirrored: each concrete enum lands in a
subpackage named for its source module, so consumers reference the
fully-qualified `mat2doc.enum.<module>.<NAME>`:

| Source module | Mat2Doc subpackage | fully-qualified reference |
|---|---|---|
| `src/docx/enum/text.py` | `+mat2doc\+enum\+text\` | `mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT` |
| `src/docx/enum/section.py` | `+mat2doc\+enum\+section\` | `mat2doc.enum.section.WD_ORIENTATION` |
| `src/docx/enum/dml.py` | `+mat2doc\+enum\+dml\` | `mat2doc.enum.dml.MSO_THEME_COLOR_INDEX` |
| `src/docx/enum/style.py` | `+mat2doc\+enum\+style\` | `mat2doc.enum.style.WD_BUILTIN_STYLE` |
| `src/docx/enum/table.py` | `+mat2doc\+enum\+table\` | `mat2doc.enum.table.WD_TABLE_ALIGNMENT` |
| `src/docx/enum/shape.py` | `+mat2doc\+enum\+shape\` | `mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE` |

This **differs from Mat2Ppt's flat `+enum\`** — a deliberate per-module layout
per the P3-3 brief. The P3-1 base stays at `+mat2doc\+enum\+base\`.

## The 12 `text` / `section` / `dml` enums (108 members) — P3-3

Each row is a distinct enumeration; the full member set (108 members total) was
frozen and proven **byte-identical to the python-docx oracle** at Gate-3 (`s0018`
`probe_diff`, 498/498 facts, exit 0), including every member's `name`,
`ms_api_value`, `xml_value`, and declaration order (order is behavioural — it
drives `from_xml` first-match resolution).

| Enum | Subpackage | Base | Members | Notes |
|---|---|---|---|---|
| `WD_PARAGRAPH_ALIGNMENT` | text | `BaseXmlEnum` | 9 | alias `WD_ALIGN_PARAGRAPH`; value gap at 6 (faithful) |
| `WD_BREAK_TYPE` | text | *plain value classdef* | 10 (+1 member-alias) | plain `enum.Enum`, no `xml_value`; `TEXT_WRAPPING` aliases `LINE_CLEAR_ALL`; alias `WD_BREAK` |
| `WD_COLOR_INDEX` | text | `BaseXmlEnum` | 18 | `INHERITED` (−1) has `xml_value=None`; alias `WD_COLOR` |
| `WD_LINE_SPACING` | text | `BaseXmlEnum` | 6 | `SINGLE`/`ONE_POINT_FIVE`/`DOUBLE` carry `xml_value="UNMAPPED"` |
| `WD_TAB_ALIGNMENT` | text | `BaseXmlEnum` | 10 | `CLEAR`/`END`/`NUM`/`START` in the 101–104 block |
| `WD_TAB_LEADER` | text | `BaseXmlEnum` | 6 | — |
| `WD_UNDERLINE` | text | `BaseXmlEnum` | 19 | `INHERITED` (−1) has `xml_value=None`; value gaps faithful |
| `WD_HEADER_FOOTER_INDEX` | section | `BaseXmlEnum` | 3 | alias `WD_HEADER_FOOTER` |
| `WD_ORIENTATION` | section | `BaseXmlEnum` | 2 | alias `WD_ORIENT` |
| `WD_SECTION_START` | section | `BaseXmlEnum` | 5 | alias `WD_SECTION` |
| `MSO_COLOR_TYPE` | dml | `BaseEnum` | 3 | no XML mapping (`RGB` 1 / `THEME` 2 / `AUTO` 101) |
| `MSO_THEME_COLOR_INDEX` | dml | `BaseXmlEnum` | 17 | **docx≠pptx** full-word tokens; `NOT_THEME_COLOR="UNMAPPED"`; no `MIXED`; alias `MSO_THEME_COLOR` |

`WD_PARAGRAPH_ALIGNMENT`, `WD_TAB_ALIGNMENT`, and `WD_HEADER_FOOTER_INDEX` also
carry the same `from_xml` / `to_xml` illustrative examples shown on the
[base-tier page](enums.md) — those examples exercise the *shape* every
`BaseXmlEnum` subclass here exposes.

## The `WD_BREAK_TYPE` member-alias idiom

`WD_BREAK_TYPE` is the one enum in this tier that is **not** a `BaseXmlEnum`: in
python-docx it is a plain `enum.Enum` whose members carry a bare integer value
(no `xml_value`, no per-member docstring, no `from_xml` / `to_xml`). It ports as
a **plain value classdef** with an `enumeration` block and a single immutable
`value (1,1) int32` property — the PROG_ID plain-enum precedent from Mat2Ppt.

Its source declares `LINE_CLEAR_ALL = 11` **and** `TEXT_WRAPPING = 11`
(text.py:80, 86). Python's `enum` makes the second same-valued member an
**alias** of the first — `WD_BREAK_TYPE.TEXT_WRAPPING is
WD_BREAK_TYPE.LINE_CLEAR_ALL` is `True`, its `.name` is `"LINE_CLEAR_ALL"`, and
iteration / `__members__` list only the 10 canonical members.

MATLAB `enumeration` members compare `==` **by member identity, not by the
`value` property** (empirically verified: two distinct members that both carry
value 11 are *not* `==`). Declaring `TEXT_WRAPPING` as a second enumeration
member would therefore make it a distinct object that is neither `==` to nor
named the same as `LINE_CLEAR_ALL` — a real divergence. So the alias is realized
exactly as Python realizes it — as a **class attribute pointing at the canonical
member**:

```matlab
properties (Constant)
    TEXT_WRAPPING = mat2doc.enum.text.WD_BREAK_TYPE.LINE_CLEAR_ALL
end
```

Consequences (all oracle-matched at Gate-3):

- `WD_BREAK_TYPE.TEXT_WRAPPING` returns the `LINE_CLEAR_ALL` member itself:
  `string(...)` is `"LINE_CLEAR_ALL"`, `... == LINE_CLEAR_ALL` is true, `.value`
  is 11.
- `... == LINE` is **false** — `LINE` is value 6, a different member (the P3-1
  brief's "aliases LINE" was a factual slip; the source aliases `LINE_CLEAR_ALL`).
- `enumeration('mat2doc.enum.text.WD_BREAK_TYPE')` yields **10** members (the
  alias is excluded from iteration, matching Python).

The consumer `Run.add_break` (`docx/text/run.py`, ported at P4-3) maps a member
to `(w:br type, clear)` attribute values by **identity** lookup, so member
identity — not the integer — is the load-bearing property the alias idiom
protects.

## `None`-valued `INHERITED` members — the P3-1 `from_xml(None)` linkage

Two members carry a Python `None` `xml_value`: **`WD_COLOR_INDEX.INHERITED`**
(−1) and **`WD_UNDERLINE.INHERITED`** (−1). In the `enumeration` declaration the
`None` is passed as `string(missing)`; the P3-1 `BaseXmlEnum.asXmlVal_`
normalizes it to a `<missing>` string (H3 tri-state — `None` → `<missing>`, `""`
→ real `""`, token → real string).

These members are **reachable** only because the docx `from_xml` has **no
None-short-circuit** (the load-bearing P3-1 Delta 1) — it is a straight
None-tolerant equality scan where `None == None` matches:

| call | result | Python equivalent |
|---|---|---|
| `WD_COLOR_INDEX.from_xml([])` | `INHERITED` | `from_xml(None)` → `INHERITED` |
| `WD_UNDERLINE.from_xml([])` | `INHERITED` | `from_xml(None)` → `INHERITED` |
| `WD_COLOR_INDEX.to_xml(INHERITED)` | raises `mat2doc:ValueError` `"WD_COLOR_INDEX.INHERITED has no XML representation"` | falsy-`xml_value` guard (Delta 2) |

The MATLAB `None` argument is `[]` (design.md §2 None sentinel); `asXmlVal_([])`
→ `<missing>`, so `from_xml([])` is the faithful `from_xml(None)` call. `[]` is
the only None form — there is no `noneArg()` convenience.

**`"UNMAPPED"` is NOT `None`** and not `""` — it is a real, truthy placeholder
string. `WD_LINE_SPACING.{SINGLE,ONE_POINT_FIVE,DOUBLE}` and
`MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR` carry it; `to_xml` returns it (the guard
passes), and `from_xml("UNMAPPED")` resolves to the **first** member carrying it
(`WD_LINE_SPACING.SINGLE`, `MSO_THEME_COLOR_INDEX.NOT_THEME_COLOR`) —
first-declared-wins, the reason declaration order is behavioural.

## `MSO_THEME_COLOR_INDEX` — the docx-vs-pptx delta

`MSO_THEME_COLOR_INDEX` shares its name with a Mat2Ppt enum but is **not** the
same enum — the Mat2Ppt version must **not** be copied. The docx member set
differs on three counts, all confirmed against the docx oracle:

| Aspect | Mat2Ppt (pptx) | Mat2Doc (docx v1.2.0) |
|---|---|---|
| xml_value tokens | abbreviations (`bg1`, `dk1`, `hlink`, `lt1`, `tx1`) | **full words** (`background1`, `dark1`, `hyperlink`, `light1`, `text1`) |
| `NOT_THEME_COLOR` xml_value | `""` | **`"UNMAPPED"`** (a truthy string) |
| `MIXED` member | present (`MIXED = -2`) | **absent** |
| member count | (pptx count) | **17** |

Because `NOT_THEME_COLOR` carries a truthy `"UNMAPPED"`,
`MSO_THEME_COLOR_INDEX.to_xml(NOT_THEME_COLOR)` (and `to_xml(0)`) returns
`"UNMAPPED"` (the `if not xml_value` guard passes), where the pptx `""` would
raise. `from_xml("")` still raises
`"MSO_THEME_COLOR_INDEX has no XML mapping for ''"` — the empty string is a real
value, not the `None` sentinel.

## The 7 module-level class aliases

The three source modules define **7** module-level aliases (`ALIAS = Canonical`).
MATLAB has no class aliasing, so each alias is a separate classdef whose
`properties (Constant)` re-export the canonical enumeration's members
(`ALIAS.X = Canonical.X`) and whose `methods (Static)` forward `from_xml` /
`to_xml` to the canonical class. Because a Constant property holds the **same
member object**, `Alias.X == Canonical.X` is identity-true, and `isa`,
`from_xml` / `to_xml`, and cross-name comparisons all behave as one enumeration.

| Alias | Canonical | source line | statics forwarded |
|---|---|---|---|
| `WD_ALIGN_PARAGRAPH` | `WD_PARAGRAPH_ALIGNMENT` | text.py:67 | `from_xml` / `to_xml` |
| `WD_BREAK` | `WD_BREAK_TYPE` | text.py:89 | none (plain enum, no XML) |
| `WD_COLOR` | `WD_COLOR_INDEX` | text.py:156 | `from_xml` / `to_xml` |
| `WD_HEADER_FOOTER` | `WD_HEADER_FOOTER_INDEX` | section.py:27 | `from_xml` / `to_xml` |
| `WD_ORIENT` | `WD_ORIENTATION` | section.py:52 | `from_xml` / `to_xml` |
| `WD_SECTION` | `WD_SECTION_START` | section.py:86 | `from_xml` / `to_xml` |
| `MSO_THEME_COLOR` | `MSO_THEME_COLOR_INDEX` | dml.py:103 | `from_xml` / `to_xml` |

`WD_BREAK` re-exports the `TEXT_WRAPPING` member-alias too, so
`string(WD_BREAK.TEXT_WRAPPING)` is `"LINE_CLEAR_ALL"` — matching Python
`WD_BREAK.TEXT_WRAPPING`. `WD_BREAK` carries no static methods because its
canonical `WD_BREAK_TYPE` is a plain enum with no XML mapping.

```matlab
% Every alias member IS the canonical member (identity), so:
mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER == ...
    mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.CENTER      % true
mat2doc.enum.dml.MSO_THEME_COLOR.to_xml( ...
    mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_6)           % "accent6"
```

(style-table-shape-enums)=
## The style / table / shape enums (P3-4)

**P3-4** — the final P3 WP — adds the remaining seven concrete enums from
`style.py`, `table.py`, and `shape.py`, plus **four** module-level class aliases,
across the new `+enum\{+style,+table,+shape}` subpackages (11 files). It closes
boundary-audit item **C2** (the previously orphaned `shape.py`) and **completes
Phase 3**. Every member was frozen and proven **byte-identical to the python-docx
oracle** at Gate-3 (`s0019` `probe_diff`, **598/598 facts, exit 0**), the crux
being **`WD_BUILTIN_STYLE`'s 132 members** — the largest single member set in the
docx enum tier, and the input to the P4-7 styles layer.

| Enum | Subpackage | Base | Members | Notes |
|---|---|---|---|---|
| `WD_BUILTIN_STYLE` | style | `BaseEnum` | **132** | built-in `WdBuiltinStyle` names; all values negative, all distinct (no internal alias); alias `WD_STYLE` |
| `WD_STYLE_TYPE` | style | `BaseXmlEnum` | 4 | `paragraph`/`character`/`numbering`(=LIST)/`table` |
| `WD_CELL_VERTICAL_ALIGNMENT` | table | `BaseXmlEnum` | 4 | `top`/`center`/`bottom`/`both`; alias `WD_ALIGN_VERTICAL` |
| `WD_ROW_HEIGHT_RULE` | table | `BaseXmlEnum` | 3 | `auto`/`atLeast`/`exact`(=EXACTLY); alias `WD_ROW_HEIGHT` |
| `WD_TABLE_ALIGNMENT` | table | `BaseXmlEnum` | 3 | `left`/`center`/`right`; `LEFT` docstring `"Left-aligned"` (no trailing period, faithful) |
| `WD_TABLE_DIRECTION` | table | `BaseEnum` | 2 | `LTR` 0 / `RTL` 1 (no XML mapping) |
| `WD_INLINE_SHAPE_TYPE` | shape | *plain value classdef* | 5 | plain `enum.Enum`; `NOT_IMPLEMENTED` = **−6** (negative int32); alias `WD_INLINE_SHAPE` |

**153 members total** across the seven enums (132 + 4 + 4 + 3 + 3 + 2 + 5),
matched in declaration order — declaration order is behavioural for the
`BaseXmlEnum` `from_xml` first-match scan. No member of the four `BaseXmlEnum`
enums carries a `None` / `""` `xml_value`, so the H3 tri-state machinery (P3-1) is
present but unexercised here; `WD_STYLE_TYPE`, `WD_CELL_VERTICAL_ALIGNMENT`,
`WD_ROW_HEIGHT_RULE`, and `WD_TABLE_ALIGNMENT` round-trip both directions over
every member (member arg and int arg — `to_xml(1)→"paragraph"`,
`to_xml(101)→"both"`, `to_xml(2)→"exact"`, `to_xml(0)→"left"`).

### `WD_BUILTIN_STYLE` — the 132-member `BaseEnum`

`WD_BUILTIN_STYLE` is a `BaseEnum` (no XML mapping, no `from_xml`/`to_xml`): each
member's integer value is the negative `WdBuiltinStyle` MS-API constant, used to
select a built-in Word style (e.g. `document.styles(WD_STYLE.BODY_TEXT)`). The
132 members are **not re-listed here** — spot values pin the set: `NORMAL` −1,
`HEADING_1`..`HEADING_9` = −2..−10, `BODY_TEXT` −67, `BLOCK_QUOTATION` −85,
`BOOK_TITLE` −265 (most-negative), `INDEX_HEADING` −34 (whose docstring
`"Index Heading"` is the one with **no** trailing period), `TOC_9` −28. All 132
carry **distinct** integer values (the oracle iterates 132 canonical members — a
Python enum would collapse any duplicate-valued member into an alias, and none
are collapsed), so — unlike `WD_BREAK_TYPE` — **no internal member-alias
`Constant` is required**; all 132 are declared as ordinary `enumeration` members.
`int(member)` ports as `double(member.value)`, and `str(member)` as `str_()` =
`"NAME (value)"` (e.g. `"BODY_TEXT (-67)"`).

### The `WD_INLINE_SHAPE_TYPE` plain-enum idiom (C2 fold-in)

`WD_INLINE_SHAPE_TYPE` — the boundary-audit **C2** fold-in from `shape.py` — is,
like `WD_BREAK_TYPE`, **not** a `Base(Xml)Enum`: in python-docx it subclasses a
plain `enum.Enum` (bare `NAME = <int>`, no `xml_value`, no per-member docstring,
no `from_xml` / `to_xml`). It ports as a **plain value classdef** with an
`enumeration` block and a single immutable `value (1,1) int32` property — the same
PROG_ID plain-enum precedent `WD_BREAK_TYPE` uses. Its five members are `CHART`
12, `LINKED_PICTURE` 4, `PICTURE` 3, `SMART_ART` 15, and `NOT_IMPLEMENTED` **−6**
(negative — stored `int32`, fine).

Where `WD_BREAK_TYPE` needed a `TEXT_WRAPPING` `Constant` aliasing
`LINE_CLEAR_ALL` (both value 11), `WD_INLINE_SHAPE_TYPE` has **no**
duplicate-valued member — the five values are distinct — so declaring all five as
ordinary `enumeration` members is exactly faithful, with **no internal alias**.
Iteration yields 5 members, matching Python.

### The 4 module-level class aliases

`style.py`, `table.py`, and `shape.py` define **four** module-level aliases
(`ALIAS = Canonical`), realized the same way as the P3-3 aliases: a separate
classdef whose `properties (Constant)` re-export the canonical members
(`ALIAS.X = Canonical.X`, the same member object → identity `==` true) and — for
the `BaseXmlEnum` aliases only — whose `methods (Static)` forward `from_xml` /
`to_xml` to the canonical class.

| Alias | Canonical | source line | statics forwarded |
|---|---|---|---|
| `WD_STYLE` | `WD_BUILTIN_STYLE` | style.py:423 | none (`BaseEnum` — 132 Constant re-exports, no XML) |
| `WD_ALIGN_VERTICAL` | `WD_CELL_VERTICAL_ALIGNMENT` | table.py:48 | `from_xml` / `to_xml` |
| `WD_ROW_HEIGHT` | `WD_ROW_HEIGHT_RULE` | table.py:82 | `from_xml` / `to_xml` |
| `WD_INLINE_SHAPE` | `WD_INLINE_SHAPE_TYPE` | shape.py:19 | none (plain enum — 5 Constant re-exports, no XML) |

`WD_STYLE` and `WD_INLINE_SHAPE` carry **no** static methods because their
canonical enums (`BaseEnum` and plain-enum respectively) expose no XML mapping —
the alias faithfully mirrors the kind of the enum it re-exports.

```matlab
% Every alias member IS the canonical member (identity), so:
mat2doc.enum.style.WD_STYLE.BODY_TEXT == ...
    mat2doc.enum.style.WD_BUILTIN_STYLE.BODY_TEXT          % true
mat2doc.enum.table.WD_ALIGN_VERTICAL.from_xml("top")      % TOP
mat2doc.enum.table.WD_ROW_HEIGHT.to_xml( ...
    mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY)        % "exact"
mat2doc.enum.shape.WD_INLINE_SHAPE.NOT_IMPLEMENTED.value  % -6
```

(phase-3-complete)=
## Phase 3 complete

With P3-4 merged, **Phase 3 is complete**: the enum base tier (P3-1
`BaseEnum` / `BaseXmlEnum`), the `ST_*` simple-type validators (P3-2), and **all**
concrete enum content (P3-3 `text`/`section`/`dml` + P3-4 `style`/`table`/`shape`)
are ported, audited, and package-equivalence-validated — **19 concrete enums,
261 members, 11 module aliases, 0 new D-numbers** across the whole of P3. Together
with P3-2's [simple-type validators](simpletypes.md), the enum + simple-type
value surface beneath the element tree is **ready for P4's oxml/text layer** (and
P5/P6/P7), where these constants become the attribute (de)serializers the
`CT_*` classes read and write. Boundary item **C2** (`shape.py`) is closed.

## Deviation posture — 0 new D-numbers

Both P3-3 and P3-4 are pure concrete value/behaviour content on the P3-1 base —
**no serialized OOXML**, so no L0–L3 ladder leg applies. The only standing convention it
exercises is **D-005** (adopt-only) — the `mat2doc:ValueError` identifier on the
`from_xml` / `to_xml` error paths, with byte-verbatim message strings. Gate-3
proved every probe fact byte-identical (`probe_diff` exit 0) against python-docx's
own enum sources — **498/498** for P3-3 (`enum/{text,section,dml}.py`) and
**598/598** for P3-4 (`enum/{style,table,shape}.py`) — including all members in
declaration order (108 + 153 = 261), the `from_xml(None)` / `UNMAPPED` /
`empty≠None` legs, the `WD_BREAK_TYPE` member-alias, the per-member
`from_xml`/`to_xml` round-trips and byte-exact `mat2doc:ValueError` error paths,
and all 11 module aliases. **No new D-number, no new ledger row.** The `doc`
member (per-member docstrings) is stored for fidelity but is non-behavioral (only
the unported `DocsPageFormatter` reads it — the P3-1 `VERIFY-E4` precedent).

---

*Ported from python-docx v1.2.0: `src/docx/enum/text.py`, `src/docx/enum/section.py`,
`src/docx/enum/dml.py`, `src/docx/enum/style.py`, `src/docx/enum/table.py`,
`src/docx/enum/shape.py` — the concrete `WD_*` / `MSO_*` enumerations and their
module aliases.*
