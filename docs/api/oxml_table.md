---
title: "mat2doc.oxml.table — the table oxml layer (7 LEAF classes + the property containers CT_TblPr · CT_TblPrEx · CT_TrPr · CT_Row)"
---

# `mat2doc.oxml.table` — the table oxml layer (the 7 leaves + the property containers)

Ported from python-docx v1.2.0 `src/docx/oxml/table.py` (in the NEW package
`+mat2doc/+oxml/+table/`) — the **seven LEAF element classes** of the table
object model: the width union `CT_TblWidth`, the row-height leaf `CT_Height`, the
column-grid pair `CT_TblGrid` / `CT_TblGridCol`, the layout-mode leaf
`CT_TblLayoutType`, the cell vertical-alignment leaf `CT_VerticalJc`, and the
vertical-merge leaf `CT_VMerge` — plus the **seven `register_element_cls` rows**
they require (`src/docx/oxml/__init__.py` :171, :174, :175, :181, :183, :185,
:186).

:::{note}
**★ Tables tier — the leaves (P6-1) AND the property containers (P6-2) are done;
the CT_Tc merge engine (P6-3a) is next.** **P6-1** (the first Phase-6 WP) ported
the seven table *leaf* elements — attribute-only (or single-list) classes with no
dependence on any other table `CT_*`. **P6-2** then added the **four property
containers** that consume those leaves: `CT_TblPr` (table properties — alignment /
autofit / style), `CT_TblPrEx` (property exceptions), `CT_TrPr` (row properties —
grid skips / row height), and `CT_Row` (the `<w:tr>` row element). P6-2 is the
**first NON-neutral table WP**: registering `w:tblPr` puts it on the **live**
`word/styles.xml` parse path (the 100 `<w:tblPr>` nodes inside the shipped table
styles now transit `CT_TblPr`), so it is a byte-critical registry-flip WP (the
P4-6 pattern) — proven byte-neutral by the M1 17/17 `styles.xml` `02d71a68…`
match. The remaining seven P6-2 tags (`w:tblStyle` / `w:tblPrEx` / `w:tr` /
`w:trPr` / `w:gridAfter` / `w:gridBefore`) have zero occurrences in `default.docx`
and never transit on M1. What remains: the **CT_Tc merge engine** (gridSpan /
vMerge, the block-item-container seam) at **P6-3a — the hardest WP of the project,
with its own dedicated pre-launch plan-audit**, `CT_TcPr` / `CT_Tbl` at **P6-3b**,
and the `Table` / `_Rows` / `_Cell` API at **P6-4a / P6-4b**.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The one-package layout

`oxml/table.py` (a `docx/oxml/` module) lands in the single new package
`+mat2doc\+oxml\+table` — following the established module→subpackage convention.
It collides with no existing package. Only the seven leaf classes land here at
P6-1; the container classes from the same module land in the same package at
P6-2 / P6-3.

| Python `src/docx/...` | MATLAB | symbols (P6-1) |
|---|---|---|
| `oxml/table.py` | `+mat2doc\+oxml\+table\` | `CT_TblWidth`, `CT_Height`, `CT_TblGrid`, `CT_TblGridCol`, `CT_TblLayoutType`, `CT_VerticalJc`, `CT_VMerge` |

**Deferred to P6-2 / P6-3 (cleanly, not stubbed here):** `CT_Row`, `CT_Tbl`,
`CT_TblPr`, `CT_TblPrEx`, `CT_TrPr`, `CT_TcPr`, `CT_Tc` — the container classes
and their accessors (`grid_span` / `vAlign_val` / `vMerge_val` / `width`, the
`merge` machinery, `tr_idx`, `iter_tcs`, `autofit`, `alignment`, `new_tbl`, …).
Nothing table-critical is stubbed inside a P6-1 leaf; each leaf ports **complete**.

(id-table-registry)=
## The 7-row registry — M1-neutral, and the `w:tblW` non-registration (C4)

This WP registers **seven** tags, exactly as python-docx registers them, in
`docx/oxml/__init__.py` source order:

| tag | class | `oxml/__init__.py` | role |
|---|---|---|---|
| `w:gridCol` | `table.CT_TblGridCol` | :171 | one column of the table grid |
| `w:tblGrid` | `table.CT_TblGrid` | :174 | the column-grid container |
| `w:tblLayout` | `table.CT_TblLayoutType` | :175 | fixed-vs-autofit layout mode |
| `w:tcW` | `table.CT_TblWidth` | :181 | a cell width (the width union) |
| `w:trHeight` | `table.CT_Height` | :183 | a row height + its height rule |
| `w:vAlign` | `table.CT_VerticalJc` | :185 | a cell's vertical alignment |
| `w:vMerge` | `table.CT_VMerge` | :186 | a cell's vertical-merge behaviour |

`default.docx` has **no table** — none of the seven tags occurs in any of its 17
parts (`w:tbl` / `w:tr` / `w:tc` are absent too), so registering the leaves
touches no M1 parse path and no M1 bytes. **Byte-neutral AND flip-neutral**: no
existing exact-class `XmlElement` pin references any of the seven tags, so unlike
a registry-adding WP on the live parse path (P4-6 `styles.xml`, P5-2a
`document.xml`), this one flips zero pins. M1 stays **17/17**
(`word/document.xml` re-derived `0e4dd503…7836327` / 1548 B, unchanged) and the
targeted regression is **132/132 with 0 flips**.

:::{note}
**C4 — `w:tblW` is NOT registered.** Upstream registers **only**
`w:tcW`→`CT_TblWidth` (`__init__.py:181`); there is **no** `w:tblW`
`register_element_cls` row (verified: `grep '"w:tblW"' docx/oxml/__init__.py` →
0 hits). A `<w:tblW>` element therefore stays a **plain `XmlElement`** in both
implementations, and `CT_TblPr` (P6-2) will carry **no** `tblW` descriptor —
`Table.autofit` reads `w:tblLayout`, not `w:tblW`. Registering `w:tblW` would be
an upstream-fidelity divergence; it is deliberately **absent** from `registry.m`
(the same class serves both `w:tblW` and `w:tcW` as a shared width type, but only
the registered `w:tcW` ever instantiates it). The P5→P6 boundary audit flagged
and corrected the brief on this point.
:::

:::{note}
**`w:vAlign` is an ELEMENT registration, disjoint from every `@w:vAlign`
attribute.** `register_element_cls("w:vAlign", CT_VerticalJc)` registers the
*element* `w:vAlign`. No prior registry row registers that element:
`CT_SectPr` only **names** `w:vAlign` inside its `_tag_seq` (H11 placement data,
no descriptor, `section.py:124`), and the many `@w:vAlign` hits elsewhere are the
*attribute* — a disjoint namespace from element-tag registration. This mirrors
upstream exactly (python-docx registers `w:vAlign`→`CT_VerticalJc` globally while
`CT_SectPr` names it only as `_tag_seq` data); the registry's duplicate-key guard
did not fire and no sectPr path is affected (default.docx carries no `w:vAlign`
anywhere).
:::

---

(id-ct_tblwidth)=
## `CT_TblWidth` — the width union (the correctness crux)

**Syntax**

```matlab
tcW = mat2doc.oxml.OxmlElement("w:tcW");     % a CT_TblWidth (registered)
tcW.type = "dxa";                            % ST_TblWidth: auto/dxa/nil/pct
tcW.w    = 2880;                             % <w:tcW w:type="dxa" w:w="2880"/>
tcW.width                                    % 1828800 EMU (a Length)
tcW.width = mat2doc.shared.Twips(1440);      % forces type->dxa, w->1440
```

**Description**

The `<w:tcW>` element (the class is registered for `w:tcW` only — see C4) — a
table-related width. Two **RequiredAttribute**s (`table.py` 405–406): `w`
(`@w:w`, `XsdInt` → a plain integer count) and `type` (`@w:type`, `ST_TblWidth` →
one of `"auto"`/`"dxa"`/`"nil"`/`"pct"`). Being REQUIRED, reading either when its
attribute is absent raises `mat2doc:InvalidXmlError` (verbatim
`required 'w:…' attribute not present on element …tcW`), never a default; the
setters always write, never remove.

(id-ct_tblwidth-union)=
**The `width` union — dxa-only `Length`, everything else `None` (H6).** Upstream
types `@w:w` as `XsdInt` (not `ST_MeasurementOrPercent`) because "only dxa
(twips) values are being used" — the verbatim upstream comment (`table.py`
402–404). The computed `width` `@property` (`table.py` 408–418) realises the union
**exactly**:

- **get:** if `type != "dxa"` → `[]` (None) — a `pct`/`auto`/`nil` width has **no
  EMU length**; else → `Twips(w)` (the twips count read straight off `@w:w`).
- **set:** `type := "dxa"`; `w := Emu(value).twips` — a `Length` is stored as its
  twips count, and the type is **forced to dxa**.

So `width` is `Length`-or-None **only** for the dxa case; every other `@w:type`
yields None on read — **not a divergence, the faithful union**. Because the getter
reads `type` **first**, a malformed element missing `@w:type` raises
`InvalidXmlError` from `width` too (faithful: Python `self.type` raises the same).
Gate-3 pinned the union byte-identical two ways: build-from-scratch serialize-hex
(dxa/pct/auto/nil + the setter) and a parse-path round-trip of python-generated
`w:tcW` bytes — dxa `w="2880"` → 1828800 EMU, pct/auto/nil → None, the setter
`width=Twips(1440)` → `type="dxa"` / `w=1440` / width 914400, all byte-for-byte.

**Example**

```matlab
tcW = mat2doc.oxml.OxmlElement("w:tcW");
tcW.type = "dxa"; tcW.w = 2880;
disp(double(tcW.width));                 % 1828800  (a Length, EMU)
tcW.type = "pct"; tcW.w = 5000;
disp(isempty(tcW.width));                % 1  (pct has no EMU length -> [])
fresh = mat2doc.oxml.OxmlElement("w:tcW");
fresh.width = mat2doc.shared.Twips(1440);  % setter forces dxa
disp(fresh.type);                        % "dxa"
disp(fresh.w);                           % 1440
```

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblWidth`*

---

(id-ct_height)=
## `CT_Height` — a row height plus its height rule

**Syntax**

```matlab
h = mat2doc.oxml.OxmlElement("w:trHeight");
h.val   = mat2doc.shared.Twips(360);                        % <w:trHeight w:val="360"/>
h.hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;    % + w:hRule="exact"
```

**Description**

The `<w:trHeight>` element — a table row's height and its height rule. Two
**OptionalAttribute**s (`table.py` 41–46), both default None: `val` (`@w:val`,
`ST_TwipsMeasure` → a `Length`, EMU held exactly in doubles per H6) and `hRule`
(`@w:hRule`, `WD_ROW_HEIGHT_RULE` → a member, dispatched by fully-qualified name
through `resolveTypeCls_` to `+enum\+table`, H10). Both are the standard H3
tri-state: `[]` (None) when the attribute is absent, and a `[]` assignment removes
the attribute. `WD_ROW_HEIGHT_RULE` carries the `AUTO` / `AT_LEAST` / `EXACTLY`
members (xml `auto` / `atLeast` / `exact`).

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_Height`*

---

(id-ct_tblgrid)=
## `CT_TblGrid` — the column grid

**Syntax**

```matlab
grid = mat2doc.oxml.OxmlElement("w:tblGrid");
c1 = grid.add_gridCol();  c1.w = mat2doc.shared.Twips(2880);
c2 = grid.add_gridCol();  c2.w = mat2doc.shared.Twips(2880);
numel(grid.gridCol_lst)                       % 2
```

**Description**

The `<w:tblGrid>` element — child of `<w:tbl>`, holding the `<w:gridCol>` children
that define the table's column count and widths. One `ZeroOrMore` descriptor
(`table.py` 268): `gridCol = ZeroOrMore("w:gridCol", successors=("w:tblGridChange",))`.
The xmlchemy `ZeroOrMore` (docx form) generates the member set `gridCol_lst` /
`new_gridCol_` / `insert_gridCol_` / `add_gridCol_` **and a public `add_gridCol`**
(the D-delta-4 public adder unique to docx `ZeroOrMore`) — no bare `gridCol`
getter, no get-or-add, no remover.

(id-ct_tblgrid-h11)=
**H11 child ordering.** `successors=("w:tblGridChange",)` → the port's Constant
`SUCCESSORS = "w:tblGridChange"`; a newly inserted `gridCol` goes **before** the
first present `<w:tblGridChange>`, else is appended. On a plain grid (no
`tblGridChange`) every `add_gridCol` **appends**, so columns preserve add order.
Gate-3 pinned both: two `add_gridCol` calls on a plain grid preserve order
byte-for-byte, and `add_gridCol` on a grid holding a `<w:tblGridChange>` inserts
the new column before it (new index 1, serialize-hex identical).

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblGrid`*

---

(id-ct_tblgridcol)=
## `CT_TblGridCol` — one table column

**Syntax**

```matlab
col = mat2doc.oxml.OxmlElement("w:gridCol");
col.w = mat2doc.shared.Twips(2880);   % <w:gridCol w:w="2880"/>
col.gridCol_idx                        % 0-based position within the parent tblGrid
```

**Description**

The `<w:gridCol>` element — child of `<w:tblGrid>`, one table column. One
`OptionalAttribute` (`table.py` 274–276): `w` (`@w:w`, `ST_TwipsMeasure` → a
`Length`, default None) — the column width in twips; H3 tri-state (`[]` when
absent, `[]`-assign removes).

(id-ct_tblgridcol-idx)=
**`gridCol_idx` — the 0-based position (H1).** The computed `gridCol_idx`
`@property` (`table.py` 278–282) returns the **0-based** position of this
`<w:gridCol>` within its parent `<w:tblGrid>`'s `gridCol_lst` — Python
`tblGrid.gridCol_lst.index(self)`. The value is **data**, so the port subtracts 1
from the 1-based `find` result (`find(lst == obj, 1) - 1`, the `% IDX` H1
conversion); the search is an element-**identity** match (H5, `==` on handles). A
`mat2doc:ValueError` (`gridCol is not in list`) mirrors Python `list.index()` on
absence — unreachable in practice (a parented `gridCol` is always in its parent's
list). Gate-3 pinned `idx0=0` / `idx1=1` off a parsed 2-column grid.

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblGridCol`*

---

(id-ct_tbllayouttype)=
## `CT_TblLayoutType` — fixed vs autofit column widths

**Syntax**

```matlab
lay = mat2doc.oxml.OxmlElement("w:tblLayout");
lay.type = "fixed";    % <w:tblLayout w:type="fixed"/>
```

**Description**

The `<w:tblLayout>` element — specifies whether column widths are fixed or
auto-adjusted to content. One `OptionalAttribute` (`table.py` 292–294): `type`
(`@w:type`, `ST_TblLayoutType` → a string, default None). H3 tri-state (`[]` when
absent, `[]`-assign removes); `ST_TblLayoutType.to_xml` validates the value is
`"fixed"` or `"autofit"`. The `CT_TblPr.autofit` accessor that reads this `.type`
(and treats absent / non-`"fixed"` as autofit) is **out of scope** here — it
lands with `CT_TblPr` at P6-2; this leaf stores only the attribute.

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblLayoutType`*

---

(id-ct_verticaljc)=
## `CT_VerticalJc` — a cell's vertical alignment

**Syntax**

```matlab
va = mat2doc.oxml.OxmlElement("w:vAlign");
va.val = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER;   % <w:vAlign w:val="center"/>
```

**Description**

The `<w:vAlign>` element — the vertical alignment of a table cell. One
**RequiredAttribute** (`table.py` 967–969): `val` (`@w:val`,
`WD_CELL_VERTICAL_ALIGNMENT` → a member, referenced by fully-qualified name so
`resolveTypeCls_` dispatches to `+enum\+table`, H10). Being REQUIRED, reading
`val` when `@w:val` is absent raises `mat2doc:InvalidXmlError`, never a default;
the setter always writes `@w:val`, never removes. `WD_CELL_VERTICAL_ALIGNMENT`
carries `TOP` / `CENTER` / `BOTTOM` / `BOTH` (xml `top` / `center` / `bottom` /
`both`), each pinned byte-identical at Gate-3. The `CT_TcPr.vAlign_val` accessor
that guards for the **absent child** (returning None when there is no `<w:vAlign>`
at all) is out of scope here → P6-3; the `@w:val` on an existing `<w:vAlign>` is
required.

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_VerticalJc`*

---

(id-ct_vmerge)=
## `CT_VMerge` — a cell's vertical merge (the sole non-None default)

**Syntax**

```matlab
vm = mat2doc.oxml.OxmlElement("w:vMerge");
vm.val                    % "continue"  (default; @w:val absent)
vm.val = "restart";       % <w:vMerge w:val="restart"/>
vm.val = "continue";      % strips @w:val -> <w:vMerge/>
```

**Description**

The `<w:vMerge>` element — the vertical-merge behaviour of a table cell. One
`OptionalAttribute` **with a non-None default** (`table.py` 975–977): `val`
(`@w:val`, `ST_Merge`, `default=ST_Merge.CONTINUE`) — i.e. the default is the
**string** `"continue"`, not None. This is the one leaf in the WP whose default is
not `[]`. The `OptionalAttribute` semantics (the docx `value == self._default`
branch, realised as `isequal(value, default)`):

- **get:** `@w:val` absent → `"continue"` (the default); present → the literal
  string (`ST_Merge` inherits an identity `convert_from_xml`).
- **set:** `[]` (None) **or** `"continue"` (the default) → **remove** `@w:val`;
  `"restart"` → write `@w:val="restart"`.

So a `<w:vMerge/>` with no `@w:val` reads `"continue"`, and assigning `"continue"`
strips the attribute. The empty string `""` is a real string, **not** the default,
so it never removes — instead `ST_Merge` **rejects** it verbatim
(`must be one of ('continue', 'restart'), got ''`, a `ValueError`; the port's
`tupleRepr_` reproduces the Python tuple repr exactly). Gate-3 pinned each branch
byte-identical.

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_VMerge`*

---

# The P6-2 property containers — `CT_TblPr` · `CT_TblPrEx` · `CT_TrPr` · `CT_Row`

The four classes below are the **property containers** of the table object model:
`<w:tblPr>` (table properties), `<w:tblPrEx>` (property exceptions), `<w:trPr>`
(row properties) and `<w:tr>` (the row element). They consume the P6-1 leaves
(`CT_TblLayoutType`, `CT_Height`) and the shared leaves `CT_String`, `CT_Jc` and
`CT_DecimalNumber`. Their **cell** contents (`<w:tc>` → `CT_Tc`) and the
containing `<w:tbl>` → `CT_Tbl` land at P6-3a / P6-3b, so `CT_Row` ships the two
cell-dependent members (`_new_tc`, the `grid_span` step of `tc_at_grid_offset`) as
clean, tag-based boundary handlers — see the CT_Row section.

(id-ct_tblpr)=
## `CT_TblPr` — table properties (`<w:tblPr>`): alignment · autofit · style

**Syntax**

```matlab
tblPr = mat2doc.oxml.OxmlElement("w:tblPr");                    % a CT_TblPr (registered)
tblPr.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER; % <w:jc w:val="center"/>
tblPr.style     = "LightGrid";                                 % <w:tblStyle w:val="LightGrid"/>
tblPr.autofit   = false;                                       % <w:tblLayout w:type="fixed"/>
```

**Description**

The `<w:tblPr>` element — child of `<w:tbl>`, holding the table's style reference,
justification (alignment) and layout (autofit) among its 18-tag `_tag_seq`
(`table.py` 309–328, transcribed VERBATIM into the Constant `TAG_SEQ`). Four
`ZeroOrOne` descriptors ride the non-contiguous H11 successor slices — `tblStyle`
(`_tag_seq[1:]`→`TAG_SEQ(2:end)`), `bidiVisual` (`[4:]`→`(5:end)`), `jc`
(`[8:]`→`(9:end)`), `tblLayout` (`[13:]`→`(14:end)`) — each generated with the
generic docx family (`get.x` / `get_or_add_x` / `new_x_` / `insert_x_` / `add_x_`
/ `remove_x_`). Three computed `@property` members sit on top:

- **`alignment`** (get/set) — the table justification. **See the A2 cross-enum note
  below**; the getter returns a `WD_PARAGRAPH_ALIGNMENT` member.
- **`autofit`** (get/set) — `false` **iff** a `<w:tblLayout>` child has
  `@type="fixed"`; otherwise `true` (including when `<w:tblLayout>` is absent). The
  setter writes `@type = "autofit"` if truthy else `"fixed"` (H4).
- **`style`** (get/set) — `./w:tblStyle/@val` (a `CT_String`) or `[]` when absent.
  The setter uses the **private** `_add_tblStyle` adder (not `get_or_add`):
  `_remove_tblStyle(); if value is None: return; _add_tblStyle().val = value`.

:::{important}
**★ A2 — `CT_TblPr.alignment` returns a `WD_PARAGRAPH_ALIGNMENT` member; compare by
NAME or `.value`, never cross-class `==`.** `CT_TblPr` **reuses the same registered
`CT_Jc`** that paragraphs use (`w:jc` → `CT_Jc`, registered once at P4-2 — **one
element class, two context enums**; P6-2 does **not** add a second `w:jc` row).
`CT_Jc.val` is typed `WD_ALIGN_PARAGRAPH` (`parfmt.py:49`), so `jc.val` yields a
`WD_PARAGRAPH_ALIGNMENT` member. python-docx's `alignment` getter wraps it in
`cast("WD_TABLE_ALIGNMENT | None", jc.val)` — a **static-type no-op** that at
runtime returns `jc.val` unchanged. The port replicates this **verbatim**: the
getter returns `obj.jc.val`, a `mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT` member —
**exactly Python's runtime object**.

A *converting* getter (`WD_TABLE_ALIGNMENT(jc.val)`) would be **provably wrong**:
the legal `<w:jc w:val="both"/>` (Word's Justify) reads back as
`WD_PARAGRAPH_ALIGNMENT.JUSTIFY` (int 3), which `WD_TABLE_ALIGNMENT` cannot
represent — `WD_TABLE_ALIGNMENT.from_xml("both")` **raises**. So a literal
`jc.val` port is the only faithful getter (this exact read-back is frozen as the
Gate-3 `tblPr_both` byte fixture). The setter accepts a `WD_TABLE_ALIGNMENT` value
and writes it through `CT_Jc.val`; the cross-enum `to_xml` resolves by **int
value** (`WD_TABLE_ALIGNMENT.CENTER.value == 1` → `WD_ALIGN_PARAGRAPH.CENTER` →
`"center"`), so the **bytes are correct** (`alignment = CENTER` → `<w:jc
w:val="center"/>`, byte-identical to python-docx).

**Binding user idiom** — because MATLAB enum `==` is **class-scoped** (no `eq`
override; `member == <int>` and `SomeEnum.X == OtherEnum.X` both return *silent
false*, `member == "NAME"` returns true), compare an alignment read **by name** or
by `.value`, never by cross-class `==`:

```matlab
tblPr.alignment == "CENTER"                 % true  (name compare — the recommended idiom)
double(tblPr.alignment.value) == 1          % true  (value compare)
tblPr.alignment == WD_TABLE_ALIGNMENT.CENTER % FALSE in MATLAB (cross-class ==), true in Python
```

This mirrors an upstream typing quirk through the project's ratified class-scoped
enum design; it is **not a deviation** (bytes and the returned member's name/value
are identical to Python) and carries **no D-number** — ruled *no-D* at Gate-2 and
recorded in **design.md §2** (the A2 note). See also the
[A2 cross-enum idiom row](../python_matlab_mapping.md) in the mapping reference.
:::

:::{note}
**`w:tblStyle` → `CT_String` — a brief-under-specified functional dependency
(closed here).** `CT_TblPr.style` / its setter read and write `.val` on the
`<w:tblStyle>` child, so `w:tblStyle` **must** resolve to `CT_String` — an
unregistered child would be a generic `XmlElement` with no `.val`. This is a
genuine upstream `register_element_cls` (`oxml/__init__.py:178`) that P6-1
explicitly **deferred "to P6-2/P6-3"**; P6-2 owns `CT_TblPr` (its consumer), so the
deferral closes here (the same "brief-under-specified hard functional dependency"
pattern as P5-1's `w:evenAndOddHeaders`). **M1-neutral:** `default.docx` contains
**zero** `<w:tblStyle>` reference elements, so nothing transits `CT_String` via
this row on the M1 path.
:::

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblPr`*

---

(id-ct_tblprex)=
## `CT_TblPrEx` — table-property exceptions (`<w:tblPrEx>`)

**Syntax**

```matlab
ex = mat2doc.oxml.OxmlElement("w:tblPrEx");   % a CT_TblPrEx (registered)
```

**Description**

The `<w:tblPrEx>` element — table-property *exceptions*, applied at a lower level
(e.g. a `<w:tr>`) to override a table's appearance (used chiefly when two tables
merge). Python's `CT_TblPrEx` (`table.py` 390–396) is a **bare container**: no
child descriptors, no attributes, no `@property` members — only a docstring. So
the port is a **pure pass-through** subclass of `BaseOxmlElement` that adds nothing
beyond the transparent constructor. It exists so that a parsed `<w:tblPrEx>`
element (and a `CT_Row.tblPrEx` child) resolves to a **named** class rather than a
generic `XmlElement`, matching the upstream registration (`oxml/__init__.py:177`).
Serialize/parse are byte-identical to the generic element (no accessors run on
parse), so registering it is **byte-neutral**. **M1-neutral:** `default.docx`
contains zero `<w:tblPrEx>` elements.

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TblPrEx`*

---

(id-ct_trpr)=
## `CT_TrPr` — table-row properties (`<w:trPr>`): grid skips · row height

**Syntax**

```matlab
trPr = mat2doc.oxml.OxmlElement("w:trPr");            % a CT_TrPr (registered)
trPr.grid_before                                      % 0 (no <w:gridBefore>)
trPr.get_or_add_gridBefore().val = 1;                 % <w:gridBefore w:val="1"/>
trPr.trHeight_val = mat2doc.shared.Twips(360);        % <w:trHeight w:val="360"/>
```

**Description**

The `<w:trPr>` element — a table row's grid-before / grid-after skip counts and its
row height + height rule, over a 15-tag `_tag_seq` (`table.py` 897–913). Three
`ZeroOrOne` descriptors ride the H11 slices — `gridBefore` (`[3:]`→`(4:end)`),
`gridAfter` (`[4:]`→`(5:end)`), `trHeight` (`[8:]`→`(9:end)`). **Note the
declaration-vs-schema-order trap:** `gridBefore` is *declared after* `gridAfter` in
the Python source, yet its successor slice places it **before** `gridAfter` — the
re-sort is driven by the slice, not by declaration order. Four `@property` members:

- **`grid_after`** / **`grid_before`** (read-only) — `0` when the child is absent
  (H3), else `child.val` (a `CT_DecimalNumber`, H6).
- **`trHeight_hRule`** / **`trHeight_val`** (get/set) — `[]` (None) when
  `<w:trHeight>` is absent; the setter guards **`value is None and self.trHeight is
  None → return`** (do **not** create an empty `<w:trHeight>` just to assign None),
  then `get_or_add_trHeight` + assign (`CT_Height.hRule` / `CT_Height.val`).

:::{note}
**`w:gridAfter` / `w:gridBefore` → `CT_DecimalNumber` — two P4-6 deferrals closed.**
`grid_after` / `grid_before` read `.val` on these children (a decimal number), so
they must resolve to `CT_DecimalNumber` (registered by this WP;
`oxml/__init__.py:169` / `:170`). `w:trHeight` → `CT_Height` was already registered
at P6-1. **M1-neutral:** `default.docx` has zero `<w:trPr>` (and zero
gridAfter/gridBefore) elements.
:::

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_TrPr`*

---

(id-ct_row)=
## `CT_Row` — a table row (`<w:tr>`): row accessors · custom inserters · the CT_Tc boundary

**Syntax**

```matlab
tr = mat2doc.oxml.OxmlElement("w:tr");                 % a CT_Row (registered)
tr.grid_before                                         % 0 (no <w:trPr>/<w:gridBefore>)
tr.trHeight_val = mat2doc.shared.Twips(360);           % <w:trPr><w:trHeight .../></w:trPr>
tr.tr_idx                                              % 0-based index among sibling <w:tr>
```

**Description**

The `<w:tr>` element — a table row holding its optional `<w:tblPrEx>` (property
exceptions) and `<w:trPr>` (row properties) plus its `<w:tc>` cells. Descriptors
(`table.py` 58–61): `tblPrEx` / `trPr` are `ZeroOrOne` with **custom inserters**,
`tc` is `ZeroOrMore`. Row-level `@property` members delegate to `<w:trPr>`:

- **`grid_after`** / **`grid_before`** (read-only) — `0` when `<w:trPr>` is absent,
  else `trPr.grid_after` / `grid_before`.
- **`tr_idx`** (read-only) — the **0-based** index of this `<w:tr>` among its
  parent's `<w:tr>` siblings. Python is `tbl.tr_lst.index(self)`; because
  `CT_Tbl.tr_lst` is exactly `findall(qn("w:tr"))`, the port reads the parent's
  `w:tr` children directly (tag-based, **needs no CT_Tbl** and no stub — works now
  and after `CT_Tbl` registers), with the H1 `find(...) - 1` and H5 identity match.
- **`trHeight_hRule`** / **`trHeight_val`** (get/set) — `[]` when `<w:trPr>` absent
  (get); the setter uses `get_or_add_trPr` then delegates to `CT_TrPr`.

**H11 custom inserters (`table.py` 132–140).** `tblPrEx` and `trPr` do **not** use
the generic successor-slice engine; each overrides its inserter (like `CT_P`/`CT_R`
force `pPr`/`rPr` to the front): `_insert_tblPrEx` = `self.insert(0, tblPrEx)` →
`obj.insert(1, …)` (H1); `_insert_trPr` = `tblPrEx.addnext(trPr)` if a `tblPrEx` is
present, else `self.insert(0, trPr)`. So `tblPrEx` is forced to index 0 and `trPr`
lands immediately after it — a row built cells-first still serializes
`[tblPrEx, trPr, tc]`. `get_or_add_x` / `add_x_` route through the **own** override,
not the generic `addChild`.

:::{important}
**The CT_Tc boundary — split at the `grid_span` dependency (P6-3a).** `CT_Tc`
(`w:tc`) lands at P6-3a, so `<w:tc>` children resolve to **generic `XmlElement`**
here. Everything in `CT_Row` that does **not** need `CT_Tc` is ported LIVE
(grid_after/before via trPr, `tr_idx`, trHeight_*, the inserters, `tc_lst`). The
two cell-dependent members are handled per **design.md §4** (clean handler naming
the target symbol + owning WP):

- **`_new_tc`** (`table.py` 142–143) = `return CT_Tc.new()`. `CT_Tc.new` is P6-3a,
  so `new_tc_` raises `mat2doc:notYetPorted` (owner P6-3a). The `tc` `ZeroOrMore`
  adders (`add_tc_` / `add_tc`) route through this `_new_tc` override, so they raise
  too — **faithful**: Python's `add_tc` also reaches `CT_Tc.new()` via `_new_tc`.
- **`tc_at_grid_offset`** (`table.py` 79–98) walks `tc_lst` summing `tc.grid_span`
  (a `CT_Tc` accessor) to find the cell at an exact grid offset. Ported
  **verbatim**, with the `grid_span` access **isa-guarded**: when a `tc` is not yet
  a `CT_Tc` it raises `mat2doc:notYetPorted` at exactly that dependency point (so it
  never silently mis-walks). The fast paths that never reach `grid_span`
  (`grid_offset == grid_before` → the first cell; `grid_offset < grid_before` →
  `ValueError`) already work now. Once `CT_Tc` registers at P6-3a the `isa` guard
  passes and the method runs unchanged — **no rework**.
:::

*Ported from python-docx v1.2.0: `src/docx/oxml/table.py::CT_Row`*

---

(id-table-byte-proof)=
## Byte proof — M1-neutral, and the leaf subtrees round-trip byte-identical

The seven leaves emit no bytes into any M1 part, and the novel paths are the
table-leaf subtrees of a **real** python-docx table. All are frozen as permanent
Gate-3/Gate-4 references (co-located `.gitattributes` `* binary` pins) and are
**byte-identical to python-docx**:

| leg | method | result |
|---|---|---|
| M1 17/17 byte-neutrality | `mat2doc.Document().save()` vs `references\s0001` | L0 PASS + 16 XML L1 + 1 bin; `document.xml` `0e4dd503…` / 1548 B unchanged |
| full leaf surface | `probe_diff` s0059 (values + serialize-hex + error messages) | MATCH (exit 0) |
| table-leaf round-trip | 10 frozen fixtures parse→serialize vs oracle bytes (`references\s0060`) | 10/10 byte-identical |
| width union | dxa/pct/auto/nil + setter, build-from-scratch **and** parse-path | byte-identical both ways |
| registry C4 | 7 tags resolve; `w:tblW` returns `""` (unregistered) | PASS |
| targeted regression | 5 named classes, one R2024b session, foreground | 132/132, 0 flips |

The round-trip fixtures come from an independent real 3×2 python-docx table
(fixed layout; column widths dxa 2880/1440; row-0 height `Twips(500)` +
`WD_ROW_HEIGHT_RULE.EXACTLY`; cell(0,0) `vAlign=CENTER`; a `cell(0,1).merge` →
`vMerge` restart/continue) — the subtrees that transit the new CT leaves plus four
loose width-union `w:tcW` elements, each re-serialized through the leaf classes
byte-for-byte. **Zero new D-numbers** — every equivalence leg is L1 byte-identical
or value-exact; the standing adopt-only deviations exercised are **D-001** (own
parser/serializer), **D-delta-4** (the ZeroOrMore public `add_gridCol`), and
**D-zip-time** (envelope only).

(id-table-byte-proof-p62)=
### P6-2 byte proof — `styles.xml` stays byte-identical through `CT_TblPr`, and the property subtrees round-trip

P6-2 is the **first non-neutral** table WP, so its byte proof carries an extra
leg: `word/styles.xml` / `word/stylesWithEffects.xml` each hold **100
`<w:tblPr>` nodes** (inside table styles) that now transit `CT_TblPr` on the M1
parse path. All P6-2 equivalence legs are **byte-identical to python-docx**:

| leg | method | result |
|---|---|---|
| **★ M1 17/17** (styles.xml now transits CT_TblPr) | `mat2doc.Document().save()` vs `references\s0001` | L0 PASS + 16 XML L1 + 1 bin; `styles.xml` `02d71a68…bc30e384` / 349 458 B **unchanged** |
| full props+row surface | `probe_diff` s0061 (values + serialize-hex; A2 by-name/value compares) | MATCH (exit 0) |
| table-props round-trip | 6 frozen fixtures parse→serialize vs oracle (`references\s0062`) | 6/6 byte-identical |
| **★ `w:jc w:val="both"` edge** | parse → `WD_PARAGRAPH_ALIGNMENT.JUSTIFY` (int 3), re-serialize | byte-identical (`tblPr_both` `620dea7a…`) |
| alignment cross-enum | set CENTER/LEFT/RIGHT → `w:jc` bytes; None → `w:jc` removed | byte-identical (A2 int-value setter) |
| registry | 7 P6-2 rows resolve; **`w:jc` NOT re-registered** (single `CT_Jc`, A2) | PASS |
| targeted regression | 5 named classes, one R2024b session, foreground | 107/107, 0 flips |

The styles.xml SHA being unchanged to the byte **proves** the `CT_TblPr`
registration is byte-neutral (registering changes only a parsed node's class;
`CT_TblPr` adds no parse-time behavior). The 6 round-trip fixtures include loose
`tblPr` / `trPr` / `tr` (the custom-inserter order `[tblPrEx, trPr]`) and two
real-`add_table` subtrees with the full 18-decl nsmap. **Zero new D-numbers** —
every leg is L1 byte-identical or value-exact; the A2 cross-class `==` gap is
non-byte / non-output and is ruled **no-D** (design.md §2).

:::{note}
**Milestone flag (COM).** The s0060 (leaf) and s0062 (property) table subtrees emit
output-visible table structure (`w:tblGrid` / `w:gridCol` / `w:tblLayout` /
`w:trHeight` / `w:tcW` / `w:vAlign` / `w:vMerge` / `w:tblPr` with jc/tblStyle /
`w:trPr` with gridBefore/gridAfter/trHeight / `w:tr` with the tblPrEx/trPr
inserters) but a full table is **not yet reachable through a Mat2Doc-generated
package** — the cell/table containers `CT_Tc` / `CT_TcPr` / `CT_Tbl` (P6-3a /
P6-3b) are needed to build a real table end-to-end. The frozen s0060 + s0062
fixtures are routed to mso-office-verifier (the Word COM oracle) at the next
milestone sweep once the containers land; the full table Word-COM sweep (a plain
grid **and** a merged-cell package) is booked at **P6-4b**.
:::

---

(id-table-next)=
## Leaves + property containers done — the CT_Tc merge engine (P6-3a) next

The table oxml **leaf** layer (P6-1) and the **property-container** tier (P6-2) are
now both whole:

- **P6-1 leaves** — the width union (`CT_TblWidth`), the row-height leaf
  (`CT_Height`), the column grid (`CT_TblGrid` / `CT_TblGridCol`), the layout-mode
  leaf (`CT_TblLayoutType`), and the two cell leaves (`CT_VerticalJc` /
  `CT_VMerge`) — registered M1-neutrally.
- **P6-2 containers** — `CT_TblPr` (alignment / autofit / style), `CT_TblPrEx`
  (bare exceptions container), `CT_TrPr` (grid skips + row height) and `CT_Row`
  (the `<w:tr>` row with its custom inserters and the CT_Tc boundary handlers) —
  the first **non-neutral** table WP (`w:tblPr` transits the live `styles.xml`
  parse path, proven byte-neutral by the M1 17/17 `02d71a68…` match), replicating
  the A2 `CT_Jc` cross-enum reuse and closing the `w:tblStyle` → `CT_String` and
  `w:gridAfter` / `w:gridBefore` → `CT_DecimalNumber` deferrals.

Both landed **byte-proven** on real + loose table subtrees with **zero new
D-numbers**. What remains in **Phase 6**:

- **P6-3a** — the **CT_Tc merge engine** (gridSpan / vMerge, the block-item
  container seam). This is the **hardest WP of the project**; it gets its own
  **dedicated pre-launch plan-audit** (risk-register #2), which also adjudicates
  the Row ↔ Tc ↔ TcPr ↔ Tbl mutually-recursive dependency seam and un-stubs
  `CT_Row._new_tc` + the `grid_span` step of `tc_at_grid_offset`.
- **P6-3b** — `CT_TcPr` / `CT_Tbl` + the registry completion and the tbl
  un-defer sweep (the tag-based `w:tbl`→generic sites auto-upgrade when
  `w:tbl`→`CT_Tbl` registers).
- **P6-4a / P6-4b** — the `Table` / `_Rows` / `_Columns` / `_Cell` API (which
  discharges the P5-3a `Section.iter_inner_content` `w:tbl`→`Table` debt and the
  `Document.add_table` / `tables` stubs), then `_Cell.merge` + `add_row` /
  `add_column` and the table Word-COM sweep.
