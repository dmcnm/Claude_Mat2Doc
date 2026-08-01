---
title: "mat2doc.table — the Table/_Rows/_Columns/_Row/_Column/_Cell API tier (Table · the row/column collections · the cell · the add_table/add_row/add_column authoring path · cell merge — the COMPLETE table tier)"
---

# `mat2doc.table` — the Table / row / column / cell API tier (`Table` + `_Rows`/`_Columns` + `_Row`/`_Column` + `_Cell` + the `add_table`/`add_row`/`add_column`/`merge` authoring path)

Ported from python-docx v1.2.0 `src/docx/table.py::Table` /
`::_Columns` / `::_Column` / `::_Rows` / `::_Row` (in the NEW package
`+mat2doc/+table/`), plus the **authoring un-stubs** they light up —
`blkcntnr.py::BlockItemContainer.add_table` / `iter_inner_content`
(`+mat2doc/BlockItemContainer.m`), `document.py::Document.add_table` /
`iter_inner_content` (`+mat2doc/+document/Document.m`) and the
`section.py::Section.iter_inner_content` `w:tbl` branch
(`+mat2doc/+section/Section.m`). This is the **API surface** over the P6-1..P6-3b
table-oxml layer: `Table` is a proxy over an already-registered `CT_Tbl`,
`_Rows`/`_Columns` are the row/column sequences, `_Row`/`_Column` are the
single-row/column proxies, and `add_table` is the **first end-to-end table
authoring byte path**.

:::{note}
**★ The FIRST end-to-end table is now reachable and byte-proven.** P6-4a ports the
user-facing `Table` proxy and the four collection/leaf proxies
(`_Rows`/`_Columns`/`_Row`/`_Column`) and **un-stubs the table-authoring path**
(`Document.add_table` / `BlockItemContainer.add_table` + the three
`iter_inner_content` sites + the `Section.iter_inner_content` `w:tbl` branch), so
`Document().add_table(rows, cols)` now builds a real `<w:tbl>` in a Mat2Doc
package. It adds **no `register_element_cls` row and no serialization code** —
equivalence is **behavioral** (proxy value parity) plus **serialized-bytes**
parity on the `add_table` authoring path (the constructed `word/document.xml` is a
byte oracle — `CT_Tbl.new_tbl` was byte-proven at P6-3b) — so it is
**byte-neutral** (`Document().save()` stays M1 17/17, **zero new D-numbers**). The
**table READ/authoring API is now live**; `_Cell.merge` + `add_row` / `add_column`
+ the table Word-COM sweep (**P6-4b**) closed Phase 6 (see the next note).
:::

:::{important}
**★ PHASE 6 COMPLETE** — the full table tier (oxml 14 classes +
`Table`/`_Rows`/`_Columns`/`_Row`/`_Column`/`_Cell` API +
`add_table`/`add_row`/`add_column`/`merge`) is **byte-proven end-to-end and
COM-verified**. **P6-4b** lands the last piece — `_Cell` (with `.merge`),
`Table.add_row`/`add_column`, and `_Row.cells`/`_Column.cells` (the cells grid
walk) — un-stubbing the 8 P6-4a table stubs plus the required
`BlockItemContainer.tables`. Every cell-authoring path (content, `add_paragraph`,
nested `add_table`, `add_row`/`add_column`) and the **full cell-merge matrix**
(horizontal `gridSpan`, vertical/block `vMerge`, merge-then-set-text, the 3×3
mixed merge) is **byte-identical to python-docx v1.2.0** (Gate-3 `s0073`–`s0082`,
each part-level 17/17 L1) and **COM-verified in real Microsoft Word** (Word
16.0.20228 — all five table packages open silently, merges honored, cell text
intact, clean round-trip; `com_verify_P6_tables.md`). **Zero new D-numbers across
all of Phase 6.** See the [`_Cell`](#id-cell) and
[cell-merge authoring path](#id-cell-merge) sections below.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The one-package layout

`table.py` (a top-level `docx/` module) lands in the single new package
`+mat2doc\+table` — following the established module→subpackage convention. It
collides with no existing package. The oxml element classes it reads
(`CT_Tbl` / `CT_Row` / `CT_Tc` / `CT_TblGridCol` + the property containers) live
separately in `+mat2doc\+oxml\+table` (P6-1..P6-3b — see
[`oxml_table.md`](oxml_table.md)).

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `table.py` | `+mat2doc\+table\` | `Table`, `_Columns`→`Columns_`, `_Column`→`Column_`, `_Rows`→`Rows_`, `_Row`→`Row_` |

**The FLAG-3 trailing-underscore renames.** The four leading-underscore
"private" proxy classes rotate their leading `_` to a **trailing** one (design.md
§2, the toolbox-wide FLAG-3 convention — the same rule that gave `_Relationship`→
`Relationship_`, `_TableStyle`→`TableStyle_`): `_Columns`→`Columns_`,
`_Column`→`Column_`, `_Rows`→`Rows_`, `_Row`→`Row_`. The public `Table` keeps its
name. `_Cell` (`table.py` 192–311) is **NOT** ported here — it lands with P6-4b
(no `Cell_.m` file exists yet).

**The tier boundary.** `Table` / `Columns_` / `Column_` / `Rows_` / `Row_` add
**NO** oxml logic, **NO** registry rows and **NO** serialization code. Every
accessor delegates one-to-one to the wrapped `CT_Tbl` / `CT_Row` /
`CT_TblGridCol` (which own all the H6/H3/H4/H10 logic — see
[`oxml_table.md`](oxml_table.md)) or up to the `DocumentPart` (style resolution).
So equivalence is **behavioral** except the `add_table` authoring path, whose
constructed `document.xml` is a **byte oracle**.

(id-table-unstub)=
## The authoring un-stub — `add_table` / `iter_inner_content` now live (C2 discharged)

Since P2-3 the block-item `add_table` / `iter_inner_content` adders and the
`Document`-level ones were `mat2doc:notYetPorted` stubs (no `Table` proxy). P6-4a
**un-stubs all five** members, whose deps (`CT_Tbl.new_tbl` byte-proven P6-3b;
`CT_Body._insert_tbl` live P2-3; the `Table` proxy now live) are all satisfied:

| member | Python | now delegates to |
|---|---|---|
| `BlockItemContainer.add_table(rows, cols, width)` | `blkcntnr.py:61-72` | `Table(CT_Tbl.new_tbl(rows,cols,width))` after `_insert_tbl` |
| `BlockItemContainer.iter_inner_content` | `blkcntnr.py:74-79` | `1×N` cell of `Paragraph`(CT_P) / `Table`(else) |
| `Document.add_table(rows, cols, style)` | `document.py:150-158` | `_body.add_table(rows,cols,_block_width)` → `table.style = style` |
| `Document.iter_inner_content` | `document.py:180-182` | `_body.iter_inner_content()` |
| `Section.iter_inner_content` (`w:tbl` branch) | `section.py:157-163` | `Table(element, self)` — **discharges the P5-3a C2 debt** |

The **P5-3a C2 debt is discharged**: `Section.iter_inner_content`'s `w:tbl`→`Table`
branch (which raised `mat2doc:notYetPorted` owner P6-4a) now returns a real
`Table`, so a table inside a section is wrapped, never silently dropped. **The
un-stub adds no registry rows and moves not a single byte of the default
`word/document.xml`** — a bare `Document().save()` fires ZERO table stubs and stays
M1 17/17 byte-identical.

(id-add-table-signature)=
## ★ The `add_table` authoring path — the two signatures

`add_table` exists at **two** levels with a **deliberately different third
argument** — a signature difference that is load-bearing and easy to get wrong:

| level | Python | MATLAB | 3rd arg |
|---|---|---|---|
| **public** (`Document`) | `Document.add_table(rows, cols, style=None)` (`document.py:150`) | `d.add_table(rows, cols, style)` | **`style`** (a table-style object or name; `[]`/None → default) |
| **container** (`_Body`, `_Cell`, …) | `BlockItemContainer.add_table(rows, cols, width)` (`blkcntnr.py:61`) | `container.add_table(rows, cols, width)` | **`width`** (a `Length`, REQUIRED) |

The public `Document.add_table` supplies the **width itself** — it computes
`self._block_width` (`document.py:156`, `Emu(page_width − left − right)` off the
last section) and passes it down to `_body.add_table(rows, cols, self._block_width)`,
then applies the style: `table.style = style`. So the user chooses **style** at
the `Document` level and never passes a width; the container level takes an
explicit **width** and no style. Mat2Doc ports both **exactly** — `Document.m`'s
`add_table` declares `style = []` (Python `style=None`, H13 default fidelity),
`BlockItemContainer.m`'s `add_table` takes a 3rd `width`.

:::{warning}
**`d.add_table(2, 3, Inches(6))` is INVALID** — `Inches(6)` is passed as the
**style**, and python-docx raises `KeyError "no style with name '5486400'"`
(oracle-confirmed). To author a table of a specific width, either omit the style
(`d.add_table(2, 3)` — the width comes from `_block_width`, which equals
`Inches(6)` for the default template) or drop to the container level
(`_body.add_table(2, 3, Inches(6))`). Every example below uses the **faithful**
public signature.
:::

(id-byte-proof)=
**The first end-to-end table — byte-proven.** The public path
`document.add_table(2, 3)` produces a `word/document.xml` **byte-identical** to
python-docx (`word/document.xml` SHA-256
`a1eda0439dbf02fb02c109ca218c5c8b0fbe7136a9e54022cd0f28a4ea1820bf`, both sides —
re-derived at Gate-2 across a 12-scenario matrix, all **full-package L1 17/17
parts each**). `styles.xml` is **unchanged** by a plain table (the table adds
nodes to `document.xml` only). When a **style** is applied
(`add_table(2, 2, "My Table Style")` on a style added via `styles.add_style`), the
table emits a `<w:tblStyle w:val="…"/>` and `styles.xml` gains the style
definition — both byte-identical to python-docx. **Zero new D-numbers** (consistent
with `new_tbl` being byte-proven at P6-3b).

---

(id-table)=
## `Table`

**Syntax**

```matlab
d   = mat2doc.Document();
tbl = d.add_table(2, 3);                                         % a Table (style=[], width=_block_width)
tbl.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;    % <w:jc w:val="center"/>
n   = tbl.rows.len_();                                           % number of rows (2)
c0  = tbl.columns.getitem_(0);                                   % Python table.columns[0] -> a Column_
tbl.autofit = false;                                            % <w:tblLayout w:type="fixed"/>
```

**Description**

`Table` is the proxy over a WordprocessingML `<w:tbl>` element — the public API
tier over the P6-3b `CT_Tbl` oxml layer. In python-docx it is
`class Table(StoryChild)` (`table.py:29`), so it is
`classdef Table < mat2doc.shared.StoryChild` — **NOT** an `ElementProxy` — with
default handle identity (Python default object identity) and defines **no
`eq`/`ne`** (a `Table` is not compared by wrapped-element identity). It stores the
`<w:tbl>` twice under two names — `element_` (Python `_element`, used by the
`table_direction` setter) and `tbl_` (Python `_tbl`, every other method) — both
the same `CT_Tbl` handle (the `Paragraph` `p_`/`element_` precedent).

(id-table-members)=
**The live member surface (P6-4a).** Everything that reads the table's structure
ports fully:

| member | kind | delegates to | note |
|---|---|---|---|
| `alignment` | get/set | `_tblPr.alignment` | **A2 cross-enum** — see below |
| `autofit` | get/set | `_tblPr.autofit` | `false` iff `<w:tblLayout @type="fixed">` (H4) |
| `columns` | get (`@lazyproperty`) | `Columns_(_tbl, self)` | cached (computed-flag idiom); read-only |
| `rows` | get (`@lazyproperty`) | `Rows_(_tbl, self)` | cached; read-only |
| `style` | get/set | `_tbl.tblStyle_val` + `part.get_style/id(…, TABLE)` | live style resolution (P4-7a) |
| `table` | get | `obj` (self) | the child→parent terminus |
| `table_direction` | get/set | `_tbl.bidiVisual_val` | `WD_TABLE_DIRECTION` — see below |
| `_column_count` | `column_count_()` | `_tbl.col_count` | public method (upstream tests read it) |
| `_tblPr` | `tblPr_()` | `_tbl.tblPr` | private→public method; a `CT_TblPr` |

The `columns`/`rows` `@lazyproperty` accessors are **cached on first access** (the
Dependent read-only property + private cache + computed-flag idiom — never
`isempty` as the sentinel, H3), so `tbl.rows == tbl.rows` is true.

(id-table-a2)=
:::{important}
**★ A2 — `Table.alignment` returns a `WD_PARAGRAPH_ALIGNMENT` member; compare by
NAME or `.value`, never cross-class `==`.** `Table.alignment` get →
`self._tblPr.alignment` = `CT_TblPr.alignment` (P6-2), which returns a
**paragraph-alignment** member (`CT_Jc.val`) because python-docx's
`cast("WD_TABLE_ALIGNMENT | None", jc.val)` is a **runtime no-op**. This port
returns that member **verbatim** — it does NOT convert to `WD_TABLE_ALIGNMENT` (a
converting getter would crash on the legal `<w:jc w:val="both">` = Justify that
`WD_TABLE_ALIGNMENT` cannot represent). At runtime the member is a
`mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT` (of which `WD_ALIGN_PARAGRAPH` is an
alias class), **exactly** what `Paragraph.alignment` returns.

Because MATLAB enum `==` is **class-scoped**, compare an alignment read **by name**
or by `.value`, never by cross-class `==`:

```matlab
tbl.alignment == "CENTER"                              % true  (name compare — the recommended idiom)
double(tbl.alignment.value) == 1                       % true  (value compare)
tbl.alignment == mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER  % FALSE in MATLAB, true in Python
```

The **setter** accepts a `WD_TABLE_ALIGNMENT` value and writes it through
`CT_TblPr` (the cross-enum `to_xml` resolves by **int value** → the correct
`"left"`/`"center"`/`"right"` bytes). This is the same ratified A2 note as
`CT_TblPr.alignment` (design.md §2) — **not a deviation, no D-number** (bytes and
the member's name/value are identical to Python).
:::

(id-table-direction)=
**`table_direction` — a real `WD_TABLE_DIRECTION` (no A2 issue), returned as a
logical.** `table_direction` get → `self._tbl.bidiVisual_val`. python-docx's
`cast("WD_TABLE_DIRECTION | None", …)` is again a runtime no-op, and
`bidiVisual_val` returns `[]` (None) or a **logical** (`CT_OnOff.val`), so the
getter returns that verbatim (as python-docx returns the bool cast). The setter →
`self._element.bidiVisual_val = value` (a `WD_TABLE_DIRECTION` member; `CT_Tbl`
applies `bool(value)` — LTR=0 → `"0"`, RTL=1 → a bare `<w:bidiVisual/>`). If a user
compares against a `WD_TABLE_DIRECTION` member, the same by-name/`.value` idiom
applies.

(id-table-p6-4b)=
**★ The P6-4b cell/mutator members are now LIVE.** The eight members that needed
`_Cell` are all un-stubbed at P6-4b: `add_row` / `add_column` (the mutators — each
returns a `Row_` / `Column_` over the new `<w:tr>` / `<w:gridCol>`), `cell(row,
col)` (the grid-address accessor → a `Cell_`), `row_cells(row)` / `column_cells(col)`
(one row's / column's cells), and `_cells` (`cells_()`, the full row-major grid).
`_Cell` itself lands here (`Cell_.m`); the grid accessors and merge are documented
in the [`_Cell`](#id-cell) and [cell-merge](#id-cell-merge) sections below.

**Example**

```matlab
d   = mat2doc.Document();
tbl = d.add_table(2, 3);                                    % public API: (rows, cols, style=[])
disp(tbl.rows.len_());                                      % 2
disp(tbl.columns.len_());                                   % 3
tbl.alignment = mat2doc.enum.table.WD_TABLE_ALIGNMENT.CENTER;
disp(tbl.alignment == "CENTER");                           % 1  (by-name idiom — NEVER cross-class ==)
disp(double(tbl.alignment.value) == 1);                    % 1  (value compare)
tbl.autofit = false;
disp(tbl.autofit);                                         % 0  (<w:tblLayout w:type="fixed"/>)
disp(tbl.table == tbl);                                    % 1  (the child->parent terminus)
disp(tbl.column_count_());                                 % 3
```

*Ported from python-docx v1.2.0: `src/docx/table.py::Table`*

---

(id-rows-columns)=
## `_Rows` / `_Columns` — the row and column collections

**Syntax**

```matlab
rows = tbl.rows;                  % a Rows_
rows.len_()                       % number of rows
r0   = rows.getitem_(0);          % Python rows[0] -> a Row_
mid  = rows.getitem_(struct("start", 1, "stop", 3, "step", []));   % rows[1:3]
for r = rows.to_array(); disp(r.index_()); end
cols = tbl.columns;               % a Columns_
c0   = cols.getitem_(0);          % Python columns[0] -> a Column_
```

**Description**

`Columns_` (`_Columns`, `table.py:347`) and `Rows_` (`_Rows`, `table.py:507`) are
the column and row **sequences** — length, iteration and indexed access. Both are
`class _Columns(Parented)` / `class _Rows(Parented)` — plain parented proxies
holding the `<w:tbl>` and the parent `Table`, with no single wrapped element of
their own — so they derive from `mat2doc.shared.Parented` with default handle
identity and **no `eq`/`ne`**.

(id-collections-sequence)=
**The Sequence surface — explicit methods, the 0-based `getitem_` convention.**
The shared 1-based `()` collection base (`RedefinesParen`) is a **future work
package** (the standing VERIFY-COLLECTION flag). Per the established Mat2Doc
precedent (`Sections`, `TabStops`, `Styles`), the Python `Sequence` dunders are
ported as **explicit methods**:

| Python | MATLAB | note |
|---|---|---|
| `columns[key]` / `rows[key]` | `getitem_(key)` | key is the **Python 0-based** int (or a slice `struct`, `_Rows` only) |
| `for x in columns` / `rows` | `for x = getitem.to_array()` | `__iter__` → materialized `1×N` |
| `len(columns)` / `len(rows)` | `len_()` | `__len__` |

`getitem_` takes the **Python 0-based** key (the `Sections`/`TabStops` precedent,
NOT the Mat2Ppt 1-based `Slides` convention), then hits the 1-based MATLAB list
with an explicit `+1` at the single indexing site (H1). A **negative** key counts
from the end. The two collections differ in their **out-of-range message and their
key kinds**:

- **`Columns_`** is **INT-ONLY** (no slice overload). Out-of-range → `mat2doc:IndexError`
  with the **custom** message `"column index [%d] is out of range"` formatted with
  the **ORIGINAL** idx (H1 — e.g. `columns.getitem_(-9)` on a 3-column table →
  `"column index [-9] is out of range"`).
- **`Rows_`** supports **INT + SLICE** (`list(self)[idx]`). An out-of-range int →
  the standard CPython `"list index out of range"`. A **slice** key is a `struct`
  with fields `start`/`stop`/`step` (each a scalar double or `[]` for None) — the
  interim currency until `RedefinesParen` lands; the private `sliceIndices_` is a
  faithful port of CPython `slice.indices(n)` + `range(...)` (byte-identical to
  `Sections.sliceIndices_`), returning a `1×N` `Row_` array (empty slice → `1×0`).

**H5 identity.** Every `getitem_` / `to_array` element mints a **fresh** `Row_` /
`Column_` view of its `<w:tr>` / `<w:gridCol>` (python-docx does not cache them);
the wrapped `CT_Row` / `CT_TblGridCol` is the shared identity. The `columns`/`rows`
collections themselves ARE cached on the `Table` (`@lazyproperty`). Each
collection's `table()` property-as-method returns the owning `Table`.

**Example**

```matlab
cols = tbl.columns;                       % a Columns_
disp(cols.len_());                        % 3
c1   = cols.getitem_(1);
disp(c1.index_());                        % 1  (0-based grid data, H1)
disp(cols.getitem_(-1).index_());         % 2  (negative wrap -> last)
try
    cols.getitem_(9);                     % out of range
catch e
    disp(e.identifier);                   % mat2doc:IndexError
end
rows = tbl.rows;                          % a Rows_
mid  = rows.getitem_(struct("start", 0, "stop", 2, "step", []));   % rows[0:2]
disp(numel(mid));                         % 2  (a 1x2 Row_ array)
disp(rows.getitem_(-1).table() == tbl);   % 1  (the owning Table)
```

*Ported from python-docx v1.2.0: `src/docx/table.py::_Columns` / `src/docx/table.py::_Rows`*

---

(id-row-column)=
## `_Row` / `_Column` — the single-row / single-column proxies

**Syntax**

```matlab
row = tbl.rows.getitem_(0);      % a Row_
row.height      = mat2doc.shared.Pt(20);
row.height_rule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
row.index_()                     % 0
col = tbl.columns.getitem_(0);   % a Column_
col.width = mat2doc.shared.Inches(1.5);
col.index_()                     % 0
```

**Description**

`Row_` (`_Row`, `table.py:387`) and `Column_` (`_Column`, `table.py:314`) are
parented proxies over a single `<w:tr>` / `<w:gridCol>`. Both are
`class …(Parented)` → `mat2doc.shared.Parented` (default handle identity, no
`eq`/`ne`). `Row_` stores the `<w:tr>` under two names (`tr_` = Python `_tr`,
`element_` = Python `_element`); `Column_` stores the `<w:gridCol>` as `gridCol_`.

(id-row-column-members)=
**The member surface.** Every accessor delegates one-to-one to the wrapped
`CT_Row` / `CT_TblGridCol`, which own all the H3/H6/H10 logic:

| proxy | member | kind | delegates to | note |
|---|---|---|---|---|
| `Column_` | `width` | get/set | `_gridCol.w` | `Length \| []` (H6 EMU, H3 tri-state — `[]`-assign removes `@w:w`) |
| `Column_` | `_index` (`index_()`) | method | `_gridCol.gridCol_idx` | 0-based grid data (H1); public method (upstream reads it) |
| `Column_` | `cells` | method (RO) | `Table._cells` column stride | **LIVE @ P6-4b** — `1×N` `Cell_` (the column's cells top→bottom; see [cells grid walk](#id-cell-merge)) |
| `Row_` | `height` | get/set | `_tr.trHeight_val` | `Length \| []` (H6/H3) |
| `Row_` | `height_rule` | get/set | `_tr.trHeight_hRule` | `WD_ROW_HEIGHT_RULE \| []` (H10/H3) |
| `Row_` | `grid_cols_after` | get (RO) | `_tr.grid_after` | int — unpopulated grid-cols after the last cell |
| `Row_` | `grid_cols_before` | get (RO) | `_tr.grid_before` | int — unpopulated grid-cols before the first cell |
| `Row_` | `_index` (`index_()`) | method | `_tr.tr_idx` | 0-based row data (H1) |
| `Row_` | `cells` | method (RO) | `iter_tc_cells_` grid walk | **LIVE @ P6-4b** — `1×N` `Cell_` (the row's cells left→right; span-identity, see [cells grid walk](#id-cell-merge)) |
| both | `table` | method | `_parent.table` | the owning `Table` (property-as-method) |

`_index` is 0-based **data** (Python `list.index()`), already `-1`-adjusted inside
the oxml layer; ported via a **public** `index_()` method (upstream tests read
`_Row._index` / `_Column._index`). The single leading-underscore rotations are
`_index`→`index_`. `Row_.cells` / `Column_.cells` are **LIVE at P6-4b** — they
build `Cell_` objects over the grid (the [cells grid walk](#id-cell-merge)).

**Example**

```matlab
row = tbl.rows.getitem_(0);                 % a Row_
row.height      = mat2doc.shared.Pt(20);
row.height_rule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
disp(double(row.height));                   % 254000  (20 pt in EMU)
disp(row.height_rule == mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY);   % 1
disp(row.index_());                         % 0
col = tbl.columns.getitem_(0);              % a Column_
col.width = mat2doc.shared.Inches(1.5);
disp(double(col.width));                    % 1371600  (1.5 in in EMU)
col.width = [];                             % H3 -> removes @w:w
disp(isempty(col.width));                   % 1
disp(col.index_());                         % 0
```

*Ported from python-docx v1.2.0: `src/docx/table.py::_Row` / `src/docx/table.py::_Column`*

---

(id-cell)=
## `_Cell` — the table cell (`Cell_`)

**Syntax**

```matlab
t = mat2doc.Document().add_table(2, 2, mat2doc.shared.Inches(4));
c = t.cell(0, 0);                      % Python table.cell(0, 0) -> a Cell_
c.text = "A";                          % replace all content with one <w:p>/<w:r> "A"
p = c.add_paragraph("second line");    % a Paragraph appended to the cell
inner = c.add_table(2, 2);             % a NESTED table inside the cell
c.vertical_alignment = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER;
c.width = mat2doc.shared.Inches(1.5);  % <w:tcW w:w="…" w:type="dxa"/>
n = c.grid_span;                       % columns this cell spans (1 normal)
```

**Description**

`Cell_` is the FLAG-3 trailing-underscore rename of python-docx's `_Cell`
(`table.py:192`). In python-docx it is `class _Cell(BlockItemContainer)`, so it is
`classdef Cell_ < mat2doc.BlockItemContainer` — a **block-item container** (it
contributes the paragraph/table add+read surface from
[`BlockItemContainer`](document.md)) **plus** the cell-specific members
(`grid_span` / `merge` / `text` / `vertical_alignment` / `width`).

**The `_tc` / `_element` dual name.** python-docx stores the wrapped `<w:tc>`
twice — `self._tc = self._element = tc` (`table.py:198`). Mat2Doc holds it as the
private `tc_` field (every `self._tc` site) **and** through the inherited
`element_()` C3 seam (every `self._element` site); both are the **one** `CT_Tc`.
Unlike `_BaseHeaderFooter` (which lazily overrides `element_()`), `_Cell` stores a
**concrete** element, so the base `BlockItemContainer.element_()` returning the
stored tc is exactly right — the inherited container surface (`add_paragraph` /
`paragraphs` / `add_table` / `tables`) operates directly on the tc.

(id-cell-members)=
**The member surface.**

| member | kind | delegates to | note |
|---|---|---|---|
| `text` | get/set | `paragraphs` join / `tc.clear_content()`+`add_p`+`add_r` | **get** = `"\n".join(p.text …)` (**no strip**, H16 latent); **set** replaces ALL content with one `<w:p>`/`<w:r>` |
| `add_paragraph(text, style)` | method | `super().add_paragraph` | docstring override only — verbatim `BlockItemContainer` behavior (H13 defaults `""`/`[]`) |
| `add_table(rows, cols)` | method | `super().add_table(rows, cols, width)` | a **NESTED** table; width = the cell's width, or `Inches(1)` when the cell has none (H3); a trailing empty `<w:p>` is appended (Word requires a paragraph as a cell's last block child) |
| `grid_span` | get (RO) | `tc.grid_span` | columns this cell spans (1 normal, 2+ horizontally merged) |
| `paragraphs` | get (RO) | `super().paragraphs` | `1×N` Paragraph; each read mints fresh views (H5) |
| `tables` | get (RO) | `super().tables` | `1×N` Table — the base `BlockItemContainer.tables` is **un-stubbed at P6-4b** (a required dependency of `_Cell.tables`) |
| `vertical_alignment` | get/set | `./w:tcPr/w:vAlign/@w:val` | `WD_CELL_VERTICAL_ALIGNMENT` \| `[]` (H10/H3 — get uses the RAW `tcPr`, a read never creates it; `[]`-assign removes) |
| `width` | get/set | `tc.width` | `Length` (EMU) \| `[]` (H6/H3 — dxa twips via `CT_TblWidth`) |
| `merge(other)` | method | `tc.merge(other._tc)` | a NEW merged `Cell_` — see [cell merge](#id-cell-merge) |

:::{note}
**`_Cell.tables` needed one dependency un-stubbed.** `_Cell.tables` is
`super(_Cell, self).tables` (`table.py:262`), so the shared
`BlockItemContainer.tables` (stubbed since P2-3, all deps live once `Table` landed
at P6-4a) is un-stubbed at P6-4b. It is the single shared implementation that
`_Cell.tables`, `Body_.tables` (inherited) all route through. `Document.tables`
(its **own** separate stub, `document.py:221`) is left stubbed — out of the
P6-4b 8-stub scope; its concrete owner is re-assigned to **P8-3** (the
blkcntnr-remainder WP, the `Document.paragraphs`→P8-3 precedent).
:::

**Example**

```matlab
d = mat2doc.Document();
t = d.add_table(2, 2);
c = t.cell(0, 0);                              % a Cell_
c.text = "A";                                  % one <w:p>/<w:r> "A"
disp(c.text);                                  % A
c.add_paragraph("p2");                         % append a second paragraph
disp(c.text);                                  % A<newline>p2   (H16 newline-join, no strip)
disp(numel(c.paragraphs));                     % 2
inner = c.add_table(1, 1);                     % a NESTED table
disp(numel(c.tables));                         % 1   (tables() un-stubbed @ P6-4b)
disp(c.grid_span);                             % 1
c.vertical_alignment = mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER;
disp(c.vertical_alignment == mat2doc.enum.table.WD_CELL_VERTICAL_ALIGNMENT.CENTER);  % 1
c.vertical_alignment = [];                     % H3 -> removes w:vAlign
disp(isempty(c.vertical_alignment));           % 1
c.width = mat2doc.shared.Inches(1.5);
disp(double(c.width));                          % 1371600  (1.5 in EMU)
```

*Ported from python-docx v1.2.0: `src/docx/table.py::_Cell`*

---

(id-add-row-column)=
## `Table.add_row` / `Table.add_column` / `Table.cell` — grow the table, address a cell

`add_row()` appends a `<w:tr>` (one `<w:tc>` per grid column, each carrying the
column's width) and returns the new `Row_`; `add_column(width)` appends a
`<w:gridCol>` **and** a `<w:tc>` to **every** existing row, returning the new
`Column_`. `cell(row, col)` addresses the `_cells` grid (`col + row*_column_count`,
H1 `+1`) and returns the `Cell_` at that grid position (a spanned cell is returned
at each position it covers — see [span identity](#id-cell-merge)).

**Example**

```matlab
d = mat2doc.Document();
t = d.add_table(2, 2);
r = t.add_row();                               % a Row_ (a 3rd row)
disp(t.rows.len_());                           % 3
rc = r.cells;                                  % the new row's cells (1x2 Cell_)
rc(1).text = "X";  rc(2).text = "Y";
col = t.add_column(mat2doc.shared.Inches(1));  % a Column_ (a 3rd column)
disp(t.columns.len_());                        % 3
disp(t.cell(0, 0).text);                       % (row 0, col 0)
```

*Ported from python-docx v1.2.0: `src/docx/table.py::Table.add_row` / `::Table.add_column` / `::Table.cell`*

---

(id-cell-merge)=
## ★ Cell merge (`_Cell.merge`) + the cells grid walk (`_Row.cells` / `_Column.cells`)

**Syntax**

```matlab
t  = mat2doc.Document().add_table(2, 2);
m  = t.cell(0, 0).merge(t.cell(0, 1));   % horizontal merge -> a merged Cell_ (gridSpan=2)
m.text = "merged";                        % write into the merged cell
disp(m.grid_span);                        % 2
```

**The merge authoring path — `cell1.merge(cell2)` → `gridSpan`/`vMerge` in the
package.** `_Cell.merge(other)` (`table.py:237`) is the faithful three-liner
`tc = self._tc; tc_2 = other._tc; merged_tc = tc.merge(tc_2); return _Cell(merged_tc,
self._parent)` — Mat2Doc's `merge` calls the **byte-proven `CT_Tc.merge` engine**
(the P6-3a merge engine, risk-register #2) and returns a NEW `Cell_` over the
top-left merged `<w:tc>`, parented to `self._parent`. The engine writes `w:gridSpan`
(horizontal spans) and `w:vMerge` (`restart` on the top cell, a **bare `<w:vMerge/>`**
continuation below) into the package. If the requested region is not rectangular it
raises `mat2doc:InvalidSpanError` ("requested span not rectangular"), verbatim.

:::{important}
**★ The cell merge is byte-identical to python-docx and COM-verified.**
`cell1.merge(cell2)` through the full public API produces the **exact**
python-docx `word/document.xml` bytes across every merge geometry (Gate-3
`s0077`–`s0082`, each part-level 17/17 L1): horizontal `gridSpan`
(`9f626e2b…`), vertical `vMerge` (`e833fac8…`), 2×2 block (`ac155531…`),
merge-then-set-text (`c96a3351…`), and the 3×3 mixed merge with populated span
cells (`63b25b1a…`). **All merge shapes open clean in real Microsoft Word** (Word
16.0.20228; `com_verify_P6_tables.md`): the horizontal `gridSpan=2` reads as one
wide cell over two, the 2×2 block collapses to a single merged cell, and the 3×3
mixed merge lays out with the four labels (TL/TR/BLK/V) in the correct span
geometry — no repair prompt, no dropped content, clean round-trip. **Zero new
D-numbers.**
:::

**The cells grid walk — a merged cell is the SAME handle at each spanned grid
position (H5).** `_Row.cells` and `Table._cells` build the grid by **repeating one
`Cell_` handle** across every grid position a spanned cell covers, exactly as
python-docx's `_cells` repeats one `_Cell` reference. `Cell_` is a **handle class**,
so two accesses of a spanned position compare `==` (handle identity), matching
lxml's same-tc-same-`_Cell` semantics. A horizontal span appends the one handle
`grid_span` times; a `vMerge="continue"` continuation recurses to the `_tc_above`
span root (the private `iter_tc_cells_`), so the continuation position yields the
**root** cell's handle. The span-identity partition is independently proven at both
the handle level and the byte level — the 3×3 mixed-merge grid partition
`[0,0,2,3,3,5,3,3,5]` (first-identity index repeated at each spanned position) is
byte-companioned by `s0082` (17/17 L1 vs python-docx), and text written through a
**non-root** grid position lands in the span-root `<w:tc>` in both the handle read
and the serialized bytes.

:::{note}
**Fresh wrappers per access (VERIFY-2 — faithful, not a deviation).** Each
`cells` / `cell(…)` / `*_cells` call mints **fresh** `Cell_` wrappers, so cells
compared **across two separate accesses** are `~=` — exactly as python-docx's
recomputed `_cells` `@property` gives `table.cell(1,2) is table._cells[5]` →
`False`. Within a **single** access, span identity is preserved exactly. The
serialized bytes are unaffected. NOT a D-number.
:::

**Example**

```matlab
d = mat2doc.Document();
t = d.add_table(2, 2);
m = t.cell(0, 0).merge(t.cell(0, 1));          % horizontal merge (gridSpan=2)
disp(m.grid_span);                             % 2
m.text = "merged";                             % write into the merged cell
disp(t.cell(0, 0).text);                       % merged
row0 = t.rows.getitem_(0).cells;               % the merged row -> SAME handle twice
disp(row0(1) == row0(2));                       % 1   (span identity, H5)
row1 = t.rows.getitem_(1).cells;               % the unmerged row -> distinct handles
disp(row1(1) == row1(2));                       % 0
```

*Ported from python-docx v1.2.0: `src/docx/table.py::_Cell.merge` / `::_Row.cells` / `::_Column.cells`*

---

(id-iter-inner-content)=
## `iter_inner_content` — a table in the document / a section yields a `Table`

`iter_inner_content` (on `Document`, `_Body`/`BlockItemContainer`, and `Section`)
yields each `Paragraph` **or** `Table` in document order. Because `Paragraph` and
`Table` are distinct proxy types sharing no `matlab.mixin.Heterogeneous` base, the
result is a **`1×N` cell** in document order (mirroring
`Paragraph.iter_inner_content` / `Section.iter_inner_content`). The `CT_P` →
`Paragraph` wrapping and the `<w:tbl>` → `Table` wrapping are **both** ported fully
now — **the P5-3a C2 debt (the `w:tbl` branch that used to raise
`mat2doc:notYetPorted` owner P6-4a) is discharged**. H9: the Python generator is
materialized (no tree mutation during iteration); H10: `isinstance(element, CT_P)`
→ `isa(element, "mat2doc.oxml.text.CT_P")`.

**Example**

```matlab
d = mat2doc.Document();
d.add_table(1, 1);                          % one table, no leading paragraph on the default body
items = d.iter_inner_content();             % a 1xN cell (Paragraph | Table)
disp(class(items{end}));                    % mat2doc.table.Table
sec = d.sections.getitem_(0);               % the one default section
sitems = sec.iter_inner_content();          % C2 debt discharged: the w:tbl yields a Table
disp(class(sitems{1}));                     % mat2doc.table.Table (no notYetPorted)
```

---

## ★ PHASE 6 COMPLETE — the full table tier is byte-proven end-to-end and COM-verified

P6-4a landed the table **API surface** (the `Table` proxy + the `Rows_`/`Columns_`
collections + the `Row_`/`Column_` proxies + the `add_table` authoring path);
**P6-4b** closes it with `_Cell` (with `.merge`), `Table.add_row`/`add_column`, and
`_Row.cells`/`_Column.cells` — un-stubbing the 8 P6-4a table stubs plus the required
`BlockItemContainer.tables`. Both are behavioral + un-stub with **no registry rows**,
byte-neutral (M1 17/17, `document.xml` unchanged on a bare save), and carry **zero
new D-numbers**.

With P6-4b, **Phase 6 (tables) is COMPLETE**: the table **oxml layer** (14
`oxml/table.py` element classes — 7 leaves P6-1 + 4 property containers P6-2 +
`CT_Tc`/`CT_TcPr` the merge engine P6-3a + `CT_Tbl` P6-3b), the table **API tier**
(`Table` / `_Rows` / `_Columns` / `_Row` / `_Column` / `_Cell`), and the full
**authoring surface** (`add_table` / `add_row` / `add_column` / `merge`) are all
byte-proven end-to-end against python-docx v1.2.0 and **COM-verified in real
Microsoft Word** (all five table packages — plain, styled, horizontal-merge,
block-merge, 3×3 mixed-merge — open clean, merges honored, cell text intact;
`com_verify_P6_tables.md`). The **`CT_Tc` merge engine** — the project's
risk-register #2 and the single hardest WP — is byte-identical across the entire
merge matrix. **Zero new D-numbers across all of Phase 6.** Next: **Phase 7**
(images / drawing / inline pictures).
