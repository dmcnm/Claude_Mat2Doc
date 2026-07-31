---
title: "mat2doc.oxml.table — the table oxml LEAF layer (CT_TblWidth · CT_Height · CT_TblGrid · CT_TblGridCol · CT_TblLayoutType · CT_VerticalJc · CT_VMerge)"
---

# `mat2doc.oxml.table` — the table oxml LEAF layer (the 7 width / height / grid / merge leaves)

Ported from python-docx v1.2.0 `src/docx/oxml/table.py` (in the NEW package
`+mat2doc/+oxml/+table/`) — the **seven LEAF element classes** of the table
object model: the width union `CT_TblWidth`, the row-height leaf `CT_Height`, the
column-grid pair `CT_TblGrid` / `CT_TblGridCol`, the layout-mode leaf
`CT_TblLayoutType`, the cell vertical-alignment leaf `CT_VerticalJc`, and the
vertical-merge leaf `CT_VMerge` — plus the **seven `register_element_cls` rows**
they require (`src/docx/oxml/__init__.py` :171, :174, :175, :181, :183, :185,
:186).

:::{note}
**★ Tables tier opens — the table LEAF classes are done; the container classes
(TblPr / TrPr / Row → the CT_Tc merge engine) come next.** This is the **first
work package of Phase 6** (tables). It ports the seven table *leaf* elements —
attribute-only (or single-list) classes with no dependence on any other table
`CT_*` — and registers their tags. Because `default.docx` **contains no table**,
this WP is **M1-NEUTRAL and flip-neutral**: no default.docx part carries any of
the seven tags, so nothing transits the new classes on the M1 parse path (M1
stays 17/17, **zero re-pins**). The container classes follow: `CT_TblPr` /
`CT_TblPrEx` / `CT_TrPr` / `CT_Row` at **P6-2**, then the **CT_Tc merge engine**
(gridSpan / vMerge, the block-item-container seam) at **P6-3a — the hardest WP of
the project**, `CT_TcPr` / `CT_Tbl` at **P6-3b**, and the `Table` / `_Rows` /
`_Cell` API at **P6-4a / P6-4b**.
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

:::{note}
**Milestone flag (COM).** The s0060 table subtrees emit output-visible new table
structure (`w:tblGrid` / `w:gridCol` / `w:tblLayout` / `w:trHeight` / `w:tcW` /
`w:vAlign` / `w:vMerge`) but are **not yet reachable through a Mat2Doc-generated
package** — the container classes `CT_Tbl` / `CT_Row` / `CT_Tc` (P6-2 / P6-3) are
needed to build a real table end-to-end. The frozen fixtures are routed to
mso-office-verifier (the Word COM oracle) at the next milestone sweep once the
containers land; the full table Word-COM sweep (a plain grid **and** a merged-cell
package) is booked at **P6-4b**.
:::

---

(id-table-next)=
## Table leaves done — the containers (P6-2) and the CT_Tc merge engine (P6-3a) next

The table oxml **leaf** layer is now whole: the width union (`CT_TblWidth`), the
row-height leaf (`CT_Height`), the column grid (`CT_TblGrid` / `CT_TblGridCol`),
the layout-mode leaf (`CT_TblLayoutType`), and the two cell leaves (`CT_VerticalJc`
/ `CT_VMerge`) — registered M1-neutrally, byte-proven on a real table's subtrees,
with **zero new D-numbers**.

What remains in **Phase 6** is the container tier that consumes these leaves:

- **P6-2** — `CT_TblPr` / `CT_TblPrEx` / `CT_TrPr` / `CT_Row`. This is the first
  **non-neutral** table WP: registering `w:tblPr` puts it on the **live**
  `word/styles.xml` parse path (multiple `w:tblPr` subtrees live in the shipped
  styles), so P6-2 is a byte-critical registry-flip WP (the P4-6 pattern), not
  M1-neutral like P6-1. Its porter must also replicate the `CT_Jc` cross-enum
  reuse (`CT_TblPr.alignment` reads the same registered `CT_Jc` typed
  `WD_ALIGN_PARAGRAPH`, cast to `WD_TABLE_ALIGNMENT`).
- **P6-3a** — the **CT_Tc merge engine** (gridSpan / vMerge, the block-item
  container seam). This is the **hardest WP of the project**; it gets its own
  **dedicated pre-launch plan-audit** (risk-register #2), which also adjudicates
  the Row ↔ Tc ↔ TcPr ↔ Tbl mutually-recursive dependency seam.
- **P6-3b** — `CT_TcPr` / `CT_Tbl` + the registry completion and the tbl
  un-defer sweep (the tag-based `w:tbl`→generic sites auto-upgrade when
  `w:tbl`→`CT_Tbl` registers).
- **P6-4a / P6-4b** — the `Table` / `_Rows` / `_Columns` / `_Cell` API (which
  discharges the P5-3a `Section.iter_inner_content` `w:tbl`→`Table` debt and the
  `Document.add_table` / `tables` stubs), then `_Cell.merge` + `add_row` /
  `add_column` and the table Word-COM sweep.
