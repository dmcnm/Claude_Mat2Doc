---
title: "mat2doc.text.ParagraphFormat / mat2doc.text.TabStops / mat2doc.text.TabStop — the text/parfmt + tabstops API tier"
---

# `mat2doc.text.ParagraphFormat` / `mat2doc.text.TabStops` / `mat2doc.text.TabStop` — the paragraph-format API tier

Ported from python-docx v1.2.0 `src/docx/text/parfmt.py::ParagraphFormat` and
`src/docx/text/tabstops.py::TabStops` / `::TabStop` (package `+mat2doc/+text/`).
`ParagraphFormat` is the **third P4 API-tier proxy** after [`Font`](text_font.md)
and [`Run`](text_run.md); `TabStops`/`TabStop` are the **first proxy *sequence***
of the port.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## What "API tier" means here — behaviorally equivalent, byte-neutral

`ParagraphFormat`, `TabStops` and `TabStop` are `ElementProxy` subclasses. Like
every P4 API proxy before them they add **no `register_element_cls` row, no oxml
class, and no serialization code** — nothing on the save path moves. Their entire
job is to **get/set** the correct `<w:pPr>` / `<w:tabs>` / `<w:tab>` XML by
delegating to the `CT_PPr` / `CT_TabStops` / `CT_TabStop` helpers byte-validated
at **P4-2**. So the equivalence bar for this WP is **behavioral**, not
byte-fixture: every get returns the same value python-docx returns, and every set
produces the same serialized bytes.

**M1 is trivially preserved.** The default template instantiates none of the
three classes and the save path is unchanged, so the M1 17/17 byte-neutrality
sweep holds (`mat2doc.Document().save()` → 17/17 byte-identical to the frozen
`references\s0001` reference). There is no stale-pin risk (no registry rows) and
no new D-number. Gate-3 confirmed the whole surface byte/value-identical
(`probe_diff` MATCH **120/120**) with regression **611/611**.

| Python `src/docx/...` | MATLAB | class |
|---|---|---|
| `text/parfmt.py::ParagraphFormat` | `+mat2doc\+text\` | `ParagraphFormat` |
| `text/tabstops.py::TabStops` | `+mat2doc\+text\` | `TabStops` |
| `text/tabstops.py::TabStop` | `+mat2doc\+text\` | `TabStop` |

`+mat2doc\+text` (`mat2doc.text.ParagraphFormat`) is distinct from the oxml
`+mat2doc\+oxml\+text` (`mat2doc.oxml.text.CT_PPr`); `ParagraphFormat` **wraps** a
pPr-owner, it is not a `CT_PPr`.

## The proxy shape — `ParagraphFormat` wraps the pPr *owner*, not the `pPr`

Unlike `Font` (which wraps a `<w:r>`), `ParagraphFormat(element, parent)` wraps
the **pPr's owner** — a `<w:p>` (`CT_P`) or a style pPr owner — and reaches its
properties through `self._element.pPr` on read (returning `[]`/None if the owner
has no `<w:pPr>` yet) and `self._element.get_or_add_pPr()` on write. Every
accessor delegates to a P4-2 `CT_PPr` `@property` helper (`jc_val`,
`first_line_indent`, `ind_left`/`ind_right`, `keepLines_val`/`keepNext_val`,
`pageBreakBefore_val`, `widowControl_val`, `spacing_*`); `ParagraphFormat` adds
**no oxml logic**.

`ElementProxy` reference semantics and the H5 element-identity `eq`/`ne` are
inherited unchanged.

(id-parfmt-tri-state)=
## The tri-state (H3) — `[]` is the "inherited" sentinel

Every `ParagraphFormat` property is **None-vs-value**. A `None`/`[]` return means
"this setting is inherited from the style hierarchy" (the property is not set on
this paragraph); assigning `[]` clears it, via inline `isequal(x, [])` (the
established Mat2Doc None idiom — no shared `isNone` helper). The tri-state
*removal* logic lives in the P4-2 `CT_PPr`/`CT_Ind`/`CT_Spacing` setters (unchanged
here): e.g. `alignment = []` removes `<w:jc>`; `line_spacing = []` leaves an
**empty** `<w:spacing/>` (attrs removed, element retained).

---

## `ParagraphFormat`

**Syntax**

```matlab
p  = mat2doc.oxml.OxmlElement("w:p");        % a CT_P (pPr owner)
pf = mat2doc.text.ParagraphFormat(p);        % wrap the paragraph's formatting
pf.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
pf.left_indent = mat2doc.shared.Pt(36);      % <w:ind w:left="720"/>
pf.line_spacing = 1.5;                        % <w:spacing w:line="360" w:lineRule="auto"/>
ts = pf.tab_stops;                            % a TabStops view (read-only, lazy)
```

**Description**

Paragraph formatting: alignment, indentation, spacing, line spacing, widow/orphan
control, keep-together/keep-with-next, page-break-before, and tab stops. An
`ElementProxy` over a pPr owner; every accessor delegates to the P4-2 `CT_PPr`
helpers. Reads a `<w:pPr>` lazily (absent → `[]`); writes `get_or_add_pPr()`
first.

### `alignment` + the bool tri-states

`alignment` is a `WD_ALIGN_PARAGRAPH` member or `[]` (`<w:jc w:val="...">`). Each
of **`keep_together`** (`<w:keepLines>`), **`keep_with_next`** (`<w:keepNext>`),
**`page_break_before`** (`<w:pageBreakBefore>`), and **`widow_control`**
(`<w:widowControl>`) is a tri-state bool delegating to `CT_OnOff` (D-delta-1):
`true`→a bare tag, `false`→`@w:val="0"`, `[]`→removed.

### The indents (sign both ways) + spacing

- **`first_line_indent`** — a `Length` or `[]`. A **positive** value writes
  `<w:ind w:firstLine="...">`; a **negative** value writes
  `<w:ind w:hanging="...">` (the sign split, and the `Length(-hanging)` negation
  on get, live in `CT_PPr`/`CT_Ind`, P4-2). `Pt(24)`→`w:firstLine="480"` (get
  304800); `Pt(-24)`→`w:hanging="480"` (get −304800); `Pt(0)`→`w:firstLine="0"`.
- **`left_indent`** / **`right_indent`** — signed `Length`/`[]` (`@w:left` /
  `@w:right`); `Pt(-36)`→`w:left="-720"`.
- **`space_before`** / **`space_after`** — `Length`/`[]` (`@w:before` /
  `@w:after`); `Pt(6)`→`w:before="120"`.

### `line_spacing` / `line_spacing_rule` — the float-vs-`Length` detection (the meaty logic, `parfmt.py:102-160,256-286`)

`line_spacing` is polymorphic: a **float** multiple of single-spacing, an
**absolute `Length`**, or `[]`. The get/set logic keys off the wrapped
`<w:spacing>` `@w:line` + `@w:lineRule`:

**On get** (`_line_spacing`):

| `<w:spacing>` state | `line_spacing` returns | `line_spacing_rule` returns |
|---|---|---|
| no `@w:line` | `[]` | (from `@w:lineRule`, or `[]`) |
| `@w:lineRule="auto"` (MULTIPLE) | `line / Pt(12)` — a **float** (D-STYPE-1) | special-member map below |
| any other rule (`exact`/`atLeast`) | the absolute `Length` (EMU) | `EXACTLY` / `AT_LEAST` |

The MULTIPLE→rule special-member map (`_line_spacing_rule`): `@w:line="240"`
(≡ `Twips(240)`) → **`SINGLE`**, `"360"` → **`ONE_POINT_FIVE`**, `"480"` →
**`DOUBLE`**; any other MULTIPLE value → bare **`MULTIPLE`**.

**On set**, `line_spacing` dispatches on the **argument type**:

| assigned | `<w:spacing>` result | rule |
|---|---|---|
| a `Length` (e.g. `Pt(18)`) | `@w:line="360" @w:lineRule="exact"` | EXACTLY — **unless already `AT_LEAST`**, then preserved |
| a non-`Length` numeric (a float multiple, e.g. `1.75`) | `@w:line="420" @w:lineRule="auto"` (`Emu(value·Twips(240))`) | MULTIPLE |
| `[]` | attrs removed, empty `<w:spacing/>` retained | — |

`line_spacing_rule` set writes the matching Twips line for `SINGLE`/
`ONE_POINT_FIVE`/`DOUBLE` (each + MULTIPLE); `AT_LEAST`/`EXACTLY` write
`@w:lineRule` only; `[]` removes the rule.

:::{note}
`Pt(12) == Twips(240) == 152400 EMU` — the single-space unit. `parfmt.py` uses
`Pt(12)` on the **get** side and `Twips(240)` on the **set** side; both are ported
verbatim. Under **D-STYPE-1** the MULTIPLE get returns a plain MATLAB `double`
(e.g. `1.5`) where Python returns a `float`; the value is identical (`%.12g`) and
the type distinction is unobservable in MATLAB by design — output-invisible,
zero new D-number.
:::

### `tab_stops` (read-only, lazy)

`tab_stops` is a `@lazyproperty` returning a [`TabStops`](#tabstops) view of the
paragraph's tab stops, **computed and cached** on first access (the
`mat2doc.shared.lazyproperty` cache + logical computed-flag idiom — never the
`isempty` sentinel). `get_or_add_pPr()` runs once and adds a `<w:pPr>` if absent
(Python lazyproperty body semantics). Read-only (no `set.tab_stops`), mirroring
Python's `AttributeError`.

**Example**

```matlab
p  = mat2doc.oxml.OxmlElement("w:p");
pf = mat2doc.text.ParagraphFormat(p);
pf.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;   % <w:jc w:val="center"/>
pf.first_line_indent = mat2doc.shared.Pt(-18);                % <w:ind w:hanging="360"/>
pf.line_spacing = 1.75;                                       % <w:spacing w:line="420" w:lineRule="auto"/>
disp(pf.line_spacing);                                        % 1.75  (a double, D-STYPE-1)
disp(string(pf.line_spacing_rule));                           % MULTIPLE
pf.line_spacing = mat2doc.shared.Pt(18);                      % <w:spacing w:line="360" w:lineRule="exact"/>
disp(string(pf.line_spacing_rule));                           % EXACTLY
```

*Ported from python-docx v1.2.0: `src/docx/text/parfmt.py::ParagraphFormat`*

---

(tabstops)=
## `TabStops`

**Syntax**

```matlab
pf = mat2doc.text.ParagraphFormat(mat2doc.oxml.OxmlElement("w:p"));
ts = pf.tab_stops;                             % a TabStops view over the w:pPr
ts.add_tab_stop(mat2doc.shared.Inches(1));     % LEFT / SPACES defaults
n     = ts.len_();                             % Python len(ts)
first = ts.getitem_(0);                         % Python ts[0]  (H1, 0-based arg)
ts.clear_all();
```

**Description**

A **sequence** of [`TabStop`](#tabstop) objects — the tab stops of a paragraph or
style. In python-docx it is a collections-style sequence reached via
`ParagraphFormat.tab_stops`, supporting iteration, indexed access, `del`, and
`len()`. An `ElementProxy` over the **`w:pPr`** (`CT_PPr`), reaching the
`<w:tabs>` container on demand via `pPr.tabs` / `pPr.get_or_add_tabs()` /
`pPr._remove_tabs()`.

(id-tabstops-interim-surface)=
### First proxy sequence — the interim explicit surface (VERIFY-COLLECTION)

`TabStops` is the **first proxy sequence** in Mat2Doc. The shared collection base
(`RedefinesParen`; design.md §2's `x[i] → x(i+1)`, `len(x) → numel(x)`,
`for s in x → x.to_array()` mapping) is a **future work package** and does not
exist yet. Following the established cross-toolbox precedent — Mat2Ppt's
`_GradientStops` → `GradientStops_.m`, which ports the sequence surface as
explicit methods and defers native `()` indexing with the identical note — the
Python sequence dunders are ported here as **explicit methods**:

| Python dunder | MATLAB method | note |
|---|---|---|
| `__getitem__(idx)` | `getitem_(idx)` | keeps the **Python 0-based** argument (H1) |
| `__delitem__(idx)` | `delitem_(idx)` | Python 0-based |
| `__len__()` | `len_()` | `numel(to_array())` |
| `__iter__()` | `to_array()` | materialized 1×N `TabStop` array (H9) |

:::{important}
**Interim surface, accepted at Gate 2.** When the shared `RedefinesParen`
collection base lands, `TabStops` should **derive from it** and expose native
1-based indexing — `tab_stops(1)` ≡ `getitem_(0)`. Do **not** retrofit now;
`getitem_` intentionally keeps the Python 0-based argument for line-for-line
fidelity. This is the one architectural hand-off recorded for the future
collection WP.
:::

### H1 indexing + the sequence operations

`getitem_`/`delitem_` take the **Python 0-based** index and hit the 1-based MATLAB
child list with an explicit `+1` (the H1 site). Full Python `list` int-index
semantics are replicated: a **negative** index counts from the end; an
**out-of-range** index raises `mat2doc:IndexError` with the faithful message.

- **`add_tab_stop(position, alignment, leader)`** — creates a `<w:tab>` and inserts
  it in **position order** (`insert_tab_in_order`), so `<w:tabs>` stays sorted by
  `@w:pos`. `alignment` defaults to `WD_TAB_ALIGNMENT.LEFT`, `leader` to
  `WD_TAB_LEADER.SPACES`. Adding `Inches(2), Inches(0.5), Inches(1)` (out of order)
  yields `<w:tab w:pos="720"/>`, `1440`, `2880` — pos-sorted.
- **`getitem_(idx)`** — `getitem_(0)` → first stop; `getitem_(-1)` → last (negative
  wrap); an out-of-range index → `mat2doc:IndexError "list index out of range"`
  (verbatim); on an empty `TabStops` → `"TabStops object is empty"`.
- **`delitem_(idx)`** — removes the stop at `idx`; out-of-range →
  `mat2doc:IndexError "tab index out of range"` (verbatim); deleting down to empty
  **removes the `<w:tabs>` element** entirely.
- **`clear_all()`** — removes all stops (drops `<w:tabs>`).
- **`len_()`** / **`to_array()`** — element child count / a fresh 1×N `TabStop`
  array (each a fresh view — `TabStop` objects are not cached, H5).

**Example**

```matlab
pf = mat2doc.text.ParagraphFormat(mat2doc.oxml.OxmlElement("w:p"));
ts = pf.tab_stops;
ts.add_tab_stop(mat2doc.shared.Inches(2));                       % pos 2880
ts.add_tab_stop(mat2doc.shared.Inches(0.5));                     % pos 720   (sorts first)
ts.add_tab_stop(mat2doc.shared.Inches(1), ...                    % pos 1440
    mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER, ...
    mat2doc.enum.text.WD_TAB_LEADER.DOTS);
disp(ts.len_());                                                 % 3
disp(double(ts.getitem_(0).position));                          % 457200   (0.5 in, sorted first)
disp(double(ts.getitem_(-1).position));                         % 1828800  (2 in, negative wrap)
ts.delitem_(1);                                                  % remove the middle stop
```

*Ported from python-docx v1.2.0: `src/docx/text/tabstops.py::TabStops`*

---

(tabstop)=
## `TabStop`

**Syntax**

```matlab
tab = ts.getitem_(0);                                    % from a TabStops
pos = tab.position;                                      % a Length (signed twips)
tab.alignment = mat2doc.enum.text.WD_TAB_ALIGNMENT.CENTER;
tab.leader    = mat2doc.enum.text.WD_TAB_LEADER.DOTS;
tab.position  = mat2doc.shared.Inches(2);                % re-sorted in place
```

**Description**

An individual tab stop, wrapping a single `<w:tab>` (`CT_TabStop`). Accessed via
list semantics on its containing `TabStops`. All three properties are read/write
and delegate to the P4-2 `CT_TabStop` descriptors:

- **`position`** — a `Length` in signed twips (`@w:pos`, `RequiredAttribute`).
- **`alignment`** — a `WD_TAB_ALIGNMENT` member (`@w:val`, `RequiredAttribute`).
- **`leader`** — a `WD_TAB_LEADER` member (`@w:leader`, `OptionalAttribute`, default
  `SPACES`). The default is written **attribute-absent**: `add_tab_stop(Inches(1))`
  with the default leader serializes `<w:tab w:pos="1440" w:val="left"/>` — no
  `@w:leader` — and reads back `SPACES` (the `OptionalAttribute` default-removal,
  D-delta-1/-2). Explicit `DOTS` → `@w:leader="dot"`.

(id-tabstop-position-quirk)=
### The `position`-setter identity quirk (faithful, `tabstops.py:118-123`)

Setting `position` is the one non-trivial accessor: it **re-inserts** a fresh
`<w:tab>` at the new pos-order position and removes the old element, so the parent
`<w:tabs>` stays sorted. python-docx reassigns **only `self._tab`** (to the
re-inserted element), leaving `self._element` on the **old, now-detached**
element. Mat2Doc ports this verbatim: after a `position` set, `element()`/`eq()`
still reflect the **original** element while the three properties reflect the
**new** one.

So with a fresh proxy `t2 = ts.getitem_(0)` over the same element, `t == t2` is
`true`; after `t.position = Inches(1.5)`, `t == t2` **stays `true`** (both proxies'
eq basis is the old detached element), yet `t.position` reads the **new** value
and `to_array` re-sorts. This is a **genuine python-docx behavior**, correctly
**not** "corrected" (Gate-2 confirmed against `tabstops.py:118-123`).

**Example**

```matlab
pf = mat2doc.text.ParagraphFormat(mat2doc.oxml.OxmlElement("w:p"));
ts = pf.tab_stops;
ts.add_tab_stop(mat2doc.shared.Inches(0.5));
ts.add_tab_stop(mat2doc.shared.Inches(1));
ts.add_tab_stop(mat2doc.shared.Inches(2));
t  = ts.getitem_(0);                          % the 0.5 in stop
t2 = ts.getitem_(0);                          % fresh proxy, same element
disp(t == t2);                                % true
t.position = mat2doc.shared.Inches(1.5);      % re-inserts at 1.5 in, removes 0.5 in
disp(t == t2);                                % true  (eq still tracks the OLD element)
disp(double(t.position));                     % 1371600  (the NEW value, via _tab)
```

*Ported from python-docx v1.2.0: `src/docx/text/tabstops.py::TabStop`*

---

## The G-scenario — the closest guard yet to M2's `add_paragraph` path

Gate-3 froze an end-to-end **document byte pin** (`s0027`): a fresh
`mat2doc.Document()` with **four formatted paragraphs**, each `<w:p>` formatted
through a `ParagraphFormat` (1.75-multiple; `AT_LEAST` then `Pt(18)`;
hanging-indent + three pos-sorted tab stops with `<w:tabs>` before `<w:ind>`;
`ONE_POINT_FIVE` + center + widow + keep-together in CT_PPr schema order), saved
and compared part-by-part against a frozen python-docx reference. Result:
**17/17 parts byte-identical**, with `word/document.xml` byte-identical.

This exercises `ParagraphFormat`/`TabStops`/`TabStop` **writing real `w:pPr` into a
real `word/document.xml` on the actual save pipeline** — including the CT_PPr
child-ordering hazard (H11) and the pos-sorted `<w:tabs>`. It is the **closest
proof to M2's `add_paragraph` formatting path** available before the `Paragraph`
proxy lands (P4-6).

## Deviation posture — 0 new D-numbers

`ParagraphFormat`, `TabStops` and `TabStop` add no output-visible divergence. The
two standing adopt-only classes that touch this WP:

- **D-STYPE-1** (int/float indistinguishability on `Length`) — **engaged and
  value-confirmed**: `line_spacing` MULTIPLE get returns a MATLAB `double` where
  Python returns a `float`; the value is identical (`%.12g`) and the type
  distinction is unobservable in MATLAB by design (no output byte differs — the
  set path serializes through the P4-2 `ST_` simpletypes → integer `@w:line`).
- **D-serializer-nsdecl** — **does not engage**: every wrapped element is
  `w:`-only (no foreign namespace), so no `ns0` is ever minted and all serhex byte
  pins (including the XML prolog) are byte-identical; the G-scenario/M1 document
  parts are byte-identical.

Gate-3 recorded `PASS-DEVIATION(D-STYPE-1, D-serializer-nsdecl)` as adopt-only
context only — operational verdict a clean **PASS** with zero measured
byte/value deviation, **zero new D-numbers**.
