---
title: "mat2doc.oxml.text — the run-properties (rPr) element tier"
---

# `mat2doc.oxml.text` — the text/font element tier (`CT_RPr` + rPr children)

Ported from python-docx v1.2.0 `src/docx/oxml/text/font.py` (7 element classes,
package `+mat2doc/+oxml/+text/`) plus the two shared leaves
`CT_OnOff` / `CT_String` from `src/docx/oxml/shared.py`
(package `+mat2doc/+oxml/+shared/`), and the 28 font-block
`register_element_cls` rows of `src/docx/oxml/__init__.py:198-225`.

This is the **first real `w:`-content element tier** in the port. Everything
before it was infrastructure — the parser, the xmlchemy descriptor engine
(`oxml_xmlchemy`), the simple-type validators (`simpletypes`), the enums
(`enums` / `enums_content`), and the OPC/document plumbing. `CT_RPr` (the
`<w:rPr>` run-properties container) is the first class where the child-descriptor
engine meets a real, strictly-ordered OOXML schema sequence — and it sits
**directly on the M2 byte-critical path**.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The two-package layout

`font.py` (a `text/` sub-module) lands in `+oxml\+text`; `shared.py` (a top-level
oxml module) lands in `+oxml\+shared` — following the established
module→subpackage convention (`document.py`→`+document`, `coreprops.py`→
`+coreprops`). `+oxml\+shared` (`mat2doc.oxml.shared.CT_*`) does **not** collide
with the top-level `mat2doc.shared` (RGBColor / the `Length` family) — they are
distinct packages.

| Python `src/docx/...` | MATLAB | classes |
|---|---|---|
| `oxml/text/font.py` | `+mat2doc\+oxml\+text\` | `CT_RPr`, `CT_Color`, `CT_Fonts`, `CT_Highlight`, `CT_HpsMeasure`, `CT_Underline`, `CT_VerticalAlignRun` |
| `oxml/shared.py` (2 of its leaves) | `+mat2doc\+oxml\+shared\` | `CT_OnOff`, `CT_String` |

**C3 fold-in.** `CT_OnOff` / `CT_String` are the target of 21 of the 28
font-block registry rows (`w:b`, `w:i`, `w:rStyle`, …), a leaf pair with no
unported dependencies of their own, so the P3→P4 boundary audit folded them into
this WP (`CT_DecimalNumber`, the third `shared.py` leaf, waits for its first
consumer at P4-2). All 9 classes plus the 28 registry rows are one work package.

## First real content oxml — M1 stays byte-identical

Registering the 28 font-block tags changes the **parse path** of parts that
already exist. `word/styles.xml` in the shipped `default.docx` (349 458 B)
contains **464 real `<w:rPr>` blocks**; once the registry rows are live those
blocks parse to `CT_RPr` / `CT_OnOff` / `CT_Color` / … instead of a generic
`XmlElement`. Because every `CT_*` class extends `BaseOxmlElement` and reserializes
through the same `serialize_part_xml` walk with **no serialization override**,
registering the tags changes only the parsed node **class**, never its **bytes**.

The **M1 17/17 byte-neutrality sweep holds**: `mat2doc.Document().save()` →
unzip → all 17 parts byte-identical to the frozen `references\s0001` reference,
with `word/styles.xml` (349 458 B) and `word/document.xml` (1 548 B) L1
byte-identical (Gate-3 `pkgcompare` L0 PASS + 16 XML L1 + 1 bin). The
464-block `styles.xml` also round-trips parse→serialize L1 (349 458 B in ==
out). This is the standing obligation every future `CT_*` WP inherits.

(id-h11-successor-ordering)=
## The H11 successor-ordering mechanism — the M2-critical pattern

`CT_RPr` declares its child elements as **27 `ZeroOrOne` descriptors** over a
single **39-entry `_tag_seq`** (`font.py:79-119`, ported verbatim as the Constant
`TAG_SEQ`). OOXML mandates that a `<w:rPr>`'s children appear in **schema order** —
`w:rStyle`, `w:rFonts`, `w:b`, `w:bCs`, `w:i`, … — regardless of the order the API
creates them. Get the order wrong and Word shows a **repair prompt = FAIL**, byte
results notwithstanding. This is the whole M2 byte-risk, and P4-1a is where the
descriptor engine first meets it.

**How the re-sort works.** Each descriptor carries, as its `successors` list, the
schema slice `_tag_seq[N:]` — every tag that must sort **after** it. When the API
adds a child, `insertChildInSequence` (the P1-3b engine method) hands those
successor tags to the P1-3a tree-op `insert_element_before`, which scans them in
order, `addprevious`-es the **first successor already present** in the tree, and
**appends** when none is present. So a new child always lands *before the first
schema-later sibling already there* — the children stay in schema order no matter
the creation order.

**The H1 base shift, applied once.** A Python slice `_tag_seq[N:]` (0-based)
becomes the MATLAB slice `TAG_SEQ(N+1:end)` (1-based). The `+1` is applied **once,
at the slice declaration** in `CT_RPr`; the engine itself contains no index
arithmetic. `CT_RPr`'s help header carries the full per-descriptor slice table.

**The non-contiguous slices.** Several `_tag_seq` tags between `w:color` and
`w:oMath` have **no descriptor on `CT_RPr`** (`w:spacing`, `w:w`, `w:kern`,
`w:position`, `w:szCs`, `w:effect`, `w:bdr`, `w:shd`, `w:fitText`, `w:em`,
`w:lang`, `w:eastAsianLayout`), so the descriptor slices **jump** across them —
`sz = _tag_seq[24:]`, `highlight = _tag_seq[26:]`, `u = _tag_seq[27:]`,
`vertAlign = _tag_seq[32:]`, `specVanish = _tag_seq[38:]`. The intervening tags
are **kept in the successor lists, not collapsed** — so when a parsed `<w:rPr>`
already holds one of them (a real `w:spacing`/`w:szCs`/`w:lang`), a newly-added
`w:sz`/`w:u`/`w:vertAlign` still lands in exactly the right place relative to it.
`oMath`'s slice `_tag_seq[39:]` → `TAG_SEQ(40:end)` is a legal empty `1×0` slice
(no successors → always appends last).

**Why this is byte-exact.** Gate-2 built five full-27-descriptor orderings
(reversed, two shuffles, tail-first adversarial, canonical) in **both** python-docx
1.2.0 and Mat2Doc via identical `get_or_add_*` sequences: every ordering
converged to the same canonical schema order and the serialized bytes matched
PY==ML for all five. Gate-3 added the decisive `parsed_nondesc` case — 8
descriptors inserted into a **parsed** `<w:rPr>` already holding the
non-descriptor `w:spacing`/`w:szCs`/`w:lang` — yielding
`rStyle, b, color, spacing, sz, szCs, highlight, u, vertAlign, lang, specVanish`,
**byte-identical to the python-docx oracle**. This exercises the non-contiguous
slices against real intervening siblings. `insert_element_before` /
`first_child_found_in` supply the ordering logic; P4-1a supplies only the correct
successor tag-lists, and they are correct.

:::{note}
`font.py`'s own docstring calls `_tag_seq` a "40-element tuple"; the literal tuple
is **39 entries** (`font.py:80-118`). The port uses the actual 39. The slices
resolve identically under either count (`specVanish=[38:]`→`(w:oMath,)`,
`oMath=[39:]`→`()`), and the live scrambled-order oracle match is the definitive
proof. Gate-2 recounted the AST mechanically — 39.
:::

## `_new_color` — the metaclass creator override

`CT_RPr` defines `_new_color` (`font.py:149-151`), which xmlchemy's metaclass lets
**win** over the generated default creator. The port carries this as `new_color_`
(design.md §2 OVERRIDE alert, the same shape as Mat2Ppt's `CT_Chart._new_title`):
`get_or_add_color` / `add_color_` route through `new_color_`, which seeds a
`<w:color w:val="000000"/>` (RGB black) fragment via `parse_xml`, rather than the
generic engine creator that would make an empty `<w:color/>`. The other 26
descriptors use the generic `BaseOxmlElement` engine
(`getChild`/`getOrAddChild`/`newChild`/`addChild`/`insertChildInSequence`/
`removeChild`). On a fresh `rPr`, `get_or_add_color()` returns a `CT_Color` whose
`val` is `RGBColor(0,0,0)`, and it is idempotent (the same live child handle on
re-call, H5).

## `CT_OnOff` / `CT_String` tri-state (the C3 leaves)

**`CT_OnOff`** (`shared.py:27-36`) backs the boolean run/paragraph properties
(`w:b`, `w:i`, `w:caps`, …) — the target of **20** of the 28 registry rows. Its
`val` is an `OptionalAttribute("w:val", ST_OnOff, default=True)` — a **non-None
default**, which drives the D-delta-1 tri-state through
`BaseOxmlElement.setAttrTyped`:

- **get when `@w:val` absent** → `True` (the default; a bare `<w:b/>` means
  bold-on);
- **set `val = False`** → `ST_OnOff.to_xml(False)` = `"0"` → `<w:b w:val="0"/>`;
- **set `val = True`** (== default) → `@w:val` **removed** → bare `<w:b/>`
  (D-delta-1: `value == default` erases the attribute);
- **set `val = []`** (None) → `@w:val` **removed** (D-delta-1: `value is None`).

`ST_OnOff.from_xml` maps `1`/`true`/`on` → `true` and `0`/`false`/`off` → `false`.
All four states and the six `from_xml` tokens are Gate-3 `probe_diff`-exact vs
the oracle.

**`CT_String`** (`shared.py:39-52`) backs the `w:rStyle` row: a
`RequiredAttribute("w:val", ST_String)` plus a static factory
`new(nsptagname, val)` that builds a loose element and sets its `val`. The Python
`cast(CT_String, …)` is a type hint only and is dropped.

---

## `CT_RPr`

**Syntax**

```matlab
rPr = mat2doc.oxml.OxmlElement("w:rPr");   % a CT_RPr (registered)
child = rPr.get_or_add_x()                 % x in {rStyle,rFonts,b,...,oMath}
val   = rPr.style;                         % @property helpers (font.py 153-305)
```

**Description**

The `<w:rPr>` run-properties container — the M2 byte-critical element. Its 27
`ZeroOrOne` child descriptors ride the non-contiguous H11 successor slices of the
39-entry `TAG_SEQ` (see [the H11 mechanism above](#id-h11-successor-ordering)),
so scrambled `get_or_add_*` adds always reserialize in schema order. Each
descriptor generates the docx family `get.x` / `get_or_add_x` / `new_x_` /
`insert_x_` / `add_x_` / `remove_x_` (underscore rotation of Python's
`_new_x`/`_insert_x`/`_add_x`/`_remove_x`; `get_or_add_x` is public). `w:color`
overrides its creator via `_new_color` (black-seed).

The 8 `@property` helpers (`font.py:153-305`) are all ported live:
`highlight_val`, `rFonts_ascii`, `rFonts_hAnsi`, `style`, `subscript`,
`superscript`, `sz_val`, `u_val`. Note the **asymmetry**: the `rFonts_hAnsi`
setter does **not** remove `w:rFonts` on `None` — it sets `hAnsi = None`,
removing only `@w:hAnsi` (`font.py:201-206`), unlike `rFonts_ascii`, which removes
the whole `w:rFonts`. The `_get_bool_val` / `_set_bool_val` helpers
(`font.py:307-319`) port to `get_bool_val_` / `set_bool_val_` (consumed by the
`Font` proxy, a later WP).

**Example**

```matlab
rPr = mat2doc.oxml.OxmlElement("w:rPr");
rPr.get_or_add_u().val = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
rPr.get_or_add_b();                          % w:b created AFTER w:u
order = arrayfun(@(e) string(e.nsptag_str), rPr.xpath("./*"));
disp(order);   % "w:b"    "w:u"  -- w:b re-sorted BEFORE w:u (H11 schema order)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/font.py::CT_RPr`*

---

## `CT_Color`

**Syntax**

```matlab
c = mat2doc.oxml.OxmlElement("w:color");
c.val = mat2doc.shared.RGBColor(60, 47, 128);   % RequiredAttribute (ST_HexColor)
c.themeColor = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;  % OptionalAttribute
```

**Description**

The `<w:color>` element. `val` is a `RequiredAttribute("w:val", ST_HexColor)` →
an `RGBColor` (or the literal string `"auto"`); a missing `@w:val` raises
`mat2doc:InvalidXmlError`. `themeColor` is an
`OptionalAttribute("w:themeColor", MSO_THEME_COLOR)` with default `None` → an
`MSO_THEME_COLOR` member, or `[]` when absent (set `[]` removes the attribute,
H3). The enum simple-type is named by its **fully-qualified** class name
(`mat2doc.enum.dml.MSO_THEME_COLOR`) so `resolveTypeCls_` dispatches to `+enum`
verbatim.

**Example**

```matlab
c = mat2doc.oxml.OxmlElement("w:color");
c.val = mat2doc.shared.RGBColor(60, 47, 128);   % serializes @w:val="3C2F80"
disp(class(c.val));                              % 'mat2doc.shared.RGBColor'
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/font.py::CT_Color`*

---

## `CT_Fonts` / `CT_Highlight` / `CT_HpsMeasure` / `CT_Underline` / `CT_VerticalAlignRun`

**Syntax**

```matlab
sz = mat2doc.oxml.OxmlElement("w:sz");   sz.val = mat2doc.shared.Pt(12);   % -> "24"
u  = mat2doc.oxml.OxmlElement("w:u");    u.val  = mat2doc.enum.text.WD_UNDERLINE.SINGLE;
```

**Description**

The five leaf children of `CT_RPr`, each a thin attribute holder delegating to a
P3-2 simple type or a P3-3 enum:

- **`CT_Fonts`** (`w:rFonts`) — `ascii` / `hAnsi` = `OptionalAttribute` of
  `ST_String`, default `None` (`[]` when absent).
- **`CT_Highlight`** (`w:highlight`) — `val` = `RequiredAttribute` of
  `WD_COLOR_INDEX` (e.g. `"yellow"` ↔ `YELLOW`).
- **`CT_HpsMeasure`** (`w:sz`) — `val` = `RequiredAttribute` of `ST_HpsMeasure`
  (half-points): `Pt(12)` serializes `w:val="24"`; reading back gives a `Length`
  (152 400 EMU).
- **`CT_Underline`** (`w:u`) — `val` = `OptionalAttribute` of `WD_UNDERLINE`,
  default `None`. Its setter removes **all** existing `w:u` before re-adding
  (a duplicated `<w:u>` input collapses to a single value).
- **`CT_VerticalAlignRun`** (`w:vertAlign`) — `val` = `RequiredAttribute` of
  `ST_VerticalAlignRun` (`baseline` / `superscript` / `subscript`).

Required-attribute getters raise `mat2doc:InvalidXmlError` when the attribute is
absent; optional-default-`None` getters return `[]`, and setting `[]` removes the
attribute (H3).

**Example**

```matlab
sz = mat2doc.oxml.OxmlElement("w:sz");
sz.val = mat2doc.shared.Pt(12);   % serializes @w:val="24" (half-points)
disp(double(sz.val));              % 152400  (a Length/Emu)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/font.py::CT_Fonts` /
`CT_Highlight` / `CT_HpsMeasure` / `CT_Underline` / `CT_VerticalAlignRun`*

---

## `CT_OnOff`

**Syntax**

```matlab
b = mat2doc.oxml.OxmlElement("w:b");   % a CT_OnOff
tf = b.val;                            % @w:val absent -> True (default)
b.val = false;                         % -> <w:b w:val="0"/>
```

**Description**

The boolean run/paragraph property element (`w:b`, `w:i`, `w:caps`, …), the target
of 20 of the 28 font-block registry rows. `val` is an
`OptionalAttribute("w:val", ST_OnOff, default=True)` — a non-None default, which
drives the D-delta-1 tri-state (see [the tri-state section above](#ct_onoff--ct_string-tri-state-the-c3-leaves)):
absent → `True`; `False` → `w:val="0"`; `True` (== default) or `[]` (None) →
`@w:val` removed. `from_xml` maps `1`/`true`/`on` → `true`, `0`/`false`/`off`
→ `false`.

**Example**

```matlab
b = mat2doc.oxml.OxmlElement("w:b");
disp(b.val);      % 1  (true — @w:val absent; default True: a bare <w:b/> is bold-on)
b.val = false;    % serializes <w:b w:val="0"/>
disp(b.val);      % 0  (false)
b.val = true;     % == default -> @w:val removed -> a bare <w:b/> again
disp(b.val);      % 1  (reads back True from the default)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/shared.py::CT_OnOff`*

---

## `CT_String`

**Syntax**

```matlab
rStyle = mat2doc.oxml.shared.CT_String.new("w:rStyle", "Emphasis");
name   = rStyle.val;                         % "Emphasis"
```

**Description**

The single-required-string element backing `w:rStyle`. `val` is a
`RequiredAttribute("w:val", ST_String)`; the static `new(nsptagname, val)` builds
a loose element of the given tag and sets its `val`. The Python `cast(CT_String,
…)` is a type hint only (dropped).

**Example**

```matlab
rStyle = mat2doc.oxml.shared.CT_String.new("w:rStyle", "Emphasis");
disp(rStyle.val);   % "Emphasis"  (serializes <w:rStyle w:val="Emphasis"/>)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/shared.py::CT_String`*

---

## Deviation posture — 0 new D-numbers

P4-1a opened **no new deviation**. Every divergence maps to a ruling already
adopted and carried `mat2doc:`-namespaced: **D-001** (own OOXML
parser/serializer), **D-serializer-nsdecl** (loose-element `xmlns` emission — the
`_new_color` fragment relies on it), **D-STYPE-3/4** (via the P3-2
`ST_HexColor` / `ST_HpsMeasure` types), **D-delta-1/-2/-3** (the docx attribute
engine — D-delta-1's True-default removal is exercised here and passes), and
**D-005** (error-path message tokens). Gate-3 proved the whole s0020 `probe_diff`
**MATCH (exit 0)** with the M1 17/17 byte-neutrality gate and the H11
scrambled-order oracle both byte-exact; no L2 canonical-only result surfaced
anywhere. No ledger row was added.
