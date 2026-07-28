---
title: "mat2doc.oxml.text — the run (`CT_R`) + run-properties (`CT_RPr`) element tier"
---

# `mat2doc.oxml.text` — the run / run-properties element tier (`CT_R` + `CT_RPr` + children)

Ported from python-docx v1.2.0 `src/docx/oxml/text/font.py` (7 element classes,
the `<w:rPr>` run-properties tier) **and** `src/docx/oxml/text/run.py` (the
`<w:r>` run tier — 6 element classes + one helper), both in package
`+mat2doc/+oxml/+text/`, plus the two shared leaves `CT_OnOff` / `CT_String`
from `src/docx/oxml/shared.py` (package `+mat2doc/+oxml/+shared/`), and the
34 element `register_element_cls` rows they register
(`src/docx/oxml/__init__.py:198-225` font block + `:72-78` run block).

:::{note}
This page is built in two passes. **P4-1a** ported the run-**properties** tier
(`font.py` → `CT_RPr` + rPr children); **P4-1b** adds the **run** tier
(`run.py` → `CT_R` + run inner-content + `CT_Text`), which sits one level up:
`<w:r>` is the run element the `document.xml` body is built from
(`w:p`/`w:r`/`w:t`), and `CT_RPr` is its first child. Both are on the **M2
byte-critical path**. The run-tier material begins at
[The run tier — `CT_R`](#the-run-tier-ct_r) below.
:::

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

(the-run-tier-ct_r)=
## The run tier — `CT_R`

`run.py` ports the **run element** `<w:r>` and its inner-content leaves. `<w:r>`
is the container the `document.xml` body is actually made of — a paragraph
`<w:p>` holds runs `<w:r>`, and a run holds an optional `<w:rPr>` (the P4-1a
run-properties element) followed by text and break content
(`<w:t>`/`<w:tab>`/`<w:br>`/`<w:cr>`/…). This tier is the second half of the M2
byte-critical path.

**`CT_R` child descriptors.** `CT_R` declares its children as **one `ZeroOrOne`**
(`rPr`) plus **five `ZeroOrMore`** content descriptors (`br`, `cr`, `drawing`,
`t`, `tab`), all with `successors=()` (`run.py:33-38`). Two hand-routed members
survive the generic engine because they are defined explicitly on the class
(xmlchemy skips a name already defined): `add_t(text)` (byte-critical, below) and
`add_drawing(...)`.

**H11 — `rPr` first, content in insertion order.** OOXML requires `<w:rPr>` to be
the **first** child of a run. `run.py` does not express this with a successor
list (`rPr` has `successors=()`); it **overrides the inserter**
(`run.py:145-147`: `self.insert(0, rPr)`). The port carries this as the
`insert_rPr_` override — `obj.insert(1, rPr)` (the [H1 base shift](#id-h11-successor-ordering)
applied once) — so `get_or_add_rPr` always forces `rPr` to index 0, while every
content descriptor (`successors=()`) simply **appends**. Adding content and *then*
`get_or_add_rPr()` still yields `[rPr, t, …]`; the run inner-content keeps the
order it was added.

### `add_t` and `xml:space="preserve"` — the H16 whitespace hazard

`CT_R.add_t(text)` (`run.py:40-45`) adds a `<w:t>` whose char data is `text`,
then sets `@xml:space="preserve"` **iff the text has leading or trailing
whitespace** — Python's exact test is `len(text.strip()) < len(text)`. Without
`xml:space="preserve"`, Word collapses those boundary spaces on load, so this
attribute is **byte-load-bearing**: `add_t(" hello ")` must serialize
`<w:t xml:space="preserve"> hello </w:t>`, while `add_t("hello")` must serialize
`<w:t>hello</w:t>` with no attribute. The reserved `xml` prefix is **never**
`xmlns`-declared, so no `xmlns:xml` appears anywhere in the output.

:::{warning}
**H16 — Python `str.strip()` strips a *wider* whitespace set than MATLAB
`strip()`.** A naive port of the boundary test as
`strlength(strip(text)) < strlength(text)` is a **real byte divergence**: MATLAB
`strip()` does **not** remove U+0085 (NEL), U+00A0 (NBSP), U+2007 (figure space),
U+202F (narrow NBSP), U+2028 (line separator) and several others, whereas
CPython `str.strip()` removes all of them. So `add_t(char(160) + "x")` (an NBSP
lead — not exotic; it appears in pasted text) would omit `xml:space="preserve"`
while python-docx emits it. This was caught by the Gate-2 audit (probe
`xmlspace_10/11/13`) and fixed: `CT_R` carries a `Constant PY_STRIP_WS` holding
the **exact 29-code-point CPython strip set** (enumerated mechanically over all
of Unicode — `09-0D, 1C-1F, 20, 85, A0, 1680, 2000-200A, 2028, 2029, 202F, 205F,
3000`; all BMP, so a UTF-16 code-unit membership test is surrogate-safe), and the
condition is rewritten as a first/last-code-unit membership test — exactly
equivalent to Python's length-decrease test. Any future port whose **output**
depends on a `.strip()` boundary must reuse this set, never native `strip()`.
Gate-3 froze the five non-ASCII boundary cases (NBSP/NEL/figure-space/
narrow-NBSP/line-sep → `preserve`) as byte-pinned L1 regression fixtures.
:::

### `_RunContentAppender` — text → run content

`CT_R.text = "…"` (the setter) and the `_RunContentAppender` helper
(`RunContentAppender_`, module-private, `run.py:262-307`) translate a plain
string into run-content elements via a small finite-state machine that buffers
regular characters and flushes them as `<w:t>` at each break boundary:

| input character | action | element |
|---|---|---|
| `"\t"` (0x09) | flush buffer, then `add_tab()` | `<w:tab/>` |
| `"\r"` or `"\n"` (0x0D/0x0A) | flush buffer, then `add_br()` | `<w:br/>` |
| any other | accumulate into the buffer | (pending `<w:t>`) |

So `"a\tb\nc"` → `[w:t "a", w:tab, w:t "b", w:br, w:t "c"]` and `"x\ry"` →
`[w:t "x", w:br, w:t "y"]`. Note `\r`/`\n` map to **`w:br`** (`add_br`), not
`w:cr` — the class docstring's mention of `w:cr` is an upstream inaccuracy; the
**code** (`run.py:294-299`) calls `add_br`. The buffer flush uses
`if strlength(text) > 0` (Python `if text:`, H4) so an empty run adds no `<w:t>`.
H2: the FSM iterates the UTF-16 code units of `char(text)`; astral characters
split into surrogate halves that never match `\t\r\n` and are rejoined by
`string(bfr_)` on flush — byte-equivalent to Python's per-code-point loop.

### `CT_Text` and the run-content `str_` (H3)

`CT_Text` (`<w:t>`) is a **text-bearing** element: its content is the lxml
element `.text` (char data), not an attribute. It does **not** shadow `.text` —
it only adds `__str__` → `str_()`, which returns the element text or `""` when
empty. Python `return self.text or ""` maps **both** the `None` case (lxml
`.text` is `None` when empty) and the falsy-`""` case to `""` — `str_` **never**
returns `None`/`[]` (H3/H4). The break/positional-tab leaves supply the same
`str_` accessor so a run's whole inner-content text can be read uniformly:
`CT_Br` → `"\n"` for a text-wrapping break (the absent-default) and `""` for a
page/column break; `CT_Cr` → `"\n"`; `CT_NoBreakHyphen` → `"-"`; `CT_PTab` →
`"\t"`. (H2: the `"\n"`/`"\t"` are produced via `newline`/`char(9)`, since a
MATLAB double-quoted `"\n"` is the two literal characters backslash-n.)

:::{note}
`CT_R.text` (the getter) joins `str_()` over `w:br|w:cr|w:noBreakHyphen|w:ptab|
w:t|w:tab`. `w:tab` maps to `CT_TabStop`, which is **not ported until the parfmt
WP** — so a `.text` read over a run that *contains* a `<w:tab>` errors until then
(the parfmt WP must add `CT_TabStop.str_() → "\t"`). This is off the M2 **write**
path: `add_tab()` still creates a correct `<w:tab/>`, and the byte-critical
`add_t` / `_RunContentAppender` write path is fully live.
:::

### The `XmlParser` engine fix (registered-class URI forwarding)

Registering `w:r → CT_R` exposed — and P4-1b fixed — a **latent P1-2 parser
defect**. `XmlParser.parseElement_` builds an element two ways: the generic
fallback `XmlElement(name, ownDecls, uri)` **forwards** the parser-resolved
namespace URI, but the registered-class branch called
`feval(cls, name, ownDecls)` **without** it and set the URI *after*
construction. When an element's prefix is **ancestor-declared and not in the
fixed nsmap** (e.g. the `a03` deviation fixture, where `<w:p>` binds a second
prefix `q` to the `w` URI and the child is `<q:r>` = `{wURI}r`), the `XmlElement`
constructor re-resolved the prefix itself and threw `mat2doc:KeyError 'q'` before
the post-construction setter could run. Pre-P4-1b, `{wURI}r` was unregistered and
took the working generic path; registering `CT_R` moved it onto the broken
registered path. The one-line fix makes the registered branch forward the URI
exactly as the generic fallback does — `elm = feval(cls, name, ownDecls, uri)`
(the `CT_*` constructors are transparent pass-throughs, so `uri` reaches
`XmlElement`'s early-return) — and drops the now-redundant setter. It **restores
the parser's intended behaviour** (the in-scope resolved URI is authoritative for
every element, registered or not); the `a03` D-nsprefix-rewrite deviation bytes
are byte-identical before and after (proven by direct pre/post comparison), and
the full regression returned to 550/550.

---

## `CT_R`

**Syntax**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");        % a CT_R (registered)
r.get_or_add_rPr();                         % <w:rPr/> forced to the FRONT (H11)
t = r.add_t(" hi ");                        % <w:t xml:space="preserve"> hi </w:t>
r.text = "a" + char(9) + "b";               % <w:t>a</w:t><w:tab/><w:t>b</w:t>
```

:::{note}
A tab / newline in the run text must be a **real** control character — in MATLAB
a double-quoted `"a\tb"` is the four literal characters `a \ t b`, not a tab.
Use `char(9)` (tab), `char(10)` (LF), or `char(13)` (CR), as above and in the
examples below. (The header-comment shorthand `"a\tb"` denotes intent only.)
:::

**Description**

The `<w:r>` **run** element — the container `document.xml` body content is built
from (`w:p`/`w:r`/`w:t`). One `ZeroOrOne` `rPr` descriptor plus five `ZeroOrMore`
content descriptors (`br`/`cr`/`drawing`/`t`/`tab`), all `successors=()`. `rPr`
is forced to the front by the `insert_rPr_` override (H11); content descriptors
append in insertion order. `add_t(text)` sets `@xml:space="preserve"` when `text`
has leading/trailing whitespace, using the `PY_STRIP_WS` 29-code-point set
([H16 above](#the-run-tier-ct_r)), **not** MATLAB `strip()`. The `text` getter
joins `str_()` over the run inner-content; the `text` setter drives
`_RunContentAppender`. `CT_R.text` shadows the lxml `.text` attribute and is
ported by overriding the protected `getText_`/`setText_` (the serializer reads
`text_raw_`, bypassing the shadow, so the run's own char data serializes
unchanged — byte-verified in the M1 sweep).

**Example**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");
r.add_t("a"); r.get_or_add_rPr();               % rPr re-sorts BEFORE the text
order = arrayfun(@(e) string(e.nsptag_str), r.xpath("./*"));
disp(order);                                      % "w:rPr"  "w:t"   (H11)
t = r.add_t(" hello ");                           % boundary space -> xml:space
disp(t.get(mat2doc.oxml.qn("xml:space")));        % "preserve"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/run.py::CT_R`*

---

## `CT_Text`

**Syntax**

```matlab
t = mat2doc.oxml.OxmlElement("w:t");
t.text = "hello";
s = t.str_();                                     % "hello"
```

**Description**

The `<w:t>` text element (a sequence of characters within a run). Content is the
element char data, not an attribute; `CT_Text` does not shadow `.text`, adding
only `__str__` → `str_()`. `str_()` returns the text, or `""` when the element
has no content — it **never** returns `None`/`[]` (H3: lxml `.text` is `None`
when empty; H4: Python `self.text or ""`). Whitespace preservation is driven by
`CT_R.add_t` setting `@xml:space="preserve"`; `CT_Text` itself treats
`xml:space` as a plain stored attribute.

**Example**

```matlab
t = mat2doc.oxml.OxmlElement("w:t");
t.text = "hi";      disp(t.str_());               % "hi"
e = mat2doc.oxml.OxmlElement("w:t");
disp(e.str_());                                   % ""  (never [] / None)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/run.py::CT_Text`*

---

## `CT_Br` / `CT_Cr` / `CT_NoBreakHyphen` / `CT_PTab`

**Syntax**

```matlab
br = mat2doc.oxml.OxmlElement("w:br");   br.str_();   % "\n" (LF; type defaults to textWrapping)
br.type = "page";                        br.str_();   % ""
```

**Description**

The four break / positional-tab run-content leaves, each supplying a `str_()`
"text equivalent" so a run's inner-content text reads uniformly:

- **`CT_Br`** (`<w:br>`) — a line/page/column break. Descriptors
  `type = OptionalAttribute("w:type", ST_BrType, default="textWrapping")` and
  `clear = OptionalAttribute("w:clear", ST_BrClear)`. `type` has a **non-None
  default**, so (D-delta-1 tri-state) an absent `@w:type` reads `"textWrapping"`
  and setting the value to `None`/`[]` **or** to the default removes the
  attribute. `str_()` → `"\n"` for a text-wrapping break (incl. the
  absent-default), `""` for page/column.
- **`CT_Cr`** (`<w:cr>`) — a soft carriage return; `str_()` → `"\n"`. The
  complex-type name is distinguished only to give `w:cr` its `"\n"` behaviour
  (the schema maps `w:cr` to `CT_Empty`).
- **`CT_NoBreakHyphen`** (`<w:noBreakHyphen>`) — a non-breaking hyphen; `str_()`
  → `"-"`.
- **`CT_PTab`** (`<w:ptab>`) — an absolute-position tab; `str_()` → `"\t"`.

**Example**

```matlab
cr = mat2doc.oxml.OxmlElement("w:cr");            disp(cr.str_());   % "\n"
nb = mat2doc.oxml.OxmlElement("w:noBreakHyphen"); disp(nb.str_());   % "-"
pt = mat2doc.oxml.OxmlElement("w:ptab");          disp(pt.str_());   % "\t"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/run.py::CT_Br` / `CT_Cr` /
`CT_NoBreakHyphen` / `CT_PTab`*

---

## `RunContentAppender_`

**Syntax**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");
mat2doc.oxml.text.RunContentAppender_.append_to_run_from_text(r, "a" + char(9) + "b");
% r now has <w:t>a</w:t><w:tab/><w:t>b</w:t>
```

**Description**

The module-private `_RunContentAppender` (leading underscore rotated to the end,
design.md §2) — a `handle` FSM that translates a string into run-content
elements in a `<w:r>`. Contiguous regular characters accumulate in a single
`<w:t>`; a `"\t"` flushes and appends `<w:tab>` (`add_tab`); a `"\r"`/`"\n"`
flushes and appends `<w:br>` (`add_br`); end of text flushes any pending `<w:t>`
(only when non-empty, H4). Reference semantics: `add_char`/`flush` mutate the
shared buffer and the shared `CT_R`. The private `_r`/`_bfr` rotate to
`r_`/`bfr_`; the public classmethod entry `append_to_run_from_text(r, text)`
ports to a static method.

**Example**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");
r.text = "x" + char(13) + "y";                     % CR -> the setter drives the appender
names = arrayfun(@(e) string(e.nsptag_str), r.xpath("./*"));
disp(names);                                         % "w:t"  "w:br"  "w:t"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/run.py::_RunContentAppender`*

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

**P4-1b (the run tier) likewise opened 0 new D-numbers.** The two corrections it
made are **fixes to match python-docx**, not deviations: the H16 `PY_STRIP_WS`
whitespace-set replication (so `add_t`'s `xml:space` decision is byte-identical to
CPython) and the `XmlParser` registered-class URI-forwarding fix (a latent P1-2
engine bug — the signed `a03` **D-nsprefix-rewrite** deviation bytes are proven
byte-identical before and after). Gate-3 was byte-perfect on the whole run
surface (`s0021` `probe_diff` MATCH exit 0 — the xml:space battery incl. the five
non-ASCII AUD-1 cases, `_RunContentAppender`, H11 rPr-first, `CT_Text`,
run-content `str_`, `CT_R.text`, comment helpers), the M1 17/17 byte-neutrality
gate held (`document.xml` 1548 B & `styles.xml` 349458 B L1), and the full
regression is back to 550/550. The standing adopt-only umbrella (D-001 /
D-serializer-nsdecl / D-delta-1 / D-STYPE-\* / D-zip-time) covers everything; no
ledger row was added.
