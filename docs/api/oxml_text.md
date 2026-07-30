---
title: "mat2doc.oxml.text — the complete text-oxml element layer (CT_P · CT_R · CT_RPr · CT_PPr · CT_Hyperlink · CT_LastRenderedPageBreak + children)"
---

# `mat2doc.oxml.text` — the complete text-oxml element layer (`CT_P` + `CT_R` + `CT_PPr`/`CT_RPr` + `CT_Hyperlink` + `CT_LastRenderedPageBreak` + children)

Ported from python-docx v1.2.0 `src/docx/oxml/text/font.py` (7 element classes,
the `<w:rPr>` run-properties tier) **and** `src/docx/oxml/text/run.py` (the
`<w:r>` run tier — 6 element classes + one helper), both in package
`+mat2doc/+oxml/+text/`, plus the two shared leaves `CT_OnOff` / `CT_String`
from `src/docx/oxml/shared.py` (package `+mat2doc/+oxml/+shared/`), and the
34 element `register_element_cls` rows they register
(`src/docx/oxml/__init__.py:198-225` font block + `:72-78` run block).

:::{note}
This page is built in four passes. **P4-1a** ported the run-**properties** tier
(`font.py` → `CT_RPr` + rPr children); **P4-1b** added the **run** tier
(`run.py` → `CT_R` + run inner-content + `CT_Text`), which sits one level up:
`<w:r>` is the run element the `document.xml` body is built from
(`w:p`/`w:r`/`w:t`), and `CT_RPr` is its first child. **P4-2** adds the
**paragraph-properties + paragraph** tier (`parfmt.py` → `CT_PPr` + its
para-property children `CT_Ind`/`CT_Jc`/`CT_Spacing`/`CT_TabStop`/`CT_TabStops`,
and `paragraph.py` → `CT_P`), which sits one level up again: `<w:p>` is the
paragraph the body is a sequence of, and `<w:pPr>` is its first child and the
`add_heading` write target. **P4-3** adds the **hyperlink + rendered-page-break**
tier (`hyperlink.py` → `CT_Hyperlink`, and `pagebreak.py` →
`CT_LastRenderedPageBreak`) — the two remaining run-level inner-content classes:
`<w:hyperlink>` is the second thing (besides `<w:r>`) a paragraph can hold, and
`<w:lastRenderedPageBreak>` is the renderer-inserted break inside a run that
`Paragraph.rendered_page_breaks` splits paragraphs on. **P4-3 completes the
text-oxml layer** (font · run · run-properties · paragraph · paragraph-properties
· hyperlink · page-break). The first three tiers are on the **M2 byte-critical
path**; the hyperlink/page-break tier is a **pure lookup addition** (the shipped
template carries zero of either). The run-tier material begins at
[The run tier — `CT_R`](#the-run-tier-ct_r); the paragraph-tier material begins at
[The paragraph tier — `CT_PPr` and `CT_P`](#the-paragraph-tier-ct_ppr-and-ct_p);
the hyperlink/page-break material begins at
[The hyperlink + page-break tier](#the-hyperlink-and-pagebreak-tier).
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

(the-paragraph-tier-ct_ppr-and-ct_p)=
## The paragraph tier — `CT_PPr` and `CT_P`

`parfmt.py` ports the **paragraph-properties element** `<w:pPr>` (`CT_PPr`) plus
its five para-property children, and `paragraph.py` ports the **paragraph
element** `<w:p>` (`CT_P`). `<w:p>` is the block the `document.xml` body is a
sequence of; `<w:pPr>` is its **first** child and carries the `<w:pStyle>` that
`add_heading` writes. This is the **third and highest** slice of the M2
byte-critical path: `add_paragraph`/`add_heading` write `w:p` → `w:pPr` →
`w:pStyle` into `document.xml`. Registering the 12 parfmt/paragraph tags is
**byte-neutral** — the M1 17/17 sweep holds and the **204 `<w:pPr>` blocks** in
the shipped `styles.xml` (349 458 B) now parse to `CT_PPr` with identical bytes
(Gate-3 `pkgcompare` L0 PASS + 16 XML L1 + 1 bin; `document.xml` 1 548 B &
`styles.xml` 349 458 B L1; the 204-`w:pPr` `styles.xml` round-trips
parse→serialize L1).

(id-ct_ppr-h11)=
### `CT_PPr` — the 36-tag `_tag_seq`, pStyle-first (the `add_heading` crux)

`CT_PPr` is the paragraph analogue of `CT_RPr`: **12 `ZeroOrOne` descriptors**
over a single **36-entry `_tag_seq`** (`parfmt.py:64-101`, ported verbatim as the
Constant `TAG_SEQ`). OOXML mandates that a `<w:pPr>`'s children appear in schema
order — `w:pStyle`, `w:keepNext`, …, `w:spacing`, `w:ind`, `w:jc`, …, `w:sectPr` —
regardless of the order the API creates them; get it wrong and Word shows a
**repair prompt = FAIL**. The re-sort mechanism is exactly the
[H11 successor-ordering mechanism](#id-h11-successor-ordering) documented for
`CT_RPr`: each descriptor carries the schema slice `_tag_seq[N:]` (0-based
Python) → `TAG_SEQ(N+1:end)` (1-based MATLAB, the H1 base shift applied once at
the slice declaration), and `insertChildInSequence` lands a new child *before the
first schema-later sibling already present*, else appends.

The **12 descriptor slices** (own 1-based index → successor slice):

| prop | own idx | Python (parfmt.py) | MATLAB slice |
|---|---|---|---|
| `pStyle` | 1 | `_tag_seq[1:]` | `TAG_SEQ(2:end)` |
| `keepNext` | 2 | `_tag_seq[2:]` | `TAG_SEQ(3:end)` |
| `keepLines` | 3 | `_tag_seq[3:]` | `TAG_SEQ(4:end)` |
| `pageBreakBefore` | 4 | `_tag_seq[4:]` | `TAG_SEQ(5:end)` |
| `widowControl` | 6 | `_tag_seq[6:]` | `TAG_SEQ(7:end)` (framePr@5 skipped) |
| `numPr` | 7 | `_tag_seq[7:]` | `TAG_SEQ(8:end)` |
| `tabs` | 11 | `_tag_seq[11:]` | `TAG_SEQ(12:end)` |
| `spacing` | 22 | `_tag_seq[22:]` | `TAG_SEQ(23:end)` |
| `ind` | 23 | `_tag_seq[23:]` | `TAG_SEQ(24:end)` |
| `jc` | 27 | `_tag_seq[27:]` | `TAG_SEQ(28:end)` |
| `outlineLvl` | 31 | `_tag_seq[31:]` | `TAG_SEQ(32:end)` |
| `sectPr` | 35 | `_tag_seq[35:]` | `TAG_SEQ(36:end)` |

Like `CT_RPr`, the slices are **non-contiguous** — `w:framePr` (idx 5),
`w:suppressLineNumbers` (8), `w:pBdr` (9), `w:shd` (10) and a dozen more intervening
`_tag_seq` tags have **no descriptor on `CT_PPr`**, so the slices jump across
them; the intervening tags are **kept in the successor lists, not collapsed**, so
a newly-added child still lands correctly relative to a parsed non-descriptor
sibling. **`pStyle` (`successors=_tag_seq[1:]`) sorts before *everything*** — this
is the `add_heading` case: whatever order the API builds a `<w:pPr>`, the
`<w:pStyle>` re-sorts to the front. Gate-2/Gate-3 built four scrambled
`get_or_add_*` orderings (incl. `pStyle`-**last** and a full 12-child scramble) in
both python-docx 1.2.0 and Mat2Doc; every ordering converged to the canonical
`w:pStyle, …, w:spacing, w:ind, w:jc, …, w:sectPr` and the serialized bytes were
**byte-identical PY==ML**.

All 12 descriptors use the **generic** `BaseOxmlElement` engine (no `_new_x` /
`_insert_x` override on `CT_PPr`, unlike `CT_RPr._new_color` or `CT_R._insert_rPr`).
Each generates the docx family `get.x` / `get_or_add_x` / `new_x_` / `insert_x_` /
`add_x_` / `remove_x_`. The **13 `@property` helpers** (`parfmt.py:122-339`) are
all ported live: `first_line_indent`, `ind_left`, `ind_right`, `jc_val`,
`keepLines_val`, `keepNext_val`, `pageBreakBefore_val`, `spacing_after`,
`spacing_before`, `spacing_line`, `spacing_lineRule`, `style`, `widowControl_val`.
Note `spacing_lineRule` maps a **present `@w:line` with absent `@w:lineRule`** to
`WD_LINE_SPACING.MULTIPLE` (`parfmt.py:296-298`), and `first_line_indent` uses the
signed/unsigned split of `CT_Ind` (a negative value → `w:hanging`, positive →
`w:firstLine`, `[]` clears both).

:::{note}
Three `_tag_seq` tags have descriptors on `CT_PPr` whose **child class is owned by
a later WP** and is therefore left generic here (byte-neutral): `w:numPr`→
`CT_NumPr` (P8), `w:sectPr`→`CT_SectPr` (P5), and `w:outlineLvl`→`CT_DecimalNumber`
(`oxml/shared.py`, a numbering/shared WP). The `w:outlineLvl` **registry row is
deferred** with them — `CT_DecimalNumber` is not yet ported, so its tag resolves
to a generic `XmlElement`. This is byte-neutral (round-trip is class-independent —
proven by the 204-`w:pPr` `styles.xml`, which contains `outlineLvl`-bearing `pPr`,
round-tripping L1) and behavior-neutral (no ported `CT_PPr` accessor reads `.val`
on `numPr`/`outlineLvl`/`sectPr`). The future `CT_DecimalNumber` WP must add the
`w:outlineLvl` row.
:::

### The para-property children — `CT_Ind` / `CT_Jc` / `CT_Spacing`

The three plain attribute-holder children of `CT_PPr`, each delegating to a P3-2
simple type or a P3-3 enum:

- **`CT_Ind`** (`w:ind`) — four `OptionalAttribute`s, **all default `None`** (`[]`
  when absent): `left` / `right` = `ST_SignedTwipsMeasure` (**signed**),
  `firstLine` / `hanging` = `ST_TwipsMeasure` (**unsigned**). The signed-vs-unsigned
  split is per-attribute and verified against `parfmt.py:32-43` — `left`/`right`
  serialize a leading `-` for negative twips, `firstLine`/`hanging` never do. Each
  reads back a `Length`; setting `[]` removes the attribute (H3).
- **`CT_Jc`** (`w:jc`) — `val` = `RequiredAttribute("w:val", WD_ALIGN_PARAGRAPH)`
  (a.k.a. `WD_PARAGRAPH_ALIGNMENT`), the enum named by its fully-qualified class so
  `resolveTypeCls_` dispatches to `+enum`. A missing `@w:val` raises
  `mat2doc:InvalidXmlError`. `CENTER` ↔ `w:val="center"`.
- **`CT_Spacing`** (`w:spacing`) — four `OptionalAttribute`s, all default `None`:
  `after` / `before` = `ST_TwipsMeasure` (unsigned), `line` = `ST_SignedTwipsMeasure`
  (signed), `lineRule` = `WD_LINE_SPACING`. `EXACTLY` ↔ `w:lineRule="exact"`; a
  present `@w:line` with an absent `@w:lineRule` reads back as `MULTIPLE` (the
  `CT_PPr.spacing_lineRule` mapping above).

### `CT_TabStop` / `CT_TabStops` — and the `str_`→`"\t"` carry-forward (P4-1b gap CLOSED)

**`CT_TabStop`** (`<w:tab>`) is **overloaded**: the same `w:tab` tag serves both a
real **tab stop** (inside `<w:tabs>`) and a **tab character** within a run. Its
tab-stop usage uses three attributes — `val` = `RequiredAttribute` of
`WD_TAB_ALIGNMENT`, `pos` = `RequiredAttribute` of `ST_SignedTwipsMeasure`, and
`leader` = `OptionalAttribute` of `WD_TAB_LEADER` with a **non-None default**
`WD_TAB_LEADER.SPACES`. That non-None default drives the D-delta-1 tri-state
(`LEADER_DEFAULT` holds the actual `SPACES` member, not `[]`): assigning `[]`
(None) **or** the `SPACES` member removes `@w:leader`, and reading `@w:leader` when
absent returns `SPACES`; `leader = DOTS` writes `w:leader="dot"`.

Its **run usage** needs only `str_()` → `"\t"` (`char(9)`, H2 — never the literal
two-char `"\t"`). This **closes the P4-1b carry-forward gap**: `CT_R.text` (and
`CT_P.text`) joins `str_()` over `w:br|w:cr|w:noBreakHyphen|w:ptab|w:t|w:tab`; at
P4-1b `w:tab` was unregistered (generic `XmlElement`, no `str_`), so a `.text` read
over a run *containing* a `<w:tab>` **errored**. Registering `w:tab`→`CT_TabStop`
(with this `str_`) means a run with children `[t, tab, t]` now reads back
`"a\tb"` — the exact behaviour the P4-1b page's note flagged as pending.

**`CT_TabStops`** (`<w:tabs>`) is the sorted container: `tab` =
`OneOrMore("w:tab", successors=())` → `tab_lst` / `new_tab_` / `insert_tab_` /
`add_tab_` / `add_tab` (public), `successors=()` → append.
`insert_tab_in_order(pos, align, leader)` (`parfmt.py:383-392`) creates a `w:tab`,
sets `pos`/`val`/`leader`, then inserts it **before the first existing tab whose
`pos` is greater** (keeping the sequence sorted by position), else appends — so
inserting 720, 240, 480 yields serialized positions 240, 480, 720.

### `CT_P` — the paragraph, pPr-first

`CT_P` (`<w:p>`) declares one `ZeroOrOne` `pPr` plus two `ZeroOrMore` content
descriptors (`hyperlink`, `r`), all with `successors=()` (`paragraph.py:29-31`).
OOXML requires `<w:pPr>` to be the **first** child of a paragraph. Like
`CT_R._insert_rPr`, `paragraph.py` expresses this **not** with a successor list but
by overriding the inserter (`paragraph.py:104-106`: `self.insert(0, pPr)`); the
port carries this as the `insert_pPr_` override — `obj.insert(1, pPr)` (the
[H1 base shift](#id-h11-successor-ordering) applied once) — so `get_or_add_pPr`
always forces `pPr` to index 0, while `r`/`hyperlink` (`successors=()`) simply
**append** in insertion order. Appending a `<w:r>` and *then* setting `p.style`
(which adds `pPr`) still yields child order `[pPr, r]`.

`CT_P.text` (`paragraph.py:95-102`) **shadows** the lxml `.text` attribute — it is
the paragraph's concatenated inner-content text (`"".join(e.text for e in
w:r|w:hyperlink)`), not element char data — ported by overriding the protected
`getText_` (D10, the same shape as `CT_R.text`); it is **read-only** in Python, so
`setText_` is overridden to raise `mat2doc:AttributeError`. `CT_P` also ports
`add_p_before`, `alignment` (via `pPr.jc_val`), `clear_content`
(`./*[not(self::w:pPr)]`, keeps only `pPr`), `inner_content_elements`
(`./w:r | ./w:hyperlink`), `lastRenderedPageBreaks`, `set_sectPr`, and `style`
(via `pPr.style`).

:::{note}
`CT_P.text` over a paragraph **containing a `<w:hyperlink>`** does not yet
reproduce python-docx: `w:hyperlink`→`CT_Hyperlink` is owned by the **P4-3**
hyperlink WP and resolves to a generic `XmlElement` until then (its `.text` is the
element's own char data, not `CT_Hyperlink`'s concatenated run text). This is the
same dependency-order shape as the `w:tab` gap this WP closed, and it is **off the
M2 write path** — the default body is one empty `w:p`, and
`add_paragraph`/`add_heading` write `w:pPr` + `w:r` only. `inner_content_elements`
/ `lastRenderedPageBreaks` likewise return generic elements for `w:hyperlink` until
P4-3 (byte/structure identical; only the element class differs). P4-3 must register
`w:hyperlink` and re-probe `CT_P.text` over a hyperlink-bearing paragraph.
:::

---

## `CT_PPr`

**Syntax**

```matlab
pPr = mat2doc.oxml.OxmlElement("w:pPr");            % a CT_PPr (registered)
pPr.style = "Heading1";                             % <w:pStyle w:val="Heading1"/>
pPr.get_or_add_jc().val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
val = pPr.spacing_after;                            % @property helpers (parfmt.py 122-339)
```

**Description**

The `<w:pPr>` paragraph-properties container — the M2 `add_heading` write target.
Its 12 `ZeroOrOne` child descriptors ride the non-contiguous H11 successor slices
of the 36-entry `TAG_SEQ` (see [the H11 table above](#id-ct_ppr-h11)), so scrambled
`get_or_add_*` adds always reserialize in schema order and `pStyle` sorts first.
Each descriptor generates the docx family; all 12 use the generic engine (no
creator/inserter override). The 13 `@property` helpers are ported live, incl.
`spacing_lineRule`'s line-present/lineRule-absent → `MULTIPLE` mapping and
`first_line_indent`'s signed `w:hanging`/`w:firstLine` split. `w:numPr`/`w:sectPr`/
`w:outlineLvl` children are left generic (owned by P8/P5/a shared WP) — byte-neutral.

**Example**

```matlab
pPr = mat2doc.oxml.OxmlElement("w:pPr");
pPr.get_or_add_jc().val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER; % jc added FIRST
pPr.style = "Heading1";                          % pStyle added AFTER jc
order = arrayfun(@(e) string(e.nsptag_str), pPr.xpath("./*"));
disp(order);   % "w:pStyle"  "w:jc"  -- pStyle re-sorted BEFORE jc (H11 schema order)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/parfmt.py::CT_PPr`*

---

## `CT_Ind` / `CT_Jc` / `CT_Spacing`

**Syntax**

```matlab
ind = mat2doc.oxml.OxmlElement("w:ind");   ind.left = mat2doc.shared.Twips(720);
jc  = mat2doc.oxml.OxmlElement("w:jc");     jc.val   = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
sp  = mat2doc.oxml.OxmlElement("w:spacing"); sp.after = mat2doc.shared.Twips(120);
```

**Description**

The three plain attribute-holder children of `CT_PPr`:

- **`CT_Ind`** (`w:ind`) — `left` / `right` = `OptionalAttribute` of
  `ST_SignedTwipsMeasure` (**signed**); `firstLine` / `hanging` = `OptionalAttribute`
  of `ST_TwipsMeasure` (**unsigned**); all default `None` (`[]`). Reads back a
  `Length`; setting `[]` removes the attribute (H3).
- **`CT_Jc`** (`w:jc`) — `val` = `RequiredAttribute` of `WD_ALIGN_PARAGRAPH`
  (`CENTER` ↔ `"center"`); absent `@w:val` → `mat2doc:InvalidXmlError`.
- **`CT_Spacing`** (`w:spacing`) — `after` / `before` = `ST_TwipsMeasure`, `line`
  = `ST_SignedTwipsMeasure`, `lineRule` = `WD_LINE_SPACING`; all default `None`.
  `EXACTLY` ↔ `"exact"`; a present `@w:line` with absent `@w:lineRule` reads
  `MULTIPLE`.

**Example**

```matlab
ind = mat2doc.oxml.OxmlElement("w:ind");
ind.left  = mat2doc.shared.Twips(720);        % <w:ind w:left="720" ...
ind.right = mat2doc.shared.Twips(-60);         % ... w:right="-60"/>  (signed)
disp(double(ind.left));                         % 457200  (a Length/Emu)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/parfmt.py::CT_Ind` / `CT_Jc` /
`CT_Spacing`*

---

## `CT_TabStop` / `CT_TabStops`

**Syntax**

```matlab
tab = mat2doc.oxml.OxmlElement("w:tab");
tab.val = mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT;
tab.pos = mat2doc.shared.Twips(720);           % <w:tab w:val="left" w:pos="720"/>
s = tab.str_();                                 % "\t"  (run tab-character usage)

tabs = mat2doc.oxml.OxmlElement("w:tabs");
tabs.insert_tab_in_order(mat2doc.shared.Twips(720), ...
    mat2doc.enum.text.WD_TAB_ALIGNMENT.LEFT, mat2doc.enum.text.WD_TAB_LEADER.SPACES);
```

**Description**

**`CT_TabStop`** (`<w:tab>`) is overloaded — a tab stop (in `<w:tabs>`) *and* a
run tab-character. Tab-stop attrs: `val` (`RequiredAttribute` of
`WD_TAB_ALIGNMENT`), `pos` (`RequiredAttribute` of `ST_SignedTwipsMeasure`),
`leader` (`OptionalAttribute` of `WD_TAB_LEADER`, **non-None default `SPACES`** —
assigning `[]` or `SPACES` removes `@w:leader`, absent reads back `SPACES`,
`DOTS`→`"dot"`). Run usage: `str_()` → `"\t"` (`char(9)`), which **closes the
P4-1b carry-forward gap** so `CT_R.text`/`CT_P.text` over a run holding a `<w:tab>`
returns the tab character. **`CT_TabStops`** (`<w:tabs>`) holds `tab` =
`OneOrMore(successors=())`; `insert_tab_in_order(pos, align, leader)` keeps the
sequence sorted by `pos`.

**Example**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");
r.text = "a" + char(9) + "b";                   % <w:t>a</w:t><w:tab/><w:t>b</w:t>
disp(r.text);                                    % "a	b"  (tab char reproduced -- P4-1b gap closed)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/parfmt.py::CT_TabStop` /
`CT_TabStops`*

---

## `CT_P`

**Syntax**

```matlab
p = mat2doc.oxml.OxmlElement("w:p");            % a CT_P (registered)
p.add_r();                                       % append a <w:r/>
p.style = "Heading1";                            % adds <w:pPr><w:pStyle .../></w:pPr>, FORCED to front
txt = p.text;                                    % concatenated inner-content text (shadow, D10)
```

**Description**

The `<w:p>` **paragraph** element — the block the `document.xml` body is a
sequence of, and the `add_paragraph`/`add_heading` target. One `ZeroOrOne` `pPr`
descriptor (forced to the **front** by the `insert_pPr_` override, H11) plus two
`ZeroOrMore` content descriptors (`r`, `hyperlink`, `successors=()` → append).
`CT_P.text` shadows the lxml `.text` (a computed inner-content string over
`w:r|w:hyperlink`, ported via a `getText_` override, D10); it is **read-only** —
`setText_` raises `mat2doc:AttributeError`. Also ports `add_p_before`,
`alignment` (via `pPr.jc_val`), `clear_content` (keeps only `pPr`),
`inner_content_elements`, `lastRenderedPageBreaks`, `set_sectPr`, and `style`.
`w:hyperlink` children stay generic until the **P4-3** hyperlink WP.

**Example**

```matlab
p = mat2doc.oxml.OxmlElement("w:p");
p.add_r();                                        % content added FIRST
p.style = "Heading1";                             % pPr added AFTER, forced to index 0
order = arrayfun(@(e) string(e.nsptag_str), p.xpath("./*"));
disp(order);                                       % "w:pPr"  "w:r"   (pPr-first, H11)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/paragraph.py::CT_P`*

---

(the-hyperlink-and-pagebreak-tier)=
## The hyperlink + page-break tier — `CT_Hyperlink` and `CT_LastRenderedPageBreak`

`hyperlink.py` ports the **hyperlink element** `<w:hyperlink>` (`CT_Hyperlink`) and
`pagebreak.py` ports the **rendered page-break element**
`<w:lastRenderedPageBreak>` (`CT_LastRenderedPageBreak`) — the two remaining
run-level inner-content classes and the **last oxml work package of P4**. A
paragraph's inner content is `w:r | w:hyperlink` (a hyperlink is a run container
that carries its own display text and address); a rendered page-break is the
empty element a Word *renderer* drops inside a run when it runs out of page, and
`Paragraph.rendered_page_breaks` splits a paragraph on it.

Unlike the font/run/paragraph tiers, this tier is **not on the M2 byte-critical
write path**: the shipped `default.docx` carries **zero** `<w:hyperlink>` and
**zero** `<w:lastRenderedPageBreak>` in `document.xml`, `styles.xml`,
`stylesWithEffects.xml`, and `numbering.xml`, so registering the two tags is a
**pure element-lookup addition** — neither class is instantiated during the M1
sweep, and **M1 stays 17/17 byte-identical** (Gate-3 `pkgcompare` L0 PASS + 16 XML
L1 + 1 bin; `document.xml` 1548 B & `styles.xml` 349458 B L1). Registering
`w:hyperlink` also **closes the P4-2 `CT_P.text`-over-hyperlink VERIFY** (below)
with the registry row alone — no `CT_P` edit.

**This completes the text-oxml element layer** — `font` (P4-1a) · `run` (P4-1b) ·
`parfmt`/`paragraph` (P4-2) · `hyperlink`/`pagebreak` (P4-3). The text API proxies
(Font / Run / Paragraph) that consume these elements begin at P4-4.

### `CT_Hyperlink` — attributes, run text, and the `history` non-None default

`CT_Hyperlink` is a run container: it declares `r = ZeroOrMore("w:r")`
(`r_lst` / `new_r_` / `insert_r_` / `add_r_` / `add_r`, `successors=()` → append)
plus three `OptionalAttribute`s:

- **`rId`** (`r:id`, `XsdString`, default `None` → `[]`) — the relationship id of
  an **external** hyperlink target.
- **`anchor`** (`w:anchor`, `ST_String`, default `None` → `[]`) — the bookmark name
  of an **internal** hyperlink target.
- **`history`** (`w:history`, `ST_OnOff`, **`default=True`**) — a **non-None
  default**, so the D-delta-1 tri-state applies exactly as for `CT_OnOff`: an absent
  `@w:history` reads back **`true`**; setting `history = false` writes
  `w:history="0"`; setting `history = true` (== the default) **or** `[]` (None)
  **removes** `@w:history`. `ST_OnOff.from_xml` maps `1`/`true`/`on` → `true` and
  `0`/`false`/`off` → `false`.

**`.text`** shadows the lxml `.text` attribute (D10, the same mechanism as
`CT_R.text` / `CT_P.text`): it is the hyperlink's concatenated **run** text,
`"".join(r.text for r in self.xpath("w:r"))` over the **direct** `w:r` children,
ported by overriding the protected `getText_`. It is **read-only** — `setText_`
raises `mat2doc:AttributeError`. The serializer reads `text_raw_` (bypassing the
shadow), so the hyperlink's own char data (normally none) serializes unchanged.

**`lastRenderedPageBreaks`** returns every `<w:lastRenderedPageBreak>` descendant
via `self.xpath("./w:r/w:lastRenderedPageBreak")` (ported as a no-arg method so
display never evaluates it); now that `CT_LastRenderedPageBreak` is registered in
the same WP, these resolve to `CT_LastRenderedPageBreak` instances.

### `CT_LastRenderedPageBreak` — the precedes/follows detection algorithm

`CT_LastRenderedPageBreak` (`<w:lastRenderedPageBreak>`) is an **empty** element
(no attributes, no children) that carries **behaviour**, not data. The schema maps
`w:lastRenderedPageBreak` to `CT_Empty`; python-docx gives it a distinguished
complex-type name purely to attach the page-break split machinery. The whole class
is xpath-driven sibling/descendant traversal, ported faithfully over the `+oxml`
XPath engine (which already supports the `parent::`, `ancestor::`,
`following-sibling::`, `preceding-sibling::` axes, `last()`, and `self::` unions in
predicates — the engine battery was built ahead of this WP).

Four **public** accessors (ported as no-arg **methods** so display never evaluates
them and the `ValueError` guards fire only on an explicit call):

- **`precedes_all_content`** (common — a break on an even paragraph boundary):
  `false` if the break is inside a hyperlink (a hyperlink is *atomic*, there is
  always hyperlink text before it); else non-empty
  `./w:r[1]/w:lastRenderedPageBreak[not(preceding-sibling::*[<run-inner-content>])]`
  — i.e. the break is in the **first** run of the paragraph and nothing
  content-bearing precedes it in that run.
- **`follows_all_content`** (rare, hand-edited XML): `false` inside a hyperlink;
  else non-empty
  `(./w:r)[last()]/w:lastRenderedPageBreak[not(following-sibling::*[<run-inner-content>])]`
  — the break is in the **last** run and nothing content-bearing follows.
- **`preceding_fragment_p`** / **`following_fragment_p`**: a loose `CT_P` clone
  with, respectively, everything **from the break onward** or everything **up to and
  including the break** removed — the two halves `rendered_page_breaks` yields. Each
  guards with `self == _first_lrpb_in_p(_enclosing_p)` (H5 handle identity) and
  raises `mat2doc:ValueError` *"only defined on first rendered page-break in
  paragraph"* if this is not the paragraph's first break.

The `<run-inner-content>` predicate is the Constant `RUN_INNER_CONTENT_XPATH`,
ported verbatim (7-term `self::` union, exact `" | "` separators):
`self::w:br | self::w:cr | self::w:drawing | self::w:noBreakHyphen | self::w:ptab | self::w:t | self::w:tab`.

**Detection truth table** (Gate-3 `s0023`, 8 sample trees, MATLAB == python-docx on
all — the first break of each is probed):

| tree | shape | `precedes` | `follows` | note |
|---|---|---|---|---|
| t1 | `<r><lrpb/><t>text</t></r>` | **true** | false | break precedes all content (common even-boundary case) |
| t2 | `<r><t>text</t><lrpb/></r>` | false | **true** | break follows all content |
| t3 | `<r><t>a</t><lrpb/><t>b</t></r>` | false | false | mid-run |
| t4 | hyperlink-enclosed lrpb | false | false | in a hyperlink → both false (atomic) |
| t5 | `<r/><r><lrpb/><t>x</t></r>` | false | false | **the `w:r[1]`-only subtlety** — empty run-1, break at start of run-2 |
| t6 | two breaks in one run | false | false | first break; second → the `ValueError` guard |
| t7 | `<r><lrpb/></r>` | **true** | **true** | sole-content break (contrived T/T, not precluded by spec) |
| t8 | `<pPr>…</pPr><r><lrpb/><t>text</t></r>` | **true** | false | break-precedes with a `w:pPr` present |

### Fragment splits — atomic hyperlink vs bare run, byte-identical

`preceding_fragment_p` / `following_fragment_p` build the two halves on a
**`deepcopy` clone** of the enclosing `w:p` (mutations never touch the live tree).
The split differs by whether the break is inside a hyperlink:

- **Bare run** (`_preceding_frag_in_run` / `_following_frag_in_run`): remove all
  `w:p` inner-content on the far side of the break's **run** (keeping `w:pPr`), then
  remove all run inner-content on the far side of the break **within** its run
  (keeping `w:rPr`), then remove the break itself.
- **Atomic hyperlink** (`_preceding_frag_in_hlink` / `_following_frag_in_hlink`): a
  page-break inside a hyperlink is not "split" — the **whole hyperlink** goes into
  the *preceding* fragment (only the break is removed from inside it), and the
  *following* fragment drops the hyperlink and everything before it.

Gate-3 froze six split trees and compared **both** the fragment `.text` **and** the
raw-UTF-8 `serialize_part_xml` bytes (`serhex`, a byte pin) — all MATLAB ==
python-docx: t1 `""`/`"text"`, t2 `"text"`/`""`, t3 `"a"`/`"b"`, **t4 (atomic
hyperlink) `"prexy"`/`"post"`**, t7 `""`/`""`, and **t8 (`w:pPr` retained in both
fragments** — the `not(self::w:pPr)` filter proven at the byte level). The atomic
t4 case also serializes the parsed `xmlns:w`/`xmlns:r` declarations
**byte-identically** — confirming the D-serializer-nsdecl verbatim-until-moved fix
holds for these clone-and-mutate splits.

The six `@lazyproperty` members (`_is_in_hyperlink`, the four frag builders, and
`_run_inner_content_xpath` — a lazyproperty over a fixed literal → a `Constant`)
are cached with a **value + logical computed-flag** pair (design.md §2 — never
`isempty` as the sentinel). Caching is **faithful**, not merely an optimization:
each public Python `@property` re-reads the cached fragment, so after a hypothetical
tree mutation both Python and this port return the **stale** first fragment (Gate-2
proved this live on both sides).

### The P4-2 `CT_P.text`-over-hyperlink closure

At P4-2, `w:hyperlink` was unregistered, so `CT_P.text` over a paragraph
*containing* a hyperlink returned the generic element's own char data. Registering
`w:hyperlink` → `CT_Hyperlink` this WP flips it to the hyperlink's concatenated run
text — **with no `CT_P` edit**. Gate-3 `s0023`:
`<w:p><w:r>Before </w:r><w:hyperlink r:id="rId1"><w:r>link</w:r></w:hyperlink><w:r> after</w:r></w:p>`
→ `p.text == "Before link after"` (MATLAB == python-docx). This is the same
dependency-order shape as the `w:tab` gap P4-2 closed.

:::{note}
**Accepted deviation — the s7 created-element nsdecl residual (no new D-number).**
On a **loose created** `<w:hyperlink>` serialized standalone, setting `rId` mints an
`xmlns:ns0` for the relationships URI; **clearing `rId`** then leaves it orphaned —
lxml keeps the now-unused `xmlns:ns0` while the Mat2Doc serializer recomputes the
used namespaces and drops it. The two forms have **identical elements, attributes,
text, and expanded names** (exclusive-C14N-equal); the difference is
namespace-declaration-only. It is a **created-element manifestation of the signed
`D-serializer-nsdecl`** (whose declaration-emission fix covered *parsed* verbatim
nodes only), carried as an **accepted, deferred residual — no new D-number**. It is
**unreachable on every real path**: real hyperlinks keep their `rId`, and a
hyperlink inside `document.xml` is rooted under `w:document` (which already declares
`xmlns:r`) so `r:id` uses the existing prefix and no namespace is ever minted — M1
stays 17/17 byte-neutral. Reopen-check booked at **P4-5b** (`add_hyperlink`) and
**P7** (`add_picture`). See
`validation\summary\decision_2026-07-28_nsdecl_created_element_orphan.md` and the
`D-serializer-nsdecl` proof.
:::

---

## `CT_Hyperlink`

**Syntax**

```matlab
h = mat2doc.oxml.OxmlElement("w:hyperlink");   % a CT_Hyperlink (registered)
h.rId    = "rId7";                              % external target relationship id
h.anchor = "section1";                          % OR an internal bookmark name
tf       = h.history;                           % true when @w:history absent (default)
r = h.add_r();                                   % append a <w:r/>
txt = h.text;                                    % concatenated run text (shadow, D10)
```

**Description**

The `<w:hyperlink>` element — a run container carrying a hyperlink's display text
and address. `r = ZeroOrMore("w:r")` (`r_lst` / `new_r_` / `insert_r_` / `add_r_` /
`add_r`, `successors=()` → append). Three `OptionalAttribute`s: `rId` (`r:id`,
`XsdString`, default `None` → `[]`), `anchor` (`w:anchor`, `ST_String`, default
`None` → `[]`), and `history` (`w:history`, `ST_OnOff`, **`default=True`**) — a
non-None default driving the D-delta-1 tri-state (absent → `true`; `false` →
`w:history="0"`; `true` (== default) or `[]` → `@w:history` removed). `.text`
shadows the lxml `.text` (a computed join of the direct `w:r` children's live
`CT_R.text`, ported via `getText_`, D10) and is **read-only** — `setText_` raises
`mat2doc:AttributeError`. `lastRenderedPageBreaks` returns the
`./w:r/w:lastRenderedPageBreak` descendants. Registering this class closes the P4-2
`CT_P.text`-over-hyperlink VERIFY.

**Example**

```matlab
h = mat2doc.oxml.OxmlElement("w:hyperlink");
disp(h.history);                 % 1  (true — @w:history absent; default True)
h.rId = "rId7";                  % <w:hyperlink r:id="rId7">
r = h.add_r();  r.add_t("here");
disp(h.text);                    % "here"  (concatenated run text)
h.history = false;               % serializes @w:history="0"
disp(h.history);                 % 0  (false)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/hyperlink.py::CT_Hyperlink`*

---

## `CT_LastRenderedPageBreak`

**Syntax**

```matlab
% a break at the start of a paragraph's first run
brks = p.xpath("./w:r/w:lastRenderedPageBreak");  % index in two steps —
lrpb = brks(1);                                    % MATLAB has no f(...)(1) chained index
tf   = lrpb.precedes_all_content();    % often true (even paragraph boundary)
frag = lrpb.preceding_fragment_p();    % loose CT_P: the break and all that follows removed
```

**Description**

The `<w:lastRenderedPageBreak>` element — a renderer-inserted page break, an empty
element that is a child of `CT_R` (peer to `CT_Text`). It carries the
precedes/follows **detection** and the paragraph fragment-**split** algorithm that
`Paragraph.rendered_page_breaks` (a later WP) is built on. `precedes_all_content`
is `false` inside a hyperlink, else true iff the break is in the paragraph's first
run with no content-bearing sibling before it; `follows_all_content` is the
last-run mirror. `preceding_fragment_p` / `following_fragment_p` return loose `CT_P`
clones split at the break (atomic when the break is inside a hyperlink — the whole
hyperlink stays with the preceding page), each raising `mat2doc:ValueError` unless
this is the paragraph's first break. `RUN_INNER_CONTENT_XPATH` (the 7-term `self::`
union) is a verbatim `Constant`; the six `@lazyproperty` members use the design.md
§2 value + logical-flag cache.

**Example**

```matlab
% <w:p><w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r></w:p>
xml = "<w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">" + ...
      "<w:r><w:lastRenderedPageBreak/><w:t>text</w:t></w:r></w:p>";
p = mat2doc.oxml.parse_xml(xml);
brks = p.xpath("./w:r/w:lastRenderedPageBreak");
lrpb = brks(1);                            % MATLAB has no f(...)(1) chained index
disp(lrpb.precedes_all_content());         % 1  (true — break precedes all content)
disp(lrpb.follows_all_content());          % 0  (false)
fp = lrpb.following_fragment_p();
disp(fp.text);                             % "text"  (content after the break)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/pagebreak.py::CT_LastRenderedPageBreak`*

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

**P4-2 (the paragraph tier) also opened 0 new D-numbers.** Every P4-2-surface
fact is L1 byte-identical or value-exact (Gate-3 `s0022` `probe_diff` **MATCH
exit 0** — the `CT_PPr` H11 scrambled-add-order crux 4/4 byte-identical, the
`CT_TabStop` `str_`→`"\t"` carry-forward, `CT_Ind`/`CT_Jc`/`CT_Spacing` + all 13
`CT_PPr` accessors, `CT_P` pPr-first, and the 204-`w:pPr` `styles.xml` round-trip);
the M1 17/17 byte-neutrality gate held (`document.xml` 1548 B & `styles.xml`
349458 B L1); no L2 canonical-only result surfaced. The standing adopt-only
umbrella (**D-001** / **D-serializer-nsdecl** / **D-delta-1/-2** — exercised by
`CT_TabStop.leader` default `SPACES` and every default-`None` attr — / **D-delta-4**
— public `add_tab`/`add_r`/`add_hyperlink` — / **D-STYPE-\*** — the signed/unsigned
twips types — / **D10** — the `CT_P.text` shadow — / **D-zip-time**) covers
everything; no ledger row was added.

:::{note}
**Gate-3 verdict was FAIL on regression neutrality (558/566), then routed to a
Gate-4 test-pin update — not a P4-2 source change.** Registering the 12 tags
correctly flipped 8 pre-existing class-specificity pins
(`verifyClass(...,'XmlElement')` on xpath-result arrays and created children) RED,
because those node-sets now resolve to their proper registered `CT_*` subclasses
instead of a generic `XmlElement` — the *intended, byte-neutral* effect of
registration. Every node **value** assertion (tag/path/count/dedup) stayed green,
and one test's own comment (*"created w:p resolves to generic XmlElement (CT_P is
P4)"*) had explicitly predicted the flip. Gate-4 re-pinned the 8 (relaxed the 7
`Test_p1_3x_xpath` umbrella pins to `IsInstanceOf`, re-pinned `cP`→`CT_P` in
`Test_p2_3_document_shell`) and added 9 new permanent P4-2 tests — **cold total
566 → 575**. The equivalence surface was byte-correct throughout.
:::

**P4-3 (the hyperlink + page-break tier) completes the text-oxml layer with 0 new
D-numbers.** The entire behavioural + byte surface is equivalent to python-docx
1.2.0 — M1 17/17 L1, the `CT_LastRenderedPageBreak` detection over 8 trees, the
fragment-split byte pins (incl. the atomic-hyperlink and `w:pPr`-retention cases),
the second-break `ValueError` guard, the `CT_Hyperlink` `history` H3 tri-state
(7/8 mutation steps byte-identical), `CT_Hyperlink.text`, and the `CT_P.text`-over-
hyperlink closure — with **regression 575/575 GREEN** and **zero stale-pin flips**
(the `w:hyperlink` / `w:lastRenderedPageBreak` xpath results are not exact-class-
pinned anywhere). The **one** divergence is the accepted **s7 created-element
`ns0` residual** documented in the tier note above — a **manifestation of the
signed `D-serializer-nsdecl`** (created-element minted-then-orphaned sub-case,
which its parsed-verbatim fix did not cover), **no new D-number**. It is
namespace-declaration-only (exclusive-C14N-equal), reads **L3 only** under
`pkgcompare`'s full-in-scope-nsmap canonicalizer, and is **unreachable on every
real path** (real hyperlinks keep their `rId`; document-tree hyperlinks use the
root's `xmlns:r` and never mint `ns0`). Recorded **PASS-DEVIATION(D-serializer-
nsdecl)**, ACCEPTED/deferred with a P4-5b/P7 reopen-check
(`decision_2026-07-28_nsdecl_created_element_orphan.md`); Gate-4 pins the port's
actual s7 output with a known-deviation guard. No ledger row was added.
