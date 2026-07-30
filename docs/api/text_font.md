---
title: "mat2doc.text.Font / mat2doc.dml.ColorFormat — the text/font API tier (P4 API tier OPENS)"
---

# `mat2doc.text.Font` / `mat2doc.dml.ColorFormat` — the text/font API tier

Ported from python-docx v1.2.0 `src/docx/text/font.py::Font` (package
`+mat2doc/+text/`) and `src/docx/dml/color.py::ColorFormat` (package
`+mat2doc/+dml/`). These are the **first two API-tier proxies** of the port.

:::{note}
**This WP (P4-4a) opens the P4 API tier.** Everything above it in P4 was the
**oxml element layer** — `CT_RPr`, `CT_R`, `CT_PPr`, `CT_P`, the run/paragraph
content classes, and the `<w:color>`/`<w:sz>`/`<w:u>`/… leaves (see
[`api/oxml_text.md`](oxml_text.md)) — which the text-oxml layer **completed** at
P4-3. `Font` and `ColorFormat` are the first **user-facing proxies** that sit on
top of that layer: they hold no XML of their own, they wrap a `w:r` and reach
through it to the P4-1a `CT_RPr` / `CT_Color` helpers. The remaining P4 API
proxies (`Run`, `Paragraph`, the styles chain) build up from here toward **M2**.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## What "API tier" means here — behaviorally equivalent, byte-neutral

`Font` and `ColorFormat` are `ElementProxy` subclasses. Unlike every P4 class
before them, they add **no `register_element_cls` row, no oxml class, and no
serialization code** — nothing on the save path moves. Their entire job is to
**get/set** the correct `<w:rPr>` XML by delegating to the CT_RPr / CT_Color
helpers byte-validated at P4-1a. So the equivalence bar for this WP is
**behavioral**, not byte-fixture: every get returns the same value python-docx
returns, and every set produces the same serialized `<w:r>` bytes.

**M1 is trivially preserved.** The default template instantiates neither class
and the save path is unchanged, so the M1 17/17 byte-neutrality sweep holds
unchanged (`mat2doc.Document().save()` → 17/17 byte-identical to the frozen
`references\s0001` reference). There is no stale-pin risk (no registry rows) and
no new D-number.

| Python `src/docx/...` | MATLAB | class |
|---|---|---|
| `text/font.py` | `+mat2doc\+text\` | `Font` |
| `dml/color.py` | `+mat2doc\+dml\` | `ColorFormat` |

Both packages are new. `+mat2doc\+text` (`mat2doc.text.Font`) is distinct from
the oxml `+mat2doc\+oxml\+text` (`mat2doc.oxml.text.CT_*`); `+mat2doc\+dml`
(`mat2doc.dml.ColorFormat`) is the first `+dml` class.

## The proxy shape — a `Font` wraps a run, not an `rPr`

`Font(r, parent=None)` wraps a **run** element (`w:r` / `CT_R`), not the
`<w:rPr>` it carries. Every accessor reaches the run-properties lazily —
`self._element.rPr` on read (returning `[]`/None if the run has no `<w:rPr>`
yet) and `self._element.get_or_add_rPr()` on write (materializing one). `Font`
adds no oxml logic: each property delegates to a `CT_RPr` helper ported at
P4-1a. The constructor is faithful to `font.py` including its two redundant/dead
members — the `self._element = r` re-assign (already set by the `ElementProxy`
super) and `self._r = r` (`-> r_`, set from the run but **never read** anywhere
in `font.py`).

`ColorFormat(rPr_parent)` likewise wraps the **run** (the `<w:rPr>` *parent*),
not the `<w:color>` element, and reaches the color through
`self._element.rPr.color` (the private `_color` helper).

`ElementProxy` reference semantics and the H5 element-identity `eq`/`ne` are
inherited unchanged: two `Font` proxies over the same run compare **equal**;
`Font.color` returns a **fresh** `ColorFormat` each access (not cached), and it
compares equal to the run's `Font` because both wrap the same element.

(id-font-tri-state)=
## The tri-state (H3) — `[]` is the "inherited" sentinel

All ~27 `Font` properties and all 3 `ColorFormat` properties are **None-vs-value**
tri-states. In python-docx a `None` return means "this setting is inherited from
the style hierarchy" (the property is not set on this run); assigning `None`
clears it. Mat2Doc renders `None` as the empty `[]` sentinel on both get and set,
via inline `isequal(x, [])` (the established Mat2Doc None idiom — no shared
`isNone` helper). So `f.bold` on a run with no `<w:b>` returns `[]`; `f.bold = []`
removes it.

---

## `Font`

**Syntax**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");   % a CT_R (run)
f = mat2doc.text.Font(r);              % wrap the run's character properties
f.bold = true;                         % <w:rPr><w:b/></w:rPr>
tf  = f.italic;                        % bool | [] (inherited)
cf  = f.color;                         % a fresh ColorFormat (read-only)
```

**Description**

Character properties of a run: font name, size, and the whole boolean-toggle
family (bold, italic, all-caps, …), plus color, highlight, underline, and
sub/superscript. An `ElementProxy` over a `w:r`; every accessor delegates to the
P4-1a `CT_RPr` helpers. Reads a `<w:rPr>` lazily (absent → `[]`); writes
`get_or_add_rPr()` first.

### The 20 boolean properties → `w:rPr` child tags

Each boolean property is a tri-state (`true` / `false` / `[]`) delegating to
`CT_RPr.get_bool_val_` / `set_bool_val_` (P4-1a) via the private
`_get_bool_prop` / `_set_bool_prop` (`font.py:418-428`). The `name` argument each
passes is the **CT_RPr property name** (the `w:` local tag minus prefix). The
full map (verified vs `font.py` line-by-line):

| Font property | `w:rPr` child tag | Font property | `w:rPr` child tag |
|---|---|---|---|
| `all_caps` | `w:caps` | `math` | `w:oMath` |
| `bold` | `w:b` | `no_proof` | `w:noProof` |
| `complex_script` | `w:cs` | `outline` | `w:outline` |
| `cs_bold` | `w:bCs` | `rtl` | `w:rtl` |
| `cs_italic` | `w:iCs` | `shadow` | `w:shadow` |
| `double_strike` | `w:dstrike` | `small_caps` | `w:smallCaps` |
| `emboss` | `w:emboss` | `snap_to_grid` | `w:snapToGrid` |
| `hidden` | `w:vanish` | `spec_vanish` | `w:specVanish` |
| `imprint` | `w:imprint` | `strike` | `w:strike` |
| `italic` | `w:i` | `web_hidden` | `w:webHidden` |

The tri-state fires through `CT_OnOff`'s D-delta-1 mechanism (see
[the oxml page](oxml_text.md#ct_onoff--ct_string-tri-state-the-c3-leaves)):
`bold = true` → a bare `<w:b/>`; `bold = false` → `<w:b w:val="0"/>`; `bold = []`
→ the tag removed; get on an absent tag → `[]`.

### The non-boolean properties

- **`color`** (read-only, `font.py:50-54`) — returns a **fresh** `ColorFormat`
  wrapping the run each access (`return ColorFormat(self._element)`; not cached).
- **`highlight_color`** (`font.py:133-144`) — a `WD_COLOR_INDEX` member or `[]`,
  via `CT_RPr.highlight_val` (`<w:highlight>`).
- **`name`** (`font.py:184-200`) — get reads `rFonts_ascii`; **set writes BOTH**
  `rFonts_ascii` **and** `rFonts_hAnsi` (the ascii/hAnsi asymmetry — and the
  hAnsi-doesn't-remove-`w:rFonts` quirk — live inside `CT_RPr`, P4-1a).
  `name = "Calibri"` → `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>`.
- **`size`** (`font.py:254-278`) — get returns `CT_RPr.sz_val`, a `Length` in
  **EMU** (the `<w:sz>` value is in **half-points**, converted by `ST_HpsMeasure`
  at P4-1a); set wraps the arg in `Emu()` and assigns `sz_val` (`[]` → remove).
  `size = Pt(12)` → `<w:sz w:val="24"/>` (24 half-points); get → `Length` 152400
  EMU → `.pt` = 12.0. **D-STYPE-1** (int/float indistinguishability on `Length`)
  applies here — adopt-only, non-engaged (all integral-EMU sizes serialize
  byte-identically).
- **`subscript` / `superscript`** (`font.py:334-367`) — tri-state booleans via
  `CT_RPr.subscript` / `.superscript`; the `<w:vertAlign>` truth table (setting
  one switches `@w:val`; clearing the *opposite* does **not** remove; clearing
  the *set* one removes) lives in `CT_RPr` (P4-1a).

### `underline` — the `WD_UNDERLINE` tri-state (`font.py:369-403`)

`underline` is not a plain bool: it is a **tri-state-plus-enum**. On **get**,
`CT_RPr.u_val` maps as:

| `u_val` | `f.underline` |
|---|---|
| `[]` (`<w:u>` absent) or `INHERITED` | `[]` (inherited) |
| `SINGLE` | `true` |
| `NONE` | `false` |
| any other `WD_UNDERLINE` member (e.g. `DOUBLE`, `WAVY`) | the member itself |

On **set**, the mapping uses **strict boolean identity** (not truthiness), so
that `true → SINGLE`, `false → NONE`, while a `WD_UNDERLINE` member or `[]` passes
through unchanged:

| assigned | `<w:u>` result |
|---|---|
| `true` | `@w:val="single"` |
| `false` | `@w:val="none"` |
| `WD_UNDERLINE.DOUBLE` | `@w:val="double"` |
| `WD_UNDERLINE.WAVY` | `@w:val="wave"` |
| `[]` | `<w:u>` removed |

The get uses an explicit `isequal(val, [])` first branch (the `[]`-vs-member
compare in MATLAB yields an *empty* logical, not `false`); this is behaviorally
identical to Python's `==`-fall-through — both the None and INHERITED cases
collapse to `[]` (byte-verified in the Gate-3 battery, VERIFY-1 closed). The set
uses `islogical(value) && isscalar(value) && value` guards (Python `value is True`
/ `is False`) so a member/`[]` is never coerced.

**Example**

```matlab
r = mat2doc.oxml.OxmlElement("w:r");
f = mat2doc.text.Font(r);
f.bold = true;                                        % <w:rPr><w:b/></w:rPr>
f.size = mat2doc.shared.Pt(12);                       % <w:sz w:val="24"/>  (half-points)
f.name = "Calibri";                                   % <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri"/>
f.underline = mat2doc.enum.text.WD_UNDERLINE.DOUBLE;  % <w:u w:val="double"/>
f.color.rgb = mat2doc.shared.RGBColor(255, 0, 0);     % <w:color w:val="FF0000"/>
disp(double(f.size));                                 % 152400   (a Length/Emu)
```

*Ported from python-docx v1.2.0: `src/docx/text/font.py::Font`*

---

## `ColorFormat`

**Syntax**

```matlab
r  = mat2doc.oxml.OxmlElement("w:r");
cf = mat2doc.dml.ColorFormat(r);                       % or: cf = Font(r).color
cf.rgb = mat2doc.shared.RGBColor(60, 47, 128);         % <w:color w:val="3C2F80"/>
t  = cf.type;                                          % MSO_COLOR_TYPE member | []
th = cf.theme_color;                                   % MSO_THEME_COLOR member | []
```

**Description**

Access to a run's color settings: `rgb` (an `RGBColor` value), `theme_color`
(an `MSO_THEME_COLOR` member), and `type` (read-only, an `MSO_COLOR_TYPE`
member). An `ElementProxy` over the run; reaches the color via
`self._element.rPr.color` (the private `_color` helper). Every absent-child /
absent-attribute path returns `[]` (None); setting `[]` removes.

### The `type`-detection precedence — THEME > AUTO > RGB (`color.py:86-101`)

`type` reads the `<w:color>` child once and classifies in a **fixed order**:

1. no color at all → `[]` (None);
2. `@w:themeColor` present → **THEME** (checked **first**);
3. else `@w:val == "auto"` → **AUTO**;
4. else → **RGB**.

The order matters: a color carrying **both** a `themeColor` **and** an RGB
`@w:val` reports **THEME**. Word writes the RGB as a "good guess" alongside the
theme color, but the theme color wins — so `theme_color = ACCENT_1` after
`rgb = RGBColor(FF,00,00)` yields `<w:color w:val="FF0000" w:themeColor="accent1"/>`
whose `type` is **THEME**, while `rgb` still reads back `FF0000` (the value is
retained, not cleared).

### The `InvalidXmlError` edge — `rgb` on a themeColor-only color

`rgb` reads `color.val` directly, and `CT_Color.val` is a **required** `@w:val`
attribute. On a **parsed** `<w:color w:themeColor="accent1"/>` — a `themeColor`
present but **no `@w:val`** — reading `rgb` **raises `mat2doc:InvalidXmlError`**,
byte-identically to python-docx (which raises the same required-attribute error;
the message is compared verbatim). `type` and `theme_color` on that **same**
element are **safe** — `type` checks `themeColor` before `val`, so only `rgb`
trips the required-attr path. This faithfully reproduces python-docx's
required-attribute behavior on the `@w:val`-absent path.

### `rgb` / `theme_color` set semantics

- `rgb = RGBColor(...)` — removes any existing `<w:color>`, then adds a fresh one
  and sets `@w:val`. `rgb = []` on a run with **no** color is a **no-op**
  (leaves `<w:r/>` — no `<w:rPr>` created); `rgb = []` on an existing color
  removes it.
- `theme_color = MEMBER` — forces a `<w:color>` and sets `@w:themeColor`.
  `theme_color = []` removes the color (only when a color **and** an `rPr` both
  already exist; otherwise a no-op).

`color.val` returns **either** an `RGBColor` object **or** the string `"auto"`;
the `== ST_HexColorAuto.AUTO` comparison ports to `isequal(color.val, AUTO)` —
`isequal(RGBColor, "auto")` is false (different class), `isequal("auto", "auto")`
is true — reproducing the Python `==` for both branches.

**Example**

```matlab
r  = mat2doc.oxml.OxmlElement("w:r");
cf = mat2doc.dml.ColorFormat(r);
cf.rgb = mat2doc.shared.RGBColor(60, 47, 128);          % <w:color w:val="3C2F80"/>
cf.type == mat2doc.enum.dml.MSO_COLOR_TYPE.RGB          % true
cf.theme_color = mat2doc.enum.dml.MSO_THEME_COLOR.ACCENT_1;
cf.type == mat2doc.enum.dml.MSO_COLOR_TYPE.THEME        % true (checked before val)
disp(cf.rgb.str_());                                    % "3C2F80"  (RGB val retained)
```

*Ported from python-docx v1.2.0: `src/docx/dml/color.py::ColorFormat`*

---

## Deviation posture — 0 new D-numbers

`Font` and `ColorFormat` add no output-visible divergence. The two standing
adopt-only classes touch this WP but **do not engage**:

- **D-STYPE-1** (int/float indistinguishability on `Length`/`size`) — applies to
  `Font.size` set via `Emu(emu)`; every size in the validation battery is
  integral-EMU and serializes byte-identically.
- **D-serializer-nsdecl** — **does not engage**: every `Font`/`ColorFormat` run
  is `w:`-only (no `r:` attribute), so no `ns0` is ever minted and all serhex
  byte pins (including the XML prolog) are byte-identical.

Gate-3 recorded `PASS-DEVIATION(D-STYPE-1, D-serializer-nsdecl)` as adopt-only
context only — both non-engaged, operational verdict a clean **PASS** with zero
measured deviation, **zero new D-numbers**.
