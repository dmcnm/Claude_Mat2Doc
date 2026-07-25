---
title: "mat2doc.shared — length units and RGBColor"
---

# `mat2doc.shared` — length units and RGBColor

Ported from python-docx v1.2.0 module `src/docx/shared.py`
(package `+mat2doc/+shared/`). The `Length` base class and its six convenience
constructors (`Emu`, `Inches`, `Cm`, `Mm`, `Pt`, `Twips`) express lengths in
English Metric Units (EMU); the `RGBColor` value object holds a red/green/blue
byte triplet. One project-added package-private helper, `pyIntArg`, has no
python-docx counterpart.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## Where this differs from python-pptx (the design guide)

python-docx v1.2.0 is the **source of truth**. Mat2Doc shares no code with
Mat2Ppt (the python-pptx port); the design is re-implemented in the `mat2doc:`
namespace. The docx `shared.py` structure differs from the pptx `util.py`
design guide in exactly these ways, all honored here:

- **`Twips` present, `Centipoints` absent.** docx defines `Twips`
  (`_EMUS_PER_TWIP = 635`; a `twips` property) and has **no** `Centipoints`
  class, no `_EMUS_PER_CENTIPOINT`, and no `centipoints` property (grep-verified
  over `src/docx`). The WP-brief scope line listing "Centipoints" was a pptx
  carry-over; adding it would be an unprovenanced feature, so it is **not**
  ported. The property set is `{cm, emu, inches, mm, pt, twips}`.
- **`RGBColor` lives in `shared.py`** (python-pptx had it in `dml/color.py`),
  so the port lives in `+mat2doc\+shared\RGBColor.m`.
- **`RGBColor` raises a two-way error split** — `TypeError` for a non-int
  component, `ValueError` for an out-of-range one (python-pptx conflated both to
  `ValueError`). Both are honored as `mat2doc:TypeError` / `mat2doc:ValueError`.
- **`RGBColor.__repr__` is present** in docx (`"RGBColor(0x%02x, 0x%02x, 0x%02x)"`)
  and is ported as `repr_()`.

## Naming — no renames (FLAG-3-docx resolved by convention)

The docx→MATLAB package layout introduces **no same-package name collisions**
(the P0 collision scan found zero), so **no symbol was renamed**. Every symbol
keeps its python-docx name under `+mat2doc\+shared\`. This closes **FLAG-3-docx**
"resolved by convention" per
`decision_2026-07-25_mat2doc_naming_collision_policy.md`.

## Deviation posture (pre-adopted, carried `mat2doc:`-namespaced)

All construction-domain divergences are on **dead paths** and are ledgered
against the pre-adopted Mat2Doc rulings
(`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`):

- **D-STYPE-1..4** — a MATLAB `double` cannot distinguish `914400` (int) from
  `914400.0` (float), nor `255` from `255.0`; an integer-valued double is
  accepted as the Python int.
- **D-002** — exotic string inputs (unicode digits, digit strings `> realmax`
  or `> 4300` digits, and the `Twips` repunit case) are rejected where CPython's
  arbitrary-precision `int` would build a value; ASCII-only int-literal grammar.
- **D-003** — the string-rejection error **message** of the multiplier
  constructors (`Inches`/`Cm`/`Mm`/`Pt`/`Twips`) is port-authored; the exception
  **class** (`mat2doc:ValueError`) is faithful to CPython.
- **D-004** — non-finite / wrong-type inputs raise MATLAB-native identifiers
  (`MATLAB:validators:mustBeFinite`, `mat2doc:TypeError`) where CPython raises
  `ValueError` / `OverflowError` / `TypeError`; the **raising** decision is
  faithful and no upstream `except` observes the identifier.

Gate-3 (`validation\mat2doc\reports\p1_1_validation.md`) confirmed all four as
**recurrences** — **0 new D-numbers**, 88/88 records bit-exact.

---

## `Length`

**Syntax**

```matlab
len = mat2doc.shared.Length(emu)
```

**Description**

Base class for the length classes. Constructs a length of `emu` English Metric
Units (914400 EMU per inch, 635 EMU per twip). `emu` may be a numeric scalar
(truncated toward zero like Python `int()`), a logical scalar (`true` → 1), or a
base-10 integer string/char (parsed like Python `int(str)`; `'2.5'` errors — a
string is parsed, never truncated).

`Length` subclasses `double`. Because the subclass defines properties, every
built-in numeric operation (`+`, `-`, `*`, `/`, `floor`, ...) returns a plain
`double` (the `Length` type is lost) — exactly Python's `class Length(int)`
degradation, where arithmetic returns a plain `int` and only the constructors
re-wrap. `isa(x,'double')` and `isa(x,'mat2doc.shared.Length')` are both true;
`double(x)` is the EMU value, held exactly (EMU magnitudes are far below 2^53).
One consequence: concatenating `Length` objects (`[a b]`) **errors** on R2024b —
hold them in cell arrays or `double()` them first. Not API-visible (Python never
relies on `Length` surviving aggregation).

:::{note}
**The string-parse mode is faithful but DEAD in docx.** The `Length`/`Emu`
constructors accept a base-10 string per Python `int(str)`, but no docx path
hands them a raw string: every `oxml/simpletypes.py` `ST_` length site applies
`int()` / `round()` to the raw XML attribute **before** construction — e.g.
`Emu(int(str_value))`, `Twips(int(str_value))`,
`Twips(int(round(float(str_value))))`, `Emu(int(round(quantity*multiplier)))`
(simpletypes.py:204/350/366/402/424). Keeping the parse mode is correct
(it matches CPython) and **strengthens** D-002.
:::

Read-only dependent properties mirror the Python `@property` bodies:

| Property | Meaning |
|---|---|
| `inches` | floating-point length in inches (true division) |
| `cm` | floating-point length in centimeters |
| `emu` | integer length in EMU (returns the instance itself, Python `return self`) |
| `mm` | floating-point length in millimeters |
| `pt` | floating-point length in points |
| `twips` | **integer** length in twips (1/20 point, 635 EMU); `int(round(emu/635))` |

**The `twips` half-to-even + signed-zero note.** The `twips` getter is
`int(round(self / float(635)))`. Python 3 `round()` with no `ndigits` rounds
**half-to-even** and returns an `int`; MATLAB `round()` rounds half **away from
zero**, so the port replicates CPython `float.__round__` verbatim
(`r = round(x); if abs(x-r)==0.5, r = 2*round(x/2)`). For integer EMU an exact
`.5` tie is mathematically impossible (635 is odd), but the algorithm is
reproduced so the result is provably identical for every input double.
Separately, MATLAB `round()` manufactures IEEE **`-0.0`** for inputs in
`(-0.5, 0)` (the reachable EMU band `-1 … -317`), whereas Python `int()` always
yields a true `0`. The getter normalizes with `+ 0` (`-0.0 → +0.0`) — the
Gate-3 comparator is signed-zero-aware and pins `Emu(-1..-317).twips == +0`.
Downstream this matters because `ST_SignedTwipsMeasure.convert_to_xml`
serializes `str(emu.twips)`, so a stray `-0` would have surfaced as `'-0'` in
XML.

**Example**

```matlab
len = mat2doc.shared.Length(914400);
len.inches   % 1
len.cm       % 2.54
len.twips    % 1440
len.emu      % 914400  (returns the Length itself)
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Length`*

---

## `Emu`

**Syntax**

```matlab
len = mat2doc.shared.Emu(emu)
```

**Description**

Convenience constructor for a length given in English Metric Units. `int(emu)`
is applied to the raw argument, so the full `int()` domain applies: numeric
truncates (toward zero, `fix`), logical maps to 0/1, and a base-10 string is
parsed (faithful but dead in docx — see the `Length` note). Returns an `Emu`,
which is a `Length`.

**Example**

```matlab
len = mat2doc.shared.Emu(457200);
len.inches   % 0.5
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Emu`*

---

## `Inches`

**Syntax**

```matlab
len = mat2doc.shared.Inches(inches)
```

**Description**

Convenience constructor for a length given in inches (914400 EMU per inch).
Python computes `int(inches * 914400)` — multiply first, then truncate toward
zero (`fix`). Accepts numeric or logical input; string input raises
`mat2doc:ValueError` (exception class faithful; message port-authored, D-003 —
Python reaches `int(str * 914400)`, string repetition, which raises for every
string in CPython 3.13).

**Example**

```matlab
len = mat2doc.shared.Inches(1);
len.emu   % 914400
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Inches`*

---

## `Cm`

**Syntax**

```matlab
len = mat2doc.shared.Cm(cm)
```

**Description**

Convenience constructor for a length given in centimeters (360000 EMU per
centimeter). Python computes `int(cm * 360000)`. Accepts numeric or logical
input; string input raises `mat2doc:ValueError` (class faithful; message
port-authored, D-003).

**Example**

```matlab
len = mat2doc.shared.Cm(2.54);
len.emu   % 914400
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Cm`*

---

## `Mm`

**Syntax**

```matlab
len = mat2doc.shared.Mm(mm)
```

**Description**

Convenience constructor for a length given in millimeters (36000 EMU per
millimeter). Python computes `int(mm * 36000)`. Accepts numeric or logical
input; string input raises `mat2doc:ValueError` (class faithful; message
port-authored, D-003).

**Example**

```matlab
len = mat2doc.shared.Mm(25.4);
len.emu   % 914400
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Mm`*

---

## `Pt`

**Syntax**

```matlab
len = mat2doc.shared.Pt(points)
```

**Description**

Convenience constructor for a length given in points (12700 EMU per point).
Python computes `int(points * 12700)`. Accepts numeric or logical input
(`Pt(true)` → 12700, since Python `int(True * 12700)`); string input raises
`mat2doc:ValueError` (class faithful; message port-authored, D-003).

**Example**

```matlab
len = mat2doc.shared.Pt(72);
len.emu   % 914400
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Pt`*

---

## `Twips`

**Syntax**

```matlab
len = mat2doc.shared.Twips(twips)
```

**Description**

Convenience constructor for a length given in twips — a twip is a twentieth of a
point, 635 EMU. **docx-specific** (python-pptx has no `Twips`). Python computes
`int(twips * 635)`. Accepts numeric or logical input; string input raises
`mat2doc:ValueError`. For the multiplier `K = 635`, a `<=6`-digit numeric string
is the single spot where CPython's `int(str * 635)` succeeds with a garbage
repunit integer (the docx analogue of pptx `Centipoints`' `<=33`-digit case at
`K = 127`); the port rejects it (dead path, recorded D-002).

**Example**

```matlab
len = mat2doc.shared.Twips(1440);
len.inches   % 1
len.emu      % 914400
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Twips`*

---

## `RGBColor`

**Syntax**

```matlab
c   = mat2doc.shared.RGBColor(r, g, b)
c   = mat2doc.shared.RGBColor.from_string(rgb_hex_str)
s   = c.str_()      % "3C2F80"   (Python __str__)
s   = c.repr_()     % "RGBColor(0x3c, 0x2f, 0x80)"  (Python __repr__)
ch  = c.char()      % '3C2F80'   (convenience char row)
tf  = (a == b)      % value equality
```

**Description**

Immutable value object holding a red/green/blue triplet of integers in the range
0–255. The Python original is a `tuple` subclass
(`class RGBColor(Tuple[int, int, int])`), an immutable value with tuple
equality-by-value; the port is a MATLAB **value class** (not a `handle`) holding
the three components in immutable properties, with `eq` / `ne` comparing by value
(two instances with the same triplet are equal; a comparison against a
non-`RGBColor` is `false` — the MATLAB analogue of tuple `==`).

**The TypeError / ValueError split.** The constructor iterates `(r, g, b)` and,
for each component **in order**, raises `mat2doc:TypeError` if it is not an
integer (Python `isinstance(val, int)`) and then `mat2doc:ValueError` if it is
out of range (`val < 0` or `val > 255`). Both use the byte-exact message
`RGBColor() takes three integer values 0-255`. Because a MATLAB `double` cannot
distinguish `255` from `255.0` (**D-STYPE-1**), an integer-**valued** real
numeric scalar is accepted as the Python int; a non-integral value is "not an
int" → `TypeError`.

| input | class | message |
|---|---|---|
| `RGBColor(256,0,0)` / `(-1,0,0)` / `(0,300,0)` | `mat2doc:ValueError` | `RGBColor() takes three integer values 0-255` |
| `RGBColor(2.5,0,0)` (non-int) / `RGBColor("ff",0,0)` | `mat2doc:TypeError` | `RGBColor() takes three integer values 0-255` |

**Hex formatting (H14).** `str_` is `"%02X%02X%02X"` — fixed-width **UPPERCASE**
hex (e.g. `"3C2F80"`). `repr_` is `"RGBColor(0x%02x, 0x%02x, 0x%02x)"` —
lowercase with a `0x` prefix per component. Both are direct `printf`-style hex
formats, **not** decimal serialization, so they correctly do **not** route
through a `pyStr` numeric helper (H14 governs `str()`/`repr()` of numbers).
`char` returns the same value as `str_` as a char row vector for display.

**`from_string`.** `RGBColor.from_string("3C2F80")` slices the 6-char hex string
into three base-16 byte values (Python `int(rgb_hex_str[:2], 16)` etc.). A
non-hex slice raises `mat2doc:ValueError` with CPython's exact message
(`invalid literal for int() with base 16: '...'`); `hex2dec` is base MATLAB
(no toolbox) and case-insensitive, matching `int(_, 16)`. For a live 6-hex
`w:color/@val` the two implementations agree byte-for-byte; a `<4`-char input
raises (identifier differs from Python's `ValueError`, dead path, D-004 shape).

**Example**

```matlab
c = mat2doc.shared.RGBColor(60, 47, 128);
c.str_()                                 % "3C2F80"
c.repr_()                                % "RGBColor(0x3c, 0x2f, 0x80)"
d = mat2doc.shared.RGBColor.from_string("3C2F80");
c == d                                   % true (value equality)
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::RGBColor`*

---

## `pyIntArg` (package-private)

**Syntax**

```matlab
d = pyIntArg(value, strMode)   % strMode = "parse" | "reject"
```

**Description**

Package-private helper (`+shared\private\pyIntArg.m`) — not part of the public
API; callable only from the seven `Length`-family constructors. Replicates the
CPython 3.13 `int()`-argument domain: numeric → `double` (constructor applies
`fix`); logical → 0/1; `"parse"` mode (`Length`, `Emu`) parses a base-10 int
literal per `int(str)` and raises `mat2doc:ValueError` on a bad literal with
CPython's exact message; `"reject"` mode (the five multiplier constructors)
raises `mat2doc:ValueError` for any string (Python reaches `int(str * K)`,
string repetition — D-003); anything else → `mat2doc:TypeError`; non-finite →
`MATLAB:validators:mustBeFinite` (D-004). ASCII-only digit grammar (D-002). The
design is carried from the Mat2Ppt `+util\private\pyIntArg` — **re-implemented,
no shared code**, `mat2doc:`-namespaced.

**Example**

```matlab
% exercised indirectly:
mat2doc.shared.Emu("914400")   % 914400  (Python int('914400'))
```

*Mat2Doc infrastructure (shared package-private helper), no python-docx
counterpart; replicates CPython 3.13 `int()` construction semantics.*
