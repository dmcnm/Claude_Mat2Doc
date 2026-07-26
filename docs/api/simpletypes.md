---
title: "mat2doc.oxml.simpletypes — the ST_* attribute-value validator tier"
---

# `mat2doc.oxml.simpletypes` — the simple-type validator tier

Ported from python-docx v1.2.0 `src/docx/oxml/simpletypes.py` (35 classes,
package `+mat2doc/+oxml/+simpletypes/`, plus two package-private helpers). A
**simple type** validates and converts the string form of a single XML
attribute value: it turns the on-the-wire text into a MATLAB value on read
(`convert_from_xml`) and a MATLAB value back into the on-the-wire text on write
(`convert_to_xml`), rejecting out-of-domain values (`validate`). Every P4 / P5 /
P6 element class that reads or writes a typed `w:` attribute delegates to one of
these classes through the `BaseOxmlElement.getAttrTyped` / `setAttrTyped`
engine. This tier **emits no serialized output of its own** — it is the value
layer beneath the element tree.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The Base / Xsd / ST_ hierarchy

The 35 classes form three tiers. The **Base** tier holds the abstract machinery;
the **Xsd** tier ports the XSD primitive types; the **ST_** tier ports the 21
WordprocessingML `ST_*` attribute types the docx element classes actually name.

| Tier | Classes | Role |
|---|---|---|
| **Base (4)** | `BaseSimpleType`, `BaseIntType`, `BaseStringType`, `BaseStringEnumerationType` | abstract roots: the `validate_int` / `validate_string` / `validate_enum` primitives and the `from_xml` / `to_xml` templates |
| **Xsd (10)** | `XsdAnyUri`, `XsdBoolean`, `XsdId`, `XsdInt`, `XsdLong`, `XsdString`, `XsdStringEnumeration`, `XsdToken`, `XsdUnsignedInt`, `XsdUnsignedLong` | XSD primitives: string pass-throughs, the four bounded integer types, and the boolean |
| **ST_ (21)** | `ST_BrClear`, `ST_BrType`, `ST_Coordinate`, `ST_CoordinateUnqualified`, `ST_DateTime`, `ST_DecimalNumber`, `ST_DrawingElementId`, `ST_HexColor`, `ST_HexColorAuto`, `ST_HpsMeasure`, `ST_Merge`, `ST_OnOff`, `ST_PositiveCoordinate`, `ST_RelationshipId`, `ST_SignedTwipsMeasure`, `ST_String`, `ST_TblLayoutType`, `ST_TblWidth`, `ST_TwipsMeasure`, `ST_UniversalMeasure`, `ST_VerticalAlignRun` | the WordprocessingML attribute types |

Two package-private helpers sit under `+simpletypes/private/`:
`intFromXml` (Python `int(str_value)` → a plain double, via the audited
`Emu` string-parse path) and `pyRound` (CPython 3 round-half-to-even). Both are
**Mat2Doc infrastructure with no python-docx counterpart** (the `pyIntArg` /
`pyStr` precedent) and so carry an infrastructure header rather than a
`Ported from … ::symbol` line.

## The convert / validate contract

Every simple type presents the same four-method surface (some inherited, some
overridden):

- **`convert_from_xml(str_value)`** — parse the XML string into a MATLAB value
  (a double for the int types, a `string` for the string types, an `Emu`/`Length`
  for the measures, an `RGBColor` or `"auto"` for `ST_HexColor`, a `datetime`
  for `ST_DateTime`, a `logical` for the booleans).
- **`convert_to_xml(value)`** — format a MATLAB value back to its XML string
  (always through `mat2doc.shared.pyStr(…, "int")` at a numeric-serialize site,
  H14 — never `num2str`/`sprintf('%g')`; the sole exception is the `ST_HexColor`
  `"%02X%02X%02X"` hex format, which is not a decimal serialize).
- **`validate(value)`** — raise unless `value` is in the type's domain (range,
  membership, or Python-type check).
- **`from_xml` / `to_xml`** — the thin templates the engine calls; each routes to
  the class's own `convert_*` / `validate`.

### The H10 re-declaration rule (why `from_xml`/`to_xml` are repeated)

Python's `from_xml` / `to_xml` are `@classmethod`s that call
`cls.convert_from_xml` / `cls.validate`, which **late-bind to the most-derived
class** at call time. MATLAB static methods have no such late binding. The port
therefore **re-declares** `from_xml` on every class that overrides
`convert_from_xml`, and `to_xml` on every class that overrides `validate` or
`convert_to_xml`, so each routes to its own converter. Pass-through leaves
(`XsdString`, `ST_DecimalNumber`, `ST_String`, …) inherit the template from the
branch base (`BaseIntType` / `BaseStringType`) unchanged, which is correct
because they do not override any converter. `ST_OnOff` deliberately inherits
`XsdBoolean.to_xml` (it overrides only `convert_from_xml`). This re-declaration
matrix was verified class-by-class at Gate 2.

## D-STYPE recurrences (adopt-verify — ZERO new D-numbers)

P3-2 opened **no new deviation**. Every divergence maps to a ruling already
adopted for the Mat2Ppt simpletypes tier and pre-adopted `mat2doc:`-namespaced
for docx (`decision_2026-07-25_mat2doc_deviation_preadoption.md`). Gate 3 proved
**175/175 probe facts** against python-docx's own `simpletypes.py`: 154
byte-identical, 21 divergences each on an unreachable / schema-invalid input,
all dispositioned to an adopted ruling.

### D-STYPE-2 — the float-parse RE-HOME (no `BaseFloatType`)

python-pptx housed float parsing in a `BaseFloatType` base. **python-docx has no
`BaseFloatType`** — the only `float(...)` in the whole tier is
`ST_UniversalMeasure.convert_from_xml` (simpletypes.py:415), reused by
`ST_SignedTwipsMeasure` (:366). The port therefore homes the float parse **in
those two classes** and creates **no `BaseFloatType` class** (37 files = 35
classes + 2 helpers, confirmed). `float()` is realized with `str2double`. On
every well-formed OOXML measure literal the two agree exactly; the lexical
divergence on malformed input (underscore, thousands-comma) is folded under
D-STYPE-2 / D-002 / D-004 — see the [D-STYPE-2 explainer](../../../validation/summary/proofs/D-STYPE-2_explained.md)
(the docx-recurrence section records the three vectors below).

### D-STYPE-1 — int / float indistinguishability

`BaseSimpleType.validate_int` accepts any finite integral-**valued** real numeric
(and `Length`, a Python int subclass, and `logical`) where CPython
`isinstance(2.0, int)` is `False`. Flows to all four Xsd int types and their ST_
descendants. Divergence only for a caller programming error (`validate_int(2.0)`
accepts where CPython raises `TypeError`); every valid integer attribute is
byte-identical.

### D-STYPE-3 — the rounded long bound (dead upper edge)

`XsdLong.validate` (±2^63) and the docx-only `XsdUnsignedLong.validate`
(0..2^64−1) carry the range bound as a double, which is not exactly
representable at the 2^63 / 2^64 edge (`…808` vs `…807`, `…616` vs `…615`).
These are **dead upper edges**: every concrete subclass
(`ST_CoordinateUnqualified` / `ST_PositiveCoordinate` and the twips/hps
derive-counts) sits far below 2^53, where every double is an exact integer.

### D-STYPE-4 + the `ST_HexColor` / `ST_HexColorAuto` split

python-pptx had one `ST_HexColorRGB` taking and emitting a 6-hex string.
python-docx **splits it in two**, and the port mirrors the docx pair:

- **`ST_HexColor`** (`< BaseStringType`) — `convert_from_xml` returns an
  `RGBColor` for a hex value **or** the literal string `"auto"`;
  `convert_to_xml` formats an `RGBColor` via `"%02X%02X%02X"`; `validate`
  requires an `RGBColor`. Hex parsing is delegated to
  `mat2doc.shared.RGBColor.from_string`, which accepts exactly `[0-9A-Fa-f]`
  (the RGB contract) — narrower than CPython `int(_, 16)` (which also takes
  sign / `0x` / underscore / surrounding space). Those extra acceptances are
  non-RGB programming errors, folded under adopted **D-STYPE-4**. The docx
  class does **not** length-guard (the pptx 6-char guard was correctly *not*
  ported — it would be an added feature).
- **`ST_HexColorAuto`** — a separate one-member `XsdStringEnumeration`
  (`AUTO = "auto"`, `_members = ("auto",)`) supplying the constant `ST_HexColor`
  returns. Its bad-value message carries the singleton tuple repr
  `('auto',)`.

## ST_UniversalMeasure unit table

`ST_UniversalMeasure.convert_from_xml` splits the value into a numeric prefix and
a **two-character** unit suffix (`str_value[:-2]` / `str_value[-2:]`), parses the
prefix as a float (D-STYPE-2 re-home), multiplies by the EMU-per-unit factor,
and wraps `Emu(int(round(quantity * multiplier)))` (round-half-to-even):

| Unit | EMU per unit | `convert_from_xml` example → EMU |
|---|---|---|
| `mm` | 36 000 | `"10mm"` → 360 000 |
| `cm` | 360 000 | `"2.5cm"` → 900 000 |
| `in` | 914 400 | `"1in"` → 914 400 |
| `pt` | 12 700 | `"12pt"` → 152 400 |
| `pc` | 152 400 | `"1pc"` → 152 400 |
| `pi` | 152 400 | `"1pi"` → 152 400 |

An unknown unit raises `mat2doc:KeyError` with the unit quoted (`'xy'`) — the
Python dict-subscript `KeyError`. The multiplier lookup is a fixed `switch` (not
a `containers.Map`): the set is closed and order-irrelevant (no H11 concern).

## ST_OnOff — the docx-novel boolean

`ST_OnOff` (`< XsdBoolean`) accepts a wider truthy vocabulary than plain XSD
boolean. `convert_from_xml` accepts exactly `{'1','0','true','false','on','off'}`
and returns `true` for `{'1','true','on'}`, else raises **InvalidXmlError** (see
the raise-map below) with the verbatim message
`value must be one of '1', '0', 'true', 'false', 'on', or 'off', got '<x>'`
(the literal is split across two source lines; the concatenation is faithful).
It is **case-sensitive** (`'On'` raises). `validate` / `convert_to_xml` / `to_xml`
are inherited from `XsdBoolean` (`true`/`false` → `"1"`/`"0"`).

## Banker's rounding (round-half-to-even)

The two `int(round(...))` sites (`ST_UniversalMeasure`, `ST_SignedTwipsMeasure`)
route through `pyRound`, which replicates **CPython 3 round-half-to-even**.
MATLAB's built-in `round` rounds half *away from zero* and is a defect at the
`.5` boundary (H6). Verified vectors: `0.5` → `0`, `1.5` → `1`, `2.5` → `2`,
`-0.5` → `0`. `ST_SignedTwipsMeasure` `"12.5"` → 7620 (12 twips × 635) and
`"-360"` → −228600 confirm the tie-to-even path end-to-end.

## Raise-type map — InvalidXmlError vs ValueError vs TypeError vs KeyError

The port raises the **same exception class as the docx source, line for line**.
Identifiers are flat `mat2doc:<Name>`:

| Identifier | Raised at | Python source class |
|---|---|---|
| `mat2doc:InvalidXmlError` | `XsdBoolean.convert_from_xml` (:116), `ST_OnOff.convert_from_xml` (:340) — the **only two** sites | `docx.exceptions.InvalidXmlError` (routed through the canonical `mat2doc.exc.InvalidXmlError` raiser) |
| `mat2doc:ValueError` | enum membership, range, hex reject, `str2double`-parse | `ValueError` |
| `mat2doc:TypeError` | `validate_int` / `validate_string`, `XsdBoolean.validate`, `ST_DateTime.validate` | `TypeError` |
| `mat2doc:KeyError` | unknown `ST_UniversalMeasure` unit | dict-subscript `KeyError` |

The two `InvalidXmlError` sites are notable because they raise from
`docx.exceptions`, **not** `ValueError` — a docx design choice the port
preserves exactly (name and message byte-verbatim, confirmed at Gate 3).

## ST_DateTime — the fromisoformat under-accept (deferred to P8-2)

`ST_DateTime.convert_from_xml` parses an `xsd:dateTime` to a MATLAB `datetime`;
`convert_to_xml` normalizes to UTC and formats `%Y-%m-%dT%H:%M:%SZ`. All canonical
Word forms round-trip byte-identically, and a non-UTC offset input is stored as
the equivalent **UTC instant** (VERIFY-tz: the instant is compared, never the
zone label — since serialize normalizes to UTC, emitted bytes are identical).

The port implements a **narrow** `fromisoformat` subset (the canonical
`YYYY-MM-DD[T ]hh:mm:ss[.ffffff][±hh:mm|Z]` grammar real Word emits). CPython
3.11+ `datetime.fromisoformat` is far more lenient, so a handful of exotic inputs
that Python parses fall through to the **1970 epoch** in the port
(`2023-06-15T10:14` seconds-omitted; `2023-1-2T3:4:5Z` non-zero-padded; the
basic no-separator and ISO-week forms). Ground truth measured at Gate 3
**corrects** the earlier framing: Python *parses* these, the port *epochs* — a
genuine under-accept, not a mutual epoch fallback (the true mutual-epoch case is
`…T10:14Z`, secondless *with* the `Z`).

This is **SAFE and unreachable at P3-2**: all under-accepted inputs are lexically
**invalid `xsd:dateTime`** (the schema mandates zero-padded fields and seconds),
the failure mode is the epoch and never a raise or a plausible-wrong value, and
the only consumer of `ST_DateTime` is `w:comment/@w:date` — which is **P8-2
(comments), not yet ported**. No code path calls `ST_DateTime.convert_from_xml`
today. The final ruling is deferred to P8-2, where the subset is to be
re-verified against a corpus of real `w:date` values; see
`validation\summary\decision_2026-07-26_st_datetime_underaccept.md`. Same
"safe ASCII/canonical-subset under-accept" class as the D-002 Arabic-Indic-digit
case (`int("٥")` → 5 in Python, rejected by the port).

## Deviation posture — 0 new D-numbers

Pure value/behaviour machinery; **no serialized OOXML**, so no L0–L3 ladder leg
applies. Equivalence is proven by an exact behavioural `probe_diff` of every
returned value and raised identifier+message against the oracle: **175/175 probe
facts across all 35 classes**, 154 byte-identical, 21 adopt-verify divergences
(D-005 type-token ×7, D-STYPE-3 dead-edge ×4, D-STYPE-2/D-004/D-002 float-parse &
non-ASCII ×4, VERIFY-fromiso schema-invalid ×5, one dead-abstract-base
observation). ZERO divergences on any valid, reachable input; no ledger row
added. The `BaseSimpleType.from_xml`/`to_xml` templates are intentionally hosted
on the branch bases (`BaseIntType` / `BaseStringType`), so
`BaseSimpleType.from_xml("42")` raises rather than returning `int(42)` — a **dead
abstract-base path** (every concrete class descends via a branch base or
re-declares its own `from_xml`), noted for posterity, no D-number.

---

## `BaseSimpleType`, `BaseIntType`, `BaseStringType`, `BaseStringEnumerationType`

**Syntax**

```matlab
mat2doc.oxml.simpletypes.BaseSimpleType.validate_int(value)
mat2doc.oxml.simpletypes.BaseSimpleType.validate_int_in_range(value, lo, hi)
v = mat2doc.oxml.simpletypes.BaseSimpleType.validate_string(value)
```

**Description**

The abstract roots. `BaseSimpleType` owns the type-check primitives
(`validate_int`, `validate_int_in_range`, `validate_string`) — but **not** the
`from_xml`/`to_xml` templates, which the H10 rule hosts on the branch bases.
`BaseIntType` (`< BaseSimpleType`) hosts `from_xml`/`to_xml` for the integer
types and a default `convert_from_xml = int(str_value)` / `convert_to_xml =
str(value)`. `BaseStringType` does the same for the string types (identity
convert). `BaseStringEnumerationType` (`< BaseStringType`) adds `validate_enum`
(membership against a class `_members` tuple, with the Python tuple-repr
`ValueError` message). docx `BaseSimpleType` has **no** `validate_float` /
`validate_float_in_range` (the D-STYPE-2 re-home) — only the integer and string
validators are ported.

**Example**

```matlab
mat2doc.oxml.simpletypes.BaseSimpleType.validate_string("rId7")   % "rId7"
try
    mat2doc.oxml.simpletypes.BaseSimpleType.validate_int(3.5);
catch e
    disp(e.message)   % "value must be <type 'int'>, got double"
end
```

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py::BaseSimpleType`
(and `BaseIntType` / `BaseStringType` / `BaseStringEnumerationType`)*

---

## The Xsd tier — `XsdInt` / `XsdLong` / `XsdUnsignedInt` / `XsdUnsignedLong` / `XsdBoolean` / string pass-throughs

**Syntax**

```matlab
mat2doc.oxml.simpletypes.XsdInt.from_xml("42")                 % 42  (double)
mat2doc.oxml.simpletypes.XsdInt.to_xml(42)                     % "42"
mat2doc.oxml.simpletypes.XsdBoolean.from_xml("1")              % true
mat2doc.oxml.simpletypes.XsdBoolean.to_xml(true)              % "1"
```

**Description**

`XsdInt` (±2^31), `XsdLong` (±2^63, D-STYPE-3), `XsdUnsignedInt` (0..2^32−1) and
the docx-only `XsdUnsignedLong` (0..2^64−1, D-STYPE-3) are the four bounded
integer types: `< BaseIntType`, each re-declaring `to_xml` so `validate` resolves
its own range. `XsdBoolean` (`< BaseSimpleType`, owns all four methods) accepts
`{'1','0','true','false'}` → `true` for `{'1','true'}`, raising
**InvalidXmlError** otherwise; `to_xml([])` (the None analogue) raises
`mat2doc:TypeError` with the verbatim
`only True or False (and possibly None) may be assigned, got '<x>'` (the
Gate-2 **F1** render fix — the pre-fix `pyStr:unsupportedType` crash is gone).
`XsdAnyUri` / `XsdId` / `XsdString` / `XsdToken` / `XsdStringEnumeration` are
string pass-throughs (identity convert, inherited templates).

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py` (the `Xsd*`
classes)*

---

## `ST_UniversalMeasure` — measure-with-unit → EMU (the D-STYPE-2 re-home)

**Syntax**

```matlab
emu = mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml("1in")
```

**Description**

Parses a number plus a two-letter unit into an `Emu` via the EMU-per-unit table
above. Read-only in the source (only `convert_from_xml` is defined — no
`to_xml`), so the port provides `convert_from_xml` plus a `from_xml` template and
leaves `to_xml` intentionally undefined (calling it fails, as in Python). This is
the **sole home of the docx float parse** (`str2double`, D-STYPE-2 re-home). An
unknown unit raises `mat2doc:KeyError`.

**Example**

```matlab
double(mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml("1in"))    % 914400
double(mat2doc.oxml.simpletypes.ST_UniversalMeasure.convert_from_xml("2.5cm"))  % 900000
```

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py::ST_UniversalMeasure`*

---

## `ST_HexColor` / `ST_HexColorAuto` — the RGBColor-or-"auto" split

**Syntax**

```matlab
mat2doc.oxml.simpletypes.ST_HexColor.convert_from_xml("3c2f80")   % RGBColor(0x3c2f80)
mat2doc.oxml.simpletypes.ST_HexColor.convert_from_xml("auto")     % "auto"
mat2doc.oxml.simpletypes.ST_HexColor.convert_to_xml(rgb)          % "3C2F80"
```

**Description**

`ST_HexColor` returns an `RGBColor` for a hex value or the constant string
`"auto"` (supplied by `ST_HexColorAuto.AUTO`); `convert_to_xml` formats an
`RGBColor` as uppercase `"%02X%02X%02X"`; `validate` requires an `RGBColor`
(else `ValueError`). `ST_HexColorAuto` is a one-member `XsdStringEnumeration`
whose bad-value message carries the singleton tuple repr `('auto',)`. Hex
parsing accepts exactly `[0-9A-Fa-f]` (D-STYPE-4); no length guard (docx has
none).

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py::ST_HexColor`
(and `::ST_HexColorAuto`)*

---

## `ST_OnOff` — the six-token docx boolean

**Syntax**

```matlab
mat2doc.oxml.simpletypes.ST_OnOff.convert_from_xml("on")    % true
mat2doc.oxml.simpletypes.ST_OnOff.convert_from_xml("off")   % false
```

**Description**

`< XsdBoolean`, accepting `{'1','0','true','false','on','off'}` (case-sensitive),
`true` for `{'1','true','on'}`, else **InvalidXmlError**. `validate` /
`convert_to_xml` / `to_xml` inherited (`true`/`false` → `"1"`/`"0"`).

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py::ST_OnOff`*

---

## `ST_DateTime` — xsd:dateTime ↔ datetime

**Syntax**

```matlab
dt = mat2doc.oxml.simpletypes.ST_DateTime.convert_from_xml("2023-06-15T03:04:05Z")
s  = mat2doc.oxml.simpletypes.ST_DateTime.convert_to_xml(dt)   % "2023-06-15T03:04:05Z"
```

**Description**

Parses `xsd:dateTime` (narrow `fromisoformat` subset + strptime templates + a
`Z`-branch), storing a non-UTC offset as the equivalent UTC instant;
`convert_to_xml` normalizes to UTC before formatting; `validate` requires a
`datetime` (else `TypeError`). Exotic schema-invalid inputs fall to the 1970
epoch (the fromisoformat under-accept, deferred to P8-2). Unreachable until P8-2
(comments); see the ST_DateTime section above and the decision doc.

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py::ST_DateTime`*

---

## The measure and enum ST_ tiers (summary)

The remaining ST_ classes port verbatim member sets and measure grammars:

- **Measures** — `ST_Coordinate` / `ST_CoordinateUnqualified` /
  `ST_PositiveCoordinate` (EMU coordinates), `ST_TwipsMeasure` /
  `ST_SignedTwipsMeasure` (twips, half-to-even), `ST_HpsMeasure`
  (half-point: `Pt(int(str)/2.0)`). The unit-detection letter sets differ **by
  class** and are ported per the exact source letters (`ST_HpsMeasure`
  `'m'`/`'n'`/`'p'`; the twips/coordinate classes `'i'`/`'m'`/`'p'`).
- **Enumerations** — `ST_BrClear` (none/left/right/all), `ST_BrType`
  (page/column/textWrapping), `ST_Merge` (continue/restart), `ST_TblWidth`
  (auto/dxa/nil/pct), `ST_TblLayoutType` (fixed/autofit), `ST_VerticalAlignRun`
  (baseline/superscript/subscript). Each raises the Python tuple-repr
  `ValueError` byte-verbatim.
- **Pass-throughs** — `ST_String` / `ST_RelationshipId` / `ST_DecimalNumber` /
  `ST_DrawingElementId` (identity or inherited range).

*Ported from python-docx v1.2.0: `src/docx/oxml/simpletypes.py` (the ST_ measure,
enumeration, and pass-through classes)*
