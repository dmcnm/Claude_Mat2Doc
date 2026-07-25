---
title: "Python → MATLAB mapping"
---

# Python → MATLAB mapping

Two living reference tables for the Mat2Doc port: **module → package** (where
each python-docx source module's symbols land in the `+mat2doc` namespace) and
**dunder / idiom** (how recurring Python constructs are rendered in base
MATLAB R2024b). Rows are appended per work package; this page starts with
**P1-1** (`docx/shared.py`).

## Module → package

| python-docx module `src/docx/...` | Mat2Doc package / files | WP | notes |
|---|---|---|---|
| `shared.py` — `Length`, `Inches`, `Cm`, `Emu`, `Mm`, `Pt`, `Twips`, `RGBColor` | `+mat2doc\+shared\` (one file per class) + `+shared\private\pyIntArg.m` | P1-1 | `Length < double` (int-subclass degradation); `RGBColor` is a value class (docx home package, not `dml/color.py`). `Twips` present / `Centipoints` absent per docx v1.2.0. `pyIntArg` is project infrastructure (no Python counterpart). |

**Not ported from `shared.py` in P1-1** (proxy tier, deferred to P2-1):
`lazyproperty`, `write_only_property`, `ElementProxy`, `Parented`,
`StoryChild`, `TextAccumulator`.

## Dunder / idiom

| Python idiom | MATLAB rendering | first WP |
|---|---|---|
| `class Length(int)` (int subclass; arithmetic degrades to a plain `int`) | `classdef Length < double` — defining properties forces every op to return a plain `double`; `isa(x,'double')` true; EMU held exactly (magnitudes ≪ 2^53, design.md §8) | P1-1 |
| `int.__new__(cls, emu)` | `obj@double(fix(emu)+0)` after `pyIntArg(emu,"parse")`; the `+0` normalizes IEEE `-0.0 → +0.0` | P1-1 |
| `int(x)` on a constructor argument | `fix(x)` (truncate toward zero — H6) after domain coercion by `pyIntArg` | P1-1 |
| `self / float(_EMUS_PER_*)` (true division) | `double(obj) / <Const>` (dependent-property getter) | P1-1 |
| `int(round(self / float(635)))` (Python-3 `round()` = half-to-even, returns int) | static `pyRoundHalfToEven_` replicating CPython `float.__round__` (`r=round(x); if abs(x-r)==0.5, r=2*round(x/2)`) + `+0` signed-zero normalize — H6/H14 | P1-1 |
| `@property` `emu` `return self` | `value = obj` (the getter returns the `Length` instance itself) | P1-1 |
| private class attrs `_EMUS_PER_*` | `Constant` properties `EMUS_PER_*_` (leading underscore rotated to trailing, design.md §2) | P1-1 |
| `class RGBColor(Tuple[int,int,int])` (immutable tuple, equality-by-value) | value `classdef` with immutable `r/g/b` properties + by-value `eq`/`ne`; comparison to a non-`RGBColor` is `false` | P1-1 |
| `__str__` / `__repr__` = `"%02X..."` / `"RGBColor(0x%02x...)"` | `str_()` / `repr_()` returning a `string` via `sprintf` (dunders rotate to trailing-underscore methods) — direct `%X`/`%x` hex, **not** routed through `pyStr` (H14) | P1-1 |
| `int(rgb_hex_str[:2], 16)` etc. | `parseHexByte_(s(1:2))` → validated `hex2dec` (base MATLAB, case-insensitive) | P1-1 |
| `raise TypeError(msg)` / `raise ValueError(msg)` | `error("mat2doc:TypeError", "%s", msg)` / `error("mat2doc:ValueError", "%s", msg)` — one identifier per Python exception type (D-004 namespace) | P1-1 |
| `isinstance(val, int)` where a MATLAB `double` cannot distinguish int from float | integer-**valued** real numeric scalar accepted as the Python int (D-STYPE-1); non-integral → `TypeError` | P1-1 |

## Namespace policy

Mat2Doc uses the `mat2doc:` error-identifier namespace and the `mat2doc.*`
package namespace throughout — **no shared code with Mat2Ppt**. Designs common
to both toolboxes (the `Length` family, `pyIntArg`) are re-implemented, not
copied. The P0 same-package name-collision scan found **zero** collisions, so
**no symbol was renamed** (FLAG-3-docx resolved by convention).
