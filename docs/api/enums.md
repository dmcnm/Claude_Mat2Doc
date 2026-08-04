---
title: "mat2doc.enum.base — the enumeration base tier"
---

# `mat2doc.enum.base` — the enumeration base tier

Ported from python-docx v1.2.0 `src/docx/enum/base.py` (the `BaseEnum` /
`BaseXmlEnum` base classes, package `+mat2doc/+enum/+base/`), plus the two
equality roots realized in that package — **`BaseIntEnum`** (value-based `==`
for the int-enums) and **`BasePlainEnum`** (identity `==` for the two plain
`enum.Enum` ports). This is the
**base machinery every docx enumeration extends** — the MS-API-value enums
(`BaseEnum`) and the XML-attribute-mapping enums (`BaseXmlEnum`) consumed by
**P3-3 / P3-4** (the concrete `WD_*` enums) and by every **P4 / P5 / P6**
element class that (de)serializes an enumerated attribute. It emits **no
serialized output of its own** — it supplies the value/name/`str` surface, the
`from_xml` / `to_xml` translation, and the enum-with-associated-data idiom that
concrete enums declare against.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## Where this differs from python-pptx (the design guide)

python-docx v1.2.0 is the **source of truth**; Mat2Ppt supplied only the MATLAB
*idiom* (the value-classdef + `enumeration` block realization). The docx base
enums are **NOT the pptx base enums** — the boundary audit
(`audit_P3-1_enum_base.md`) found four semantic deltas, each ported to the docx
form (**do not copy the Mat2Ppt `BaseXmlEnum` verbatim**):

| # | Aspect | python-pptx | python-docx v1.2.0 (SoT) | Resolution in Mat2Doc |
|---|---|---|---|---|
| 1 | `from_xml` None/empty guard | short-circuits (`if xml_value else None` → raise) | **no guard** — a straight None-tolerant equality scan | `from_xml_` scans every member with `xmlEq_`; NO guard on the query |
| 2 | `to_xml` reject test | `if xml_value is None` | **`if not xml_value`** (Python falsy — catches None **and** `""`) | `ismissing(xv) \|\| strlength(xv)==0` |
| 3 | `validate()` classmethod | present | **absent** | not ported (no `validate_`/`isMember_`) |
| 4 | bad-value / no-mapping message | uses `repr()` | **str-interpolation** for the docx-authored messages | `%s` with `queryStr_`; the stdlib `cls(value)` message keeps `%r` (F1) |

### Delta 1 — `from_xml` has NO None short-circuit (the load-bearing delta)

docx `from_xml` (base.py:63-66) is a plain equality scan:

```python
member = next((m for m in cls if m.xml_value == xml_value), None)
if member is None:
    raise ValueError(f"{cls.__name__} has no XML mapping for '{xml_value}'")
return member
```

There is **no** `if xml_value` short-circuit (the pptx one is absent), so a
member whose `xml_value is None` is **reachable** via `from_xml(None)`, and a
member whose `xml_value == ""` is reachable via `from_xml("")`. This is
**load-bearing in real docx**: `WD_COLOR_INDEX.INHERITED` (text.py:101) and
`WD_UNDERLINE.INHERITED` (text.py:282) both carry `xml_value=None` — the pptx
guard would raise on loading them. The port's `from_xml_` reproduces the scan
with a None-tolerant comparison (`xmlEq_`) and **no guard on the query**.

### Delta 2 — `to_xml` rejects falsy (None **and** `""`)

docx `to_xml` (base.py:73-77): `member = cls(value); if not xml_value: raise`.
The Python-falsy guard catches **both** a `None` `xml_value` and an empty-string
`""` one. The port raises when `ismissing(xv) || strlength(xv) == 0`, so a member
with either a `<missing>` (None) or `""` `xml_value` raises
`<Cls>.<NAME> has no XML representation`.

## The MATLAB enum-with-associated-data idiom

Python's base enums are `int` subclasses whose members carry an integer `value`
(`_value_ = ms_api_value`), a docstring, and — for `BaseXmlEnum` — an
`xml_value`. A MATLAB `enumeration` class that subclasses a numeric builtin
**cannot add properties**, and MATLAB static methods have **no `cls` binding**,
so the design (design.md §2) realizes each enum in two parts:

- A **plain value classdef** base (`BaseEnum` / `BaseXmlEnum` — *not* a `handle`;
  enum members are value objects) that owns the immutable properties `value`
  `(1,1) int32`, `doc` `(1,1) string`, and (XML variant) `xml_value`
  `(1,1) string`, plus a pass-through constructor.
- A **concrete enum** that subclasses the base and declares an `enumeration`
  block; each member passes `(ms_api_value, [xml_value,] docstr)` to the
  constructor:

```matlab
classdef WD_TAB_ALIGNMENT < mat2doc.enum.base.BaseXmlEnum   % illustrative (P3-3/P3-4)
    enumeration
        LEFT   (0, "left",   "Left-aligned tab stop.")
        CENTER (1, "center", "Center-aligned tab stop.")
        % ...
    end
    methods (Static)
        function m = from_xml(xml_value)
            m = mat2doc.enum.base.BaseXmlEnum.from_xml_("mat2doc.enum.WD_TAB_ALIGNMENT", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_("mat2doc.enum.WD_TAB_ALIGNMENT", value);
        end
    end
end
```

`from_xml` / `to_xml` are shared **Hidden static helpers** on the base taking the
concrete class name (`fullClassName`); each concrete subclass exposes thin public
statics that forward with **their own name hardcoded** (the explicit,
no-metaclass, no-runtime-dispatch pattern). Members are enumerated via
`enumeration(fullClassName)`; member-lookup-by-value (`cls(intValue)`, the
`to_xml` path) ports as `find(double([members.value]) == v, 1)` —
first-declared-wins, matching Python enum alias resolution.

**`int(member)` sites port as `double(member.value)`.** MATLAB enum members are
not numeric, so an explicit `int(member)` in the library port keeps the
`double(member.value)` spelling at the call site (design.md §2).

**Enum `==` / `~=` is VALUE-BASED (the `BaseIntEnum` root).** python-docx's
`BaseEnum(int, enum.Enum)` and `BaseXmlEnum(int, enum.Enum)` are **int
subclasses**, so their members compare by the MS-API integer value. The port
replicates this: both bases now derive from the shared root
`mat2doc.enum.base.BaseIntEnum`, whose inherited `eq`/`ne` compare
`double(value)`. See the [`BaseIntEnum`](#baseintenum) section below for the full
operand matrix and the `string(member) == "NAME"` name-compare idiom that this
mandates. Plain `enum.Enum` ports (`WD_BREAK_TYPE`, `WD_INLINE_SHAPE_TYPE`) do
**not** derive from `BaseIntEnum`; they derive from the identity sibling
[`BasePlainEnum`](#baseplainenum) (Python plain-Enum **identity** `==`).

### H3 tri-state — None vs `""` vs `<missing>`

Python types `xml_value` as `str | None`; the two are distinct values, and `""`
is a third, distinct real value. The port stores Python `None` as a MATLAB
`<missing>` string and Python `""` as the real empty string `""`:

| Python `xml_value` | MATLAB store | `from_xml` query that matches | `to_xml` result |
|---|---|---|---|
| `None` | `<missing>` (`ismissing` true) | `from_xml(None)` / `from_xml([])` / `from_xml(missing)` | raises `…has no XML representation` |
| `""` | `""` (`strlength` 0, not missing) | `from_xml("")` | raises `…has no XML representation` |
| `"center"` | `"center"` | `from_xml("center")` | `"center"` |

`asXmlVal_` normalizes on both store (ctor) and query: the `[]` double None
sentinel and a `<missing>` string both become `<missing>`; a `''`/`""`/token
becomes a real string. `xmlEq_` gives Python `==` (None==None → true; None vs
string → false; else string equality). This distinctness is the central hazard
of the WP — a `from_xml(None)` matches only a `<missing>` member and a
`from_xml("")` matches only a `""` member.

## Deviation posture — 0 new D-numbers

Pure value/behaviour machinery; **no serialized OOXML**, so no L0–L3 ladder leg
applies. Every emitted error uses the standing **D-005** (adopt-only)
`mat2doc:ValueError` identifier with byte-verbatim message strings — a recurring
convention needing no new sign-off. Gate-3 proved **34/34 probe facts
byte-identical** (`probe_diff` exit 0) against python-docx's own `enum/base.py`,
including the non-ASCII (`café`) UTF-8 leg and the F1 repr-quote leg. **No new
D-number, no new ledger row.** (`D-STYPE` is a P3-2 simpletypes concern, not
this WP.) The `DocsPageFormatter` class (base.py:80-151) is **not ported** — it
generates a ReStructuredText doc page for an enum and is docs-tooling with no
runtime consumer (grep-confirmed; Mat2Ppt **VERIFY-E4** precedent). The `doc`
property is still stored on every member for fidelity but is non-behavioral.

---

(baseintenum)=
## `BaseIntEnum`

**Syntax**

```matlab
% Shared root of BaseEnum / BaseXmlEnum — never instantiated or subclassed directly.
% Gives every int-enum member value-based == / ~= (Python int-subclass equality).
memberA == memberB     % double(a.value) == double(b.value)  (cross-class included)
member  == 1           % value compare (Python int(member) == 1)
member  == "CENTER"    % false           (Python int != str)
string(member) == "CENTER"   % the NAME-compare idiom (Python member.name == "CENTER")
```

**Description**

`BaseIntEnum` is the value-class root introduced (2026-08-03) to replicate
python-docx's **int-subclass equality**. In python-docx, `BaseEnum(int,
enum.Enum)` and `BaseXmlEnum(int, enum.Enum)` construct their members with
`int.__new__(cls, ms_api_value)` (`base.py:15,23,33,43`) — the members **are
ints**, so Python compares them by their MS-API integer value. MATLAB
`enumeration` classes compare `==` by **member identity** by default, which made a
cross-class comparison silently false and a member-vs-int comparison an
error/false. `BaseIntEnum` restores the Python semantics by overriding `eq`/`ne`
so that every `BaseEnum`/`BaseXmlEnum` member (17 concrete int-enums + 11 aliases)
compares by `double(value)`. Because a single inherited `eq` lives at the shared
root, two members of **any** two int-enum classes — same base or the two different
bases — dispatch to the same method with no ambiguity.

**Operand matrix** (element-wise, native `==` broadcasting):

| Right operand | Result vs an int-enum member | Python parity |
|---|---|---|
| int-enum member (any `BaseEnum`/`BaseXmlEnum`) | `double(a.value) == double(b.value)` — **cross-class equal-by-value → true** | `int == int` |
| numeric (`double`/`int32`/…) | value compare (`== 1` → true iff `value == 1`) | `int == int` |
| `logical` | value compare (`CENTER(1) == true`, `LEFT(0) == false`) | Python `bool` is an `int` subclass |
| `string` / `char` | **false** | `int == str` → False |
| `[]` (None sentinel) / `missing` | **false** (`~=` → true); the `[]` sentinel is treated as a **scalar** | `member == None` → False |
| genuinely empty enum/numeric **array** (`X.empty`) | native empty logical (broadcasting preserved) | n/a (MATLAB extension) |

A NaN comparison-vector encodes every non-numeric/None operand: `NaN == anything`
is false (even `NaN == NaN`), which is exactly "an int is never equal to a
str/None". `isequal` (hence `verifyEqual`) is **not** loosened — value-object
`isequal` stays class-strict.

**The name-compare idiom flips to `string(member) == "NAME"`.** Because
`member == "NAME"` is now **false** for an int-enum, compare a member's name via
`string(member) == "NAME"` (Python `member.name == "NAME"`) or its value via
`double(member.value) == N`. `switch member` now routes through the overridden
`eq` (a `case OtherClass.CENTER` matches cross-class by value; a `case "CENTER"`
string case is dead) — faithful to Python `if/elif ==`, and no library or test
code switches on an int-enum.

**Plain-enum identity `==` (`BasePlainEnum`) — residual CLOSED.**
`WD_BREAK_TYPE` and `WD_INLINE_SHAPE_TYPE` are declared `class X(enum.Enum)` in
python-docx — **not** int subclasses — so they compare by **identity** and are
**not** equal to their int. They do not derive from `BaseIntEnum`; instead they
derive from the dedicated identity sibling root
[`BasePlainEnum`](#baseplainenum) (2026-08-03), whose `eq`/`ne` make a member
equal ONLY to the same member of the same plain-enum class. So
`WD_BREAK_TYPE.LINE == "LINE"` (and `== 'LINE'`, `== 6`, a cross-class member,
`[]`/None, `missing`) is now **false** — matching python-docx exactly. **The
enum-vs-string residual named by the value-eq WP is CLOSED for Mat2Doc**: with
int-enums value-based (`BaseIntEnum`) and plain enums identity-based
(`BasePlainEnum`), no known enum-`==` divergence from python-docx remains. Use
`string(member) == "NAME"` for a name compare (Python `member.name == "NAME"`),
exactly as for the int-enums. (Mat2Ppt's own plain-enum `PROG_ID` residual is
handled by a sibling WP, in progress.)

**Byte-neutral, 0 D-number.** Only `eq`/`ne` were added; `value`, `xml_value`,
`from_xml`, `to_xml` and every serialization path are untouched, so saved `.docx`
output is unchanged. This change **resolves the A2 §2 cross-enum divergence** (see
[Table API](table_api.md) — `Table.alignment == WD_TABLE_ALIGNMENT.CENTER` now
returns **true**, matching python-docx). The identical change is replicating to
Mat2Ppt in a sibling WP (in progress); shipped Mat2Ppt (M3) still has the old
class-scoped identity `==`.

**Example**

```matlab
import mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT
import mat2doc.enum.table.WD_TABLE_ALIGNMENT
WD_PARAGRAPH_ALIGNMENT.CENTER == WD_TABLE_ALIGNMENT.CENTER    % true  — cross-class equal-by-value (the fix)
WD_PARAGRAPH_ALIGNMENT.JUSTIFY == WD_TABLE_ALIGNMENT.CENTER   % false — value 3 ~= value 1
WD_PARAGRAPH_ALIGNMENT.CENTER == 1                            % true  — value compare
WD_PARAGRAPH_ALIGNMENT.CENTER == "CENTER"                     % false — int != str
string(WD_PARAGRAPH_ALIGNMENT.CENTER) == "CENTER"            % true  — the name-compare idiom
WD_PARAGRAPH_ALIGNMENT.CENTER == []                           % false (scalar); ~= [] -> true
```

*Ported from python-docx v1.2.0: `src/docx/enum/base.py::BaseEnum` /
`BaseXmlEnum` (the `int`-subclass equality of `int.__new__(cls, ...)`)*

---

(baseplainenum)=
## `BasePlainEnum`

**Syntax**

```matlab
% Identity sibling of BaseIntEnum — the root of the two PLAIN enum.Enum ports.
% Never instantiated or subclassed except by WD_BREAK_TYPE / WD_INLINE_SHAPE_TYPE.
memberA == memberB            % same member of the same plain-enum class -> true
WD_BREAK_TYPE.LINE == "LINE"  % false   (Python plain Enum: member != str)
WD_BREAK_TYPE.LINE == 6       % false   (plain Enum is NOT an int subclass)
string(member) == "LINE"      % the NAME-compare idiom (Python member.name == "LINE")
```

**Description**

`BasePlainEnum` is the value-class root introduced (2026-08-03) to give the two
plain `enum.Enum` ports **Python identity** `==` / `~=`. python-docx declares
`WD_BREAK_TYPE` (`enum/text.py:70`) and `WD_INLINE_SHAPE_TYPE` (`enum/shape.py:6`)
as `class X(enum.Enum)` — **not** int subclasses — so a member compares by
IDENTITY: equal only to the same member, never to a string, its own int, another
class's member, or `None`. These two were correctly excluded from `BaseIntEnum`
(they are not ints), but MATLAB's built-in `enumeration` `==` compared a member
to a string by NAME, so `WD_BREAK_TYPE.LINE == "LINE"` used to return **true** —
the last enum-`==` divergence from python-docx. `BasePlainEnum` overrides
`eq`/`ne` so a member equals ONLY a member of the SAME plain-enum class with the
SAME name (Python member identity).

**Operand matrix** (element-wise, native `==` broadcasting):

| Right operand | Result vs a plain-enum member | Python parity |
|---|---|---|
| same member, same class | **true** | `member is member` |
| other member / other plain-enum class (even equal `value`) | **false** | `member is other` → False |
| `string` / `char` (e.g. `"LINE"`) | **false** — THE FIX (was true) | `Enum member == str` → False |
| numeric / `logical` (incl. its own value `6`) | **false** | plain Enum is **not** an int |
| `[]` (None sentinel) / `missing` | **false** (`~=` → true); `[]` treated as a **scalar** | `member == None` → False |
| genuinely empty enum/string **array** | native empty logical (broadcasting preserved) | n/a (MATLAB extension) |

Within one plain-enum class member names are unique (an alias such as
`WD_BREAK_TYPE.TEXT_WRAPPING` resolves to the canonical `LINE_CLEAR_ALL` member,
name `"LINE_CLEAR_ALL"`), so `string(a) == string(b)` **is** member identity. The
cross-class and non-member cases map both operands to NaN comparison vectors
(`NaN == NaN` is false) — exactly "a plain Enum member is never equal to a
non-member". `ne = ~eq` is the literal negation.

**Divergence from the int-enum sibling (intentional).** For a numeric operand
`BaseIntEnum` returns value equality (`CENTER == 1` → true) whereas
`BasePlainEnum` returns **false** (a plain Enum is not its int); for same-class
members `BaseIntEnum` compares `double(value)` whereas `BasePlainEnum` compares
the NAME (identity), so two members that share a `value` only via an alias still
resolve to one canonical name and cross-class value collisions never match.

**Residual CLOSED.** With int-enums value-based (`BaseIntEnum`) and plain enums
identity-based (`BasePlainEnum`), **no known enum-`==` divergence from
python-docx remains for Mat2Doc** — the enum-vs-string residual named by the
value-eq WP is retired. Use `string(member) == "NAME"` for a name compare, the
same idiom as the int-enums.

**Byte-neutral, 0 D-number.** Only `eq`/`ne` were added; the `value` property,
member declarations, constructors and the `TEXT_WRAPPING` alias are untouched,
and plain enums carry no `xml_value`/`from_xml`/`to_xml` serialization path, so
saved `.docx` output is unchanged. A faithful side effect: `Run.add_break("PAGE")`
now raises `mat2doc:KeyError` like python-docx (the old name-compare silently
accepted the string). Mat2Ppt's own plain-enum `PROG_ID` residual is handled by a
sibling WP (in progress).

**Example**

```matlab
import mat2doc.enum.text.WD_BREAK_TYPE
import mat2doc.enum.shape.WD_INLINE_SHAPE_TYPE
WD_BREAK_TYPE.LINE == WD_BREAK_TYPE.LINE                 % true  — same member
WD_BREAK_TYPE.LINE == WD_BREAK_TYPE.PAGE                 % false — other member
WD_BREAK_TYPE.LINE == "LINE"                             % false — Python member != str (THE FIX)
WD_BREAK_TYPE.LINE == 6                                  % false — plain Enum is not its int
WD_BREAK_TYPE.TEXT_WRAPPING == WD_BREAK_TYPE.LINE_CLEAR_ALL  % true  — alias is the canonical member
string(WD_BREAK_TYPE.LINE) == "LINE"                     % true  — the name-compare idiom
WD_BREAK_TYPE.LINE == WD_INLINE_SHAPE_TYPE.PICTURE       % false — cross plain-enum class
WD_BREAK_TYPE.LINE == []                                 % false (scalar); ~= [] -> true
```

*Ported from python-docx v1.2.0: `src/docx/enum/text.py::WD_BREAK_TYPE` /
`src/docx/enum/shape.py::WD_INLINE_SHAPE_TYPE` (plain `enum.Enum` identity `==`)*

---

## `BaseEnum`

**Syntax**

```matlab
% Never instantiated directly — subclassed by a concrete enumeration.
double(member.value)   % the MS API integer value  (Python int(member))
string(member)         % the member NAME           (Python member.name)
member.str_()          % "NAME (value)"            (Python str(member))
```

**Description**

Base class for enumerations that do **not** map an XML attribute value —
MS-API-only enums whose members carry an integer `value` (the constant assigned
the same-named member in the Microsoft API enum) and a docstring. A plain
**value class** (not `handle`); a concrete enum subclasses it with an
`enumeration` block passing `(ms_api_value, docstr)` per member.

`double(member.value)` is the Python `int(member)`; `string(member)` is the
enumeration member name (Python `member.name`, deliberately **not** overridden so
the name stays directly accessible); `member.str_()` is Python `str(member)` =
`f"{name} ({value})"`. The constructor `.strip()`s the docstring.

**Example**

```matlab
% Exercised through a concrete BaseEnum subclass (never instantiated directly).
% If WD_SECTION_START were such a subclass:
a = mat2doc.enum.WD_SECTION_START.NEW_PAGE;
double(a.value)   % 2            (Python int(WD_SECTION_START.NEW_PAGE))
string(a)         % "NEW_PAGE"   (member name, Python self.name)
a.str_()          % "NEW_PAGE (2)"   (Python str(member))
```

*Ported from python-docx v1.2.0: `src/docx/enum/base.py::BaseEnum`*

---

## `BaseXmlEnum`

**Syntax**

```matlab
% Concrete subclass forwards to these; never instantiated directly.
m = <Enum>.from_xml(xml_value)   % member whose xml_value == xml_value (None-tolerant)
s = <Enum>.to_xml(value)         % XML attribute string for a member / its int
double(m.value)                  % MS API integer value
m.xml_value                      % the XML attribute string ("" or <missing> possible)
```

**Description**

Base class for enumerations that **also** map an XML attribute value. Each member
carries the MS API integer `value` **plus** an `xml_value` typed `str | None` in
Python — a real token, the empty string `""`, or `None` (stored here as a
`<missing>` string; see H3 above). Concrete subclasses pass
`(ms_api_value, xml_value, docstr)` per member and expose thin static
`from_xml` / `to_xml` forwarders.

**`from_xml(xml_value)`** — returns the member whose `xml_value` equals the
query, with None-tolerant comparison and **no short-circuit guard** (Delta 1).
`from_xml(None)` returns the `<missing>`-`xml_value` member if one exists (else
raises); `from_xml("")` returns the `""`-`xml_value` member if one exists. No
match raises `mat2doc:ValueError`:
`<Cls> has no XML mapping for '<query>'` — str-interpolated (None renders the
literal `None`, `""` renders empty), **not** `repr`.

**`to_xml(value)`** — `value` may be a member, its integer, or `None`. It
resolves the member (`cls(value)`: identity for a member, value-lookup for an
int) and returns its `xml_value`, **unless** that is falsy (`<missing>` **or**
`""` — Delta 2), which raises
`<Cls>.<NAME> has no XML representation`. A bad integer raises
`<v> is not a valid <Cls>`; `to_xml(None)` raises `None is not a valid <Cls>`.

**F1 (repr quotes on the stdlib message).** CPython's `EnumCls(value)` raises
`%r is not a valid %s`; for a **string** value the `%r` yields single quotes.
The port's `valueStr_` string branch reproduces this (`'center' is not a valid
<Cls>`), while int / None / member inputs render unquoted (`999` / `None` /
`NAME (value)`). Gate-3 confirmed the string branch byte-equal to CPython.

**No `validate()`** (Delta 3) — docx v1.2.0 has no such classmethod, so the port
has none (`methods` lists only `BaseXmlEnum` and `str_`).

**Example**

```matlab
% Exercised through a concrete BaseXmlEnum subclass (never instantiated directly).
% If WD_TAB_ALIGNMENT were such a class:
m = mat2doc.enum.WD_TAB_ALIGNMENT.from_xml("center");
string(m)                                     % "CENTER"
double(m.value)                               % 1
mat2doc.enum.WD_TAB_ALIGNMENT.to_xml(m)       % "center"
mat2doc.enum.WD_TAB_ALIGNMENT.to_xml(1)       % "center"  (by int value)
% from_xml(None) -> the xml_value=None member if the enum declares one (e.g. INHERITED),
% else mat2doc:ValueError "…has no XML mapping for 'None'".
```

*Ported from python-docx v1.2.0: `src/docx/enum/base.py::BaseXmlEnum`*
