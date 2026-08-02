---
title: "numbering — the numbering tier (CT_Num · CT_NumLvl · CT_NumPr · CT_Numbering · NumberingPart · _NumberingDefinitions)"
---

# `numbering` — the numbering-definitions tier (the P8-1 WP that opens Phase 8)

Ported from python-docx v1.2.0 `src/docx/oxml/numbering.py`
(`CT_Num` / `CT_NumLvl` / `CT_NumPr` / `CT_Numbering`),
`src/docx/parts/numbering.py` (`NumberingPart.numbering_definitions` /
`_NumberingDefinitions`; `NumberingPart.new`) and
`src/docx/parts/document.py::DocumentPart.numbering_part`. This is the **first
work package of Phase 8** — it lands the four numbering `CT_*` element classes
(the `+mat2doc\+oxml\+numbering` package), un-stubs
`NumberingPart.numbering_definitions` and `DocumentPart.numbering_part`, and
registers the **8** numbering tags — all **byte-neutral** for M1.

:::{important}
**★ Registering the 8 numbering tags is byte-neutral — M1 stays 17/17.** The
numbering registration flips **two** M1-central parts onto the new parse path:
`word/styles.xml` (its 7 `<w:numPr>` — 6 `<w:numId>` + 1 `<w:ilvl>` — now transit
`CT_NumPr`/`CT_DecimalNumber`) **and** `word/numbering.xml` (its root
`<w:numbering>` + 9 `<w:num>` now transit `CT_Numbering`/`CT_Num`). Registering a
content type changes only a parsed node's **class**, never its
content/attribute-order/child-order, so `mat2doc.Document().save()` is
byte-identical to python-docx — `styles.xml`
`02d71a68ddb92c055e84526d6a9e45700be3872781ababf759241694bc30e384` / 349458 B and
`numbering.xml`
`70976f19cbcd896e51890859fe6ecb3467a5a7ad0c040160fedc5a1993cb09ce` / 5513 B, both
byte-identical to the frozen `s0001` round-trip oracle. `<w:abstractNum>` is
**deliberately unregistered** (python-docx has no `register_element_cls`
for it), so it stays a generic `XmlElement` on both sides. **Zero new
D-numbers.**
:::

:::{note}
**The 8 registry rows** (`docx/oxml/__init__.py:105-112`, byte-neutral):
`w:abstractNumId` / `w:ilvl` / `w:numId` / `w:startOverride` →
`mat2doc.oxml.shared.CT_DecimalNumber` (**reused** from P4-6, not re-ported);
`w:lvlOverride` → `CT_NumLvl`; `w:num` → `CT_Num`; `w:numPr` → `CT_NumPr`;
`w:numbering` → `CT_Numbering`.
:::

---

## `CT_Numbering` — the `<w:numbering>` root

**Syntax**

```matlab
numbering = mat2doc.oxml.OxmlElement("w:numbering");   % registry -> CT_Numbering
num  = numbering.add_num(abstractNum_id);   % a new <w:num>, auto-assigned numId
lst  = numbering.num_lst;                    % the <w:num> children (document order)
hit  = numbering.num_having_numId(numId);    % the <w:num> with that @w:numId, else KeyError
n    = numbering.next_numId_;                 % _next_numId: first unused numId (>=1)
```

**Description**

`CT_Numbering` (`numbering.py::CT_Numbering`) is the root of a numbering part
(`word/numbering.xml`). Its one descriptor is
`num = ZeroOrMore("w:num", successors=("w:numIdMacAtCleanup",))`, so a new
`<w:num>` inserts **before** the first `<w:numIdMacAtCleanup>` (H11), else appends.

**`add_num(abstractNum_id)`** is the auto-numbering path: it allocates the next
free numId via `_next_numId`, builds a `CT_Num` (`CT_Num.new(next_num_id,
abstractNum_id)`) and inserts it in sequence — the generated public `add_num` is
**suppressed** by this explicit method (xmlchemy no-ops a generated member whose
name is already defined on the class). **`next_numId_`** (the `_next_numId`
`@property`, ported as a **Dependent** property per design.md §2) returns the
first numId ≥ 1 not already used by a `<w:num>`, **filling gaps** — it is data
arithmetic on the numId *values* (`range(1, N+2)` → `1:(N+1)`), never a 0/1 index
shift, and always resolves (pigeonhole: among `1..N+1` at least one value is
free). **`num_having_numId(numId)`** returns the first `<w:num>` whose `@w:numId`
matches (`xpath(...)[0]`), and raises `mat2doc:KeyError`
`no <w:num> element with numId <n>` on a miss (Python `IndexError` → `KeyError`).

**Example**

```matlab
numbering = mat2doc.oxml.OxmlElement("w:numbering");
n1 = numbering.add_num(0);          % <w:num w:numId="1"><w:abstractNumId w:val="0"/>
n2 = numbering.add_num(1);          % <w:num w:numId="2"><w:abstractNumId w:val="1"/>
fprintf('add_num,add_num -> numId %d, %d\n', n1.numId, n2.numId);
% add_num,add_num -> numId 1, 2

same = numbering.num_having_numId(2);
fprintf('num_having_numId(2) == n2 : %d\n', same == n2);   % 1 (same handle, H5)

try
    numbering.num_having_numId(99);
catch ME
    fprintf('%s : %s\n', ME.identifier, ME.message);
end
% mat2doc:KeyError : no <w:num> element with numId 99
```

The `_next_numId` gap-fill:

```matlab
gb = mat2doc.oxml.OxmlElement("w:numbering");
fprintf('empty -> %d\n', gb.next_numId_);           % empty -> 1
gb.add_num(0); gb.add_num(0); gb.add_num(0);         % numId 1, 2, 3
gb.remove(gb.num_having_numId(2));                   % now {1, 3}
fprintf('{1,3} -> %d\n', gb.next_numId_);            % {1,3} -> 2
```

*Ported from python-docx v1.2.0: `src/docx/oxml/numbering.py::CT_Numbering`*

---

## `CT_Num` — a `<w:num>` list-definition instance

**Syntax**

```matlab
num = mat2doc.oxml.numbering.CT_Num.new(num_id, abstractNum_id);  % <w:num w:numId=...>
aid = num.abstractNumId;         % the required <w:abstractNumId> child (InvalidXmlError if absent)
v   = num.numId;                 % the required @w:numId (a double int)
lo  = num.add_lvlOverride(ilvl); % a new CT_NumLvl <w:lvlOverride w:ilvl=...>
lst = num.lvlOverride_lst;       % the <w:lvlOverride> children (document order)
```

**Description**

`CT_Num` (`numbering.py::CT_Num`) is a concrete list-definition instance. It has a
**required** `<w:abstractNumId>` child referencing an abstract numbering
definition (`abstractNumId` = `OneAndOnlyOne`, so the getter raises
`mat2doc:InvalidXmlError` when absent), an optional repeating `<w:lvlOverride>`
list (`lvlOverride` = `ZeroOrMore`, `successors=()` → append), and a **required**
`@w:numId` attribute (`RequiredAttribute("w:numId", ST_DecimalNumber)`).
**`new(num_id, abstractNum_id)`** builds `<w:num w:numId="…">` and appends a
`CT_DecimalNumber` `<w:abstractNumId w:val="…"/>`. **`add_lvlOverride(ilvl)`**
adds a `CT_NumLvl` with `@w:ilvl` set — it wins over the generated
`add_lvlOverride` (the same xmlchemy suppression as `add_num`).

**Example**

```matlab
num = mat2doc.oxml.numbering.CT_Num.new(1, 0);    % <w:num w:numId="1">
fprintf('numId=%d  abstractNumId.val=%d\n', num.numId, num.abstractNumId.val);
% numId=1  abstractNumId.val=0
lo = num.add_lvlOverride(0);                       % a CT_NumLvl <w:lvlOverride w:ilvl="0">
fprintf('lvlOverride ilvl=%d  class=%s\n', lo.ilvl, class(lo));
% lvlOverride ilvl=0  class=mat2doc.oxml.numbering.CT_NumLvl
```

*Ported from python-docx v1.2.0: `src/docx/oxml/numbering.py::CT_Num`*

---

## `CT_NumLvl` — a `<w:lvlOverride>` level override

**Syntax**

```matlab
lvl = lvlOverride;                          % a CT_NumLvl (e.g. from CT_Num.add_lvlOverride)
so  = lvl.add_startOverride(val);           % a CT_DecimalNumber <w:startOverride w:val=...>
so2 = lvl.get_or_add_startOverride();       % ZeroOrOne get-or-create
v   = lvl.ilvl;                             % the required @w:ilvl
```

**Description**

`CT_NumLvl` (`numbering.py::CT_NumLvl`) identifies a list level to override with
the settings it contains. It carries an optional `<w:startOverride>`
(`ZeroOrOne`, `successors=("w:lvl",)`) and a **required** `@w:ilvl`
(`RequiredAttribute("w:ilvl", ST_DecimalNumber)`). **`add_startOverride(val)`**
adds a `CT_DecimalNumber` `<w:startOverride>` with `@w:val` set (`w:startOverride`
→ `CT_DecimalNumber` in the registry).

**Example**

```matlab
num = mat2doc.oxml.numbering.CT_Num.new(1, 0);
lo  = num.add_lvlOverride(0);          % a CT_NumLvl <w:lvlOverride w:ilvl="0">
so  = lo.add_startOverride(5);         % <w:startOverride w:val="5"/>
fprintf('startOverride val=%d\n', so.val);   % startOverride val=5
```

*Ported from python-docx v1.2.0: `src/docx/oxml/numbering.py::CT_NumLvl`*

---

## `CT_NumPr` — the paragraph `<w:numPr>` numbering properties

**Syntax**

```matlab
numPr = mat2doc.oxml.OxmlElement("w:numPr");  % registry -> CT_NumPr
ilvl  = numPr.get_or_add_ilvl();               % <w:ilvl> child (get-or-create)
numId = numPr.get_or_add_numId();              % <w:numId> child (get-or-create)
a = numPr.ilvl;      % the <w:ilvl> child or []  (read-only ZeroOrOne getter)
b = numPr.numId;     % the <w:numId> child or []
```

**Description**

`CT_NumPr` (`numbering.py::CT_NumPr`) is the numbering-properties container that
appears inside `<w:pPr>`, carrying the list level (`<w:ilvl>`) and
numbering-definition reference (`<w:numId>`) for a numbered/bulleted paragraph.
Both are `ZeroOrOne` (child class `CT_DecimalNumber`) with the byte-critical H11
successor slices ported verbatim:

| descriptor | successors |
|---|---|
| `ilvl` | `("w:numId", "w:numberingChange", "w:ins")` |
| `numId` | `("w:numberingChange", "w:ins")` |

Because `ilvl`'s successors list `w:numId`, adding `numId` **first** and `ilvl`
second still serializes `<w:ilvl>` **before** `<w:numId>` — the successor slice
re-sorts it. The commented-out `@ilvl.setter` / `@numId.setter` blocks in the
python-docx source (`numbering.py:60-75`) are **inactive in v1.2.0** and are **not
ported** — there is no `set.ilvl` / `set.numId` member.

**Example** (the H11 ordering — `numId` added first, `ilvl` serializes first):

```matlab
numPr = mat2doc.oxml.OxmlElement("w:numPr");
numPr.get_or_add_numId().val = 2;    % added FIRST
numPr.get_or_add_ilvl().val  = 0;    % added second -> sorts BEFORE numId
kids = numPr.getchildren();
local = strings(1, numel(kids));
for k = 1:numel(kids); local(k) = extractAfter(string(kids(k).tag), "}"); end
fprintf('child order: %s\n', strjoin(local, " "));
% child order: ilvl numId
```

*Ported from python-docx v1.2.0: `src/docx/oxml/numbering.py::CT_NumPr`*

---

## `NumberingPart` — the numbering.xml part

**Syntax**

```matlab
nd = numberingPart.numbering_definitions;   % a _NumberingDefinitions (@lazyproperty)
mat2doc.parts.NumberingPart.new();          % raises mat2doc:NotImplementedError (faithful)
```

**Description**

`NumberingPart` (`parts/numbering.py::NumberingPart`) is the `XmlPart` for
`word/numbering.xml`. **`numbering_definitions`** (a `@lazyproperty`, un-stubbed
at P8-1) returns a `NumberingDefinitions_` over the parsed `<w:numbering>` root
(a `CT_Numbering`), cached via a logical computed-flag (design.md `@lazyproperty`
rule — never an `isempty` sentinel).

**`new()` is a *faithful* `NotImplementedError`, not a port stub.** python-docx
v1.2.0 itself declares `def new(cls): raise NotImplementedError`
(`numbering.py:11-14`), so the MATLAB `NumberingPart.new()` reproduces this
verbatim as `mat2doc:NotImplementedError`. This is the one `notYetPorted`-shaped
site that is faithful upstream behavior — it is exempt from the zero-stub exit
criterion.

**Example**

```matlab
try
    mat2doc.parts.NumberingPart.new();
catch ME
    fprintf('%s\n', ME.identifier);   % mat2doc:NotImplementedError
end
```

*Ported from python-docx v1.2.0: `src/docx/parts/numbering.py::NumberingPart`*

---

## `NumberingDefinitions_` — the `<w:num>` definitions collection

**Syntax**

```matlab
nd = mat2doc.parts.NumberingDefinitions_(numbering_elm);  % wrap a <w:numbering> (CT_Numbering)
n  = nd.len_();                                            % __len__: number of <w:num> defs
```

**Description**

`NumberingDefinitions_` (`parts/numbering.py::_NumberingDefinitions`, leading-`_`
→ trailing-`_` per the FLAG-3 rotation) is a plain `handle` object (**not** an
`ElementProxy`) that wraps the `<w:numbering>` element and reports how many
`<w:num>` definitions it holds. The v1.2.0 surface is **exactly** `__init__` and
`__len__` — there is **no** `__getitem__` / `__iter__` / `add`, so no
`RedefinesParen` `()` indexing is needed. `len(nd)` is ported as `nd.len_()` (the
`LatentStyles` collection idiom).

**Example**

```matlab
d  = mat2doc.Document();
np = d.part().numbering_part;                 % a NumberingPart (default.docx HAS one)
nd = np.numbering_definitions;                % a NumberingDefinitions_
fprintf('class=%s  len_=%d\n', class(nd), nd.len_());
% class=mat2doc.parts.NumberingDefinitions_  len_=9

nd2 = np.numbering_definitions;               % lazyproperty -> same handle
fprintf('lazyproperty same handle: %d\n', nd == nd2);   % 1
```

*Ported from python-docx v1.2.0: `src/docx/parts/numbering.py::_NumberingDefinitions`*

---

## `DocumentPart.numbering_part` — the part accessor

**Syntax**

```matlab
np = documentPart.numbering_part;    % the NumberingPart (@lazyproperty)
```

**Description**

`DocumentPart.numbering_part` (`parts/document.py:98-109`, a `@lazyproperty`,
un-stubbed at P8-1) returns the related `NumberingPart` via
`part_related_by(RT.NUMBERING)`. On a package that **has** a numbering part
(`default.docx` does), it returns that related part. On a package **without** one,
it takes the `KeyError` branch and calls `NumberingPart.new()` — which
**faithfully raises `mat2doc:NotImplementedError`**, exactly as python-docx does
(python-docx cannot create a numbering part from nothing either). The
`@lazyproperty` caches only a **successful** return; an exception is re-raised on
every access, never cached.

**Example**

```matlab
d  = mat2doc.Document();
np = d.part().numbering_part;                 % a mat2doc.parts.NumberingPart
fprintf('numbering_part class = %s\n', class(np));
% numbering_part class = mat2doc.parts.NumberingPart
```

*Ported from python-docx v1.2.0: `src/docx/parts/document.py::DocumentPart.numbering_part`*

---

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
