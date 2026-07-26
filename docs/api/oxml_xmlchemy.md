---
title: "mat2doc.oxml — xmlchemy engine (BaseOxmlElement + XPath)"
---

# `mat2doc.oxml` — xmlchemy engine (`BaseOxmlElement` + XPath)

Ported from python-docx v1.2.0 module `src/docx/oxml/xmlchemy.py` (package
`+mat2doc/+oxml/`). This is the **declarative element-class engine** that sits
directly above `oxml_core` (P1-2): the effective base class
`BaseOxmlElement` from which every `CT_*` custom element class descends, plus
the **mini-XPath engine** (`evaluate_xpath`) that `BaseOxmlElement.xpath`
delegates to. It supplies, in one place, the tree-ops and the
attribute-descriptor machinery the schema classes reuse.

**xmlchemy ships in two slices; this page documents both.** **P1-3a** ports the
three `BaseOxmlElement` **tree-ops**, the **attribute-descriptor engine**
(`OptionalAttribute` / `RequiredAttribute`), and the **XPath mini-engine**;
Gate-3 **53/53, 0 new D-numbers** (`reports\p1_3a_validation.md` — 38 xpath
value/byte, 10 attr-descriptor, 5 tree-ops, over the parsed `w:document` tree).
**P1-3b** (this extension) ports the **child-element descriptor engine** — the
`ZeroOrOne` / `ZeroOrMore` / `OneAndOnlyOne` / `OneOrMore` / `Choice` /
`ZeroOrOneChoice` descriptor families and the 11 generic `getChild` /
`getRequiredChild` / `getChildList` / `newChild` / `insertChildInSequence` /
`addChild` / `getOrAddChild` / `removeChild` / `firstChildFoundIn` /
`removeChildren` / `getOrChangeToChild` engine methods every future `CT_*`
child-descriptor member delegates to; Gate-3 **46/46, 0 new D-numbers**
(`reports\p1_3b_validation.md` — the H11 successor-ordering battery
byte-identical to a live python-docx 1.2.0 oracle). **P1-3b completes
xmlchemy.**

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## The metaclass replacement (design.md §2)

python-docx's `MetaOxmlElement` walks each class body and, for every declared
`OptionalAttribute` / `RequiredAttribute` (and the child-element descriptors),
**generates** a read/write property whose get/set bodies reduce to a handful of
generic operations. MATLAB `classdef` has no metaclass hook of that kind, so the
port replaces the generated members with **explicit engine methods every `CT_*`
class delegates to** (design.md §2). `BaseOxmlElement` therefore carries the
generic operations as ordinary methods (`getAttrTyped` / `setAttrTyped` /
`getAttrRequired` / `setAttrRequired`, and the tree-ops); a schema class holds a
`Constant` table naming its attributes and their simple types, and its
generated-property equivalents forward into these engine methods. At P1-3a **no
`CT_*` class exists yet**, so the attribute engine is *plumbing* — wired, audited
line-by-line, and proven against a controlled test simple-type, but dormant on
the production surface until the `ST_*` simple types (P3) and the `CT_*` classes
land.

## Where this differs from python-pptx (the design guide) — the setter deltas

python-docx v1.2.0 is the **source of truth**. The xmlchemy engine is a
**byte-identical re-port** of the already-SOLVED Mat2Ppt engine (no shared
code — re-implemented in the `mat2doc:` namespace), and docx confirms the same
surface **except for three attribute-setter deltas**. These are carried as
**behavior** (the docx form is ported directly); **no new D-number** — Gate-3
confirmed 0. `isequal(x, [])` is the `x is None` analogue throughout (H3: `""`
is a real string, `isequal("", [])` is `false`, so `""` is stored, never
removed).

| Delta | python-docx (ported) | python-pptx (not ported here) | Effect |
|---|---|---|---|
| **D-delta-1** | `OptionalAttribute` setter guards `value is None OR value == default` (xmlchemy.py:203) | guards only `value == default` | Assigning `None` (`[]`) to an optional attribute with a **non-None default** **removes** it in docx; pptx would fall through to `to_xml(None)`. |
| **D-delta-2** | `OptionalAttribute` setter, **after** `to_xml`, removes the attribute when `str_value is None` (xmlchemy.py:208–211) | sets unconditionally | A simple type whose `to_xml` returns `None` **erases** the attribute rather than writing it. |
| **D-delta-3** | `RequiredAttribute` setter, **after** `to_xml`, **raises** `ValueError("cannot assign {value} to this required attribute")` when `str_value is None` (xmlchemy.py:257–258) | sets unconditionally | A required attribute whose `to_xml` returns `None` is a hard `mat2doc:ValueError`, never a silent write. |

The setter-delta message text in D-delta-3 (`"{value}"`) is rendered from a
best-effort `string(value)` (`valueRepr_`), which may differ from CPython
`str(value)` for non-string values — **error-path message text only**, carried
under the existing **D-005** message-token class (no new number).

## Gate-2 FIX-1 — the `self::` / `parent::` name-test (ISSUE-1)

The mso-auditor (Opus) caught **FIX-1**, the one engine defect in this WP, in
`evaluate_xpath.m`'s handling of **named node tests on the `self` and `parent`
axes**. The re-ported engine (built for pptx, which has *zero* `self::NAME` /
`parent::NAME` call-sites) ignored the name test on those axes: `self::w:tbl` on
a `w:p` context wrongly matched the context node, and
`./parent::w:r/parent::w:hyperlink` mis-resolved — a real docx pattern
(`pagebreak`'s `_is_in_hyperlink`), so the bug would have **inverted a
production boolean**. The fix adds the name-test filter on both axes (match only
when `tagOf(node) == resolve(name)`; a bare `.` / `..` / `self::*` / `parent::*`
still matches unchanged). Re-verified against live lxml as **F01–F05** in
Gate-3: `self::w:tbl`→∅, `self::w:p`→`[p]`, `parent::w:body`→`[body]`,
`./parent::w:r/parent::w:hyperlink` inside a hyperlink→`[hyperlink]` and on a
plain run→∅ — all lxml-identical. This closes design.md §3's
*"never silently mis-evaluated"* rule for the `self` / `parent` axes.

:::{warning}
**★ VERIFY-2 — the XPath coverage gap (completed by P1-3x).** The engine was
re-ported from Mat2Ppt, which was built against **python-pptx's** narrower
xpath call-site inventory. python-docx uses a **broader** xpath surface. A
full docx call-site grep (Gate-2) enumerated patterns the engine **does not yet
support** and which currently **RAISE `mat2doc:XPathError`** (the *safe*
direction — they raise, they never mis-evaluate):

- **bare top-level unions** `./w:p | ./w:tbl` — docx's single most common
  pattern (`CT_Body`, `CT_Tc`, block-item containers)
- `not(self::…)` (block-item iteration)
- `preceding-sibling::` / `following-sibling::` (pagebreak, section iteration)
- `preceding::` (section block iterator)
- `position()=1`, `(…)[last()]` (tables, sections)
- predicate attribute-sub-paths, a predicate on a terminal string step (`//@id[2]`)

P1-3a itself is **not blocked** — its scope is the engine as re-ported, and the
raises are faithful *"not-yet-supported"* guards, not wrong answers. But these
patterns are consumed by the docx `CT_*` API classes (`CT_Body` / `CT_Tbl` /
`CT_SectPr`, pagebreak, styles, comments), so a dedicated
**xpath-engine-extension WP — P1-3x** — is scheduled to extend `evaluate_xpath`
to cover them, each byte-verified against lxml on the frozen docx call-site
inventory. Sequencing: **P1-3a (done) → P1-3b (child descriptors) → P1-3x
(xpath extension) → P1-4 (opc)** — before the first CT_* consumer needs them
(M1 is unaffected; it only parses/serializes `document.xml`). See
`validation\summary\decision_2026-07-25_mat2doc_xpath_engine_extension.md`.
Gate-3 pins these patterns as raises-asserts **R01–R09**, frozen with lxml's
`would_be` for contrast — they go **deliberately RED when P1-3x lands**, surfacing
the intended subset-widening.
:::

---

## `BaseOxmlElement`

**Syntax**

```matlab
e = mat2doc.oxml.BaseOxmlElement(nsptag)
e = mat2doc.oxml.BaseOxmlElement(nsptag, nsmapOrDecls)
```

**Description**

Effective base class for all custom `CT_*` element classes —
`classdef BaseOxmlElement < mat2doc.oxml.XmlElement`. It adds the standardized
tree-ops and the attribute-descriptor engine to one place; every `CT_*` class
ported in later WPs extends it. The constructor is a **transparent
pass-through**: it forwards all positional args verbatim to the base
`XmlElement` constructor (design.md §2 "CT_* constructor contract"), with **no
re-validation** of the `nsmap` argument — `XmlElement` is the single point that
accepts both the struct-`nsmap` and the Nx2-string decl-pair currencies, and the
parser instantiates registered `CT_*` classes via
`feval(cls, name, ownDecls)` with the Nx2 decl-pair, so a struct-typed guard here
would break parsing.

The **child-element descriptor engine** (`getChild` / `getOrAddChild` /
`addChild` / … — the 11 methods of P1-3b) is documented in its own section
below.

**Deferred, in their owning WPs** (still not ported): the `xml` property
(`serialize_for_reading`, a pretty-printed test-only helper —
the **doc-serialize/OPC WP**; provided here as a clean `mat2doc:notYetPorted`
stub so an accidental caller gets a *named* error); and `__repr__` / `_nsptag`
(display-only — note the rotated name `nsptag_` would collide with `XmlElement`'s
private `nsptag_`, so a direct rotation is unavailable and a distinct name is
required when eventually needed).

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
pPr = p.insert_element_before( ...
    mat2doc.oxml.OxmlElement("w:pPr"), "w:r");   % pPr placed before w:r
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement`*

---

## `first_child_found_in`

**Syntax**

```matlab
child = e.first_child_found_in(tagname1, tagname2, ...)
```

**Description**

Returns the first child element whose tag is among `tagname1, tagname2, …`
(each a prefixed tag, e.g. `"w:pPr"`), or `[]` (Python `None`, H3) when none is
found. Search order is **argument order** — the first *name* that matches wins,
**not** document order — exactly the Python `for tagname in tagnames: child =
self.find(qn(tagname))` loop.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
found = p.first_child_found_in("w:pPr", "w:r");   % the w:r (argument order)
none  = p.first_child_found_in("w:pPr");          % [] — not found (H3)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement.first_child_found_in` (lines 656–662)*

---

## `insert_element_before`

**Syntax**

```matlab
elm = e.insert_element_before(elm, tagname1, tagname2, ...)
```

**Description**

Inserts `elm` immediately **before** the first child whose tag is among
`tagname1, tagname2, …` (argument-order search via `first_child_found_in`), or
**appends** it when none is found. Returns `elm`. Because `append` / `addprevious`
**move** the element (H5), `elm`'s text/tail travel with it.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
pPr = p.insert_element_before(mat2doc.oxml.OxmlElement("w:pPr"), "w:r");
% p is now <w:p><w:pPr/><w:r/></w:p>
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement.insert_element_before` (lines 664–670)*

---

## `remove_all`

**Syntax**

```matlab
e.remove_all(tagname1, tagname2, ...)
```

**Description**

Removes **all** child elements whose tag (e.g. `"w:p"`) is among the given
tagnames. `findall` returns a materialized list (H9), so removing while looping
over it is safe in both languages.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:pPr"));
p.append(mat2doc.oxml.OxmlElement("w:r"));
p.remove_all("w:r");     % p is now <w:p><w:pPr/></w:p>
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement.remove_all` (lines 672–677)*

---

## Attribute-descriptor engine — `getAttrTyped` / `setAttrTyped` / `getAttrRequired` / `setAttrRequired`

**Syntax**

```matlab
value = e.getAttrTyped(name, type, default)      % OptionalAttribute getter
        e.setAttrTyped(name, type, value, default)  % OptionalAttribute setter
value = e.getAttrRequired(name, type)            % RequiredAttribute getter
        e.setAttrRequired(name, type, value)        % RequiredAttribute setter
```

**Description**

The generic get/set operations that xmlchemy's `MetaOxmlElement` bakes into each
generated `OptionalAttribute` / `RequiredAttribute` property, exposed here as
engine methods a `CT_*` class delegates to. `name` is the schema attribute name
(`"w:val"` or a plain `"space"`); `type` is the simple-type (or enum) class-name
token; `default` is the `OptionalAttribute` default (`None` → `[]`, H3).

- **`name` → Clark key** via `attrClarkName_` (`BaseAttribute._clark_name`,
  xmlchemy.py:139–143): a prefixed name (`"w:val"`) resolves through `qn`; a
  plain name (`"space"`) is used verbatim.
- **`type` → converter** via `resolveTypeCls_` (`BaseAttribute.simple_type`,
  xmlchemy.py:117–120): a **bare** short name (`"ST_String"`) prefixes to
  `mat2doc.oxml.simpletypes.<name>`; a name that **already carries a package**
  (`"mat2doc.enum.WD_UNDERLINE"`) is used verbatim (the enum-dispatch currency).
  Both surfaces expose `from_xml(xml_value)` / `to_xml(value)` statics, so the
  `feval` dispatch is uniform.
- **Getter behavior:** `getAttrTyped` returns `default` when the attribute is
  absent, else `type.from_xml(attr)`. `getAttrRequired` **raises**
  `mat2doc:InvalidXmlError("required '%s' attribute not present on element %s")`
  (the element's Clark `tag`, matching lxml's `obj.tag`) when absent.
- **Setter behavior:** the three **docx setter deltas** above (D-delta-1/2/3) all
  live here.

:::{note}
**Dormant at P1-3a.** `resolveTypeCls_` targets `+oxml/+simpletypes` (P3) and
`+enum` — neither exists yet, and **no `CT_*` class delegates here in this
slice**, so the dispatch is not reachable on the production surface. Gate-3
exercised the full engine (get/set, both attribute kinds, all three deltas,
byte-exact serialized fixtures, and the `InvalidXmlError`/`ValueError` messages)
against a **controlled test simple-type** (`+p13probe\TST.m`, identical on both
sides) that isolates the divergence to the engine itself. The example below is
therefore **illustrative of the call pattern**, not runnable standalone until the
`ST_*` types land.
:::

**Example (illustrative — the `ST_*` simple types are P3)**

```matlab
% Once mat2doc.oxml.simpletypes.ST_String exists, a CT_* class body forwards:
%   val = obj.getAttrTyped("w:val", "ST_String", [])   % OptionalAttribute get
%   obj.setAttrTyped("w:val", "ST_String", "Heading1", [])   % -> w:val="Heading1"
%   obj.setAttrTyped("w:val", "ST_String", [], "Normal")     % D-delta-1: REMOVES w:val
%   id = obj.getAttrRequired("w:id", "ST_DecimalNumber")     % raises if absent
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::OptionalAttribute` / `RequiredAttribute` `_getter` / `_setter` (lines 154–261)*

---

## The child-element descriptor engine (P1-3b)

python-docx declares each child element on a `CT_*` class with a **child
descriptor** — `ZeroOrOne` / `ZeroOrMore` / `OneAndOnlyOne` / `OneOrMore` /
`Choice` / `ZeroOrOneChoice`. `MetaOxmlElement` walks the class body and, per
descriptor, **generates** a family of accessor members (`get.x` / `x_lst` /
`_new_x` / `_insert_x` / `_add_x` / `add_x` / `get_or_add_x` / `_remove_x` /
`get_or_change_to_x` / `_remove_eg_x`) whose bodies all reduce to a handful of
generic operations. As with the attribute engine (design.md §2), MATLAB has no
metaclass hook, so the port carries those generic operations as **11 explicit
engine methods** on `BaseOxmlElement`; a future `CT_*` class holds a `Constant`
`_tag_seq` / successor / group-member table and its generated-member equivalents
forward into these methods (no runtime metaprogramming).

**Descriptor family → engine method** (docx `src/docx/oxml/xmlchemy.py`):

| Descriptor (generated member it backs) | docx lines | engine method |
|---|---|---|
| `ZeroOrOne` / `Choice` (`get.x`) | 380–382 | `getChild` |
| `OneAndOnlyOne` (`get.x`, required) | 499–505 | `getRequiredChild` |
| `ZeroOrMore` / `OneOrMore` (`x_lst`) | 397–398 | `getChildList` |
| `_new_x` (default creator) | 366–367 | `newChild` |
| `_insert_x` | 319–321 | `insertChildInSequence` |
| `_add_x` / `add_x` | 284–291 | `addChild` |
| `ZeroOrOne` (`get_or_add_x`) | 557–562 | `getOrAddChild` |
| `ZeroOrOne` (`_remove_x`) | 572–573 | `removeChild` |
| `ZeroOrOneChoice` (`get.x`) | 622–623 | `firstChildFoundIn` |
| `ZeroOrOneChoice` (`_remove_eg_x`) | 610–612 | `removeChildren` |
| `Choice` (`get_or_change_to_x`) | 453–461 | `getOrChangeToChild` |

Because no `CT_*` class exists yet (they arrive at P4+), these methods are the
**engine contract** — audited line-by-line and frozen against a live oracle, but
with **no production caller on the current surface**. Gate-3 drove them with the
**real `CT_PPr._tag_seq`** (`oxml/text/parfmt.py` 64–119) as the successor
slices; every serialized `w:pPr` fragment was **byte-identical** to
python-docx 1.2.0.

### H11 — the successor-ordering crux

`insertChildInSequence` expands the class's `SUCCESSORS` constant (a string
array of the tags that must sort **after** the inserted child) into the
repeating tagname arguments of the P1-3a tree-op `insert_element_before`, which
scans them in **argument order**, `addprevious`-es the **first present**
successor, and otherwise **appends**. So a child always lands *before the first
successor already in the tree* — the mechanism that keeps `w:pPr` children in
schema order regardless of the order the API creates them.

The successor list for a tag at schema position *N* is the schema slice
`_tag_seq[N:]` (everything from *N* onward — the child sorts before all of it).
The H1 0→1 base shift is applied **once, at declaration**, when the slice is
encoded (`_tag_seq[N:]` → `tagSeq(N+1:end)`); the engine itself contains **no
index arithmetic**. Gate-3 hand-encoded all 12 `CT_PPr` slices on both sides and
replayed an 18-step out-of-order build (`get_or_add_*` in scrambled order, an
H5 get-existing repeat, a remove+re-add, a foreign trailing child, and a
`sectPr` re-add that takes the **append** path *after* the foreign child): the
serialized `w:pPr` matched the python-docx oracle **byte-for-byte at every
step**.

### H5 — get-or-add identity

`getOrAddChild` (and `getOrChangeToChild`) return the **live** child handle from
`find` when the child is already present — never a copy — so back-to-back calls
return the **same handle** (`a == b`, Python `is` true), and `addChild`'s return
**is** the node now in the tree.

### H9 — materialized child list

`getChildList` returns `findall(...)`, a fully **materialized** `(1,N)`
`mat2doc.oxml.XmlElement` array (Python returns a `list`, not a lazy iterator).
"None present" is a typed **`1x0`** array — a real empty list, distinct from the
`[]` (None) the single-child getters return. Because the array is a snapshot,
removing children while iterating over it is safe in both languages.

### H3 — the tri-state "absent" encoding

Three distinct encodings of "not there", matching Python exactly:

| Method | absent result | Python |
|---|---|---|
| `getChild` / `firstChildFoundIn` | `[]` (class `double`) | `None` |
| `getChildList` | `1x0` typed `XmlElement` | `[]` (empty list) |
| `getRequiredChild` | **raises** `mat2doc:InvalidXmlError` | `InvalidXmlError` |

`getRequiredChild`'s message is byte-identical to python-docx, RST
double-backticks and all — `` required ``<w:abstractNumId>`` child element not
present `` (verified against the live `CT_Num.abstractNumId`, `numbering.py:20`).

### docx-vs-pptx deltas

The child-descriptor family is a **byte-identical re-port** of the SOLVED
Mat2Ppt engine (no shared code — re-implemented `mat2doc:`-namespaced). docx
v1.2.0 confirms the same surface, with **one generated-member delta**, plus two
descriptor families that are **dead code**:

- **D-delta-4 (engine-neutral, no D-number).** docx
  `ZeroOrMore.populate_class_members` calls `_add_public_adder` (xmlchemy.py
  536; hoisted up to `_BaseChildElement`, 340–352), so a docx `ZeroOrMore`
  **also** generates a *public* `add_x()`; pptx's `ZeroOrMore` does not (only
  `OneOrMore` does). The public `add_x()` routes through the **same `addChild`
  primitive** as `_add_x`, so the engine surface here is unchanged. The delta
  only decides **which per-class delegating member a future `CT_*` WP
  scaffolds**: the genoxml scaffolder must emit a public `add_x → addChild`
  delegator at each of the **24 docx `ZeroOrMore` sites** (pptx does not). This
  is a **genoxml-scaffolder obligation**, recorded engine-neutrally — **no new
  D-number**.
- **`Choice` / `ZeroOrOneChoice` are dead code in docx v1.2.0.** A grep over
  `src/docx/**/*.py` (excluding `xmlchemy.py`) finds **0** production
  `Choice(` / `ZeroOrOneChoice(` uses, so `getOrChangeToChild`,
  `firstChildFoundIn`, and `removeChildren` have **no real-element byte-oracle**
  and cannot until a docx element ever declares a choice group. They are ported
  for **engine-contract parity** with the descriptor family (and mirror the
  *live, validated* Mat2Ppt engine, where the same methods back dml
  fill/line/color). Gate-3 covers them with a **synthetic** value probe over
  hand-built `w:gA` / `w:gB` / `w:gC` groups (5/5) — no package scenario is
  possible or needed.

:::{note}
**Fold-forward + VERIFY-3.** The engine emits **no standalone XML part** — its
accessors are consumed by `CT_*` classes that do not exist yet. So the
**package-level L1 scenario lands at the first `CT_*` WP** (a `CT_PPr`
out-of-order `get_or_add_*` build → a real `w:pPr` inside `document.xml`); this
WP's Gate-3 is the **engine-probe freeze**, not a pkgcompare L0–L3 ladder. And
because the registry is still **empty** at P1-3b, parsed trees remain
**homogeneous** `XmlElement`, so the design.md §2 Sealed-method /
heterogeneous-array risk is **not yet exercisable and MUST be re-verified at the
first `CT_*` registration** (**VERIFY-3**, carried forward from P1-3a).
:::

---

## `getChild`

**Syntax**

```matlab
child = e.getChild(tag)
```

**Description**

Returns the child element with prefixed tag `tag` (e.g. `"w:pPr"`), or `[]`
(Python `None`, H3) when absent. Backs the `ZeroOrOne` and `Choice` getters
(`get.x`) — `return obj.find(qn(nsptag))`.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:pPr"));
pPr  = p.getChild("w:pPr");    % the live w:pPr handle
none = p.getChild("w:r");      % [] (None, H3)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::_BaseChildElement._getter` (lines 380–382)*

---

## `getRequiredChild`

**Syntax**

```matlab
child = e.getRequiredChild(tag)
```

**Description**

Returns the child with prefixed tag `tag`, or **raises**
`mat2doc:InvalidXmlError` when absent — the `OneAndOnlyOne` getter. A missing
required child is malformed XML, not a sentinel. The message reproduces the
Python RST double-backticks verbatim: `` required ``<tag>`` child element not
present ``.

**Example**

```matlab
num = mat2doc.oxml.BaseOxmlElement("w:num");
try
    num.getRequiredChild("w:abstractNumId");   % raises — child absent
catch e
    disp(e.identifier)   % "mat2doc:InvalidXmlError"
end
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::OneAndOnlyOne._getter` (lines 499–505)*

---

## `getChildList`

**Syntax**

```matlab
list = e.getChildList(tag)
```

**Description**

Returns **all** children with prefixed tag `tag`, in document order, as a
materialized `(1,N)` `mat2doc.oxml.XmlElement` array (H9). Backs `ZeroOrMore` /
`OneOrMore` (`x_lst`) — `return obj.findall(qn(nsptag))`. "None present" is a
typed **`1x0`** array (an empty *list*), **not** `[]` (None) — distinct from the
single-child getters (H3). `findall` is already 1-based document order (H1 — no
index shift).

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
p.append(mat2doc.oxml.OxmlElement("w:r"));
runs = p.getChildList("w:r");         % 1x2 XmlElement array
none = p.getChildList("w:hyperlink"); % 1x0 typed empty (materialized, H9)
disp(numel(runs))                     % 2
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::_BaseChildElement._list_getter` (lines 397–398)*

---

## `newChild`

**Syntax**

```matlab
child = e.newChild(tag)
```

**Description**

Creates a **loose** child element of the correct type (via the registry — a
registered `CT_*` class, or a plain `XmlElement` fallback), with no attributes
and **not yet attached** to `e`. The default creator (`_new_x`) —
`return OxmlElement(nsptag)`. `e` is unused (matches Python: the default creator
ignores `obj`) but kept for method-call form; a `CT_*` class with a `_new_x`
override supplies its own creator instead of this generic one.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
r = p.newChild("w:r");   % loose <w:r/>, not yet in p
disp(r.nsptag_str)       % "w:r"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::_BaseChildElement._creator` (lines 366–367)*

---

## `insertChildInSequence`

**Syntax**

```matlab
child = e.insertChildInSequence(child, successors)
```

**Description**

Inserts `child` immediately **before** the first of `successors` already present
in `e` (argument-order search via the P1-3a tree-op `insert_element_before`), or
**appends** it when none is present; returns `child`. `successors` is a `(1,:)`
string array — a `CT_*` class's `SUCCESSORS` constant, the schema slice
`_tag_seq[N:]` for the child's tag. This is the mechanism behind H11
successor-ordering; the H1 base shift lives in the *declaration* of the slice,
never here.

**Example**

```matlab
pPr = mat2doc.oxml.BaseOxmlElement("w:pPr");
pPr.append(mat2doc.oxml.OxmlElement("w:jc"));
% w:spacing must precede w:jc → pass w:jc as its successor:
spacing = pPr.insertChildInSequence(mat2doc.oxml.OxmlElement("w:spacing"), "w:jc");
% pPr is now <w:pPr><w:spacing/><w:jc/></w:pPr>
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::_BaseChildElement._add_inserter` (lines 319–321)*

---

## `addChild`

**Syntax**

```matlab
child = e.addChild(tag, successors)
child = e.addChild(tag, successors, name1, value1, ...)
```

**Description**

Creates a child of type `tag` with the default creator (`newChild`), sets any
trailing `name`/`value` attribute pairs (the Python `**attrs`, applied via the
typed-property setter `child.(name) = value`), inserts it in schema sequence
(`insertChildInSequence` against `successors`), and returns it. Backs the
private `_add_x` and — in docx — the **public** `add_x` (D-delta-4); both route
through this one primitive. A `CT_*` class with a `_new_x` / `_insert_x`
override does **not** delegate here; it hand-routes through its own creator.

:::{note}
The `**attrs` typed-set path is **dormant** until a `CT_*` class with typed
attribute properties exists (no current docx `_add_x(**attrs)` call site) —
re-verified at the first such CT WP (VERIFY-3).
:::

**Example**

```matlab
pPr = mat2doc.oxml.BaseOxmlElement("w:pPr");
pPr.append(mat2doc.oxml.OxmlElement("w:jc"));
spacing = pPr.addChild("w:spacing", "w:jc");   % create + insert before w:jc
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::_BaseChildElement._add_adder` (lines 284–291)*

---

## `getOrAddChild`

**Syntax**

```matlab
child = e.getOrAddChild(tag, successors)
```

**Description**

Returns the child with prefixed tag `tag`, **creating and inserting it in
sequence** (via `addChild` against `successors`) if it is absent — the
`ZeroOrOne` `get_or_add_x`. When the child is already present, the **same live
handle** is returned on every call (H5); no duplicate is created.

**Example**

```matlab
pPr = mat2doc.oxml.BaseOxmlElement("w:pPr");
jc      = pPr.getOrAddChild("w:jc",      string.empty(1,0));  % no successors → append
spacing = pPr.getOrAddChild("w:spacing", "w:jc");             % before w:jc
same    = pPr.getOrAddChild("w:spacing", "w:jc");             % H5: same handle
disp(same == spacing)   % 1  (no new child added)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::ZeroOrOne._add_get_or_adder` (lines 557–562)*

---

## `removeChild`

**Syntax**

```matlab
e.removeChild(tag)
```

**Description**

Removes **all** children of `e` with prefixed tag `tag` (via `remove_all`) — the
`ZeroOrOne` `_remove_x`. Faithful to `remove_all`, it removes *every* match, not
just the first.

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
p.append(mat2doc.oxml.OxmlElement("w:r"));
p.removeChild("w:r");                 % removes both
disp(numel(p.getChildList("w:r")))    % 0
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::ZeroOrOne._add_remover` (lines 572–573)*

---

## `firstChildFoundIn`

**Syntax**

```matlab
child = e.firstChildFoundIn(tags)
```

**Description**

Returns the first child of `e` whose tag is among `tags` (a `(1,:)` string array
of the choice group's member tags), searched in **argument order** (first *name*
that matches wins, not document order), or `[]` (None, H3) when none is present.
The `ZeroOrOneChoice` group getter (`get.x`) — `return
obj.first_child_found_in(*member_nsptagnames)`.

:::{note}
`ZeroOrOneChoice` is **dead code in docx v1.2.0** (0 production uses); this
method is an engine-contract parity member with synthetic coverage only.
:::

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:r"));
found = p.firstChildFoundIn(["w:pPr", "w:r"]);  % the w:r (argument order)
none  = p.firstChildFoundIn("w:pPr");           % [] (H3)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::ZeroOrOneChoice._choice_getter` (lines 622–623)*

---

## `removeChildren`

**Syntax**

```matlab
e.removeChildren(tags)
```

**Description**

Removes whichever of the group-member tags `tags` are present in `e` (one
variadic `remove_all` over the whole group ≡ the Python per-tag loop) — the
`ZeroOrOneChoice` `_remove_eg_x`.

:::{note}
`ZeroOrOneChoice` is **dead code in docx v1.2.0**; engine-contract parity member
with synthetic coverage only.
:::

**Example**

```matlab
p = mat2doc.oxml.BaseOxmlElement("w:p");
p.append(mat2doc.oxml.OxmlElement("w:gA"));
p.append(mat2doc.oxml.OxmlElement("w:gB"));
p.removeChildren(["w:gA", "w:gB", "w:gC"]);   % drop whichever are present
disp(numel(p.getChildList("w:gA")))            % 0
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::ZeroOrOneChoice._add_group_remover` (lines 610–612)*

---

## `getOrChangeToChild`

**Syntax**

```matlab
child = e.getOrChangeToChild(tag, groupTags, successors)
```

**Description**

Returns the choice member with prefixed tag `tag`, **switching the group to it**
when a *different* member is currently present: if `tag` is already present it is
returned unchanged (H5 same handle); otherwise every member of `groupTags` is
removed (`removeChildren`) and `tag` is created and inserted in sequence
(`addChild` against `successors`). The `Choice` `get_or_change_to_x`.

:::{note}
`Choice` is **dead code in docx v1.2.0** (0 production uses); engine-contract
parity member with synthetic coverage only.
:::

**Example**

```matlab
rPr = mat2doc.oxml.BaseOxmlElement("w:rPr");
% choice group {w:b, w:bCs}:
b   = rPr.getOrChangeToChild("w:b",   ["w:b","w:bCs"], string.empty(1,0)); % add w:b
bCs = rPr.getOrChangeToChild("w:bCs", ["w:b","w:bCs"], string.empty(1,0)); % swap to w:bCs
disp(numel(rPr.getChildList("w:b")))   % 0  (w:b was removed)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::Choice._add_get_or_change_to_method` (lines 453–461)*

---

## `xpath` (method on `BaseOxmlElement`)

**Syntax**

```matlab
nodes = e.xpath(expr)
nodes = e.xpath(expr, ns)
```

**Description**

Evaluate an XPath expression over `e`'s tree. `e.xpath(expr)` uses the **fixed
WordprocessingML namespace map** (`mat2doc.oxml.nsmap`, the Python `nsmap`),
mirroring `BaseOxmlElement.xpath` (xmlchemy.py:687–692), which centralizes the
namespace mapping. **docx-vs-pptx:** docx injects the **PUBLIC `nsmap`**; pptx
injects the private `_nsmap`. python-docx's own override **drops** lxml's
`namespaces` kwarg and always injects `nsmap`, so no upstream call site passes a
custom map — `e.xpath(expr)` is byte-for-byte the upstream behavior. The optional
`ns` (a scalar prefix→URI struct) resolves prefixes against a custom map instead.

**Return type** mirrors lxml exactly (H3 — an empty match is a typed **empty
array**, never `[]`/None; callers use `numel` / `~isempty` / `arr(1)`):

| Expression ends in | Result |
|---|---|
| `/@attr` | `(1,N)` string (attribute values); `string.empty(1,0)` on no match |
| `/text()` | `(1,N)` string (text nodes); `string.empty(1,0)` on no match |
| otherwise | `(1,N) mat2doc.oxml.XmlElement`; `XmlElement.empty(1,0)` on no match |

The supported subset and error contract are those of `evaluate_xpath` (below);
anything outside the subset raises `mat2doc:XPathError`.

**Example**

```matlab
xml = "<w:document xmlns:w='http://schemas.openxmlformats.org/" + ...
      "wordprocessingml/2006/main'><w:body><w:p><w:r><w:t>hi</w:t>" + ...
      "</w:r></w:p></w:body></w:document>";
doc = mat2doc.oxml.parse_xml(xml);       % returns a BaseOxmlElement subtree once CT_* register
% At P1-3a the registry is empty, so drive the engine directly for a runnable demo:
body = mat2doc.oxml.evaluate_xpath(doc, "w:body", mat2doc.oxml.nsmap());
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement.xpath` (lines 687–692)*

---

## `evaluate_xpath` — the mini XPath-1.0 engine

**Syntax**

```matlab
result = mat2doc.oxml.evaluate_xpath(context, xpath_str, ns)
```

**Description**

The mini XPath-1.0 engine `BaseOxmlElement.xpath` delegates to. Evaluates
`xpath_str` against the `XmlElement` `context` with prefix→URI map `ns` (the P1-2
`nsmap()` currency), returning **exactly** what lxml's
`_Element.xpath(expr, namespaces=ns)` returns for the closed subset of
expressions the docx/lxml surface uses (design.md §XPath). This is **not** a
general XPath engine: anything outside the verified subset raises
`mat2doc:XPathError` rather than being silently mis-evaluated (design.md §3, §7)
— see the **VERIFY-2** warning above.

**Supported subset**

| Category | Examples |
|---|---|
| prefixed child paths | `w:body/w:p/w:r` |
| self / relative-descendant | `.`  `./w:x`  `.//w:x` |
| absolute / absolute-descendant | `/w:document`  `//w:p`  `//@r:id` |
| wildcard child | `./*[1]` |
| **positional predicate (1-based, H1 — never shift)** | `w:p[1]`  `w:p[2]` |
| attribute-equality predicate | `w:t[@xml:space="preserve"]`  `w:x[@w:val=5]` |
| element-existence predicate | `w:p[w:pPr]`  `w:x[w:y[@w:val="1"]]` |
| attribute result | `./w:pPr/w:rPr/@w:val`  `//@r:id` |
| **`text()` result / mixed content** | `./w:p//w:t/text()`  `./text()`  `.//text()` |
| ancestor axis | `ancestor::w:tbl` |
| **`self::` / `parent::` named tests (FIX-1)** | `self::w:p`  `./parent::w:r/parent::w:hyperlink` |
| parent step + union **group** | `(../w:x \| ../w:y)/w:z[@w:val="1"]` |

**Positional predicates are 1-based (H1).** `w:p[1]` selects the first matching
child, `w:p[2]` the second — never shift the index when porting a Python
call-site.

**`text()` and mixed content (F1 / WPC-F1 / D10).** `text()` selects **all**
text-node children of an element in document order: the element's own leading
text (its C-level text node, read via `text_raw_()` to **bypass** any `CT_*`
`getText_` property shadow — D10), then the **tail** of each child element. A
child's tail sorts **after** that child's entire subtree (WPC-F1: sort key
`[docKey(kid), Inf]`), giving lxml document order for nested `text()`. So for
`<w:p>PT1<w:r>RT</w:r>MID<w:hyperlink>HL</w:hyperlink>END</w:p>`,
`./text()`→`["PT1","MID","END"]` and `.//text()`→`["PT1","RT","MID","HL","END"]`,
both exactly lxml.

**The public-`nsmap` default.** The engine takes `ns` explicitly; the
`BaseOxmlElement.xpath` method (above) supplies the **public** `nsmap` as the
default, so `e.xpath(expr)` resolves prefixes against the same fixed
WordprocessingML map docx uses.

**Result / no-match contract.** `/@attr` and `/text()` terminals return `(1,N)`
string arrays; every other expression returns `(1,N) mat2doc.oxml.XmlElement`.
A no-match is the **typed empty array** (`string.empty(1,0)` /
`XmlElement.empty(1,0)`), never `[]` (H3). String (attribute / `text()`) results
are identity-deduped by node (F3), matching lxml node-set semantics.

**Re-port provenance.** This is the SOLVED Mat2Ppt `+oxml` evaluator
(WP5 + the WP5-C corrective fixes) re-ported verbatim with `mat2doc`
namespacing; docx v1.2.0 is the module source of truth. The evaluator is
tree/namespace-agnostic within its subset, so only the identifiers and error ids
changed — **plus FIX-1** (the `self::` / `parent::` name-test, above). The
WP5-C corrective semantics (F1 / F2 / F3 / WPC-F1) are adopted here
`mat2doc:`-namespaced (**shared deviation, no new D-number**), bringing the engine
to lxml fidelity so design.md's *"never silently mis-evaluated"* rule holds with
**zero** silent mis-evaluations even within the subset.

**Example**

```matlab
c   = mat2doc.oxml.nsmap();     % the fixed WordprocessingML prefix->URI map
p   = mat2doc.oxml.parse_xml( ...
    "<w:p xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'>" + ...
    "<w:r><w:t>hi</w:t></w:r></w:p>");
t   = mat2doc.oxml.evaluate_xpath(p, 'w:r/w:t', c);
disp(t.nsptag_str)               % "w:t"  (1x1 XmlElement)
none = mat2doc.oxml.evaluate_xpath(p, 'w:r/w:br', c);
disp(isempty(none))              % 1  -- typed EMPTY XmlElement, not [] (H3)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/xmlchemy.py::BaseOxmlElement.xpath`
(the lxml `_Element.xpath` subset it narrows; design-realization, D-001). Re-ported
from the corrected Mat2Ppt `+oxml` `evaluate_xpath` — no shared code.*
