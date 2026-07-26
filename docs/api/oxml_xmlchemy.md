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

**P1-3a is slice 1 of 2.** This page documents exactly what P1-3a ports:
the three `BaseOxmlElement` **tree-ops**, the **attribute-descriptor engine**
(`OptionalAttribute` / `RequiredAttribute`), and the **XPath mini-engine**. The
**child-element descriptor engine** (`ZeroOrOne` / `ZeroOrMore` /
`OneAndOnlyOne` / `OneOrMore` / `Choice` / `ZeroOrOneChoice` and their
`getChild` / `getOrAddChild` / `addChild` / `removeChild` / … generated
members) is **P1-3b** and is *not* on this page. Gate-3 for P1-3a:
`reports\p1_3a_validation.md` — **53/53, 0 new D-numbers** (38 xpath value/byte,
10 attr-descriptor, 5 tree-ops), run over the parsed `w:document` tree.

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

**Deferred, in their owning WPs** (not ported at P1-3a): the child-element
descriptor engine (`getChild` / `getOrAddChild` / `addChild` / … — **P1-3b**);
the `xml` property (`serialize_for_reading`, a pretty-printed test-only helper —
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
