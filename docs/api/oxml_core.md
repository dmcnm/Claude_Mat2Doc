---
title: "mat2doc.oxml — tree / parser / serializer core"
---

# `mat2doc.oxml` — tree / parser / serializer core

Ported from python-docx v1.2.0 modules `src/docx/oxml/ns.py`,
`src/docx/oxml/parser.py`, and `src/docx/opc/oxml.py::serialize_part_xml`
(package `+mat2doc/+oxml/`). This is the **byte-fidelity foundation** of
Mat2Doc — the layer beneath `xmlchemy` / `BaseOxmlElement` (P1-3): the
WordprocessingML namespace machinery, the own order-preserving OOXML parser,
the `XmlElement` tree node (lxml `_Element` replacement), the element
factory / class registry, and the lxml-matched part-XML serializer. Its
observable output is serialized XML **bytes**; the bar is **byte-identical to
lxml 5.3.0 / python-docx 1.2.0** on real docx content (Gate-3
`reports\p1_2_validation.md`: 27/27 — 23 L1 byte-identical, 1 L2
PASS-DEVIATION, 3 reject-as-designed, 0 new D-numbers).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
Keep this page's structure simple — one section per symbol: syntax,
description, example, ported-from. Edit the headers, not this page, when the
generator lands; until then this page is maintained by hand to the same shape.
:::

## Where this differs from python-pptx (the design guide)

python-docx v1.2.0 is the **source of truth**. Mat2Doc shares no code with
Mat2Ppt (the python-pptx port); the `+oxml` design is **re-implemented** in the
`mat2doc:` namespace against the SOLVED Mat2Ppt `+oxml` reference. The docx
`ns.py` / `parser.py` surface differs from the pptx design guide in exactly
these ways, all honored here:

- **`nsmap` is PUBLIC and has 16 entries** (pptx used a module-private `_nsmap`
  with 26). Because docx names it publicly there is **no underscore rotation**:
  the file is `nsmap.m`, not `nsmap_.m`. The 16 prefixes, in insertion order,
  are `a, c, cp, dc, dcmitype, dcterms, dgm, m, pic, r, sl, w, w14, wp, xml,
  xsi`. Versus pptx the map **ADDS** `dgm` (drawingml/diagram) and `w14`
  (office/word/2010/wordml), **REBINDS** `sl` to the schemaLibrary namespace,
  and **DROPS** every pptx-only prefix (`ct, ep, i, mo, mv, o, p, pd, pr, v,
  ve, w10, wne`).
- **`nspfxmap` is the subset function** (docx's name), not pptx's `namespaces`;
  there is no `nsmap` *alias* and no `nsuri` helper — docx `ns.py` has no such
  symbols, so the faithful surface omits them.
- **`OxmlElement` takes 3 arguments** — `(nsptag_str, attrs, nsdecls)` —
  where pptx's took 2 (`nsptag_str, nsmap`). See the `OxmlElement` section.
- **The OPC prefixes `ct` / `pr` / `r` used by `[Content_Types].xml` and
  `.rels`** live in a **separate** `nsmap` in `docx/opc/oxml.py` with its own
  `element_class_lookup` / `parse_xml` / `qn`. That is a distinct OPC-layer WP,
  **not** this one — so the `registry` here does not carry the 5 OPC classes,
  and this `nsmap` does not carry `ct` / `pr`.

The `w:`-default / root-declares-everything shape of `word/document.xml` is
handled correctly because the parser resolves each element / attribute prefix
from the **in-scope declaration stack**, not the fixed map — so root-declared
but non-`nsmap` prefixes (`wpc, mc, mo, mv, o, wp14, w10, wpg, wpi, wne, wps`)
resolve and round-trip verbatim. The fixed `nsmap` serves only the generation
path (`qn` / `createElement` / `OxmlElement`).

## Why an own parser (D-001)

Spike S1 proved `matlab.io.xml.dom` returns attributes **alphabetically
sorted** and its writer re-scrambles them, so byte-conservative round-tripping
of OOXML parts (attribute order identical to what Word wrote) is impossible
through the DOM. The pre-approved contingency **D-001** was adopted for
Mat2Doc: a purpose-built recursive-descent parser (`XmlParser`) that preserves
document order end-to-end, closing the parse → tree → serialize → bytes
round-trip, paired with a serializer byte-matched to lxml. `matlab.io.xml.dom`
remains available for read-only harness cross-checks, and is confirmed
**absent from the toolbox path** (Gate-3 D-001 clean-path check: the only two
matches under `+mat2doc\` are provenance comment lines in `parse_xml.m` and
`XmlParser.m`, zero code references).

## Deviation posture (pre-adopted, carried `mat2doc:`-namespaced)

Every divergence in this layer is a **recurrence of an already-adopted
Mat2Ppt deviation**, carried in the `mat2doc:` namespace per
`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`.
Gate-3 confirmed **0 new D-numbers**.

| D-number | What it is | Reachability |
|---|---|---|
| **D-001** | Own order-preserving OOXML-subset parser (`XmlParser`) + lxml-matched writer (`serialize_part_xml`) replace `matlab.io.xml.dom` (which sorts attributes). | Design baseline — the whole layer. |
| **D-006** | Parser **rejects `<!DOCTYPE>`** (`mat2doc:XMLSyntaxError`) where lxml silently ignores it — no DTD, `resolve_entities=False` could not honor DTD entities, anti-XXE. Asymmetric reject-vs-accept. | Dead on OOXML (no docx part carries a DTD). |
| **D-002** | ASCII `[0-9]` / hex digit grammar in numeric char-ref parsing: `&#xZZ;` / `&#1a;` are rejected before conversion (never `NaN`/`Inf`). | Dead on OOXML. |
| **D-005** | A non-text value passed to `set` / `text` / `tail` raises `mat2doc:TypeError` whose message reports the MATLAB type token (e.g. `'double'`) where lxml says `'int'`; exception class and template faithful. | Dead-path, API-invisible. |
| **D-serializer-nsdecl** | Verbatim-until-moved namespace-declaration model — parsed nodes serialize their decls VERBATIM; public moves clear the flag (lxml `moveNodeToDocument` reconciliation), reproduced at serialize time. | Byte-neutral on every generated part; restores L1 on parsed foreign parts. |
| **D-nsprefix-rewrite** | Residual foreign-file **prefix-rendering** divergence (L2): a parsed, never-moved element using the non-first of several same-URI bindings re-renders through the first binding — canonically equivalent, Office-safe. See the `a03` note under the serializer. | L2, dead on generation. |

## Gate-2 FIX-1 — the whitespace-only-tail-drop rule

The mso-auditor (Fable, cross-model) caught **FIX-1**, the one code defect in
this WP, in `XmlParser.m`'s `remove_blank_text` replication. libxml2's rule,
established by a 6-case live-lxml probe battery, is: **a whitespace-only
`text` / `tail` node is dropped unless the element's leading TEXT slot is
non-blank; a non-blank TAIL protects nothing.** The port's original
`finalizeBlankText_` let a non-blank *tail* latch a `mixed` flag and thereby
preserve later ws-only tails. The fix deletes the `else; mixed = true;` branch
(two lines) so that only the leading-text check can set `mixed`. Consequence,
re-verified vs live lxml and pinned in Gate-3 as `a09d` / `a09e` / `a09ctl`:
elements whose leading text is blank now **drop** their ws-only trailing tail
(byte-identical to lxml), while a non-blank-TEXT control **keeps** its ws tail
(no over-drop). All 16 real `s0001` parts stayed 16/16 byte-identical. No new
D-number.

The related **V-BLANK** class (char-ref / CDATA-derived whitespace, which lxml
keeps but the port folds into plain text and drops) is **accepted-unreachable**
— no real docx part carries inter-element char-ref/CDATA whitespace — and is
deliberately **excluded** from the frozen Gate-3 fixtures rather than pinned.

---

## Cross-cutting parser conventions

These hold across the whole read path (`parse_xml` / `XmlParser`):

- **OOXML subset (design.md §3).** The parser handles exactly what OOXML parts
  need: the XML declaration (parsed then **skipped** — the tree does not store
  it; the serializer re-emits the canonical declaration), elements and
  attributes **in document order** (H11 — the point of D-001), character data
  with the text/tail tri-state, the **five predefined entities** plus **decimal
  and hex numeric character references only** (ASCII digit grammar — D-002),
  CDATA sections (unwrapped to text), and namespace declarations recorded as
  ordered decls **stored where declared**. Anything outside the subset raises
  `mat2doc:XMLSyntaxError` — never silently mis-parsed.
- **Bytes ↔ text boundary (H2).** `parse_xml` decodes UTF-8 bytes once via
  `native2unicode(bytes, "UTF-8")` at entry, or accepts an already-decoded
  char/string literal (both currencies python-docx uses). Astral characters
  flow through as UTF-16 surrogate pairs, so a `&#128512;` char ref and a raw
  4-byte emoji (😀) both re-serialize as correct 4-byte UTF-8.
- **text / tail tri-state (H3).** An element with no character-data node has
  `text` / `tail` = `[]` (Python `None`), **not** `""`. `<w:p></w:p>` → text
  `[]` (→ serializer self-closes `<w:p/>`); char data before the first child
  becomes `element.text`, after a child becomes that child's `tail`.
- **Namespaces (H8).** Declarations are stored **where the document declared
  them**, ordered (prefix `""` = default `xmlns`). Element / attribute prefixes
  resolve from the **in-scope declaration stack** (innermost → outermost), not
  the fixed `nsmap` — so ancestor-declared and non-fixed prefixes resolve
  correctly, and a default-ns child inherits the URI without its own decl. The
  reserved `xml` prefix is always in scope, never recorded, never emitted.
- **`remove_blank_text=True` replication (incl. FIX-1).** For an element with
  ≥1 element child, a whitespace-only leading `text` is dropped, and a child's
  whitespace-only `tail` is dropped **unless the element's leading TEXT is
  non-blank** (a non-blank tail no longer protects — FIX-1). A leaf element
  keeps its text verbatim. This is a **byte-level NO-OP on all 16 real
  `s0001` parts** and matches lxml.
- **DOCTYPE rejected (D-006).** `<!DOCTYPE …>` raises `mat2doc:XMLSyntaxError`
  where lxml accepts-and-ignores — a documented, dead-on-OOXML divergence.
- **Class-registry dispatch (H10).** The parse walk applies the class registry
  keyed by the **resolved Clark name** at each element. The registry is
  **empty** at P1-2 (the `CT_*` WPs come later), so every node is a plain
  `XmlElement` now, but the hook is wired and proven.

## Cross-cutting serializer conventions

These hold across every serialized part (`serialize_part_xml`), all
byte-matched to lxml 5.3.0 (goldens `harness\common\golden\*_docx.bin`):

- **Byte-exact, no pretty-print (L1).** No indentation, no reordering, no
  `xmlns=""` un-declaration. Output is `1×N uint8` (Python `bytes`); the single
  text→bytes conversion is one `unicode2native(…, "UTF-8")` at the byte
  boundary (H2).
- **Declaration.** Every part begins with
  `<?xml version='1.0' encoding='UTF-8' standalone='yes'?>` — **single** quotes,
  exact case — followed by exactly one LF (golden `declaration_docx.bin`).
- **Attribute order.** Attributes are emitted in **insertion order**
  (golden `attr_order_docx.bin`), double-quoted.
- **Escaping (H7).** Text context escapes `&` `<` `>` and CR (as `&#13;`);
  attribute context additionally escapes `"` (`&quot;`), LF (`&#10;`) and TAB
  (`&#9;`). **Decimal** character references only (never hex); apostrophe is
  **never** escaped (values are always double-quoted); non-ASCII passes through
  as raw UTF-8. (goldens `escaping_text_docx.bin` / `escaping_attr_docx.bin`.)
- **Self-closing tri-state (H3).** `<w:p/>` (no space before `/>`) is emitted
  **only** when the element has no children **and** its raw text is `[]` (None).
  Text `""` serializes as `<w:t></w:t>`.
- **Namespace declarations (H8) — verbatim-until-moved (D-serializer-nsdecl).**
  Emitted on the owning element, in stored order, **before** ordinary
  attributes. A **verbatim** element (parsed, not yet re-moved —
  `isNsVerbatim_`) emits its stored declarations **verbatim**, never
  suppressed — this restores L1 byte-match with python-docx's own open→save on
  foreign parts carrying redundant nested decls. A **moved** element
  (everything Mat2Doc generates) **suppresses** a stored declaration whose URI
  is already reachable through a valid, non-shadowed ancestor binding — lxml's
  move-time reconciliation merge, reproduced at serialize time. Declarations on
  one element never suppress each other; suppression and prefix lookup are
  shadow-aware.

:::{note}
**Residual `D-nsprefix-rewrite` (L2, foreign-file only, dead on generation).**
Declaration *emission* is verbatim, but *prefix rendering* on a parsed,
never-moved element is still scope-derived: when the as-written tag/attr prefix
is not the port's scope-pick (an element using the non-first of several
same-URI bindings), the serialized bytes differ from lxml while remaining
**canonically equivalent** (expanded names identical, Office-safe). Gate-3's
sole L2 case is **`a03`**: input declares `xmlns:w` and `xmlns:q` to the *same*
URI with a child written `<q:r q:val="1"/>`; lxml keeps `<q:r q:val>`, the port
re-renders `<w:r w:val>` through the first same-URI binding (first byte diff at
offset 204). `pkgcompare`'s `canonical_diff` = **None** → canonically equal,
non-corrupting. Mapped to the already-SIGNED **`D-nsprefix-rewrite`**
(accepted 2026-07-18); it is frozen as a `verifyNotEqual` known-deviation
fixture, so any future byte change surfaces as a deliberate red. A full fix
needs an as-written-prefix parser data-model change (attributes are stored
under Clark keys, discarding as-written prefixes) — out of scope. See
`validation\summary\deviation_ledger.md`.
:::

---

## `nsmap`

**Syntax**

```matlab
map = mat2doc.oxml.nsmap()
```

**Description**

Returns a scalar struct whose field names are the **16** WordprocessingML
namespace prefixes and whose values are the corresponding namespace URIs
(string scalars). Field order preserves the Python source dict's insertion
order exactly (H11); iterate with `fieldnames(map)`. Built once into a
`persistent` variable. This is docx's **public** `nsmap` (no underscore
rotation) — the FIXED map used by `qn` / `nsdecls` / `NamespacePrefixedTag` /
`createElement` / `OxmlElement` for the generation path (the parse path
resolves prefixes from the in-scope declaration stack instead).

**Example**

```matlab
map = mat2doc.oxml.nsmap();
map.w                      % "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
string(fieldnames(map))'   % all 16 prefixes in insertion order
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::nsmap`*

---

## `pfxmap`

**Syntax**

```matlab
map = mat2doc.oxml.pfxmap()
```

**Description**

Reverse map: namespace URI → namespace prefix. Returns an Nx2 string array;
column 1 is the URI, column 2 the prefix, in the inversion order of `nsmap`.
Python's `pfxmap` is a dict keyed by URI; URIs are not valid MATLAB struct
field names, so it is represented as an ordered Nx2 string array (never
`containers.Map`). Its sole consumer is `NamespacePrefixedTag.from_clark_name`.

**Example**

```matlab
map = mat2doc.oxml.pfxmap();
uri = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
map(map(:, 1) == uri, 2)    % "w"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::pfxmap`*

---

## `nspfxmap`

**Syntax**

```matlab
s = mat2doc.oxml.nspfxmap(pfx1, pfx2, ...)
```

**Description**

Returns a 1×1 struct containing the subset prefix → URI mappings for the given
prefixes (any number, e.g. `nspfxmap("a", "r", "w")`) — docx's subset function
(pptx's equivalent `namespaces` is absent here). Field order is the argument
order (H11); a repeated prefix keeps its first position, exactly like a Python
dict comprehension. Zero arguments → struct with no fields (Python `{}`). An
unknown prefix raises `mat2doc:KeyError`.

**Example**

```matlab
s = mat2doc.oxml.nspfxmap("a", "r", "w");
string(fieldnames(s))'   % ["a"    "r"    "w"]  (argument order)
s.w                      % "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::nspfxmap`*

---

## `nsdecls`

**Syntax**

```matlab
decls = mat2doc.oxml.nsdecls(pfx1, pfx2, ...)
```

**Description**

Returns namespace-declaration attribute text for the given prefixes — one
`xmlns:<pfx>="<uri>"` declaration per prefix, single-space separated, in
argument order, values double-quoted. Zero arguments → `""`. Handy for adding
required namespace declarations to a tree root element. An unknown prefix
raises `mat2doc:KeyError`.

**Example**

```matlab
mat2doc.oxml.nsdecls("w", "r")
% 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
%  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::nsdecls`*

---

## `qn`

**Syntax**

```matlab
clark_name = mat2doc.oxml.qn(tag)
```

**Description**

Returns a Clark-notation qualified tag name for `tag` (e.g. `"w:p"`). `qn`
stands for *qualified name*. On the hot path, results are memoized in a
`persistent dictionary` keyed by the full prefixed tag; on a miss the value is
computed exactly as Python does — split on `":"`, look up `nsmap`, format Clark
form — so behavior (including `ValueError` on a malformed tag and `KeyError`
on an unknown prefix) is identical; only repeat-call cost differs. Errors are
never cached. The memo is a pure unordered cache over a fixed map, never
iterated (H11 not implicated).

**Example**

```matlab
mat2doc.oxml.qn("w:p")
% "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::qn`*

---

## `NamespacePrefixedTag`

**Syntax**

```matlab
obj = mat2doc.oxml.NamespacePrefixedTag(nstag)
obj = mat2doc.oxml.NamespacePrefixedTag.from_clark_name(clark_name)
```

**Description**

Value object that knows the semantics of an XML tag having a namespace prefix,
e.g. `"w:p"`. The Python class subclasses `str`; MATLAB `classdef` cannot
subclass `string` (sealed), so this is a **value class** (Python `str` is an
immutable value — value semantics match) that wraps the tag text in a private
property and exposes explicit `string(obj)` / `char(obj)` conversions to
recover str-ness (H2). Any ported call site that uses a `NamespacePrefixedTag`
directly *as* a string (formatting, comparison, dict key) must call
`string(obj)` explicitly.

Read-only properties (mirror the Python `@property`s):

| Property | Meaning |
|---|---|
| `clark_name` | `"{uri}local"` Clark notation |
| `local_part` | local part of the tag, e.g. `"foobar"` for `"f:foobar"` |
| `nsmap` | 1×1 struct with a single field `<prefix>` → URI |
| `nspfx` | namespace prefix, e.g. `"f"` |
| `nsuri` | namespace URI for the tag's prefix |

The static `from_clark_name` constructs an instance from Clark notation.

**Example**

```matlab
nsptag = mat2doc.oxml.NamespacePrefixedTag("w:p");
nsptag.clark_name   % "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
nsptag.local_part   % "p"
string(nsptag)      % "w:p"
t2 = mat2doc.oxml.NamespacePrefixedTag.from_clark_name(nsptag.clark_name);
string(t2)          % "w:p"  (round-trip)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/ns.py::NamespacePrefixedTag`*

---

## `parse_xml`

**Syntax**

```matlab
root = mat2doc.oxml.parse_xml(xml)
```

**Description**

Parse `xml` and return the root `mat2doc.oxml.XmlElement` of an
order-preserving tree — the MATLAB replacement for
`etree.fromstring(xml, oxml_parser)` under python-docx's config
`etree.XMLParser(remove_blank_text=True, resolve_entities=False)`. `xml` is
either UTF-8 **bytes** (`1×N uint8`, the on-disk part blob) or an
already-decoded **char row / string scalar** (a loose-element XML literal);
both call currencies python-docx uses are supported. The class registry is
applied at each element (registered `CT_*` classes instantiate; plain
`XmlElement` otherwise). Malformed input raises `mat2doc:XMLSyntaxError`. See
the cross-cutting parser conventions above.

**Example**

```matlab
xml  = "<w:p xmlns:w=""http://schemas.openxmlformats.org/" + ...
       "wordprocessingml/2006/main""><w:r><w:t>hi</w:t></w:r></w:p>";
root = mat2doc.oxml.parse_xml(xml);
root.tag                      % "{...wordprocessingml/2006/main}p"
% Bytes currency (an on-disk part blob) round-trips byte-for-byte:
blob  = uint8(unicode2native(xml, "UTF-8"));
bytes = mat2doc.oxml.serialize_part_xml(mat2doc.oxml.parse_xml(blob));
```

*Ported from python-docx v1.2.0: `src/docx/oxml/parser.py::parse_xml`
(lines 23–29; design-realization of the read side, D-001).*

---

## `XmlParser`

**Syntax**

```matlab
p    = mat2doc.oxml.XmlParser(text)
root = p.parse()
```

**Description**

The recursive-descent OOXML-subset parser **engine** (a `handle` class,
internal infrastructure): prolog/declaration skip, elements/attributes in
document order, char data + text/tail tri-state, entity/char-ref expansion
(ASCII digit grammar, D-002), CDATA unwrap, namespace resolution from the
in-scope stack, `remove_blank_text` replication (incl. FIX-1), DOCTYPE
rejection (D-006), malformed rejection, and per-element registry dispatch. Use
the package function `parse_xml` as the public entry point; this class is the
engine it drives. `text` is an **already-decoded** `(1,1)` string / char row
(`parse_xml` performs the bytes→string decode); `parse()` returns the root
`XmlElement`.

**Example**

```matlab
% Drive the engine directly (parse_xml is the normal entry point):
p    = mat2doc.oxml.XmlParser("<w:r xmlns:w=""urn:w""><w:t>hi</w:t></w:r>");
root = p.parse();
root.tag                 % "{urn:w}r"
kid  = root.to_array();  % child inherits prefix w from the scope stack
kid.tag                  % "{urn:w}t"
kid.text                 % "hi"
```

*Design-realization of the OOXML-subset read side (design.md §3, D-001); no
single Python def is its origin — the behavior realized is lxml's
`etree.fromstring(xml, oxml_parser)` read path
(`src/docx/oxml/parser.py::parse_xml` / `oxml_parser`, lines 18–29).*

---

## `XmlElement`

**Syntax**

```matlab
e = mat2doc.oxml.XmlElement(nsptag)
e = mat2doc.oxml.XmlElement(nsptag, nsmap)
```

**Description**

XML element node — the toolbox's lxml `_Element` replacement (design.md §3,
D-001). It is not the port of a single Python symbol; it realizes the lxml
`etree._Element` surface that python-docx v1.2.0 actually uses, matched to
lxml 5.3.0 probe evidence. Constructed with a prefixed tag (`"w:p"`) and an
optional scalar-struct `nsmap` of declarations to record on the element.

Tree shape: prefixed-tag identity plus resolved namespace URI (the `tag`
property presents lxml Clark form `"{uri}local"` so call sites comparing
against `qn(...)` work verbatim); an insertion-ordered attribute list (names
keyed exactly as lxml stores them — Clark form for namespaced attributes, plain
local name otherwise); ordered heterogeneous children (derives
`matlab.mixin.Heterogeneous` so `CT_*` subclasses mix in one vector); text/tail
as string scalar or `[]` (None, H3 — `""` is NOT None); parent backref; and an
ordered namespace-declaration list.

Key semantics:

- **Element identity = handle identity (H5).** The wrapper IS the node; `==`
  replicates lxml same-element identity (`Sealed` eq/ne/deepcopy on the
  Heterogeneous root). `append` / `insert` / `addnext` / `addprevious` **MOVE**
  an element out of its old parent, and its text/tail travel with it.
- **1-based indexing (H1).** `insert()` and `index()` are 1-based on this
  MATLAB surface; ported Python call sites shift (`elm.insert(i, c)` →
  `elm.insert(i+1, c)`, `parent.index(child)` needs `-1` for data uses).
- **Move clears the ns-verbatim flag.** A public move calls `clearNsVerbatim_`
  on the moved subtree — lxml's `moveNodeToDocument` reconciliation, which is
  what the serializer's verbatim-until-moved model keys on.
- **Type errors (D-005).** A non-text value to `set(name, value)` or assigned
  to text/tail raises `mat2doc:TypeError` reporting the MATLAB type token where
  lxml says `'int'` — dead-path, API-invisible.
- **`xpath` is NOT ported in this WP.** In python-docx `xpath()` lives on
  `BaseOxmlElement` (`xmlchemy.py`, P1-3) and the xpath engine is a later WP;
  this class has no `xpath` method yet — grep-clean, no dangling reference.

**Example**

```matlab
p = mat2doc.oxml.XmlElement("w:p");
r = mat2doc.oxml.XmlElement("w:r");
p.append(r);                                   % append MOVES r under p
tf = (p.find(mat2doc.oxml.qn("w:r")) == r);    % identity is true (H5)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/parser.py::oxml_parser`
(lxml `_Element` surface; design-realization, D-001).*

---

## `OxmlElement`

**Syntax**

```matlab
element = mat2doc.oxml.OxmlElement(nsptag_str)
element = mat2doc.oxml.OxmlElement(nsptag_str, attrs)
element = mat2doc.oxml.OxmlElement(nsptag_str, attrs, nsdecls)
```

**Description**

Return a "loose" element having the tag specified by `nsptag_str` (must contain
the standard namespace prefix, e.g. `"w:tbl"`). The result is an instance of the
registered custom class for the tag if one exists (`registry`), otherwise a
plain `XmlElement`, and carries the namespace declaration for its prefix
(`nsptag.nsmap`) so a loose subtree serializes with its own `xmlns:...`
declaration when its ancestors don't already declare it.

**The docx 3-argument signature** (`OxmlElement(nsptag_str, attrs, nsdecls)`)
differs from pptx's 2-arg form:

- `attrs` (2nd positional) — attributes to SET on the new element. Every docx
  call site keys them by a **Clark name** (`{qn("w:id"): str(id)}`); because a
  Clark name is not a valid MATLAB struct field name, `attrs` is an **Nx2
  string `[name value]`** array applied in row order (dict insertion order,
  H11). `[]` (default) = None.
- `nsdecls` (3rd positional) — a scalar struct prefix → URI; default `[]` (None)
  selects `nsptag.nsmap`, the single-prefix map of the tag.

No P1-2 caller passes `attrs` (all `attrs` callers are later `CT_*` WPs); the
signature is ported faithfully forward. For an attribute whose namespace is not
in `nsdecls`, lxml's `makeelement` would eagerly declare a generated prefix at
creation; this port instead invents the identical `nsN` at **serialize** time.
Every docx `OxmlElement` call passes an attribute whose namespace equals the
element's own prefix (e.g. `qn("w:id")` on `"w:*"`), always in `nsdecls`, so the
two are byte-equivalent on the used surface (the differing-namespace attribute
is accepted-unreachable).

**Example**

```matlab
commentId = 5;
end_ = mat2doc.oxml.OxmlElement("w:commentRangeEnd", ...
    [mat2doc.oxml.qn("w:id"), string(commentId)]);
% serializes to <w:commentRangeEnd xmlns:w="…" w:id="5"/>
```

*Ported from python-docx v1.2.0: `src/docx/oxml/parser.py::OxmlElement`
(lines 44–62).*

---

## `createElement`

**Syntax**

```matlab
element = mat2doc.oxml.createElement(nsptag_str)
element = mat2doc.oxml.createElement(nsptag_str, nsmap)
```

**Description**

Instantiate the element class registered for a tag — the makeelement-equivalent
constructor of the XML layer. For the prefixed tag `nsptag_str` (e.g. `"w:p"`)
it constructs an instance of the registered custom class when one exists
(`registry`), otherwise a plain `XmlElement` — mirroring lxml, where
`oxml_parser.makeelement` applies the class lookup and falls back to plain
`_Element` for unregistered tags. The optional scalar-struct `nsmap` stores the
namespace declarations on the new element, as
`makeelement(clark_name, nsmap=...)` does. The parse walk
(`XmlParser.parseElement_`) applies the same registry-then-fallback rule per
parsed element.

**Example**

```matlab
p = mat2doc.oxml.createElement("w:p", struct("w", ...
    "http://schemas.openxmlformats.org/wordprocessingml/2006/main"));
p.tag   % "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/parser.py::oxml_parser.makeelement`
with `element_class_lookup` applied (lines 18–20; design-realization — lxml has
no single equivalent Python def).*

---

## `registry`

**Syntax**

```matlab
cls_name = mat2doc.oxml.registry(clark_name)
```

**Description**

Custom element-class lookup: Clark tag name → MATLAB class name. Returns the
fully qualified MATLAB class name registered for the element tag `clark_name`
(`"{uri}local"`), or `""` when none is registered — the caller
(`createElement` / the parser) then falls back to plain `XmlElement`, mirroring
lxml's fallback to `_Element`. This is the MATLAB analogue of python-docx's
parse-time class lookup (`element_class_lookup =
etree.ElementNamespaceClassLookup()` on `oxml_parser`), populated by
`register_element_cls(tag, cls)` calls; lxml keys by (namespace URI, local
part), which is exactly the Clark name.

**Registration policy (H10):** an explicit static table, one line per Python
`register_element_cls` call, in Python source order, audited line-by-line. The
count guard is **120** `register_element_cls` calls total for docx (all in
`docx/oxml/__init__.py`). **At P1-2 the table is EMPTY** — no `CT_*` classes are
ported yet (`xmlchemy` / `BaseOxmlElement` is P1-3), so every parsed/created
element is a plain `XmlElement`, but the lookup hook is fully wired. The
SEPARATE 5-class OPC lookup in `docx/opc/oxml.py` is its own lookup and lands
with the OPC-layer WP.

**Example**

```matlab
cls = mat2doc.oxml.registry(mat2doc.oxml.qn("w:p"));   % "" until the CT_* WPs
```

*Ported from python-docx v1.2.0:
`src/docx/oxml/parser.py::register_element_cls` + `element_class_lookup`
(lines 18–41; registration blocks in `docx/oxml/__init__.py` pending their
`CT_*` WPs).*

---

## `serialize_part_xml`

**Syntax**

```matlab
xml_bytes = mat2doc.oxml.serialize_part_xml(part_elm)
```

**Description**

Serialize the `XmlElement` tree rooted at `part_elm` to UTF-8 file bytes,
including the XML declaration header — the MATLAB replacement for
`etree.tostring(part_elm, encoding="UTF-8", standalone=True)`. `part_elm` is a
`(1,1) mat2doc.oxml.XmlElement`; the result is `1×N uint8`, the exact bytes to
write to a `.xml` file (Python `bytes`).

`part_elm` is serialized as a **document root** (python-docx only ever passes
part roots and the `.rels` root to this `tostring` call). Like lxml's C-level
writer, serialization reads the **raw** stored element text, never a `CT_*`
subclass `.text` property shadow (D-005/D10). See the cross-cutting serializer
conventions above for the declaration, attribute-order, escaping, self-closing,
verbatim-until-moved namespace-declaration, and prefix-rendering rules, and the
`a03` / `D-nsprefix-rewrite` L2 note.

**Example**

```matlab
t = mat2doc.oxml.XmlElement("w:t", struct("w", mat2doc.oxml.nsmap().w));
t.text = "hello";
bytes = mat2doc.oxml.serialize_part_xml(t);
% <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
% <w:t xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">hello</w:t>
```

*Ported from python-docx v1.2.0: `src/docx/opc/oxml.py::serialize_part_xml`.
The serializer is XML-layer infrastructure per design.md §3; the Python symbol
lives in `opc/oxml.py` but the port is placed in `+mat2doc/+oxml/`.*
