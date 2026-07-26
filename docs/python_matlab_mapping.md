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
| `oxml/ns.py` — `nsmap`, `pfxmap`, `nspfxmap`, `nsdecls`, `qn`, `NamespacePrefixedTag` | `+mat2doc\+oxml\` (`nsmap.m`, `pfxmap.m`, `nspfxmap.m`, `nsdecls.m`, `qn.m`, `NamespacePrefixedTag.m`) | P1-2 | `nsmap` is **public** in docx (16 entries, **no** `_nsmap`→`nsmap_` rotation); adds `dgm`/`w14`, rebinds `sl`, drops the pptx-only prefixes. `nspfxmap` is docx's subset function (pptx's `namespaces`/`nsuri` absent). OPC prefixes `ct`/`pr`/`r` live in a separate `opc/oxml.py` map (OPC-layer WP). |
| `oxml/parser.py` — `parse_xml`, `OxmlElement`, `register_element_cls` + `element_class_lookup`, `oxml_parser.makeelement` | `+mat2doc\+oxml\` (`parse_xml.m`, `OxmlElement.m`, `registry.m`, `createElement.m`) | P1-2 | `OxmlElement` is **3-arg** `(nsptag_str, attrs, nsdecls)` (docx delta vs pptx 2-arg); `attrs` is an Nx2 string `[name value]` (Clark keys aren't struct fields). `registry` EMPTY at P1-2 (120-row target lands with the `CT_*` WPs). `createElement` = the makeelement-equivalent (design-realization). |
| lxml `_Element` surface + `etree.fromstring(xml, oxml_parser)` read side (design-realization, **D-001**) | `+mat2doc\+oxml\XmlElement.m`, `XmlParser.m` | P1-2 | Own order-preserving OOXML-subset parser + `_Element` tree node (Heterogeneous root, Sealed eq/ne/deepcopy). `matlab.io.xml.dom` sorts attributes → harness-only, never on the toolbox path. Rejects `<!DOCTYPE>` (**D-006**), ASCII digit char-ref grammar (**D-002**), FIX-1 ws-tail-drop rule. |
| `opc/oxml.py::serialize_part_xml` (byte-matched to lxml `etree.tostring`) | `+mat2doc\+oxml\serialize_part_xml.m` | P1-2 | Placed in `+oxml` though the Python symbol lives in `opc/oxml.py` (XML-layer infra, design.md §3). lxml-matched: single-quote declaration, insertion-order attrs, H7 escaping, verbatim-until-moved nsdecls (**D-serializer-nsdecl**); residual foreign-file prefix-rewrite L2 (**D-nsprefix-rewrite**, case `a03`). |
| `oxml/xmlchemy.py` — `BaseOxmlElement` (tree-ops + `xpath`), `OptionalAttribute` / `RequiredAttribute` descriptor closures | `+mat2doc\+oxml\BaseOxmlElement.m` | P1-3a | `classdef BaseOxmlElement < XmlElement`. Metaclass (`MetaOxmlElement`) generated members replaced by **explicit delegating engine methods** (design.md §2): tree-ops `first_child_found_in`/`insert_element_before`/`remove_all`, attribute engine `getAttrTyped`/`setAttrTyped`/`getAttrRequired`/`setAttrRequired`. **Setter deltas vs pptx** (docx form ported, **no new D#**): **D-delta-1** `value is None OR ==default`→remove; **D-delta-2** post-`to_xml` None→remove; **D-delta-3** post-`to_xml` None→`ValueError`. Attribute engine **dormant** (targets `+oxml\+simpletypes`=P3, no `CT_*` delegates yet). Child-element descriptor engine = **P1-3b**; `xml`/`serialize_for_reading` clean `notYetPorted` stub. |
| `oxml/xmlchemy.py::BaseOxmlElement.xpath` → lxml `_Element.xpath` subset (design-realization, **D-001**) | `+mat2doc\+oxml\evaluate_xpath.m` | P1-3a | Mini XPath-1.0 engine, re-ported from the SOLVED Mat2Ppt `+oxml` evaluator (WP5+WP5-C, no shared code). Closed subset (child paths, `.`/`..`/`//`, wildcard, 1-based positional/attr-eq/existence predicates, `@attr`/`text()` terminals, `ancestor::`/`self::`/`parent::` named tests, group unions); out-of-subset → `mat2doc:XPathError` (**never silently mis-evaluate**, design.md §3). **Gate-2 FIX-1** = the `self::NAME`/`parent::NAME` name-test filter (was ignored — pptx has no such call-sites). **★ VERIFY-2 coverage gap CLOSED by P1-3x** (see next row). Shared xpath deviations (F1/F2/F3/WPC-F1, D10 `text_raw_` bypass) carried `mat2doc:`, no new D#. |
| `oxml/xmlchemy.py::BaseOxmlElement.xpath` — the **docx** `.xpath()` call-site surface the pptx-derived engine did not cover (design-realization, **D-001**) | `+mat2doc\+oxml\evaluate_xpath.m` (+384/−53) | P1-3x | Extends the P1-3a mini-engine to docx's broader surface, each value/byte-verified vs lxml 5.3.0: **bare top-level unions** (`./w:p \| ./w:tbl`, doc-order + identity-dedup H11), **`not(path)`** / union-subpath predicates, **`following-sibling::`/`preceding-sibling::`/`preceding::`** axes (reverse-axis `position()` counts nearest-first, H1), **`position()=n`/`(…)[last()]`/group `(…)[1]`** predicates, **predicate attribute sub-paths** (`w:style[w:name/@w:val="…"]`), and the **`//@attr[n]` terminal** form (per-element positional → typed-empty **string**, lxml encode() convention; the sole member removed from the F2 raise-set). Adopts the shared xpath/parser deviations already namespaced `mat2doc:` — **no new D#**. VERIFY-2 gap closed; R01–R09 pins flipped (`references\p1_3x\flip\flip_types.json`). `decision_2026-07-25_mat2doc_xpath_engine_extension.md`. |
| `oxml/xmlchemy.py` — child-element descriptors `ZeroOrOne` / `ZeroOrMore` / `OneAndOnlyOne` / `OneOrMore` / `Choice` / `ZeroOrOneChoice` (the `_BaseChildElement` family) | `+mat2doc\+oxml\BaseOxmlElement.m` (+1 additive `methods` block, 11 engine members) | P1-3b | The metaclass-generated child accessors (`get.x`/`x_lst`/`_new_x`/`_insert_x`/`_add_x`/`add_x`/`get_or_add_x`/`_remove_x`/`get_or_change_to_x`/`_remove_eg_x`) → **explicit engine methods** `getChild`/`getRequiredChild`/`getChildList`/`newChild`/`insertChildInSequence`/`addChild`/`getOrAddChild`/`removeChild`/`firstChildFoundIn`/`removeChildren`/`getOrChangeToChild` every future `CT_*` member delegates to (design.md §2). **H11 successor-ordering** rides the P1-3a `insert_element_before` (insert before first present successor, else append); H5 get-or-add same-handle; H9 materialized `1x0` list vs H3 `[]`/`InvalidXmlError` tri-state. **D-delta-4** (docx form, **no new D#**): docx `ZeroOrMore` also generates a **public `add_x`** (xmlchemy.py 536 → `_BaseChildElement` 340–352) routing through the same `addChild` — engine-neutral, a **genoxml-scaffolder obligation** (24 docx `ZeroOrMore` sites). `Choice`/`ZeroOrOneChoice` = **dead code in docx v1.2.0** (parity-only, synthetic coverage; no byte-oracle). **Completes xmlchemy.** Fold-forward: package-level L1 lands at the first `CT_*` WP; VERIFY-3 heterogeneous-tree re-verify at first CT registration. |

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
| ordered `dict` prefix→URI (`nsmap`, `nspfxmap`) | scalar `struct`, insertion-ordered fields (H11); iterate `fieldnames`; never `containers.Map` | P1-2 |
| `dict` keyed by a URI or Clark name (not a valid struct field name) — `pfxmap`, `OxmlElement` `attrs` | **Nx2 `string` array** `[key, value]`, document order (H11) | P1-2 |
| `nsmap[pfx]` / dict `KeyError` | `isfield` guard → `error("mat2doc:KeyError", "%s", …)` | P1-2 |
| `etree.fromstring(xml, oxml_parser)` | `mat2doc.oxml.parse_xml(xml)` driving `XmlParser.parse` — own OOXML-subset reader (D-001), order-preserving | P1-2 |
| `oxml_parser.makeelement(clark, attrib, nsmap)` | `createElement` (+ `.set` per attr) with registry-then-`XmlElement` fallback | P1-2 |
| lxml `_Element` identity / `el1 is el2` (H5) | `handle` identity; `Sealed` `==`/`~=` on the Heterogeneous `XmlElement` root | P1-2 |
| `el.insert(i, c)` / `parent.index(child)` (0-based) | 1-based on the MATLAB surface (H1): `insert(i+1, c)`, `index(child)-1` for data | P1-2 |
| `copy.deepcopy(el)` / `__deepcopy__` | `XmlElement.deepcopy` (Sealed, structural copy of own decls + verbatim tag URI) | P1-2 |
| `etree.tostring(el, encoding="UTF-8", standalone=True)` | `mat2doc.oxml.serialize_part_xml(el)` — own writer, byte-matched to lxml (single `unicode2native` at the boundary, H2/H7) | P1-2 |
| bytes ↔ text at the XML boundary | `native2unicode` / `unicode2native('UTF-8')` only (H2); astral chars via surrogate pairs → 4-byte UTF-8 | P1-2 |
| `MetaOxmlElement` generating per-attribute get/set properties from `OptionalAttribute`/`RequiredAttribute` descriptors | explicit engine methods `getAttrTyped`/`setAttrTyped`/`getAttrRequired`/`setAttrRequired` on `BaseOxmlElement`, each `CT_*` delegates (design.md §2 — no MATLAB metaclass hook) | P1-3a |
| descriptor `simple_type` object / `_clark_name` (`BaseAttribute`) | `resolveTypeCls_` (bare `"ST_String"`→`+oxml\+simpletypes.<name>`; dotted `"mat2doc.enum.*"` verbatim) + `attrClarkName_` (prefixed→`qn`, plain→verbatim) | P1-3a |
| `attr is None` / `del obj.attrib[clark]` in a descriptor setter | `isequal(value,[])` (H3: `""`≠None) + `has_attrib`/`remove_attrib` guard; the three docx setter deltas (D-delta-1/2/3) live in `setAttrTyped`/`setAttrRequired` | P1-3a |
| `raise InvalidXmlError(...)` (required attr absent) / required-setter `ValueError` | `error("mat2doc:InvalidXmlError", ...)` (msg uses element Clark `tag`) / `error("mat2doc:ValueError", ...)` with `valueRepr_` best-effort `str(value)` (D-005 message-token class) | P1-3a |
| lxml `_Element.xpath(expr, namespaces=nsmap)` (docx injects the **public** `nsmap`) | `BaseOxmlElement.xpath(expr[, ns])` → `mat2doc.oxml.evaluate_xpath` (own mini-engine, D-001); default `ns = nsmap()` (public map, no `_nsmap`); typed-empty node-set on no match (H3), never `[]` | P1-3a |
| positional XPath predicate `[1]` (XPath is 1-based) | 1-based verbatim (H1 — never shift): `w:p[1]` selects the first match | P1-3a |
| lxml `text()` node-set (C-level text nodes) | `text_raw_()` self-text + each child's `.tail`, doc-ordered with tail-after-subtree (WPC-F1); bypasses a `CT_*` `getText_` shadow (D10, F1) | P1-3a |
| lxml bare union `A \| B \| …` (node-**set** — merged, doc-ordered, de-duplicated) | union AST node → `pathNodeset` per operand → `cellToElemArray(docSortDedupe(...))`: document order + identity-dedup (H11 union trap; `[p,p,tbl,p,p]` interleaved, never operand-grouped) | P1-3x |
| lxml `not(path)` predicate (keep where the sub-path is an empty node-set) | `notExists` predicate → `isempty(subPathNodeset(...))`; sub-path may itself be a relative union (`[not(self::w:br \| self::w:cr)]`) | P1-3x |
| lxml reverse axes `preceding-sibling::`/`preceding::` (`position()` counts nearest-first) | matches emitted **nearest-first** for the positional predicate, then `docSortDedupe` restores document order for the final node-set (H1 — position by axis direction, not doc order) | P1-3x |
| lxml `position()=n` / `(…)[last()]` / grouped `(…)[1]` | `pos`/`last` predicate kinds; a `(…)` group re-numbers positions over the whole doc-ordered grouped node-set (1-based, H1) | P1-3x |
| lxml predicate attribute sub-path `[w:name/@w:val="…"]` (existential node-set `=`) | `pathEq` predicate → `subPathStringValues` (relative sub-path terminating in `@attr`/`text()`) → `any(vals == literal)` | P1-3x |
| lxml `//@attr[n]` — positional predicate on an attribute-terminal step (one attr per element) | applied **per context element** to that element's own attr node-set (position 2 never exists → empty); result is a **typed-empty `string`** (`string.empty(1,0)`), lxml encode() convention (H3) — the only member removed from the F2 raise-set | P1-3x |
| `MetaOxmlElement` generating child accessors from `ZeroOrOne`/`ZeroOrMore`/`OneAndOnlyOne`/`OneOrMore`/`Choice`/`ZeroOrOneChoice` descriptors | 11 explicit engine methods on `BaseOxmlElement` (`getChild` … `getOrChangeToChild`); each `CT_*` generated member delegates (design.md §2 — no MATLAB metaclass hook) | P1-3b |
| successor slice `_tag_seq[N:]` (child sorts before every schema-later tag) → `obj.insert_element_before(child, *successors)` | `insertChildInSequence(child, successors)` expanding the class `SUCCESSORS` Constant via `num2cell`; **H1 0→1 base shift applied ONCE at the slice declaration** (`_tag_seq[N:]`→`tagSeq(N+1:end)`), never in the engine (H11) | P1-3b |
| `get_or_add_x`: `child = getter; if child is None: child = _add_x()` (returns the live present child, `is`-identity) | `getOrAddChild(tag, successors)`: `getChild` → `addChild` when `isequal(child,[])`; a present child returns the **same handle** back-to-back (H5) | P1-3b |
| `x_lst` `findall` **list** vs single `get.x` `find` **None** vs `OneAndOnlyOne` **raise** | `getChildList` → materialized `(1,N)`/`1x0` typed `XmlElement` (H9, snapshot-safe remove-during-iter); `getChild` → `[]` (None); `getRequiredChild` → `mat2doc:InvalidXmlError` (byte-exact RST-backtick msg) — the **tri-state** absent encoding (H3) | P1-3b |
| `ZeroOrMore` public `add_x()` (docx `_add_public_adder`, xmlchemy.py 536; pptx: `OneOrMore` only) | routes through the same `addChild` primitive as `_add_x` — **D-delta-4**, engine-neutral (**no new D#**); genoxml scaffolder emits the public delegator at each of the 24 docx `ZeroOrMore` sites | P1-3b |

## Namespace policy

Mat2Doc uses the `mat2doc:` error-identifier namespace and the `mat2doc.*`
package namespace throughout — **no shared code with Mat2Ppt**. Designs common
to both toolboxes (the `Length` family, `pyIntArg`) are re-implemented, not
copied. The P0 same-package name-collision scan found **zero** collisions, so
**no symbol was renamed** (FLAG-3-docx resolved by convention).
