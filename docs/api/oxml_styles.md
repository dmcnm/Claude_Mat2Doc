---
title: "mat2doc.oxml.styles — the styles-oxml element layer (CT_Styles · CT_Style · CT_LatentStyles · CT_LsdException · styleId_from_name + CT_DecimalNumber)"
---

# `mat2doc.oxml.styles` — the styles-oxml element layer (`CT_Styles` + `CT_Style` + `CT_LatentStyles`/`CT_LsdException` + `styleId_from_name`, with the shared `CT_DecimalNumber`)

Ported from python-docx v1.2.0 `src/docx/oxml/styles.py` (the module function
`styleId_from_name` plus four element classes — `CT_LatentStyles`,
`CT_LsdException`, `CT_Style`, `CT_Styles`), in package
`+mat2doc/+oxml/+styles/`, **and** the shared leaf
`src/docx/oxml/shared.py::CT_DecimalNumber`, in package
`+mat2doc/+oxml/+shared/`, plus the **12 styles-block `register_element_cls`
rows** they register (`src/docx/oxml/__init__.py:138-149`) and the closed
`w:outlineLvl → CT_DecimalNumber` deferral (`:244`).

:::{note}
This is the **first work package of the styles chain** — P4-6 (oxml/styles) →
P4-7a (the `styles`/`style` API) → P4-7b (latent styles + the `add_heading`
un-stub) → **M2**. It ports the `<w:styles>` element tier that the Styles API
(P4-7) reads; it registers the styles-block tags so the shipped 349 KB
`styles.xml` now parses through these classes, **byte-neutrally** (M1 stays
17/17). It is not itself an M2 or COM milestone — the M2 COM oracle fires at
P4-7b, when `add_heading` forces the whole chain.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The two-package layout

`styles.py` (a top-level `oxml/` module) lands in `+oxml\+styles`;
`shared.py::CT_DecimalNumber` (the third leaf of `oxml/shared.py`, after the
`CT_OnOff`/`CT_String` pair the P4-1a font WP folded in) lands in `+oxml\+shared`
— following the established module→subpackage convention. `styleId_from_name` is
a **module-level function** in `styles.py` (not a method), so it maps to a
package function `+oxml\+styles\styleId_from_name.m` (design.md §1: one package
function per file). Neither `+oxml\+styles` nor `+oxml\+shared` collides with the
top-level `mat2doc.shared` (the `Length` family / `RGBColor`) — they are distinct
packages.

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `oxml/styles.py` | `+mat2doc\+oxml\+styles\` | `styleId_from_name`, `CT_LatentStyles`, `CT_LsdException`, `CT_Style`, `CT_Styles` |
| `oxml/shared.py` (its third leaf) | `+mat2doc\+oxml\+shared\` | `CT_DecimalNumber` |

(id-registry-parse-path)=
## The registry-adding, byte-critical parse path — M1 stays byte-identical

This WP registers the **12 styles-block tags** atomically, exactly as
python-docx registers them (`oxml/__init__.py:138-149`), and closes the
`w:outlineLvl → CT_DecimalNumber` deferral the P4-2 parfmt WP left open (`:244`):

| tag | class | tag | class |
|---|---|---|---|
| `w:basedOn` | `shared.CT_String` | `w:semiHidden` | `shared.CT_OnOff` |
| `w:latentStyles` | `styles.CT_LatentStyles` | `w:style` | `styles.CT_Style` |
| `w:locked` | `shared.CT_OnOff` | `w:styles` | `styles.CT_Styles` |
| `w:lsdException` | `styles.CT_LsdException` | `w:uiPriority` | `shared.CT_DecimalNumber` (first registration) |
| `w:name` | `shared.CT_String` | `w:unhideWhenUsed` | `shared.CT_OnOff` |
| `w:next` | `shared.CT_String` | `w:qFormat` | `shared.CT_OnOff` |

Registering these tags changes the **parse path** of the part that already
exists: the shipped `default.docx` `word/styles.xml` (349 458 B) contains
**164 `<w:style>`** and **137 `<w:lsdException>`** blocks; once the rows are live
those blocks parse to `CT_Style` / `CT_LatentStyles` / `CT_LsdException` /
`CT_String` / `CT_OnOff` / `CT_DecimalNumber` instead of a generic `XmlElement`.
Because every `CT_*` class extends `BaseOxmlElement` and reserializes through the
same `serialize_part_xml` walk with **no serialization override**, registering
the tags changes only the parsed node **class**, never its **bytes**.

The **M1 17/17 byte-neutrality sweep holds**: `mat2doc.Document().save()` → unzip
→ all 17 parts byte-identical to the frozen `references\s0001` reference, with
`word/styles.xml` (349 458 B, SHA-256 `02d71a68…e384`, independently re-derived a
third time) L1 byte-identical (Gate-3 `pkgcompare` L0 PASS + 16 XML L1 + 1 bin).
The 349 KB `styles.xml` also round-trips parse→serialize L1. This is the standing
obligation every registry-adding `CT_*` WP inherits.

:::{note}
**F-3 (numbering tags stay generic — a P6/P8 forward flag).** `w:numId` (6×) and
`w:ilvl` (1×) **do** appear in the template `styles.xml`, inside
`w:style/w:pPr/w:numPr`; they parse as **generic `XmlElement`** here (unregistered)
and round-trip byte-identically, so leaving them deferred is byte-safe today.
Likewise the table `CT_DecimalNumber` tags (`w:gridSpan`/`w:gridBefore`/…). When
P6/P8 registers `w:numId`/`w:ilvl` → `CT_DecimalNumber`, those tags will then
transit `CT_DecimalNumber` on the `styles.xml` parse path — so that WP **must
re-run the M1 `word/styles.xml` byte gate**, not only its own parts' gate.
:::

---

## `CT_Styles`

**Syntax**

```matlab
styles = mat2doc.oxml.OxmlElement("w:styles");   % a CT_Styles (registered)
s = styles.add_style_of_type("Heading 1", ...
        mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH, true);
found = styles.get_by_id("Heading1");            % the CT_Style, or [] if absent
def   = styles.default_for(mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH);
```

**Description**

The `<w:styles>` element — the **root** of a styles part (`styles.xml`). It holds
an optional `<w:docDefaults>`, an optional `<w:latentStyles>`, and zero-or-more
`<w:style>` children, and provides the lookup surface the Styles API (P4-7a)
reads:

- **`get_by_id(styleId)`** — the `<w:style>` child whose `@w:styleId` equals
  `styleId`, via the attribute-equality XPath `w:style[@w:styleId="<id>"]`, taking
  the first match (`next(iter(xpath), None)` → `res(1)`, H1) or `[]` (H3) when
  none.
- **`get_by_name(name)`** — the `<w:style>` child with a `<w:name>` grandchild
  valued `name`, via the predicate-attribute-sub-path XPath
  `w:style[w:name/@w:val="<name>"]` (a documented `evaluate_xpath` example form),
  first match or `[]`.
- **`default_for(style_type)`** — the **last** `<w:style>` in document order with
  the matching `@w:type` and a truthy `@w:default`, or `[]` when none. The
  match test ports `s.type == style_type and s.default` as
  `isequal(s.type, style_type) && ~isequal(d, []) && d` — `isequal` gives Python
  `None == member` → `False` safely, and the `~isequal(d,[]) && d` triad
  reproduces the None/`False`/`True` truthiness (H4); **last-in-document-order**
  wins (`matches{end}`, Python `[-1]`).
- **`add_style_of_type(name, style_type, builtin)`** — appends a new `<w:style>`,
  sets `type`, sets `customStyle = [] if builtin else true` (Python
  `None if builtin else True`), sets `styleId = styleId_from_name(name)`, and
  sets `name_val = name`.

The two child descriptors are `latentStyles` (`ZeroOrOne`, successors
`_tag_seq[2:]` → `TAG_SEQ(3:end) = ["w:style"]`, the H1 base shift applied once)
and `style` (`ZeroOrMore`, `successors=()` → append). Being a docx `ZeroOrMore`,
`style` also generates the **public `add_style`** (D-delta-4), which
`add_style_of_type` calls.

**Example**

```matlab
styles = mat2doc.oxml.OxmlElement("w:styles");
s = styles.add_style_of_type("Heading 1", ...
        mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH, true);
disp(s.styleId);                                 % "Heading1"
disp(s.name_val);                                % "Heading 1"  (the stored display name)
disp(styles.get_by_id("Heading1") == s);         % true (H5 same handle)
disp(styles.get_by_name("Heading 1").styleId);   % "Heading1"  (lookup by display name)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/styles.py::CT_Styles`*

---

(id-ct_style-h11)=
## `CT_Style`

**Syntax**

```matlab
s = mat2doc.oxml.OxmlElement("w:style");         % a CT_Style (registered)
s.type    = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
s.styleId = "Heading1";
s.name_val = "heading 1";                        % <w:name w:val="heading 1"/>
s.basedOn_val = "Normal";                        % <w:basedOn w:val="Normal"/>
```

**Description**

The `<w:style>` element — a single style definition. It carries **four optional
attributes** and **ten `ZeroOrOne` child descriptors**, plus the `.._val`
`@property` members that read/write the `w:val` of the corresponding child.

**The four attributes** (`styles.py:131-138`), all default `None` (`[]`):
`type` (`OptionalAttribute` of `WD_STYLE_TYPE`, the enum named by its
fully-qualified class so `resolveTypeCls_` dispatches to `+enum`), `styleId`
(`ST_String`), `default` (`ST_OnOff`), `customStyle` (`ST_OnOff`). Setting `[]`
short-circuits and removes the attribute (D-delta-1).

**`styleId` vs. `name`.** These are distinct: `styleId` is the internal
identifier used by `get_by_id`, `base_style`, and `next_style`; `name_val` is the
human-facing display name carried by the `<w:name>` child. `add_style_of_type`
derives the id from the name via [`styleId_from_name`](#id-styleid-from-name).

**The `basedOn` / `next` chains.** `base_style` returns the sibling `CT_Style`
this one is based on — `getparent().get_by_id(basedOn.val)` — or `[]` when there
is no `<w:basedOn>` or the target id is not found; `next_style` returns the
sibling identified by `<w:next>/@w:val`, or `[]`.

:::{note}
**No `link` chain in v1.2.0 (fidelity note).** The brief's "basedOn/next/link
chains" phrasing is aspirational: python-docx v1.2.0 lists `w:link` only in the
`_tag_seq` (position 5) and declares **no** `link` `ZeroOrOne` descriptor (nor
`aliases`/`autoRedefine`/`hidden`/`personal*`/`rsid`/`tblPr`/`trPr`/`tcPr`/
`tblStylePr`). The port faithfully has **no** `link` accessor — `w:link` (38
occurrences in the template `styles.xml`) parses as a generic `XmlElement`
(unregistered) and round-trips byte-identically.
:::

(id-ct_style-tristate)=
**The H3 tri-state `.._val` members.** The Python getters split into two
absent-return conventions, ported exactly:

- `basedOn_val` / `name_val` / `uiPriority_val` → `[]` (None) when the child is
  absent;
- `locked_val` / `qFormat_val` / `semiHidden_val` / `unhideWhenUsed_val` → the
  logical `false` when absent (Python `return False`, `styles.py:172-269`).

The boolean-ish setters port `if bool(value) is True:` / `if bool(value):` as
`~isequal(value, []) && value` (guard None first, then MATLAB logical coercion),
each `remove_x_()`-ing the child first; `qFormat_val = true` adds a **bare**
`<w:qFormat/>` (no `@w:val`), matching Python's `self._add_qFormat()` with no
`.val` set. `None` is `isequal(x, [])`; `""` is a real string, never `[]`.

(id-ct_style-successor)=
**The H11 child-order successor sequence (the M2-critical pattern).** `CT_Style`
declares 10 `ZeroOrOne` descriptors over a single **22-entry `_tag_seq`**
(`styles.py:95-118`, ported verbatim as the Constant `TAG_SEQ`). OOXML mandates
that a `<w:style>`'s children appear in schema order regardless of creation order;
the re-sort mechanism is the standard one (each descriptor carries the schema
slice `_tag_seq[N:]` → `TAG_SEQ(N+1:end)`, the H1 base shift applied once at the
slice, and `insertChildInSequence` lands a new child *before the first
schema-later sibling already present*, else appends). The 10 slices — several
**non-contiguous**, jumping across the tags that carry no descriptor
(`w:aliases` at index 2; `w:link`/`w:autoRedefine`/`w:hidden` at indices 5-7;
`w:personal*`/`w:rsid` at indices 13-16):

| descriptor | own 1-based idx | Python | MATLAB slice | child class |
|---|---|---|---|---|
| `name` | 1 | `_tag_seq[1:]` | `TAG_SEQ(2:end)` | `CT_String` |
| `basedOn` | 3 | `_tag_seq[3:]` | `TAG_SEQ(4:end)` | `CT_String` |
| `next` | 4 | `_tag_seq[4:]` | `TAG_SEQ(5:end)` | `CT_String` |
| `uiPriority` | 8 | `_tag_seq[8:]` | `TAG_SEQ(9:end)` | `CT_DecimalNumber` |
| `semiHidden` | 9 | `_tag_seq[9:]` | `TAG_SEQ(10:end)` | `CT_OnOff` |
| `unhideWhenUsed` | 10 | `_tag_seq[10:]` | `TAG_SEQ(11:end)` | `CT_OnOff` |
| `qFormat` | 11 | `_tag_seq[11:]` | `TAG_SEQ(12:end)` | `CT_OnOff` |
| `locked` | 12 | `_tag_seq[12:]` | `TAG_SEQ(13:end)` | `CT_OnOff` |
| `pPr` | 17 | `_tag_seq[17:]` | `TAG_SEQ(18:end)` | `CT_PPr` (P4-2) |
| `rPr` | 18 | `_tag_seq[18:]` | `TAG_SEQ(19:end)` | `CT_RPr` (P4-1a) |

Gate-2/Gate-3 built the 10-descriptor set in **reverse** order and serialized:
the child sequence came out `w:name, w:basedOn, w:next, w:uiPriority,
w:semiHidden, w:unhideWhenUsed, w:qFormat, w:locked, w:pPr, w:rPr` — canonical
schema order, byte-identical to the python-docx oracle.

**`delete_style()` — see [the H17 note below](#id-h17-delete).**

**Example**

```matlab
s = mat2doc.oxml.OxmlElement("w:style");
s.get_or_add_rPr();                              % w:rPr added FIRST
s.get_or_add_name();                             % w:name added AFTER
order = arrayfun(@(e) string(e.nsptag_str), s.xpath("./*"));
disp(order);   % "w:name"  "w:rPr"  -- w:name re-sorted BEFORE w:rPr (H11)
disp(s.locked_val);                              % 0 (false — <w:locked> absent, H3)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/styles.py::CT_Style`*

---

## `CT_LatentStyles`

**Syntax**

```matlab
ls = mat2doc.oxml.OxmlElement("w:latentStyles"); % a CT_LatentStyles (registered)
ls.set_bool_prop("defQFormat", true);            % <w:latentStyles w:defQFormat="1"/>
lsd = ls.add_lsdException(); lsd.name = "Normal";
tf = ls.bool_prop("defSemiHidden");              % false when the attr is absent
```

**Description**

The `<w:latentStyles>` element — behavior defaults for latent styles, holding
zero-or-more `<w:lsdException>` children that each override those defaults for a
named latent style. It carries **six optional attributes** (all default `None`):
`count` / `defUIPriority` (`ST_DecimalNumber`, int) and `defLockedState` /
`defQFormat` / `defSemiHidden` / `defUnhideWhenUsed` (`ST_OnOff`, bool-ish). The
one child descriptor `lsdException` is a `ZeroOrMore` (`successors=()` → append)
that also generates the public `add_lsdException` (D-delta-4).

- **`bool_prop(attr_name)`** returns the logical `false` (not `[]`) when the named
  attribute is absent — the "effective default" the LatentStyles API reads (`[]`
  → `false`, H3).
- **`get_by_name(name)`** returns the matching `<w:lsdException>` child (first
  match, H1) or `[]` (H3).
- **`set_bool_prop(attr_name, value)`** — see [the F-1 fix below](#id-f1-boolnone).

**Example**

```matlab
ls = mat2doc.oxml.OxmlElement("w:latentStyles");
ls.count = 371;                                  % <w:latentStyles w:count="371"/>
disp(ls.bool_prop("defQFormat"));                % 0 (false — absent)
ls.set_bool_prop("defQFormat", true);
disp(ls.bool_prop("defQFormat"));                % 1 (true)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/styles.py::CT_LatentStyles`*

---

## `CT_LsdException`

**Syntax**

```matlab
lsd = mat2doc.oxml.OxmlElement("w:lsdException"); % a CT_LsdException (registered)
lsd.name = "Normal";                              % REQUIRED @w:name
lsd.semiHidden = true;   % <w:lsdException w:name="Normal" w:semiHidden="1"/>
tf = lsd.on_off_prop("semiHidden");               % true / false / [] (absent)
```

**Description**

The `<w:lsdException>` element — override visibility behaviors for one named
latent style; a child of `<w:latentStyles>` with **no child elements**. It carries
a **required** `w:name` (`ST_String`; `getAttrRequired` raises
`mat2doc:InvalidXmlError` when absent) plus **five optional** attributes:
`locked` / `qFormat` / `semiHidden` / `unhideWhenUsed` (`ST_OnOff`) and
`uiPriority` (`ST_DecimalNumber`), each default `None` (`[]`). `on_off_prop(attr)`
returns the named attribute value (`obj.(attr_name)`, the `getattr` analogue —
`[]` when absent); `set_on_off_prop(attr, value)` sets it.

**`delete_lsd_exception()` — see [the H17 note below](#id-h17-delete).**
`CT_LsdException.delete` is the **first** Python `delete()` element-removal method
ported in the whole project.

**Example**

```matlab
lsd = mat2doc.oxml.OxmlElement("w:lsdException");
lsd.name = "heading 1";
lsd.uiPriority = 9;              % <w:lsdException w:name="heading 1" w:uiPriority="9"/>
lsd.set_on_off_prop("qFormat", true);
disp(lsd.on_off_prop("qFormat"));                 % 1 (true)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/styles.py::CT_LsdException`*

---

## `CT_DecimalNumber`

**Syntax**

```matlab
up = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", 9);
n  = up.val;                                      % 9 (double holding an exact int)
```

**Description**

The `<w:val>`-bearing decimal-number element (`src/docx/oxml/shared.py`), the
third `shared.py` leaf after `CT_OnOff` / `CT_String`. `val` is a
`RequiredAttribute` of `ST_DecimalNumber` (an `xsd:int`): `getAttrRequired` raises
`mat2doc:InvalidXmlError` when `@w:val` is absent, the setter always writes it
(never removes) and range-validates via `ST_DecimalNumber.to_xml`, and `from_xml`
parses the XML integer literal exactly as Python `int(str_value)` (H6). The static
`new(nsptagname, val)` builds a loose element of the given tag with
`@w:val = str(val)` set **directly** — routed through
`mat2doc.shared.pyStr(val, "int")` (H14), **not** through the descriptor setter,
so (like Python's `str()`) it does **not** range-validate.

**First registration.** This WP is the first to register `CT_DecimalNumber` — at
`w:uiPriority` (styles) and the now-closed `w:outlineLvl` deferral (parfmt). The
numbering/table `CT_DecimalNumber` registrations (`w:numId`/`w:ilvl`/`w:gridSpan`/…)
remain **deferred to P6/P8** (see [the F-3 forward flag above](#id-registry-parse-path)).

**Example**

```matlab
up = mat2doc.oxml.shared.CT_DecimalNumber.new("w:uiPriority", 9);
disp(up.val);                                     % 9
disp(string(up.xml));                             % <w:uiPriority w:val="9" .../>
```

*Ported from python-docx v1.2.0: `src/docx/oxml/shared.py::CT_DecimalNumber`*

---

(id-styleid-from-name)=
## `styleId_from_name`

**Syntax**

```matlab
id = mat2doc.oxml.styles.styleId_from_name("heading 1");   % "Heading1"
id = mat2doc.oxml.styles.styleId_from_name("My Custom Style");  % "MyCustomStyle"
```

**Description**

The module function that returns the style id corresponding to a style `name`,
honouring the special-case names (`caption`, `heading 1` … `heading 9`). The
default for any name **not** in the special-case table is `name` with its spaces
removed (`strrep(name, " ", "")`, Python `name.replace(" ", "")`).

:::{warning}
**H15 — the lookup is Python `dict.get`, case-SENSITIVE.** The special-case table
is keyed on the **lowercase** spellings only; the port is a case-sensitive
`switch` with an `otherwise` (default) arm. Only the literal lowercase keys
(`"caption"`, `"heading 1"`, …) hit the table; any other spelling
(`"Heading 1"`, `"Normal"`) falls to the space-strip default — exactly as Python's
`dict.get(name, name.replace(" ", ""))` does. So `styleId_from_name("Heading 1")`
returns `"Heading1"` **via the default arm** (the capital-H key misses the
lowercase table entry, then the space is stripped) — coincidentally the same
result the table would give, and matching python-docx byte-for-byte. Gate-3
verified the full 11-vector against the live oracle.
:::

**Example**

```matlab
disp(mat2doc.oxml.styles.styleId_from_name("heading 1"));   % "Heading1"  (table)
disp(mat2doc.oxml.styles.styleId_from_name("Heading 1"));   % "Heading1"  (default arm, H15)
disp(mat2doc.oxml.styles.styleId_from_name("Caption"));     % "Caption"   (default arm)
disp(mat2doc.oxml.styles.styleId_from_name("heading 10"));  % "heading10" (no table key)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/styles.py::styleId_from_name`*

---

(id-h17-delete)=
## H17 — Python `delete()` vs the MATLAB `handle` destructor (resolved by rename)

`CT_Style.delete` (`styles.py:168-170`) and `CT_LsdException.delete`
(`styles.py:78-80`) both perform `self.getparent().remove(self)` — the **first**
Python `delete()` element-removal methods in the project (python-pptx's oxml had
none, so Mat2Ppt never faced this). In MATLAB, `delete` on a `handle` subclass
**is the destructor**: there is no way to have a non-destructor method named
`delete`, and MATLAB also calls it implicitly during garbage collection.

**The resolution — a naming convention (user-ratified 2026-08-03, no D-number).**
Rather than override the destructor, the `delete()` methods are renamed **by the
kind of thing each removes** so that **no method named `delete` exists**:

| Python | MATLAB |
|---|---|
| `CT_Style.delete()` | `delete_style()` |
| `CT_LsdException.delete()` | `delete_lsd_exception()` |

Each is the plain faithful port:

```matlab
p = obj.getparent();
if ~isequal(p, [])   % Python: self.getparent().remove(self)
    p.remove(obj);
end
```

Because no method is named `delete`, MATLAB's handle destructor is **never
overridden** and the GC-driven collision is **dissolved entirely** — no
`isvalid`/`try-catch` guarded-destructor machinery is needed. GC never calls these
methods, so the detached element **survives** the call exactly as in python-docx
(`self._element` survives detached). The `~isequal(p, [])` parent-guard keeps an
explicit call on an **unparented** element a no-op (Python `AttributeError` there,
unreachable in python-docx usage). The proxy layer
([`BaseStyle`](styles_api.md#id-h17-delete-proxy) / `LatentStyle_`) uses the same
convention — `delete_style()` / `delete_latent_style()` — delegating to these
element methods then setting `element_ = []`.

:::{warning}
**Binding Gate-4 rule.** Assert only the parent-side effect of the delete (child
count / serialized bytes). At the proxy layer, `element_` is `[]` (Python `None`)
after the call, so any subsequent proxy property access errors, matching python-docx
`AttributeError` — never inspect the proxy afterward.
:::

**Status: ✅ user-ratified 2026-08-03** (H17 dissolved; removed from the open
sign-off queue). Full record:
`validation\summary\decision_2026-07-30_h17_delete_destructor.md` and design.md
§9 H17; WP record `validation\mat2doc\audit_H17_delete_rename.md` (byte-neutral,
0 new D). Prior gate record: `validation\mat2doc\audit_P4-6_oxml_styles.md`
(Gate-2, block K).

(id-f1-boolnone)=
## The F-1 fix — `set_bool_prop(attr, None)` writes `"0"`, not remove

`CT_LatentStyles.set_bool_prop` (`styles.py:62-64`) ports
`setattr(self, attr_name, bool(value))`. The Gate-2 auditor caught (oracle-proven)
that the naive port `obj.(attr_name) = logical(value)` maps `[]` (None) to an
**empty** logical, which the `OnOff` `setAttrTyped` path treats as None and
**removes** the attribute — but Python `bool(None)` is `False`, and since the
`OptionalAttribute` default is None, `False ≠ default`, so python-docx **writes**
`w:def..="0"`. The fix routes the None case to the logical `false`:

```matlab
if isequal(value, [])       % Python: bool(None) -> False (writes "0", not remove)
    obj.(attr_name) = false;
else
    obj.(attr_name) = logical(value);
end
```

Gate-3 confirmed the whole truth table byte-identical to the oracle: `True`→`"1"`,
`False`→`"0"`, `1`→`"1"`, `0`→`"0"`, and **`[]` (None)→`"0"`** (written, not
removed; `bool_prop` reads back `False`). The M1 save path never hits this case
(the default template `w:latentStyles` has no `w:defQFormat`, so `styles.xml`
stays byte-clean), but the path is reachable once P4-7a ports the LatentStyles
`default_to_*` setters, where None is a legal user input.

---

## Styles-oxml layer COMPLETE — the styles API (P4-7a/b) is next → M2

This WP completes the **styles-oxml element layer**: the `<w:styles>` root, the
`<w:style>` definition, the `<w:latentStyles>`/`<w:lsdException>` latent-style
tier, the shared `<w:val>` decimal-number leaf, and the `name → styleId` mangler.
The 12 styles-block tags are registered byte-neutrally and the 349 KB `styles.xml`
now parses through these classes with M1 intact (17/17 L1) and **zero new
D-numbers**. What remains before **M2** is the **styles API chain**: **P4-7a**
ports the `styles`/`style` proxies (`Styles`, `_StyleFactory`, `BaseStyle` and its
subtypes) that read this element surface; **P4-7b** ports the latent-styles API and
un-stubs `Paragraph.style` / `Run.style`, and — with `Document.add_paragraph` /
`add_heading` — produces a real, byte-identical `document.xml` that opens clean in
Word (the M2 milestone gate + the Word-COM oracle).
