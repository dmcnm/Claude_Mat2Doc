---
title: "mat2doc.styles — the styles API/proxy tier (StyleFactory · BaseStyle hierarchy · Styles collection · BabelFish)"
---

# `mat2doc.styles` — the styles API/proxy tier (`StyleFactory` + `BaseStyle`/`CharacterStyle`/`ParagraphStyle`/`TableStyle_`/`NumberingStyle_` + `Styles` + `BabelFish`)

Ported from python-docx v1.2.0 `src/docx/styles/style.py` (the module function
`StyleFactory` plus five style classes — `BaseStyle`, `CharacterStyle`,
`ParagraphStyle`, `_TableStyle`, `_NumberingStyle`), `src/docx/styles/styles.py`
(the `Styles` collection), and `src/docx/styles/__init__.py` (`BabelFish`) — all
**flattened** into the single package `+mat2doc/+styles/`.

:::{note}
This is the **second work package of the styles chain** — P4-6 (oxml/styles) →
**P4-7a (the `styles`/`style` API)** → P4-7b (latent styles) → **M2**. It ports
the user-facing style **proxies** that read the `<w:styles>` element surface P4-6
built, and **un-stubs** the styles delegation across the document object graph
(`StylesPart`/`DocumentPart`/`Document`/`Run`/`Paragraph`). It adds **no
`register_element_cls` row and no serialization code**, so equivalence is
**behavioral** (probe value parity) plus serialized-bytes parity on the two
output-visible whole-part paths (`delete_`, `add_style`) and the end-to-end
style-by-name document byte pin — **byte-neutral** (M1 stays 17/17, **zero new
D-numbers**). It is not itself an M2 or COM milestone; the M2 COM oracle fires at
P4-7b, when `add_heading` forces the whole chain.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The module → package flattening (FLAG-3)

Python spreads the styles API across three modules: `BabelFish` lives in the
`docx.styles` package `__init__`, `StyleFactory`/`BaseStyle`/… in
`docx.styles.style`, and `Styles` in `docx.styles.styles`. The Mat2Doc port
**flattens all three into the single `+mat2doc\+styles` package** (no
module-mirroring sub-packages, per the FLAG-3 docx policy), so every symbol is
`mat2doc.styles.<Name>`.

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `styles/style.py` | `+mat2doc\+styles\` | `StyleFactory`, `BaseStyle`, `CharacterStyle`, `ParagraphStyle`, `_TableStyle`→`TableStyle_`, `_NumberingStyle`→`NumberingStyle_` |
| `styles/styles.py` | `+mat2doc\+styles\Styles.m` | `Styles` |
| `styles/__init__.py` | `+mat2doc\+styles\BabelFish.m` | `BabelFish` |

Two private classes rotate their leading underscore to a trailing one (the
toolbox-wide FLAG-3 convention `_Cell`→`Cell_`): `_TableStyle`→`TableStyle_`,
`_NumberingStyle`→`NumberingStyle_`. The non-public back-compat aliases
`_CharacterStyle = CharacterStyle` / `_ParagraphStyle = ParagraphStyle`
(`style.py:193`/`:236`) are behaviorally inert and have no MATLAB equivalent (a
class-alias file would collide) — recorded, not ported.

(id-stylefactory)=
## `StyleFactory`

**Syntax**

```matlab
s  = mat2doc.oxml.OxmlElement("w:style");
s.type = mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER;
cs = mat2doc.styles.StyleFactory(s);   % a mat2doc.styles.CharacterStyle
```

**Description**

The module function that returns a style proxy of the right `BaseStyle`
subclass for a `<w:style>` element, dispatched on its `type` attribute (a
`WD_STYLE_TYPE` member). `StyleFactory` is a **module-level function** in
`style.py` (not a method), so it maps to a package function
`+styles\StyleFactory.m` (design.md §1: one package function per file).

The H10 dispatch mirrors the Python dict literal indexed by `style_elm.type`
(`style.py:15-24`), ported as an explicit `isequal`/elseif chain over the four
`WD_STYLE_TYPE` members:

| `WD_STYLE_TYPE` (`@w:type`) | value | → MATLAB class |
|---|---|---|
| `PARAGRAPH` (`"paragraph"`) | 1 | `mat2doc.styles.ParagraphStyle` |
| `CHARACTER` (`"character"`) | 2 | `mat2doc.styles.CharacterStyle` |
| `TABLE` (`"table"`) | 3 | `mat2doc.styles.TableStyle_` |
| `LIST` (`"numbering"`) | 4 | `mat2doc.styles.NumberingStyle_` |

:::{warning}
**H3/H10 edge — `KeyError` on `None`.** `style_elm.type` reads the
`w:style/@w:type` `OptionalAttribute`, which is `[]` (None) when the attribute is
absent. Python then evaluates `{...}[None]` → `KeyError: None` (None is not a
dict key). The port faithfully reproduces this: an unmapped/`[]` type raises
`mat2doc:KeyError` with the Python repr of the key (`"None"` for `[]`). Real
`styles.xml` `<w:style>` elements always carry `@w:type`, so this path is
**unreachable** in normal document usage. Note `BaseStyle.type` (the property,
None→PARAGRAPH) is a **separate** read and is NOT used by `StyleFactory` —
faithful to Python, which dispatches on the raw oxml attribute.
:::

**Example**

```matlab
for t = [mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH, ...
         mat2doc.enum.style.WD_STYLE_TYPE.LIST]
    se = mat2doc.oxml.OxmlElement("w:style"); se.type = t;
    disp(class(mat2doc.styles.StyleFactory(se)));
end
% mat2doc.styles.ParagraphStyle
% mat2doc.styles.NumberingStyle_
```

*Ported from python-docx v1.2.0: `src/docx/styles/style.py::StyleFactory`*

---

(id-basestyle)=
## `BaseStyle`

**Syntax**

```matlab
se = mat2doc.oxml.OxmlElement("w:style");
se.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
s = mat2doc.styles.ParagraphStyle(se);
s.name = "Heading 1";          % writes internal "heading 1" via BabelFish
s.style_id = "Heading1";
s.priority = 10; s.hidden = true; s.quick_style = true;
```

**Description**

The base class for the paragraph / character / table / numbering style objects —
its properties and methods are inherited by all of them. `BaseStyle` is a
`mat2doc.shared.ElementProxy` subclass, so it has reference (handle) semantics and
the H5 element-identity `eq`/`ne` unchanged. Every read/write property reaches
through the wrapped `<w:style>` element (a P4-6 `CT_Style`) to its `.._val`
helpers; **`BaseStyle` adds no oxml logic, no registry row and no serialization
code.**

**The property surface** (`style.py:38-161`), all H3 tri-state:

- **`builtin`** (read-only) — Python `not self._element.customStyle`, a
  truthiness negation over `customStyle ∈ {None, True, False}` (H4):
  `~(~isequal(cs,[]) && cs)`.
- **`name`** — `BabelFish.internal2ui` of the `<w:name>` value, or `[]` when
  absent. The **faithful quirk**: the *setter* writes the value **verbatim** (no
  `ui2internal`) while the *getter* translates internal→UI, so
  `style.name = "Heading 1"` stores `"Heading 1"` (not the internal
  `"heading 1"`).
- **`priority`** — `w:uiPriority` value, or `[]` (None) when absent.
- **`style_id`** — the raw `@w:styleId`, or `[]`. Note the getter reads
  `_style_elm` and the setter writes `_element` (both the same `CT_Style`) —
  ported verbatim (`style.py:134`/`:138`).
- **`type`** (read-only) — `WD_STYLE_TYPE` member; **None defaults to PARAGRAPH**
  (`style.py:144-147`).
- **`hidden` / `locked` / `quick_style` / `unhide_when_used`** — the CT_OnOff-backed
  booleans; their **getters return the logical `false` (not `[]`) when absent**,
  matching python-docx (whose `semiHidden_val`/`locked_val`/`qFormat_val`/
  `unhideWhenUsed_val` return bool `False` when the child is missing). The setters
  accept `[]`/`false`/`true` and the P4-6 CT_Style setters remove/add accordingly.

(id-hetero-root)=
**The heterogeneous root (design.md §2 "Collections → Heterogeneous").**
`BaseStyle` additionally derives `matlab.mixin.Heterogeneous` so that a mixed
vector of its subclasses (`ParagraphStyle`/`CharacterStyle`/`TableStyle_`/
`NumberingStyle_`) — as produced by [`Styles.to_array`](#id-styles) and
`StyleFactory` — can be held in a single `1×N` array typed `BaseStyle`. This is
the "downstream collection WP that needs a heterogeneous proxy vector"
anticipated in the `ElementProxy` note. **Sealing:** only methods invoked
ARRAY-WISE on a heterogeneous array need `Sealed`; `Styles` uses `eq` only
scalar-to-scalar (`style == self.default(...)`), so `ElementProxy`'s unsealed
`eq`/`ne` are left unsealed (verified: scalar `eq` over a hetero array element
works; array-wise `arr == x` throws `MATLAB:class:UnsealedMethod` — unused here
and by python-docx, so acceptable and documented in the class header).

*Ported from python-docx v1.2.0: `src/docx/styles/style.py::BaseStyle`*

---

## `CharacterStyle`

**Syntax**

```matlab
se = mat2doc.oxml.OxmlElement("w:style");
se.type = mat2doc.enum.style.WD_STYLE_TYPE.CHARACTER;
cs = mat2doc.styles.CharacterStyle(se);
cs.font.bold = true;               % <w:rPr><w:b/></w:rPr> on the style
```

**Description**

A character style — applied to a `Run`, providing character-level formatting via
the `Font` object in its `font` property. `< mat2doc.styles.BaseStyle`. Adds two
members to the inherited surface:

- **`base_style`** (read/write) — the get returns a Style (via `StyleFactory`
  over the sibling `<w:style>` referenced by `<w:basedOn>`), or `[]` when there is
  no `basedOn` **or the referenced style is not found** (a dangling ref → `[]`).
  The set writes the target's `style_id` into `w:basedOn/@w:val`, or removes
  `basedOn` when the assigned style is `[]` (None).
- **`font`** (read-only) — returns a **fresh** `Font` wrapping the `<w:style>` on
  each access (Python `Font(self._element)`; not cached).

*Ported from python-docx v1.2.0: `src/docx/styles/style.py::CharacterStyle`*

---

## `ParagraphStyle`

**Syntax**

```matlab
se = mat2doc.oxml.OxmlElement("w:style");
se.type = mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH;
ps = mat2doc.styles.ParagraphStyle(se);
ps.paragraph_format.left_indent = mat2doc.shared.Pt(36);
```

**Description**

A paragraph style — provides both character formatting (inherited from
`CharacterStyle`) and paragraph formatting via the `ParagraphFormat` object in
its `paragraph_format` property. `< mat2doc.styles.CharacterStyle`. Adds:

- **`next_paragraph_style`** (read/write) — the style applied automatically to a
  new paragraph inserted after a paragraph of this style.
  - **GET** (`style.py:214-219`): returns **self** when no next style is defined
    (`<w:next>` absent) **or** when the referenced style is **not a PARAGRAPH
    type**; otherwise `StyleFactory` over the referenced sibling. The H3/H4 chain
    `next_style_elm.type != PARAGRAPH` is True when the type is None (a
    dangling/typeless ref) → also returns self.
  - **SET** (`style.py:222-226`): assigning `[]` (None) **or self** (same
    `style_id`) **removes** the `<w:next>`; any other style writes its `style_id`
    into `w:next/@w:val`. The `is None` short-circuits before `style.style_id` is
    read (`isequal(style,[]) || isequal(style.style_id, obj.style_id)`).
- **`paragraph_format`** (read-only) — a **fresh** `ParagraphFormat` wrapping the
  `<w:style>` each access (not cached).

*Ported from python-docx v1.2.0: `src/docx/styles/style.py::ParagraphStyle`*

---

## `TableStyle_` and `NumberingStyle_`

**Syntax**

```matlab
se = mat2doc.oxml.OxmlElement("w:style");
se.type = mat2doc.enum.style.WD_STYLE_TYPE.TABLE;
ts = mat2doc.styles.TableStyle_(se);   % Python _TableStyle
```

**Description**

Two thin faithful subclasses that add nothing to their parents' surface:

- **`TableStyle_`** (Python `_TableStyle`, `< ParagraphStyle`) — a table style
  provides character and paragraph formatting for its contents as well as special
  table formatting properties. python-docx v1.2.0 adds no members beyond
  `ParagraphStyle` except a `__repr__` override (MATLAB object display is a
  separate `disp`/`display` mechanism → not ported, no output-visible effect), so
  the class inherits the whole `ParagraphStyle`/`CharacterStyle`/`BaseStyle`
  surface unchanged.
- **`NumberingStyle_`** (Python `_NumberingStyle`, `< BaseStyle`) — not yet
  implemented in python-docx v1.2.0 (empty class body); it exposes only the
  inherited `BaseStyle` surface.

Both carry the FLAG-3 trailing-underscore rename (see
[the flattening note](#id-stylefactory) above and the mapping page's idiom row).

*Ported from python-docx v1.2.0: `src/docx/styles/style.py::_TableStyle` /
`::_NumberingStyle`*

---

(id-styles)=
## `Styles`

**Syntax**

```matlab
d  = mat2doc.Document();
st = d.styles;
st.len_()                          % number of styles (164 in the default template)
st.contains_("Heading 1")          % true if defined
h1 = st.getitem_("Heading 1");     % Python st["Heading 1"] -> ParagraphStyle
for s = st.to_array(); disp(s.style_id); end
```

**Description**

Provides access to the styles defined in a document, reached via the
`Document.styles` property. Wraps the `<w:styles>` root (a P4-6 `CT_Styles`).
`< mat2doc.shared.ElementProxy` (reference semantics + H5 element-identity
`eq`/`ne` inherited). In python-docx it is a collections-style sequence
supporting `len()`, iteration, `in`, and dictionary-style access by UI style
name.

(id-collection-surface)=
:::{note}
**VERIFY-COLLECTION (accepted at Gate 2).** The shared 1-based `()` collection
base (design.md §2 "Collections → `RedefinesParen` base") is a **future work
package** and does not exist yet. Following the established precedent (`TabStops`,
Mat2Ppt `_GradientStops`), the Python sequence dunder surface is ported here as
**explicit methods** keeping line-for-line fidelity:

| Python dunder | MATLAB method | usage |
|---|---|---|
| `name in styles` | `styles.contains_(name)` | membership by UI name |
| `styles[key]` | `styles.getitem_(key)` | dictionary-style access |
| `for s in styles` | `for s = styles.to_array()` | heterogeneous iteration |
| `len(styles)` | `styles.len_()` | style count |

When the collection base lands, `Styles` should derive from it and expose native
`()` access; do NOT retrofit now.
:::

**The lookup + mutation surface:**

- **`contains_(name)`** — `BabelFish.ui2internal(name)` then any `<w:style>` whose
  `name_val` equals the internal name (H4/H3 value equality; None `≠` a string).
- **`getitem_(key)`** — by UI name first
  (`get_by_name(BabelFish.ui2internal(key))`); on a miss falls back to the
  **deprecated by-id** path (`get_by_id(key)`, the key passed **raw**); else
  raises `mat2doc:KeyError` (message `no style with name '<key>'`, verbatim). The
  by-id path emits a deprecation warning — see [the warning note](#id-warning).
- **`to_array()`** — a Style per `<w:style>` child, in document order, materialized
  to a **heterogeneous `1×N` `BaseStyle` array** (no styles → a `1×0` array). Mints
  **fresh** views via `StyleFactory` each call (python-docx does not cache Style
  objects, H5).
- **`len_()`** — `numel(self._element.style_lst)`.
- **`add_style(name, style_type, builtin)`** — `builtin` defaults `false`; raises
  `mat2doc:ValueError` (`document already contains style '<name>'`) when the name
  already exists, else appends a new `<w:style>` and returns its `StyleFactory`
  proxy. (`style_name in self` re-applies `ui2internal` inside `contains_` —
  idempotent, ported verbatim.)
- **`default(style_type)`** — the default style for `style_type`, or `[]` when
  none.
- **`get_by_id(style_id, style_type)`** — the style of `style_type` matching
  `style_id`; returns the **default** for `style_type` when `style_id` is `[]`
  (None) **or** not found **or** of the wrong type. The private `_get_by_id`
  applies the H4 `if style_id` falsy test (None **and** empty string are falsy →
  `~isequal(style_id,[]) && strlength(style_id) > 0`).
- **`get_style_id(style_or_name, style_type)`** — the id of the style matching
  `style_or_name`, or `[]`. Dispatches on `isa(style_or_name,
  "mat2doc.styles.BaseStyle")` (the `isinstance(..., BaseStyle)` analogue, H10):
  a Style object → `_get_style_id_from_style` (wrong type → `mat2doc:ValueError`
  with the `NAME (value)` enum-str, default style → `[]`); None → `[]`; else a
  name → `_get_style_id_from_name`.
- **`element`** — the wrapped `<w:styles>` element (inherited `ElementProxy`
  accessor).

(id-warning)=
:::{note}
**The project's first ported `warnings.warn` — house convention
`mat2doc:UserWarning`.** `getitem_`'s deprecated by-id lookup emits
`warning("mat2doc:UserWarning", "%s", "style lookup by style_id is deprecated.
Use style name as key instead.")`, mirroring python-docx's
`warn(msg, UserWarning, stacklevel=2)`. The mechanism differs (MATLAB `warning`
vs Python `warnings.warn`) but the message text is byte-preserved and the
serialized output is unaffected. This is the **first `warnings.warn` port in
either toolbox**; the identifier form `mat2doc:<PyWarningCategory>` parallels the
established `mat2doc:<PyExceptionName>` error convention and is recommended
(Gate-2/Gate-3) as the house convention for `warnings.warn` ports (queued for
user ratification).
:::

**Example**

```matlab
d  = mat2doc.Document();
st = d.styles;
h1 = st.getitem_("Heading 1");
disp(class(h1));                   % mat2doc.styles.ParagraphStyle
disp(h1.style_id);                 % "Heading1"
nrm = st.default(mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH);
disp(nrm.name);                    % "Normal"
disp(st.get_style_id("Heading 1", ...
        mat2doc.enum.style.WD_STYLE_TYPE.PARAGRAPH));   % "Heading1"
```

*Ported from python-docx v1.2.0: `src/docx/styles/styles.py::Styles`*

---

## `BabelFish`

**Syntax**

```matlab
mat2doc.styles.BabelFish.ui2internal("Heading 1")   % "heading 1"
mat2doc.styles.BabelFish.internal2ui("heading 1")   % "Heading 1"
mat2doc.styles.BabelFish.ui2internal("Normal")      % "Normal" (passthrough)
```

**Description**

Translates special-case style names between UI form (e.g. `"Heading 1"`) and
internal/`styles.xml` form (e.g. `"heading 1"`). An all-static class. Names not in
the fixed alias table pass through unchanged, in both directions.

The 12 alias pairs (`styles/__init__.py:12-25`, verbatim) are held as a Constant
`Nx2` string array `[ui, internal]`. Python builds two dicts from the same tuple:
`internal_style_names = dict(style_aliases)` maps UI→internal (used by
`ui2internal`), `ui_style_names` is the inverse (used by `internal2ui`); both are
reproduced by scanning column 1 or column 2. The pairs are a bijection, so no
key-collision ambiguity arises.

| UI name | internal name |
|---|---|
| `Caption` | `caption` |
| `Footer` | `footer` |
| `Header` | `header` |
| `Heading 1` … `Heading 9` | `heading 1` … `heading 9` |

:::{warning}
**H15 — the lookup is Python `dict.get`, case-SENSITIVE.** The alias matches are
exact string `==` (no `lower()`/`upper()` anywhere), so `"Heading 1"` and
`"heading 1"` are distinct keys and only the listed pairs translate; an unknown
name is returned verbatim. `ui2internal("heading 1")` → `"heading 1"` (miss →
passthrough) and `internal2ui("Heading 1")` → `"Heading 1"` (miss → passthrough).
:::

*Ported from python-docx v1.2.0: `src/docx/styles/__init__.py::BabelFish`*

---

(id-h17-delete-proxy)=
## The H17 hazard at the proxy layer — `delete()` → `delete_()`

`BaseStyle.delete()` (`style.py:49-57`) performs
`self._element.delete(); self._element = None`. In MATLAB, `delete` on a `handle`
subclass **is the destructor** — MATLAB calls it implicitly during garbage
collection and there is no way to have a non-destructor method of that name.

**Why the P4-6 element-layer ruling does NOT transfer.** The design.md §9 H17
ruling (a *guarded destructor override*) was validated for the **element** layer
(`CT_Style.delete`, P4-6): a wrapper going out of scope does not fire the child's
destructor while the tree is alive, because the parent's children array holds a
**strong** reference. That argument is about the element wrapper. A `BaseStyle` is
a **separate proxy the tree does NOT reference** — it is minted transiently by
**every** `StyleFactory` call (`get_by_id` / `getitem_` / `to_array` / `default` /
`base_style` / `next_paragraph_style` / `get_style_id`) and dropped by its caller.
So overriding `delete` here with a faithful body (`obj.element_.delete()`) would
detach a **still-parented** `<w:style>` on **ordinary iteration/GC** (empirically
proven in R2024b: parent `1 kid → 0 kids`), and MATLAB gives no signal to tell an
explicit `style.delete()` from a GC destructor.

**The resolution — a FLAG-3 method rename (no D-number).** The same collision
convention as the private-class renames (`_TableStyle`→`TableStyle_`):

1. **`delete` is NOT overridden** on any style proxy — MATLAB's default `handle`
   destructor is left in place (GC-safe: no tree effect, the wrapped element stays
   parented via its parent's strong ref).
2. **`BaseStyle.delete_()`** (trailing underscore) carries the faithful
   python-docx `BaseStyle.delete()` semantics: `obj.element_.delete()` (detach via
   the P4-6 element-layer guarded `CT_Style.delete`, safe there) then
   `obj.element_ = []`. Inherited by all subclasses.

`delete_()` is only ever called **explicitly** (never by GC) and detaches a
parented element through the element-layer guard, so there is no corruption risk
and the result is **byte-identical to python-docx `style.delete()`** (Gate-3
re-derived the whole styles-part SHA `cc0bb35d…8614`, 348 872 B, `Heading1`
removed). This is a method-**naming** resolution (FLAG-3 class), **not an output
deviation → no D-number.**

:::{warning}
**Binding Gate-4/P4-7a rule.** After `delete_()` the MATLAB element handle is
**invalid** (destroyed), whereas Python's element survives detached — **never
inspect the handle after `delete_`**; assert only the parent-side effect (child
count / serialized bytes). The GC-invariance safety property was independently
re-proven at Gate-3: minting 25× transient `StyleFactory` proxies and clearing
them leaves the styles part **bit-for-bit unchanged** (`02d71a68…e384`).
:::

**Status: SIGNED-PROVISIONAL** (adopted under the overnight-decision protocol;
queued for user ratification). Full record:
`validation\summary\decision_2026-07-30_h17_delete_destructor.md` (the design.md
§9 H17 addendum); Gate records `validation\mat2doc\audit_P4-7a_styles_api.md` and
`validate_P4-7a_styles_api.md`.

---

(id-unstub)=
## The un-stub — the styles delegation goes live end-to-end

This WP wires the styles feature surface that P1-8/P2-2/P4-4b/P4-5b shipped as
`mat2doc:notYetPorted` stubs naming P4-7 as the owner. After P4-7a these paths
**resolve** over the already-live P2-2 `StylesPart` (no new part created):

| Path | Now resolves to |
|---|---|
| `Document.styles` | `mat2doc.styles.Styles` |
| `DocumentPart.styles` | `mat2doc.styles.Styles` |
| `DocumentPart.get_style(style_id, style_type)` | a `BaseStyle` subclass |
| `DocumentPart.get_style_id(style_or_name, style_type)` | `string` \| `[]` |
| `StylesPart.styles` | `mat2doc.styles.Styles` |
| `Paragraph.style` get/set | `ParagraphStyle`; writes `w:pPr/w:pStyle/@w:val` |
| `Run.style` get/set | `CharacterStyle`; writes `w:rPr/w:rStyle/@w:val` |

`Paragraph.style` and `Run.style` were already **live delegation** (P4-5b/P4-4b)
to `DocumentPart.get_style`/`get_style_id`; those targets went live here, so the
chains now **resolve end-to-end**. The Gate-3 style-by-name G-scenario proves the
whole chain (`Paragraph.style = "Heading 1"` → `DocumentPart.get_style_id` →
`Styles.get_style_id` → name lookup → `getitem_` → `StyleFactory` →
`ParagraphStyle` → `style_id`) writes a `<w:pStyle w:val="Heading1"/>` into
`word/document.xml` **byte-identical** to python-docx — the closest guard yet to
M2's `add_heading`.

:::{note}
**Still stubbed for P4-7b.** `Styles.latent_styles` remains a clean
`mat2doc:notYetPorted` stub naming P4-7b (the `LatentStyles` proxy); the
underlying `CT_Styles.get_or_add_latentStyles` descriptor is already live (P4-6).
:::

---

## Styles API un-stubbed — latent styles (P4-7b) is next → M2

This WP completes the **styles API/proxy tier**: the `StyleFactory` dispatch, the
`BaseStyle`→`CharacterStyle`→`ParagraphStyle`→`TableStyle_`/`NumberingStyle_`
hierarchy with its full property surface and `base_style`/`next_paragraph_style`
chains, the `Styles` collection, and `BabelFish`. The styles delegation is
un-stubbed across the document object graph, the H17 `delete()` collision is
resolved by the byte-faithful `delete_()` rename, and everything stays
**byte-neutral** (M1 17/17, **zero new D-numbers**). What remains before **M2** is
**P4-7b** — the latent-styles API (`LatentStyles`/`_LatentStyle`, un-stubbing
`Styles.latent_styles`) — after which `Document.add_paragraph` / `add_heading`
produce a real, byte-identical `document.xml` that opens clean in Word (the M2
milestone gate + the Word-COM oracle).
