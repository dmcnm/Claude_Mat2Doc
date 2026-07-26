---
title: "mat2doc.shared / exc / types — the proxy base tier"
---

# `mat2doc.shared` / `+exc` / `+types` — the proxy base tier

Ported from python-docx v1.2.0 `src/docx/shared.py` (the proxy classes, not the
`Length` family — that was P1-1), `src/docx/exceptions.py`, and
`src/docx/types.py`. This is the **base tier that every later API proxy
extends** (P3–P6: paragraph, run, table, section, style, and the rest of the
WordprocessingML object model all derive from `ElementProxy` / `Parented` /
`StoryChild`). It emits **no serialized output of its own** — it supplies
hierarchy, element-identity equality, `part` access, the exception raisers, and
the structural protocol contracts on top of the P1 oxml/opc machinery.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## Where this differs from python-pptx (the design guide)

python-docx v1.2.0 is the **source of truth**; Mat2Ppt supplied only the MATLAB
*idiom* (handle class, `eq`/`ne` with an `isa` guard, property-as-method,
underscore rotation). The docx **class shape differs** and was ported from the
docx source (`audit_P2-1_proxy_tier.md` §2):

| Aspect | python-docx (SoT) | python-pptx (idiom source) | Resolution in Mat2Doc |
|---|---|---|---|
| `ElementProxy` ctor | `__init__(element, parent=None)` — carries an OPTIONAL parent | `__init__(element)` — no parent | `ElementProxy(element, parent)`, `parent` defaults to `[]` (None) |
| `part` accessor home | ON `ElementProxy`, WITH a None-guard | on the sibling `ParentedElementProxy`, no guard | `ElementProxy.part()` raises `mat2doc:ValueError` when `parent_` is None, else `parent_.part()` |
| `Parented` / `StoryChild` | STANDALONE classes, NOT `ElementProxy` subclasses; hold only a parent, no element, no `eq`/`ne` | pptx's `ParentedElementProxy` DERIVES `ElementProxy` and has an element | plain `handle` classes — only `parent_` + `part()`; no element, no identity overload |
| `write_only_property` | present (shared.py 266-274) | ABSENT in pptx | new idiom guide (set-only Dependent-property pattern) |
| `TextAccumulator` | present (shared.py 356-382) | ABSENT in pptx | new faithful port |

## The Document → ElementProxy retrofit (VERIFY-M1-DOC-BASE)

At M1 (P1-8) `mat2doc.document.Document` derived `handle` only, storing
`element_`/`part_` directly, because the proxy tier did not yet exist. P2-1
supplies the tier and **retrofits the real base**:
`classdef Document < mat2doc.shared.ElementProxy`. The base now holds
`element_`, and `Document` inherits `element()` / `eq` / `ne`; `part()` remains
overridden to return the document's own `_part` (so the base None-guard never
fires for a `Document`).

**Byte-neutral — proven.** `mat2doc.Document().save` unzips **17/17
byte-identical** against the frozen M1 reference (`references\s0001`), three-way
`MATLAB ≡ python Document().save ≡ s0001`, unchanged by the retrofit — it adds
hierarchy/identity/accessor behaviour, **not one serialized byte**.

**Python-faithful `==` (the observable, non-byte change).** `Document ==` now
follows the ElementProxy H5 contract (element identity) rather than MATLAB
*instance* identity. Consequently two `DocumentPart.document` accesses — a plain
`@property` that builds a **fresh** proxy each call — compare **EQUAL**, because
both wrap the same `w:document` element. This is the python-docx result
(`part.document == part.document` is `True`). The pre-retrofit P1-8 test
asserted `verifyFalse(d1 == d2)` (MATLAB instance identity, contrary to Python);
that one assertion was corrected to `verifyTrue` and re-run green — the
**ratified VERIFY-1** edit. Non-caching (distinct instances) is preserved but,
as in Python (where it is `is`-only), is no longer observable through `==` /
`isequal`; it is pinned instead by a `delete`/`isvalid` probe.

## Deviation posture — 0 new D-numbers

Pure proxy plumbing; no output-visible divergence. The 17/17 M1 sweep
re-confirms the standing adopt-only set (D-001, D-serializer-nsdecl, D-zip-time,
D-coreprops-time) with **no new D-number and no ledger row**. The one design
divergence — flat exception identifiers (VERIFY-2, below) — is inert at runtime
(the clone has **zero** internal `except` sites for any of the three exception
classes) and needs no D.

---

## `ElementProxy`

**Syntax**

```matlab
p   = mat2doc.shared.ElementProxy(element)
p   = mat2doc.shared.ElementProxy(element, parent)
e   = p.element()      % the wrapped oxml element
pt  = p.part()         % the containing package part (raises if no parent)
tf  = (a == b)         % H5 element-identity equality
tf  = (a ~= b)
```

**Description**

Base class for lxml element proxy classes — the most common class type in
python-docx after the oxml `CT_*` classes. A **handle** class (design.md §2)
that wraps a shared element tree and holds no mutable local state: two proxy
instances wrapping the *same* element are views of one object. `Document`
derives from it (`document.py` 28).

**Element-identity equality (H5 — the key hazard).** `eq`/`ne` compare by the
**handle identity of the wrapped element**, guarded by
`isa(_, "mat2doc.shared.ElementProxy")` on **both** operands (Python's
`isinstance` guard, including the reflected-operand order). This is Python's:

```python
def __eq__(self, other):
    if not isinstance(other, ElementProxy):
        return False
    return self._element is other._element
```

The `element_ == element_` comparison bottoms out in `XmlElement`'s **Sealed**
handle-identity `eq` (≡ lxml `is`). The crucial consequence is the
**content-comparison trap**: two proxies over **distinct but byte-identical**
elements compare **NOT equal** — equality is by element identity, never by
content.

| comparison | result |
|---|---|
| two proxies, SAME element handle | `==` true / `~=` false |
| two proxies, DISTINCT byte-identical elements | `==` **false** / `~=` **true** |
| proxy vs non-proxy (either operand order) | `==` false / `~=` true |
| proxy vs `[]` (None) | `==` false / `~=` true |

**The `.part` None-guard (H3).** `parent` defaults to `[]` (the None sentinel).
`part()` reproduces `if self._parent is None: raise ValueError(...)` with an
inline typed-empty test (`isa(parent_,"double") && isempty(parent_)` — `parent_`
is only ever `[]` or a handle, so `''` can never spoof None); when absent it
raises `mat2doc:ValueError` with the byte-verbatim message
`part is not accessible from this element`, otherwise it delegates to
`parent_.part()`. Mat2Doc uses **no** shared `isNone` helper
(`decision_2026-07-26_mat2doc_none_idiom.md`).

`element` and `part` are ported property-as-method (design.md §2). `eq`/`ne` are
**not** Sealed and `ElementProxy` is **not** a Heterogeneous root — the proxy
layer never forms mixed-class arrays compared element-wise (VERIFY-4); a future
proxy-collection WP that needs a heterogeneous proxy vector must add the mixin
and seal these methods then.

**Example**

```matlab
e  = mat2doc.oxml.XmlElement("w:document");
a  = mat2doc.shared.ElementProxy(e);
b  = mat2doc.shared.ElementProxy(e);            % different proxy, same element
a == b                                          % true  (H5 element identity)
f  = mat2doc.oxml.XmlElement("w:document");
a == mat2doc.shared.ElementProxy(f)             % false (byte-identical but distinct)
mat2doc.shared.ElementProxy(e).part()           % errors: mat2doc:ValueError (no parent)
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::ElementProxy`*

---

## `Parented`

**Syntax**

```matlab
obj = mat2doc.shared.Parented(parent)
pt  = obj.part()       % delegated up: parent_.part()
```

**Description**

Common services for document elements that occur **below a part** but
occasionally need an ancestor to provide a service (adding/dropping a
relationship). Provides the `parent_` attribute and a `part` accessor delegating
to the parent.

Unlike pptx's `ParentedElementProxy`, docx's `Parented` is **NOT** an
`ElementProxy` subclass and holds **no element** — only a parent. It therefore
exposes only `part` and defines **no** `eq`/`ne`: two `Parented` instances
compare by MATLAB's default handle (instance) identity, which is exactly
Python's default object identity for a class that does not override `__eq__`.

**Example**

```matlab
obj = mat2doc.shared.Parented(some_ancestor);   % ancestor exposes part()
p   = obj.part();                               % parent_.part()
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::Parented`*

---

## `StoryChild`

**Syntax**

```matlab
obj = mat2doc.shared.StoryChild(parent)
pt  = obj.part()       % delegated up: parent_.part() (a StoryPart)
```

**Description**

Structurally identical to `Parented` but a **distinct class**: its parent is
typed `ProvidesStoryPart` and its `part` is a `StoryPart`. Kept separate to
preserve the docx hierarchy — block-level items (paragraphs, tables) derive
`StoryChild`, not `Parented` (shared.py 336-353). Like `Parented`, it holds no
element and defines no identity overload.

*Ported from python-docx v1.2.0: `src/docx/shared.py::StoryChild`*

---

## `TextAccumulator`

**Syntax**

```matlab
acc = mat2doc.shared.TextAccumulator()
acc = mat2doc.shared.TextAccumulator(separator)
acc.push(text)
s   = acc.pop()        % (1,:) string of length 0 or 1
```

**Description**

Accepts `str` fragments and joins them, in order, on `pop()`; handy when stream
text is broken up arbitrarily and you want to rejoin it within bounds. The
optional `separator` defaults to `""`. New for Mat2Doc (no pptx counterpart);
used later by the text-extraction tier.

**Generator port (H9).** Python's `pop()` is a **generator** that yields **zero
or one** str (used as `yield from accum.pop()`, so an empty accumulator produces
nothing rather than an empty string). Per design.md §2 (generators →
precomputed arrays), MATLAB `pop()` returns a `(1,:) string` of length 0 (empty
buffer) or length 1 (the joined fragments); a consumer ports
`yield from accum.pop()` as `for t = acc.pop()`, which iterates nothing on a
`1×0` result. The buffer initializes to `strings(1,0)` (Python `[]`), not `""`
(a real one-element list) — the H4 empty-list-falsy distinction.

**Example**

```matlab
acc = mat2doc.shared.TextAccumulator();
acc.push("foo"); acc.push("bar");
acc.pop()        % "foobar" (1x1); buffer now empty
acc.pop()        % 1x0 string (nothing accumulated)
```

*Ported from python-docx v1.2.0: `src/docx/shared.py::TextAccumulator`*

---

## `lazyproperty` (idiom guide — not callable)

**Syntax**

```matlab
help mat2doc.shared.lazyproperty        % documents the idiom; calling it errors
```

**Description**

Python's `@lazyproperty` is a read-only data descriptor: the getter runs only on
first access, the result is cached on the instance and returned unchanged
thereafter, and assignment raises `AttributeError` (shared.py 152-263). MATLAB
has no decorators, so this ports as an **idiom, not a callable helper** — a
non-callable file that raises `mat2doc:lazyproperty:notCallable` and exists so
`help mat2doc.shared.lazyproperty` documents the pattern:

```matlab
properties (Dependent)
    fget
end
properties (Access = private)
    fget_cache_
    fget_isComputed_ (1,1) logical = false
end
methods
    function value = get.fget(obj)
        if ~obj.fget_isComputed_
            obj.fget_cache_ = <compute>;   % Python fget body
            obj.fget_isComputed_ = true;
        end
        value = obj.fget_cache_;
    end
end
```

**Newly established for Mat2Doc.** P1 established no lazyproperty home; the
cache+computed-flag pattern already appears ad hoc across P1
(`+opc\CoreProperties.m`, `+oxml\+coreprops\CT_CoreProperties.m`). docx's
`lazyproperty` lives in `shared.py` (pptx's was in `util.py`), so its Mat2Doc
home is `mat2doc.shared.lazyproperty` (**not** a `+util` package). Read-only is
achieved by defining **no** `set.fget`. The mandatory rule (design.md §2):
back the cache with a **logical computed-flag**, **never** `isempty` on the cache
— `[]` is a legal cached value (the None sentinel, H3). Gate-3 exercised the
realized idiom on a host object: compute-count `1→1` (cached, not re-evaluated),
value stable, assignment raises.

*Mat2Doc idiom guide (no callable python-docx counterpart); mirrors design.md §2.*

---

## `write_only_property` (idiom guide — not callable)

**Syntax**

```matlab
help mat2doc.shared.write_only_property
```

**Description**

Python's `@write_only_property` is a set-only property (shared.py 266-274):
`set` defined, `get` absent. New idiom guide for Mat2Doc (no pptx counterpart),
documenting the MATLAB rendering — a Dependent property with a `set.x` method and
**no** `get.x`, so reads error while writes flow through.

*Mat2Doc idiom guide (no callable python-docx counterpart).*

---

## `mat2doc.exc` — exception raisers

**Syntax**

```matlab
mat2doc.exc.PythonDocxError(message)    % error id "mat2doc:PythonDocxError"
mat2doc.exc.InvalidSpanError(message)   % error id "mat2doc:InvalidSpanError"
mat2doc.exc.InvalidXmlError(message)    % error id "mat2doc:InvalidXmlError"
```

**Description**

python-docx's three `exceptions.py` classes — `PythonDocxError` (the base),
`InvalidSpanError`, `InvalidXmlError` — port to MATLAB **function raisers**
(the Mat2Ppt `+exc` precedent). Per design.md §2, each Python exception **class**
maps to a flat error **identifier** `mat2doc:<Name>`; the message is passed
through `"%s"` so any `%` it contains is literal (single-string rule). A
`raise X(msg)` site ports to `mat2doc.exc.X(msg)` or the equivalent inline
`error("mat2doc:X", ...)` — both emit the same greppable identifier. Gate-3
confirmed the identifier local-part (== the Python class name) and the
byte-verbatim message for each, including the `%`-literal case `span 100% invalid`.

**VERIFY-2 (ratified) — the exception hierarchy flattens to flat identifiers.**
In Python `InvalidSpanError` and `InvalidXmlError` **subclass** `PythonDocxError`
subclasses `Exception`. MATLAB error identifiers are flat — there is no
identifier-level "is-a", so `catch PythonDocxError` would **not** also catch the
two subclasses. This is **inert at runtime**: a clone-wide grep finds **zero**
`except` sites for any of the three classes anywhere in `src/` (they are
raise-only; only the upstream pytest suite asserts them via `pytest.raises`), so
no base-type catch semantics exist to lose. A public caller wanting "any
python-docx error" has the shared `mat2doc:` identifier prefix as the (broader)
analogue.

:::{note}
`PackageNotFoundError` is **not** here — it lives in `docx/opc/exceptions.py`
and was already ported under `+mat2doc\+opc\` (referenced by `PhysPkgReader`,
`api.py::Document`). The `InvalidXmlError` in `docx/oxml/exceptions.py` (a
distinct `XmlchemyError` subclass raised by `simpletypes.py`) collides onto the
same flat `mat2doc:InvalidXmlError` identifier that P1-3a already emits — a
name-collision recorded in the raiser's header, with no runtime consequence
(upstream never catches either).
:::

*Ported from python-docx v1.2.0: `src/docx/exceptions.py::{PythonDocxError, InvalidSpanError, InvalidXmlError}`*

---

## `mat2doc.types` — structural protocols (documentation-only)

**Syntax**

```matlab
help mat2doc.types.ProvidesXmlPart      % declares the `part` contract
help mat2doc.types.ProvidesStoryPart
```

**Description**

python-docx's `types.py` defines two `typing.Protocol`s — `ProvidesXmlPart`
(an object providing access to its `XmlPart`) and `ProvidesStoryPart` (one whose
part is a `StoryPart`). A Protocol is a **structural** type used only for static
type checking: an object satisfies it merely by exposing a `part` accessor,
**without inheriting from it**, and python-docx never does
`isinstance(x, ProvidesXmlPart)`.

**VERIFY-3 (ratified) — carried as documentation-only Abstract classes.** These
port to MATLAB `(Abstract)` handle classes declaring the `part` contract as an
Abstract method, so `help` and the class browser show the required surface — but
**nothing inherits from them and nothing does `isa(x, "mat2doc.types.*")`**. The
contract is satisfied structurally: any object with a `part()` method (e.g.
`mat2doc.opc.Part` and its subclasses) is a `ProvidesXmlPart`. This gives
discoverable, machine-visible contracts at zero fidelity cost, a step beyond the
Mat2Ppt precedent (which carried the pptx Protocols as prose in member help).
Do **not** subclass them for behaviour or gate logic on them.

*Ported from python-docx v1.2.0: `src/docx/types.py::{ProvidesXmlPart, ProvidesStoryPart}`*
