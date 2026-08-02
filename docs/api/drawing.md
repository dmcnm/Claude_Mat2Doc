---
title: "drawing — the DrawingML picture proxy + the run inner-content surface (Drawing · CT_R.inner_content_items · Run.iter_inner_content · the FINAL un-stub sweep)"
---

# `drawing` — the DrawingML picture proxy + the run inner-content surface (the P8-3 WP)

Ported from python-docx v1.2.0 `src/docx/drawing/__init__.py` (`Drawing`),
`src/docx/oxml/text/run.py` (`CT_R.inner_content_items`), `src/docx/text/run.py`
(`Run.iter_inner_content`) and the un-stubbed `Document.add_page_break` /
`inline_shapes` / `paragraphs` / `tables` (`src/docx/document.py`),
`DocumentPart.inline_shapes` (`src/docx/parts/document.py`) and
`OpcPackage._core_properties_part` (`src/docx/opc/package.py`). This is the
**third and FINAL port work package of Phase 8** — and the **last port WP of the
whole Mat2Doc project**.

:::{important}
**★ P8-3 is the final PORT work package.** After it, every python-docx v1.2.0
catalog symbol is ported, audited, validated, tested and documented, and the
toolbox raises **zero live `mat2doc:notYetPorted`** — the **C4 exit condition**
(a ground-truth grep of `error("mat2doc:notYetPorted"` in `+mat2doc\` returns
**0** live sites; the docstring/comment matches and `NumberingPart.new`'s
faithful `NotImplementedError` — which python-docx also raises — are not live
sites). `drawing/__init__.py::Drawing` is the **last of the 95 catalog modules**
to be ported (**C2b** — the one module with no prior owning WP). Only the
**P8-EXIT** rollup (system campaigns A/B/C + the whole-project Word-COM oracle +
the dogfooded Word user manual) remains before **Mat2Doc COMPLETE**.
:::

:::{note}
**Byte-neutral — zero new D-number.** P8-3 adds **no** registry row and **no**
serialization surface. `Drawing`, `CT_R.inner_content_items` and
`Run.iter_inner_content` are read-side accessors; the four `Document`
delegators and `DocumentPart.inline_shapes` forward to already-live parts; and
`Document.add_page_break` composes the pre-ported
`add_paragraph`→`add_run`→`add_break(WD_BREAK.PAGE)` serializers. Gate-3 proved
the port output-neutral: the `s0109` rich-doc scenario (two paragraphs incl.
non-ASCII, a hard page break, a 2×3 table, an inline picture) is **18/18 parts
byte-identical (L1)** to the python-docx oracle. **Zero L2, zero L3, zero new
D-number.**
:::

---

## `Drawing` — a `<w:drawing>` picture proxy (★ C2b — the last catalog module)

**Syntax**

```matlab
d  = items{1};        % a mat2doc.drawing.Drawing, reached via Run.iter_inner_content
tf = d.has_picture;   % bool -- true when the drawing holds an embedded picture
img = d.image;        % mat2doc.image.Image -- raises mat2doc:ValueError if not a picture
```

**Description**

`Drawing` (`drawing/__init__.py::Drawing`) wraps a `<w:drawing>` element. A
drawing can hold a picture, but it can also hold a chart, SmartArt or a drawing
canvas — the picture members qualify themselves with `has_picture`. It is a
**`Parented`** proxy (the parent-only tier): it holds **no** element for
equality and defines no `eq`/`ne`, so a `Drawing` is compared by MATLAB's
**default handle identity**, matching python-docx's default object identity
(H5). Because `Parented` stores no element handle, `Drawing` declares its own
private `drawing_` / `element_` handles (Python `self._drawing = self._element =
drawing` — one element under two names, the second set but never read upstream).

It is a **pure API/proxy tier** over the already-registered `CT_Drawing`
(`oxml/drawing.py`, P7-3): **no** oxml logic, **no** registry row, **no**
serialization code. Equivalence is therefore **behavioral**.

**`has_picture`** is `bool(self._drawing.xpath(expr))` for the inline **or**
floating picture path (`./wp:inline/…/pic:pic | ./wp:anchor/…/pic:pic`) → the
MATLAB `~isempty(...)` of the xpath result (H4 — an empty match list is falsy).
It is **false** for a linked picture, a chart, SmartArt or a canvas.
**`image`** returns an `Image` proxy for the picture's bytes: it reads the
`.//pic:blipFill/a:blip/@r:embed` relationship id (the `[0]` first xpath result
→ `(1)`, a DATA position, not a shifted index, H1), unwraps the `related_parts`
1×1-cell to the live image part (P1-5 currency, H5) and returns
`image_part.image`. When the drawing is **not** a picture it raises
`mat2doc:ValueError` with the **verbatim** message `drawing does not contain a
picture` (H4 — `if not picture_rIds` → `isempty(...)`). It is a Dependent
read-only property (matching the Python `@property`); MATLAB auto-display
catches the getter's `ValueError` and simply omits the `image` row for a
non-picture drawing (the `RenderedPageBreak` display-catch precedent — an
explicit `d.image` read raises exactly as python-docx).

**Example** (a real inline picture — `has_picture` true, `image` resolved)

```matlab
IMG = "C:\Users\dougl\Repos\python-docx\tests\test_files\python-powered.png";
doc = mat2doc.Document();
run = doc.add_paragraph().add_run();
run.add_picture(char(IMG));                 % inline <w:drawing>/<wp:inline>
items = run.iter_inner_content();
d = items{1};                               % a mat2doc.drawing.Drawing
fprintf('has_picture=%d  content_type=%s\n', d.has_picture, d.image.content_type);
% has_picture=1  content_type=image/png
fprintf('image class=%s\n', class(d.image));
% image class=mat2doc.image.Image
```

**Example** (a bare `<w:drawing>` — `image` raises `mat2doc:ValueError`)

```matlab
import mat2doc.oxml.parse_xml
import mat2doc.oxml.nsdecls
bare = parse_xml("<w:drawing " + nsdecls("w") + "/>");
D = mat2doc.drawing.Drawing(bare, []);       % [] parent (None)
fprintf('has_picture=%d\n', D.has_picture);  % has_picture=0
try
    img = D.image; %#ok<NASGU>
catch e
    fprintf('%s: %s\n', e.identifier, e.message);
    % mat2doc:ValueError: drawing does not contain a picture
end
```

*Ported from python-docx v1.2.0: `src/docx/drawing/__init__.py::Drawing`*

---

## `CT_R.inner_content_items` — a run's content in document order (oxml layer)

**Syntax**

```matlab
items = r.inner_content_items();   % 1xN cell of  str | CT_Drawing | CT_LastRenderedPageBreak
```

**Description**

`CT_R.inner_content_items` (`oxml/text/run.py::CT_R.inner_content_items`) returns
the `<w:r>` element's content as an ordered, **heterogeneous** list of
`str | CT_Drawing | CT_LastRenderedPageBreak`, with the plain-text run-content
children (`w:t`/`w:br`/`w:cr`/`w:noBreakHyphen`/`w:ptab`/`w:tab`) **coalesced**
into single `str` spans and the drawing / rendered-page-break elements
punctuating them **in document order**. UN-STUBBED at P8-3 — the reason it was
stubbed is discharged now that `CT_Drawing` (P7-3) and `CT_LastRenderedPageBreak`
(P4-3) are both registered.

The Python generator + `list(...)` is materialized into a **1×N cell array**
(the list is heterogeneous — `string` scalars interleaved with element handles —
so a cell, not a typed array, H9). `str(e)` → `e.str_()` (the element
text-equivalent, as in the `CT_R.text` getter). The tuple
`isinstance(e, (CT_Drawing, CT_LastRenderedPageBreak))` → `isa(...) || isa(...)`
(H10). A `TextAccumulator` gathers the plain-text spans; the trailing
`yield from accum.pop()` "tail" flushes the last coalesced span, so an empty run
yields the empty list `{}`.

**Example** (an interleaved run — text spans coalesce, break + drawing punctuate)

```matlab
import mat2doc.oxml.parse_xml
import mat2doc.oxml.nsdecls
frag = "<w:r " + nsdecls("w") + ">" + ...
    "<w:t>abc</w:t><w:cr/><w:t>def</w:t>" + ...     % coalesces to "abc\ndef"
    "<w:lastRenderedPageBreak/>" + ...
    "<w:t>ghi</w:t><w:tab/><w:t>jkl</w:t>" + ...     % coalesces to "ghi\tjkl"
    "<w:drawing/></w:r>";
r = parse_xml(frag);
items = r.inner_content_items();
fprintf('count=%d\n', numel(items));                 % count=4
fprintf('span1="%s"\n', replace(items{1}, newline, '\n'));   % span1="abc\ndef"
fprintf('mid  =%s\n', extractAfter(string(class(items{2})), "text."));  % mid  =CT_LastRenderedPageBreak
fprintf('span2="%s"\n', replace(items{3}, sprintf('\t'), '\t'));        % span2="ghi\tjkl"
fprintf('last =%s\n', extractAfter(string(class(items{4})), "drawing."));  % last =CT_Drawing
```

*Ported from python-docx v1.2.0: `src/docx/oxml/text/run.py::CT_R.inner_content_items`*

---

## `Run.iter_inner_content` — a run's content as API proxies (proxy layer)

**Syntax**

```matlab
items = run.iter_inner_content();   % 1xN cell of  str | RenderedPageBreak | Drawing
```

**Description**

`Run.iter_inner_content` (`text/run.py::Run.iter_inner_content`) is the proxy-tier
view over `CT_R.inner_content_items`: it yields the run's content **in document
order** as `str`, `RenderedPageBreak` and `Drawing` items (any other element type
is ignored). UN-STUBBED at P8-3 — `RenderedPageBreak` (P4-5b), `Drawing` (P8-3)
and `CT_R.inner_content_items` (P8-3) are all live.

The generator is materialized into a **1×N cell array** (heterogeneous — `str`
scalars and two proxy types, H9). `isinstance(item, str)` → `isstring(item)`;
the two element types dispatch through `isa` (H10). Each proxy is minted with
**this run** (`self`) as its parent, matching `RenderedPageBreak(item, self)` /
`Drawing(item, self)`. The trailing else-less chain drops any non-matching item
(unreachable — `inner_content_items` only emits those three types).

**Example** (the same interleaved run, at the proxy layer)

```matlab
import mat2doc.oxml.parse_xml
import mat2doc.oxml.nsdecls
frag = "<w:r " + nsdecls("w") + ">" + ...
    "<w:t>abc</w:t><w:cr/><w:t>def</w:t><w:lastRenderedPageBreak/>" + ...
    "<w:t>ghi</w:t><w:tab/><w:t>jkl</w:t><w:drawing/></w:r>";
r = parse_xml(frag);
ic = mat2doc.text.Run(r, []).iter_inner_content();
types = strings(1, numel(ic));
for k = 1:numel(ic)
    if isstring(ic{k})
        types(k) = "str";
    else
        parts = split(string(class(ic{k})), ".");
        types(k) = parts(end);
    end
end
fprintf('%s\n', join(types, " | "));
% str | RenderedPageBreak | str | Drawing
```

**Example** (an empty run → the empty cell `{}`, both layers)

```matlab
r0 = mat2doc.oxml.parse_xml("<w:r " + mat2doc.oxml.nsdecls("w") + "/>");
fprintf('oxml=%d  proxy=%d\n', numel(r0.inner_content_items()), ...
        numel(mat2doc.text.Run(r0, []).iter_inner_content()));
% oxml=0  proxy=0
```

*Ported from python-docx v1.2.0: `src/docx/text/run.py::Run.iter_inner_content`*

---

## `Document.add_page_break` — a hard page break in its own paragraph

**Syntax**

```matlab
p = document.add_page_break();   % a Paragraph containing only a page break
```

**Description**

`Document.add_page_break` (`document.py::Document.add_page_break`, un-stubbed at
P8-3) appends a new paragraph containing only a page break and returns it. It
composes the already-live serializers `self.add_paragraph()` (P4-7b) +
`paragraph.add_run().add_break(WD_BREAK.PAGE)` (P4-5b) — `WD_BREAK.PAGE` emits a
`<w:br w:type="page"/>` in a fresh run within the fresh trailing paragraph. This
is the only P8-3 member with a **write** surface, and Gate-3 proved that surface
byte-exact (the `s0109` `word/document.xml` is L1 against the python-docx oracle).

**Example**

```matlab
doc = mat2doc.Document();
p = doc.add_page_break();
fprintf('class=%s  runs=%d\n', class(p), numel(p.runs));
% class=mat2doc.text.Paragraph  runs=1
xmltxt = native2unicode(mat2doc.oxml.serialize_part_xml(doc.element()), 'UTF-8');
fprintf('emits page break: %d\n', contains(xmltxt, '<w:br w:type="page"/>'));
% emits page break: 1
```

*Ported from python-docx v1.2.0: `src/docx/document.py::Document.add_page_break`*

---

## `Document.paragraphs` / `Document.tables` — the top-level content collections

**Syntax**

```matlab
paras = document.paragraphs;   % 1xN Paragraph, document order (empty -> 1x0)
tbls  = document.tables;       % 1xN Table, document order (top-level only; empty -> 1x0)
```

**Description**

`Document.paragraphs` (`document.py:184-191`) and `Document.tables`
(`document.py:221-230`), both un-stubbed at P8-3, delegate to the body
`BlockItemContainer` (`self._body.paragraphs` / `self._body.tables`, both live).
Each returns a **1×N** array of **fresh** proxy views (H5), a **1×0** array when
empty. `paragraphs` excludes paragraphs inside revision marks (`w:ins` / `w:del`);
`tables` returns only **top-level** tables (a table nested in a cell does not
appear) and likewise excludes revision-marked tables.

**Example**

```matlab
doc = mat2doc.Document();
doc.add_paragraph("Alpha");
doc.add_paragraph("Beta");
doc.add_table(2, 3);
fprintf('paragraphs=%d  tables=%d\n', numel(doc.paragraphs), numel(doc.tables));
% paragraphs=2  tables=1
tbls = doc.tables;      % store first -- `doc.tables(1)` would parse as a method call
t = tbls(1);
fprintf('table is %dx%d\n', t.rows.len_(), t.columns.len_());
% table is 2x3
```

*Ported from python-docx v1.2.0: `src/docx/document.py::Document.paragraphs`*

---

## `Document.inline_shapes` / `DocumentPart.inline_shapes` — the inline-shape collection

**Syntax**

```matlab
shapes = document.inline_shapes;          % an InlineShapes collection
shapes = document.part.inline_shapes;     % the same collection, from the part
```

**Description**

`Document.inline_shapes` (`document.py:170-178`, un-stubbed at P8-3) delegates to
`self._part.inline_shapes`. `DocumentPart.inline_shapes`
(`parts/document.py:93-96`, un-stubbed at P8-3) is a **`@lazyproperty`**: it
mints `InlineShapes(self._element.body, self)` on first access and returns the
**same** handle on every subsequent access — ported with a
`_computed_` **logical flag** cache pair (design.md `@lazyproperty` rule; never
`isempty` as the sentinel, since an `InlineShapes` is a valid non-empty value).
An inline shape is a graphical object (a picture) contained in a run and flowed
like a character glyph.

**Example** (length parity + the shape's type; the lazyproperty returns one handle)

```matlab
IMG = "C:\Users\dougl\Repos\python-docx\tests\test_files\python-powered.png";
doc = mat2doc.Document();
doc.add_picture(char(IMG));
fprintf('doc=%d  part=%d  type=%s\n', ...
        doc.inline_shapes.len_(), doc.part.inline_shapes.len_(), ...
        string(doc.inline_shapes.getitem_(0).type));
% doc=1  part=1  type=PICTURE
a = doc.part.inline_shapes; b = doc.part.inline_shapes;
fprintf('lazyproperty same handle: %d\n', a == b);   % lazyproperty same handle: 1
```

*Ported from python-docx v1.2.0: `src/docx/parts/document.py::DocumentPart.inline_shapes`*

---

## `OpcPackage.core_properties_part_` — the default-creation branch (the silent survivor)

**Syntax**

```matlab
cp = package.core_properties;   % public accessor -> core_properties_part_.core_properties
```

**Description**

`OpcPackage.core_properties_part_` (`opc/package.py::_core_properties_part`,
un-stubbed at P8-3) returns the `CorePropertiesPart` related to the package,
**creating a default one if absent**. It is a protected accessor reached through
the public `package.core_properties`. The Python `except KeyError` →
`catch e; if e.identifier == "mat2doc:KeyError"`: only a missing part (KeyError)
triggers default creation via `CorePropertiesPart.default(self)` +
`relate_to(part, RT.CORE_PROPERTIES)`; a `ValueError` (more than one such
relationship) re-raises (`rethrow`). It is **dead on the M1 path** —
`default.docx` ships `docProps/core.xml`, so `part_related_by` succeeds and
returns it — and this default-creation branch is precisely the **silent survivor**
the C4 zero-stub exit condition had to clear.

**Example** (the found branch — a normal document already has a core-properties part)

```matlab
doc = mat2doc.Document();
cp = doc.core_properties;                 % part_related_by succeeds (found branch)
fprintf('core_properties class=%s\n', class(cp));
% core_properties class=mat2doc.opc.CoreProperties
```

**Example** (the default-creation branch — a bare package with no core rel)

```matlab
pkg = mat2doc.opc.OpcPackage();           % no relationships yet
cp = pkg.core_properties;                 % KeyError -> CorePropertiesPart.default
fprintf('default title=%s  last_modified_by=%s\n', cp.title, cp.last_modified_by);
% default title=Word Document  last_modified_by=python-docx
```

*Ported from python-docx v1.2.0: `src/docx/opc/package.py::OpcPackage._core_properties_part`*

---

:::{note}
**Hyperlink + BlockItemContainer were VERIFIED, not changed.** P8-3's scope
included a line-by-line **verify** of `text/hyperlink.py::Hyperlink` (live since
P4) and `blkcntnr.py::BlockItemContainer` — both were found fully ported and
correct, with **no edits**. Their surfaces are documented on the
[text/paragraph page](text_paragraph.md) and the
[document tier page](document.md).
:::

:::{important}
**★ PORT COMPLETE.** With P8-3, all catalog symbols are ported and there are
**zero live `mat2doc:notYetPorted`** sites. Only **P8-EXIT** — the system
campaigns, the whole-project Word-COM oracle, and the dogfooded Word user manual
— stands between here and **Mat2Doc COMPLETE**.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
