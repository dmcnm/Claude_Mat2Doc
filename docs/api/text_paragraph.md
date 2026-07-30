---
title: "mat2doc.text.Paragraph / mat2doc.text.Hyperlink / mat2doc.text.RenderedPageBreak — the paragraph/hyperlink/pagebreak API tier"
---

# `mat2doc.text.Paragraph` / `mat2doc.text.Hyperlink` / `mat2doc.text.RenderedPageBreak` — the paragraph API tier

Ported from python-docx v1.2.0 `src/docx/text/paragraph.py::Paragraph`,
`src/docx/text/hyperlink.py::Hyperlink` and
`src/docx/text/pagebreak.py::RenderedPageBreak` (package `+mat2doc/+text/`).
`Paragraph` is the **fifth P4 API-tier proxy** after [`Font`](text_font.md),
[`Run`](text_run.md), [`ParagraphFormat`](text_parfmt.md) and its
`TabStops`/`TabStop` sequence; `Hyperlink` and `RenderedPageBreak` are the read
proxies over the two non-run things a paragraph can hold. This WP is the **LAST
API-tier WP of Phase 4** — after it only the styles chain (P4-6/P4-7) remains
before **M2**.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## What "API tier" means here — behaviorally equivalent, byte-neutral

`Paragraph`, `Hyperlink` and `RenderedPageBreak` are pure API/proxy classes over
the already-byte-validated `CT_P` (P4-2) and `CT_Hyperlink` /
`CT_LastRenderedPageBreak` (P4-3). Like every P4 API proxy before them they add
**no `register_element_cls` row, no oxml class, and no serialization code** —
nothing on the save path moves. Their entire job is to **read** the correct
`<w:p>` / `<w:hyperlink>` / `<w:lastRenderedPageBreak>` XML through the P4-2/P4-3
helpers and, for the tree-mutating `Paragraph` methods (`add_run`, `clear`,
`text` setter, `insert_paragraph_before`), produce the same serialized bytes. So
the equivalence bar for this WP is **behavioral**, not byte-fixture: every get
returns the same value python-docx returns, and every mutation produces the same
bytes.

**M1 is trivially preserved.** The default template instantiates none of the
three classes and the save path is unchanged, so the M1 17/17 byte-neutrality
sweep holds (`mat2doc.Document().save()` → 17/17 byte-identical to the frozen
`references\s0001` reference). There is no stale-pin risk (no registry rows) and
no new D-number. Gate-3 confirmed the whole surface byte/value-identical
(`probe_diff` **MATCH 170/170**) with a NEW-Paragraph-API document byte pin
**17/17** and regression **57/57**.

| Python `src/docx/...` | MATLAB | class | base |
|---|---|---|---|
| `text/paragraph.py::Paragraph` | `+mat2doc\+text\` | `Paragraph` | `mat2doc.shared.StoryChild` |
| `text/hyperlink.py::Hyperlink` | `+mat2doc\+text\` | `Hyperlink` | `mat2doc.shared.Parented` |
| `text/pagebreak.py::RenderedPageBreak` | `+mat2doc\+text\` | `RenderedPageBreak` | `mat2doc.shared.Parented` |

`+mat2doc\+text` (`mat2doc.text.Paragraph`) is distinct from the oxml
`+mat2doc\+oxml\+text` (`mat2doc.oxml.text.CT_P`); `Paragraph` **wraps** a `CT_P`,
it is not one.

(id-para-instance-identity)=
## The identity regime — instance identity, not element identity (H5)

Unlike the `ElementProxy` subclasses ([`Font`](text_font.md),
[`ParagraphFormat`](text_parfmt.md), …), these three classes derive the
**standalone part-provider bases** `StoryChild` / `Parented`, neither of which
defines `eq`/`ne`. In python-docx v1.2.0 only `ElementProxy` overrides `__eq__`;
`Parented`/`StoryChild` do not — so two proxies wrapping the **same** element are
**not equal**, matching Python's default object identity. Mat2Doc reproduces this
exactly: a `Paragraph`/`Hyperlink`/`RenderedPageBreak` is compared by MATLAB's
**default handle identity**, so `Hyperlink(h, d) == Hyperlink(h, d)` is `false`
(instance identity) — the **opposite** of a `ParagraphFormat`, where two proxies
over the same element ARE equal. Both regimes are faithful and coexist correctly.
`clear()` is the one identity-bearing method: it returns the **same** handle
(`p.clear() == p` is `true`, H5), matching Python `return self`.

Because `Parented`/`StoryChild` hold no `element_`, `Hyperlink` and
`RenderedPageBreak` declare their own private handles
(`hyperlink_`/`element_`, `lastRenderedPageBreak_`/`element_`) — the `_element`
attribute is set for fidelity but, as in python-docx, never read inside those two
modules.

(id-para-list-surface)=
## The list-property surface — homogeneous `1×N` object arrays, heterogeneous `1×N` cell (auditor ACCEPT)

`Paragraph.runs` / `hyperlinks` / `rendered_page_breaks` and `Hyperlink.runs`
return **plain Python lists** (list comprehensions), NOT collection classes.
They are **homogeneous** (all one proxy type), so the faithful MATLAB surface is a
**homogeneous `1×N` object array**, seeded via `Class.empty(1,0)`:

- `numel(x) == len(x)`; native 1-based `x(i)` realizes the design.md §2 dunder
  mapping `x[i] → x(i+1)` with **no wrapper machinery**.
- An empty property is a typed `1×0` array (`numel == 0`) — the faithful `[]`.
- This mirrors the plain Python list directly and matches the
  [`TabStops.to_array()`](text_parfmt.md#tabstops) `1×N` shape (P4-5a). The
  `getitem_`/`len_` interim surface was for `TabStops`, a genuine **collection
  class** with custom `IndexError` messages — correctly **not** replicated for
  these three plain-list properties.

`iter_inner_content` is different: it yields a **mixed `Run|Hyperlink`** sequence.
`Run < StoryChild` and `Hyperlink < Parented` share **no**
`matlab.mixin.Heterogeneous` base (they share none in Python either), so they
cannot inhabit one object array. The faithful surface is a **heterogeneous `1×N`
cell array in document order** — each cell a `Run` (for a `<w:r>`) or a
`Hyperlink` (for a `<w:hyperlink>`). The Python generator is precomputed (H9: the
source performs no mutation during iteration, so laziness is unobservable); the
`./w:r | ./w:hyperlink` xpath is already document-ordered (H1).

:::{important}
**Both surface decisions were ACCEPTED at Gate 2 as the standing surface, not
merely interim.** Minting a common `matlab.mixin.Heterogeneous` base solely to
co-array `Run` and `Hyperlink` would distort the ported class hierarchy
(design.md §7, no improvements) and endanger the H5 instance-identity regime. A
homogeneous object array + a heterogeneous cell preserve order, type, and
identity with zero class-model change.
:::

---

## `Paragraph`

**Syntax**

```matlab
doc  = mat2doc.Document();                     % a real ProvidesStoryPart parent
p    = mat2doc.oxml.OxmlElement("w:p");        % a CT_P
para = mat2doc.text.Paragraph(p, doc);
run  = para.add_run("Hello");                  % append a run holding text
para.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
n    = numel(para.runs);                       % len(paragraph.runs)
txt  = para.text;                              % concatenated inner-content text
```

**Description**

A block-level paragraph — a `<w:p>` made of runs and hyperlinks (its inner
content) plus paragraph-level formatting (its `<w:pPr>`). A `StoryChild`, not an
`ElementProxy` (see [the identity regime](#id-para-instance-identity)). The
constructor stores the `<w:p>` under two names (`p_` the working handle,
`element_` what `paragraph_format` wraps) — python-docx's
`self._p = self._element = p`.

### Method / property surface

| Member | Kind | Behavior |
|---|---|---|
| `add_run(text, style)` | method | Append a run; `text` (default `[]`=None) may hold `\t`→`<w:tab/>` and `\n`/`\r`→line breaks; `if text:` is a **non-empty-string** test (H4). `style` reaches the P4-7 stub (below). |
| `alignment` | get/set | `WD_ALIGN_PARAGRAPH` member or `[]` — delegates to `CT_P.alignment` (`<w:jc>`). |
| `clear()` | method | Remove all content, **preserve `<w:pPr>`**, return the **same** handle (H5). |
| `contains_page_break` | get | `bool` — `~isempty(p_.lastRenderedPageBreaks())` (H4 `bool(list)`). |
| `hyperlinks` | get | homogeneous `1×N` `Hyperlink` array (one per `<w:hyperlink>` child). |
| `insert_paragraph_before(text, style)` | method | Return a new paragraph inserted directly before this one — see the distinct guards below. |
| `iter_inner_content()` | method | heterogeneous `1×N` cell of `Run`/`Hyperlink` in document order. |
| `paragraph_format` | get | a **fresh** [`ParagraphFormat`](text_parfmt.md) wrapping `element_` each access (not cached). |
| `rendered_page_breaks` | get | homogeneous `1×N` `RenderedPageBreak` array (usually empty). |
| `runs` | get | homogeneous `1×N` `Run` array (one per `<w:r>` child). |
| `style` | get/set | **P4-7 stub** — both raise `mat2doc:notYetPorted` (below). |
| `text` | get/set | get = concatenated inner-content text (`CT_P.text`, D10); set = `clear(); add_run(text)`. |

### The distinct truthiness guards (H4) — `add_run` vs `insert_paragraph_before`

These two methods guard `text` and `style` **differently**, and the port keeps
them distinct:

- **`add_run(text, style)`** — `if text:` (skip an empty/`None` text) and
  `if style:` (Python truthiness: `None`→skip, `""`→skip, a style object→apply).
- **`insert_paragraph_before(text, style)`** — `if text:` (truthiness) but
  **`if style is not None`** (identity, `~isequal(style, [])`). So a named or even
  empty-string style is applied — which reaches the P4-7 stub and raises.

### The `Paragraph.style` P4-7 stub delegation (faithful propagation, not a stand-in)

`style` get **and** set delegate through `self.part.get_style` /
`self.part.get_style_id` → `DocumentPart.get_style`/`get_style_id`, which are the
**P4-7 feature stubs** raising `mat2doc:notYetPorted`. The delegation is ported
**exactly**: the getter runs `p_.style` first, then hits the stub; the setter hits
`get_style_id` before `p_.style` is touched (matching Python order). Both raise
until P4-7 wires styles — this is faithful stub **propagation**, not an invented
placeholder. The same stub is reached through `add_run(text, style)` (when
`style` is truthy) and `insert_paragraph_before(text, style)` (when `style` is not
`None`). `style` is a `Dependent` read/write property so it preserves its true
Python shape — both `para.style` and `para.style = x` raise the same identifier —
and auto-display is safe (MATLAB catches the erroring getter and omits the row).
Consequently `Paragraph.text = ...` (`clear(); add_run(text)`) is safe: `add_run`
with no `style` never touches the stub.

**Example**

```matlab
doc  = mat2doc.Document();                     % a real ProvidesStoryPart parent
p    = mat2doc.oxml.OxmlElement("w:p");        % a fresh CT_P
para = mat2doc.text.Paragraph(p, doc);
para.add_run("Hello ");                        % append a run holding text
para.add_run("World");
para.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
disp(numel(para.runs));                        % 2   (runs is a 1xN Run object array)
disp(para.text);                               % "Hello World"
disp(string(para.alignment));                  % CENTER
```

`iter_inner_content` returns the **heterogeneous cell** in document order:

```matlab
doc = mat2doc.Document();
W   = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";
xml = "<w:p " + W + ">" + ...
      "<w:r><w:t>before </w:t></w:r>" + ...
      "<w:hyperlink><w:r><w:t>link</w:t></w:r></w:hyperlink>" + ...
      "<w:r><w:t> after</w:t></w:r></w:p>";
p    = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
para = mat2doc.text.Paragraph(p, doc);
items = para.iter_inner_content();             % 1x3 cell: {Run, Hyperlink, Run}
disp(numel(items));                            % 3
disp(class(items{2}));                         % mat2doc.text.Hyperlink
disp(para.text);                               % "before link after"
```

*Ported from python-docx v1.2.0: `src/docx/text/paragraph.py::Paragraph`*

---

## `Hyperlink`

**Syntax**

```matlab
doc = mat2doc.Document();
W   = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";
xml = "<w:hyperlink " + W + " w:anchor=""_Toc12345"">" + ...
      "<w:r><w:t>See intro</w:t></w:r></w:hyperlink>";
h = mat2doc.text.Hyperlink( ...
      mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8"))), doc);
frag = h.fragment;      % "_Toc12345"
addr = h.address;       % "" (internal jump; no r:id)
```

**Description**

A hyperlink occurs as a child of a paragraph, at the same level as a `Run`, and
itself contains the runs holding its visible text — "in-between", less than a
paragraph and more than a run. A `Parented`, so it has
[instance identity](#id-para-instance-identity) and a read-only surface.

### The full read surface

| Member | Behavior |
|---|---|
| `address` | the hyperlink target "URL", or `""` for an internal jump. `if rId` is a non-empty-string test (H4); resolved **LIVE** as `parent.part.rels[rId].target_ref`. |
| `fragment` | the URI fragment (named anchor), **without** the `#`; `""` when absent (`anchor or ""`, H3/H4). |
| `url` | a web-URL convenience: `""` when no `address`; `address#fragment` when a fragment is present; else `address`. |
| `runs` | homogeneous `1×N` `Run` array. **The runs get `self._parent`** (the hyperlink's parent, a `ProvidesStoryPart`) as their parent, NOT the hyperlink — ported exactly. |
| `text` | concatenated run text (`CT_Hyperlink.text`, D10). |
| `contains_page_break` | `bool` — `~isempty(hyperlink_.lastRenderedPageBreaks())`. |

### `address` — internal jump vs external target (the H4 rId test)

| `<w:hyperlink>` state | `rId` | `address` | `url` |
|---|---|---|---|
| **external** (`r:id="rId7"`) | non-empty string | `rels[rId].target_ref` (LIVE, e.g. `"https://example.com/"`) | same, or `address#fragment` |
| **external + anchor** (`r:id` + `w:anchor="foo"`) | non-empty string | the target ref | `"…#foo"` |
| **internal jump** (`w:anchor="_Toc…"`, no `r:id`) | `[]`/`""` | `""` | `""` |
| **bare** (no `r:id`, no `w:anchor`) | `[]` | `""` | `""` |

Only `address` (and therefore `url`) reads the relationships; `fragment` reads
only `@w:anchor`, so it is populated for an internal jump even though `address`
is `""`.

**Example**

```matlab
doc = mat2doc.Document();
W   = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";

% internal jump -- fragment set, address empty (no r:id to resolve)
xml = "<w:hyperlink " + W + " w:anchor=""_Toc12345"">" + ...
      "<w:r><w:t>See intro</w:t></w:r></w:hyperlink>";
h = mat2doc.text.Hyperlink( ...
      mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8"))), doc);
disp(h.fragment);        % "_Toc12345"
disp("[" + h.address + "]");   % "[]"  (empty string)
disp(h.text);            % "See intro"

% external -- a LIVE relationship resolves address/url
R   = "xmlns:r=""http://schemas.openxmlformats.org/officeDocument/2006/relationships""";
rid = doc.part().relate_to("https://example.com/", ...
        mat2doc.opc.RELATIONSHIP_TYPE.HYPERLINK, true);   % external rel
xml = "<w:hyperlink " + W + " " + R + " r:id=""" + rid + """>" + ...
      "<w:r><w:t>Example</w:t></w:r></w:hyperlink>";
h = mat2doc.text.Hyperlink( ...
      mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8"))), doc);
disp(h.address);         % "https://example.com/"
disp(h.url);             % "https://example.com/"
```

*Ported from python-docx v1.2.0: `src/docx/text/hyperlink.py::Hyperlink`*

---

## `RenderedPageBreak`

**Syntax**

```matlab
brks = para_elm.lastRenderedPageBreaks();      % CT_LastRenderedPageBreak array
rpb  = mat2doc.text.RenderedPageBreak(brks(1), someStoryParent);
pre  = rpb.preceding_paragraph_fragment;       % [] (None) or a loose Paragraph
post = rpb.following_paragraph_fragment;        % [] (None) or a loose Paragraph
```

**Description**

A `<w:lastRenderedPageBreak>` — a page-break Word inserts when it runs out of room
on a page (never a "hard" author break). python-docx never inserts these; they are
useful only for text-extraction of existing documents. A `Parented`, with
[instance identity](#id-para-instance-identity) and two read-only fragment
properties.

### `precedes_all_content` and the preceding/following fragment split

The two properties return **loose** `Paragraph` fragments **divorced from the
document body** — the content preceding / following this break. All the heavy
split machinery (the H1 xpath positional indexing, the deepcopy/remove) lives in
`CT_LastRenderedPageBreak` (P4-3, already audited); these proxy getters only
guard and wrap:

| Property | Guard (on the oxml element) | Returns |
|---|---|---|
| `preceding_paragraph_fragment` | `precedes_all_content()` true (nothing precedes the break) | `[]` (Python `None`) |
| | otherwise | `Paragraph(preceding_fragment_p(), parent)` |
| `following_paragraph_fragment` | `follows_all_content()` true (nothing follows) | `[]` (Python `None`) |
| | otherwise | `Paragraph(following_fragment_p(), parent)` |

`precedes_all_content` / `follows_all_content` are no-arg methods on the P4-3 oxml
element, used directly as booleans (H4). The getters CAN raise `mat2doc:ValueError`
(via `preceding_fragment_p`/`following_fragment_p`) when this is **not** the first
rendered page-break in its paragraph — faithful to python-docx; MATLAB
auto-display catches the erroring getter and omits the row.

(id-atomic-break-in-hyperlink)=
### The atomic break-inside-hyperlink case

When the break falls **inside a hyperlink**, the split is not a clean cut: the
**whole `<w:hyperlink>`** goes to the **preceding** fragment (localnames
`[pPr, hyperlink]`), and the **following** fragment excludes it entirely
(`[pPr, r]`). The hyperlink is treated as atomic — never split across the page
boundary. This is entirely the P4-3 `CT_LastRenderedPageBreak` algorithm; the
proxy only wraps the resulting `CT_P` fragments. Gate-3 pinned both fragments'
serialized bytes byte-identical to python-docx.

**Example**

```matlab
doc = mat2doc.Document();
W   = "xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main""";

% mid-run break -- content splits either side
xml  = "<w:p " + W + "><w:r>" + ...
       "<w:t>foo</w:t><w:lastRenderedPageBreak/><w:t>bar</w:t></w:r></w:p>";
p    = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
para = mat2doc.text.Paragraph(p, doc);
rpb  = para.rendered_page_breaks(1);           % the one RenderedPageBreak
disp(rpb.preceding_paragraph_fragment.text);   % "foo"
disp(rpb.following_paragraph_fragment.text);    % "bar"

% leading break -- nothing precedes, so the preceding fragment is [] (None)
xml  = "<w:p " + W + "><w:r>" + ...
       "<w:lastRenderedPageBreak/><w:t>foobar</w:t></w:r></w:p>";
p    = mat2doc.oxml.parse_xml(uint8(unicode2native(char(xml), "UTF-8")));
para = mat2doc.text.Paragraph(p, doc);
rpb  = para.rendered_page_breaks(1);
disp(isempty(rpb.preceding_paragraph_fragment));   % 1  (None)
disp(rpb.following_paragraph_fragment.text);        % "foobar"
```

*Ported from python-docx v1.2.0: `src/docx/text/pagebreak.py::RenderedPageBreak`*

---

## The `s7` ns-decl reopen-check — still unreachable, re-booked at P7

P4-3 accepted a `D-serializer-nsdecl` **"Family-2"** residual: a loosely-created
`w:hyperlink` whose `rId` is **set-then-cleared** leaves an orphaned `xmlns:ns0`
(lxml keeps it; our serializer recomputes used namespaces and drops it). It was
deemed unreachable via public API at P4-3, with a mandated reopen-check now that
the `Hyperlink` proxy exists.

**Verdict: still UNREACHABLE.** Independently confirmed at all three gates.
Every P4-5b public path that touches a hyperlink only **reads** it —
`Paragraph.hyperlinks`/`iter_inner_content` read existing `<w:hyperlink>`
children, and the whole `Hyperlink` proxy surface (`address`/`fragment`/`runs`/
`text`/`url`/`contains_page_break`) is read-only; no property **sets or clears**
`rId`. There is **no `add_hyperlink`** method on the `Paragraph` proxy in
python-docx v1.2.0. The `RenderedPageBreak` split deep-copies/removes elements but
never touches `rId`, and operates on fragments rooted under a `w:document` that
declares `xmlns:r`, so any retained `r:id` uses the existing prefix. Gate-3
scanned every `serialize_part_xml` byte string on both sides for `ns0`
(hex `6E7330`): **0 hits**. The set-then-clear-`rId` sequence that produces the
orphan is **not expressible** through any P4-5b public API. `D-serializer-nsdecl`
Family-2 **stands**, non-engaged; the reopen-check is **re-booked at P7**
(`add_picture` / any future `rId`-mutating public path). **No new D-number.**

## Deviation posture — 0 new D-numbers

`Paragraph`, `Hyperlink` and `RenderedPageBreak` add no output-visible
divergence. The two standing adopt-only classes that touch this WP:

- **D-serializer-nsdecl** — **non-engaged / unreachable** (the s7 reopen-check
  above): 0 orphaned `ns0` across every scenario incl. the external/internal
  hyperlink and the break-inside-hyperlink fragment split.
- **D10** (CT_P/CT_Hyperlink `.text` shadow) — consumed **read-only** via
  `para.text` / `hyperlink.text` (the already-audited `getText_` override);
  value-exact, no new manifestation.

Gate-3 recorded a clean **PASS** with zero measured byte/value deviation.

## API tier COMPLETE — styles next → M2

P4-5b closes the **P4 API tier**. The full user-facing text object model is now
ported: [`Font`](text_font.md) (P4-4a), [`Run`](text_run.md) (P4-4b),
[`ParagraphFormat`/`TabStops`/`TabStop`](text_parfmt.md) (P4-5a), and
`Paragraph`/`Hyperlink`/`RenderedPageBreak` (P4-5b) — all sitting on the complete
text-oxml element layer (P4-1a…P4-3). What remains before **M2** is the **styles
chain** (P4-6 oxml/styles + P4-7 the API), which lands `Paragraph.style` and
`Run.style` for real (retiring the `notYetPorted` stubs) and, with
`Document.add_paragraph`/`add_heading` (P4-6), produces a real, byte-identical
`document.xml` that opens clean in Word (the **M2** milestone + its Word-COM
oracle). The P4-5a/P4-5b G-scenarios — fresh documents whose paragraphs are built
**through the new Paragraph API** and compared part-by-part against a frozen
python-docx reference (**17/17** byte-identical, `word/document.xml` byte-identical)
— are the closest guard yet to that `add_paragraph` path.
