---
title: "mat2doc.text.Run / mat2doc.text.Text_ — the text/run API tier"
---

# `mat2doc.text.Run` / `mat2doc.text.Text_` — the text/run API tier

Ported from python-docx v1.2.0 `src/docx/text/run.py::Run` and
`src/docx/text/run.py::_Text` (package `+mat2doc/+text/`). `Run` is the
user-facing proxy for a `<w:r>` run element — the second P4 API-tier proxy after
[`Font`](text_font.md); `Text_` is the inert wrapper `Run.add_text` returns.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## What "API tier" means here — behaviorally equivalent, byte-neutral

`Run` is a `StoryChild` handle over the already-registered, already-byte-validated
`CT_R` (P4-1b); `Text_` is a bare `handle` over a `CT_Text`. Like `Font` and
`ColorFormat` before them, they add **no `register_element_cls` row, no oxml
class, and no serialization code** — nothing on the save path moves. Their entire
job is to **get/set** the correct `<w:r>` XML by delegating to the `CT_R` / `Font`
helpers validated at P4-1b / P4-4a. So the equivalence bar for this WP is
**behavioral**, not byte-fixture: every get returns the same value python-docx
returns, and every set produces the same serialized `<w:r>` bytes.

**M1 is trivially preserved.** The default template instantiates neither class
and the save path is unchanged, so the M1 17/17 byte-neutrality sweep holds
(`mat2doc.Document().save()` → 17/17 byte-identical to the frozen `references\s0001`
reference). There is no stale-pin risk (no registry rows) and no new D-number.
Gate-3 confirmed the whole surface byte/value-identical (`probe_diff` MATCH
**66/66**) with regression **598/598**.

| Python `src/docx/...` | MATLAB | class |
|---|---|---|
| `text/run.py::Run` | `+mat2doc\+text\` | `Run` |
| `text/run.py::_Text` | `+mat2doc\+text\` | `Text_` |

`+mat2doc\+text` (`mat2doc.text.Run`) is distinct from the oxml `+mat2doc\+oxml\+text`
(`mat2doc.oxml.text.CT_R`); `Run` **wraps** a `CT_R`, it is not one.

## The proxy shape — three names, one element; a `StoryChild` not an `ElementProxy`

`Run(r, parent)` wraps a **run** element (`w:r` / `CT_R`). The Python `__init__`
(run.py 34-36) sets `self._r = self._element = self.element = r` — **three names
for the same element** — ported verbatim as `r_` (private, working handle),
`element_` (private, what `font` wraps) and `element` (public, set but never read
inside run.py; kept for fidelity). All three hold the one `CT_R` handle.

Unlike `Font`/`ColorFormat` (which extend `ElementProxy`), `Run` extends
**`StoryChild`** — the docx block-item/run base that carries only a `parent` and
delegates `part` to it. `Run(r, parent).part` reaches `parent.part` through
`StoryChild` (verified against a `_PartStub` sentinel). `StoryChild` has no
element-identity `eq`/`ne` of its own, so two `Run` handles are compared by
plain handle identity.

## The live-vs-stub split

The Run surface is mostly **LIVE**; three members whose dependencies are not yet
ported are **stubbed** (design.md §7 — no silent approximation), each raising
`mat2doc:notYetPorted` naming the owning WP:

| member | status | owner |
|---|---|---|
| `text` get/set, `bold`/`italic`/`underline` get/set, `font`, `add_text`, `add_break`, `add_tab`, `clear`, `contains_page_break`, `mark_comment_range` | **LIVE** | — |
| `add_picture` | STUB | **P7** (`InlineShape` / `StoryPart.new_pic_inline`) |
| `iter_inner_content` | STUB | **P4-5b + P7** (`RenderedPageBreak` proxy / `Drawing`) |
| `style` get/set | STUB | **P4-7** (`CharacterStyle` / `DocumentPart.get_style`·`get_style_id`) |

`iter_inner_content` is stubbed as a **whole method** rather than partially
yielding string parts: the underlying `CT_R.inner_content_items` accessor is
itself stubbed at P4-1b (isinstance dispatch on the unregistered `CT_Drawing` /
`CT_LastRenderedPageBreak`), so there is no faithful partial path.

`style` is ported as a **Dependent read/write property** — its true Python shape —
so both `run.style` and `run.style = x` raise the *same* `notYetPorted`, and P4-7
fills the two bodies without changing the API. Auto-**display** of a `Run` is
safe: MATLAB catches the erroring `get.style` and simply **omits the style row**
(verified R2024b, Gate-2 Flag-2 RATIFIED); only an *explicit* `run.style` access
raises.

---

## `Run`

**Syntax**

```matlab
r   = mat2doc.oxml.OxmlElement("w:r");   % a CT_R (run)
run = mat2doc.text.Run(r, []);           % parent = a paragraph; [] is fine standalone
run.text = sprintf("a\tb");              % <w:t>a</w:t><w:tab/><w:t>b</w:t>
run.bold = true;                         % <w:rPr><w:b/></w:rPr>
run.add_break(mat2doc.enum.text.WD_BREAK.PAGE);  % <w:br w:type="page"/>
tf  = run.contains_page_break;           % false
```

:::{note}
The `text` setter splits on the **actual** tab (`\t` = `char(9)`) and newline
(`\n`/`\r` = `char(10)`/`char(13)`) characters. MATLAB double-quoted string
literals do **not** interpret backslash escapes, so use `sprintf("a\tb")` (or
`compose`, or explicit `char(9)`) to obtain the control characters — a bare
`"a\tb"` would set the literal five-character text `a\tb`.
:::

**Description**

Proxy for a `<w:r>` run: its text content, character formatting (via a delegated
`Font`), breaks/tabs, and comment-range marking. A `StoryChild` handle over the
P4-1b `CT_R`.

### `text` get/set (run.py 205-225)

Get returns the run's concatenated inner-content text (`CT_R.text` joins the text
equivalent of each `w:t`/`w:tab`/`w:br`/… child, P4-1b). Set replaces the run
content from a string — each `\t` → `<w:tab/>`, each `\n` **or** `\r` → `<w:br/>`;
the run's `<w:rPr>` formatting is preserved.

```matlab
run.text = sprintf("a\tb\nc\rd");
% <w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t><w:br/><w:t>d</w:t>
disp(run.text);          % a<TAB>b<LF>c<LF>d   (the \r reads back as \n)
```

An empty-string set yields a self-closing `<w:r/>` (get → `""`); a re-set
**replaces** rather than appends.

### `bold` / `italic` / `underline` — font-delegated tri-states

These do **not** hold their own state — each get/set forwards to
[`Font`](text_font.md) (run.py 98-151, 227-249):

- `bold` / `italic` → `font.bold` / `font.italic` (a tri-state `true`/`false`/`[]`);
- `underline` → `font.underline` (the `WD_UNDERLINE` tri-state-plus-enum).

`font` returns a **fresh** `Font(self._element)` each access (Python does not
cache it), so `run.bold = true` mutates the run's `<w:rPr>` through a throwaway
`Font` whose mutation persists on the shared element. Setting all three on one
run gives `<w:rPr><w:b/><w:i/><w:u w:val="single"/></w:rPr>`; `bold = []` removes
`<w:b>`; `bold = false` → `<w:b w:val="0"/>`.

### `font` (read-only, run.py 133-137)

Returns a fresh `Font` proxy for the run's character properties each access
(`Font(self._element)`; `parent` left `[]`). `run.font == run.font` is **true**
(both wrap the same `w:r`; element-identity `eq` inherited from `Font`'s
`ElementProxy`).

(id-run-add-break)=
### `add_break` — the `WD_BREAK` → `(type_, clear)` map (run.py 38-57)

`add_break(break_type)` appends a `<w:br>`; `break_type` is a `WD_BREAK` member
(P3-3), default `WD_BREAK.LINE`. It maps the member to a `(type_, clear)` pair,
then sets `br.type`/`br.clear` only when each is non-`[]`:

| `WD_BREAK` member | `(type_, clear)` | `<w:r>` inner |
|---|---|---|
| `LINE` (default) | `([], [])` | `<w:br/>` |
| `PAGE` | `("page", [])` | `<w:br w:type="page"/>` |
| `COLUMN` | `("column", [])` | `<w:br w:type="column"/>` |
| `LINE_CLEAR_LEFT` | `("textWrapping", "left")` | `<w:br w:clear="left"/>` |
| `LINE_CLEAR_RIGHT` | `("textWrapping", "right")` | `<w:br w:clear="right"/>` |
| `LINE_CLEAR_ALL` | `("textWrapping", "all")` | `<w:br w:clear="all"/>` |
| `TEXT_WRAPPING` (alias of `LINE_CLEAR_ALL`) | `("textWrapping", "all")` | `<w:br w:clear="all"/>` |

:::{important}
**The `LINE_CLEAR_*` results carry ONLY `@w:clear`, never `@w:type`.** The map
assigns `type_ = "textWrapping"`, which is the `CT_Br` `@w:type` **default**; the
`OptionalAttribute` setter removes an attribute set to its default (**D-delta-1**,
byte-confirmed). So `LINE_CLEAR_LEFT` serializes as `<w:br w:clear="left"/>`, not
`<w:br w:type="textWrapping" w:clear="left"/>` — byte-identical to python-docx.
:::

Any member outside the map (e.g. the `SECTION_*` members) reproduces Python's
`{...}[break_type]` **KeyError** via `error("mat2doc:KeyError", ...)`. The error
fires **before** `add_br`, so no partial `<w:br>` is left on the run
(`children == 0` on both sides).

### `add_tab` / `add_text` / `clear` (run.py 83-118)

- `add_tab()` appends a `<w:tab/>` (Word renders a tab character).
- `add_text(text)` appends a new `<w:t>` and returns a `Text_` wrapping it;
  leading/trailing whitespace forces `xml:space="preserve"`
  (`add_text(" hi ")` → `<w:t xml:space="preserve"> hi </w:t>`, while
  `add_text("plain")` → `<w:t>plain</w:t>`).
- `clear()` removes all run content but **preserves** the run formatting
  (`<w:rPr>`) and **returns this run** (so calls chain).

### `contains_page_break` (read-only, run.py 120-131, H4)

`true` when one or more **rendered** page-breaks (`<w:lastRenderedPageBreak>`
descendants) occur in the run — `bool(self._r.lastRenderedPageBreaks)` →
`~isempty(r_.lastRenderedPageBreaks())`. A **hard** (author) `<w:br w:type="page"/>`
is correctly **not** counted. Returns a plain logical (no `RenderedPageBreak`
proxy needed).

### `mark_comment_range(last_run, comment_id)` (run.py 176-186)

Marks the range of runs from this run to `last_run` (inclusive) as belonging to
comment `comment_id`: inserts a `<w:commentRangeStart w:id="…"/>` before this run
and a `<w:commentRangeEnd>` plus a `<w:commentReference>` run (styled
`CommentReference`) after `last_run`. For two runs in a `w:p`, the resulting child
sequence is `[commentRangeStart, r, r, commentRangeEnd, r]`, byte-identical to
python-docx.

**Example**

```matlab
p    = mat2doc.oxml.OxmlElement("w:p");
r1   = p.add_r();  r2 = p.add_r();
run1 = mat2doc.text.Run(r1, []);
run2 = mat2doc.text.Run(r2, []);
run1.add_text("hello ");
run2.add_text("world");
run1.bold = true;                                   % <w:rPr><w:b/></w:rPr> on r1
run1.add_break(mat2doc.enum.text.WD_BREAK.LINE_CLEAR_ALL);  % <w:br w:clear="all"/>
run1.mark_comment_range(run2, 3);                   % comment 3 spans run1..run2
disp(run1.contains_page_break);                     % false
```

*Ported from python-docx v1.2.0: `src/docx/text/run.py::Run`*

---

## `Text_`

**Syntax**

```matlab
r   = mat2doc.oxml.OxmlElement("w:r");
run = mat2doc.text.Run(r, []);
t   = run.add_text(" hi ");   % a mat2doc.text.Text_ over the new <w:t>
```

**Description**

The small wrapper `Run.add_text` returns. In python-docx v1.2.0 `_Text` is a bare
`object` subclass whose only state is the wrapped `<w:t>` element — it exposes
**no `text` property or any other public member** (verified against the v1.2.0
source, run.py 252-257, and the runtime oracle: `hasattr(t, "text")` is `False`).

`Text_` is ported faithfully as an inert `handle` holding a private `t_`, with
**zero public members** — no `text` accessor is added (adding one would be an
unported feature, design.md §7; Gate-2 Flag-1 RATIFIED). The underscore rotation
follows design.md §2: the module-private class `_Text` → `Text_`, and the private
`_t` → `t_`.

*Ported from python-docx v1.2.0: `src/docx/text/run.py::_Text`*

---

## Deviation posture — 0 new D-numbers

`Run` and `Text_` add no output-visible divergence. The standing adopt-only
classes that touch this WP:

- **D-delta-1** (`CT_Br` `OptionalAttribute` default-removal) — **engaged and
  byte-confirmed** by the `add_break(LINE_CLEAR_*)` cases (removes
  `@w:type="textWrapping"` so only `@w:clear` survives); adopt-only.
- **D10** (`CT_R.text` shadow) — engaged via `r_.text` get/set delegation;
  upstream (P4-1b), not re-adopted here.
- **D-serializer-nsdecl** — **does not engage**: every `Run` element is `w:`-only
  (no `r:` attribute), so no `ns0` is ever minted and all serhex byte pins
  (including the XML prolog) are byte-identical.

Gate-3 recorded `PASS-DEVIATION(D10, D-delta-1, D-serializer-nsdecl)` as
adopt-only context only — operational verdict a clean **PASS** with zero measured
deviation, **zero new D-numbers**.
