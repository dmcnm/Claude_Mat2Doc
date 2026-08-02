---
title: "comments — the comments tier (CT_Comments · CT_Comment · Comments · Comment · CommentsPart)"
---

# `comments` — the comments tier (the P8-2 WP)

Ported from python-docx v1.2.0 `src/docx/oxml/comments.py`
(`CT_Comments` / `CT_Comment`), `src/docx/comments.py`
(`Comments` / `Comment`), `src/docx/parts/comments.py` (`CommentsPart`) and the
un-stubbed `Document.add_comment` / `Document.comments`
(`src/docx/document.py`) + `DocumentPart.comments` / `_comments_part`
(`src/docx/parts/document.py`). This is the **second work package of Phase 8** —
it lands the two comments `CT_*` element classes (the
`+mat2doc\+oxml\+comments` package), the `Comments` / `Comment` API proxies (the
new `+mat2doc\+comments` package), the on-demand `CommentsPart`, registers the
**2** comment tags, and un-stubs the whole comment-authoring path — all
**byte-neutral** for M1.

:::{important}
**★ Registering `w:comments` / `w:comment` is byte-neutral — M1 stays 17/17.**
`default.docx` ships **no** `word/comments.xml` and no other M1 part contains a
`<w:comments>` / `<w:comment>` node, so the two registry rows re-class **zero**
of the 17 M1 parts (unlike the P8-1 numbering block, which re-classed
`styles.xml` + `numbering.xml`). The `WML_COMMENTS → CommentsPart` PartFactory
flip is likewise inert on `default.docx` (which has no comments part). **Zero new
D-numbers.**
:::

:::{note}
**On-demand part.** A comments part is created **only when the first comment is
added** (or the `comments` collection is first accessed):
`DocumentPart._comments_part` calls `CommentsPart.default(package)`, which
materializes `word/comments.xml`, its `[Content_Types].xml` `<Override>`, and the
`word/_rels/document.xml.rels` COMMENTS relationship. A document with no comments
emits none of these — the M1 spine is untouched.
:::

---

## `CT_Comments` — the `<w:comments>` root

**Syntax**

```matlab
cs   = mat2doc.oxml.OxmlElement("w:comments");   % registry -> CT_Comments
c    = cs.add_comment();                          % a minimum-valid <w:comment w:id="N">
lst  = cs.comment_lst;                            % the <w:comment> children (document order)
hit  = cs.get_comment_by_id(id);                  % the <w:comment> with @w:id, else [] (None)
```

**Description**

`CT_Comments` (`comments.py::CT_Comments`) is the root of a comments part
(`word/comments.xml`). It is essentially a **set** of `<w:comment>` elements
implemented as a list — a comment's offset in the collection is arbitrary; the
`w:id` is the identity. Its one descriptor is
`comment = ZeroOrMore("w:comment")` with default `successors=()` (**append**,
H11).

**`add_comment()`** appends a **minimum-valid** `<w:comment>`: a `w:id` unique
among the existing comments (via `_next_available_comment_id`), an empty required
`w:author`, and a single `CommentText` paragraph holding a `CommentReference` run
with an `<w:annotationRef/>` — content is added later through the `Comment`
proxy. The generated public `add_comment` is **suppressed** by this explicit
method (xmlchemy no-ops a generated member whose name is already defined — the
`CT_Num.add_lvlOverride` precedent). **`get_comment_by_id(id)`** returns the
first `<w:comment>` whose `@w:id` matches (`(./w:comment[@w:id='N'])[1]`), or
`[]` (None) on a miss (H3 — not an empty typed array).
**`_next_available_comment_id`** is `max(used ids) + 1`, falling back — only if
that would overflow a signed 32-bit int — to the first unused non-negative
integer; the arithmetic is on id **values**, never collection indices (H1).

**Example**

```matlab
cs = mat2doc.oxml.OxmlElement("w:comments");     % a CT_Comments
c1 = cs.add_comment();                            % <w:comment w:id="0">
c2 = cs.add_comment();                            % <w:comment w:id="1">
c1.author = "Amy"; c1.initials = "AJ";
fprintf('ids: %d, %d\n', c1.id, c2.id);
% ids: 0, 1

hit = cs.get_comment_by_id(1);
fprintf('get_comment_by_id(1) == c2 : %d\n', hit == c2);   % 1 (same handle, H5)
miss = cs.get_comment_by_id(99);
fprintf('get_comment_by_id(99) empty: %d\n', isequal(miss, []));   % 1 (None)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/comments.py::CT_Comments`*

---

## `CT_Comment` — a single `<w:comment>` (★ ST_DateTime goes live)

**Syntax**

```matlab
c = cs.comment_lst(1);          % a CT_Comment
c.id        = 0;                % RequiredAttribute w:id (ST_DecimalNumber)
c.author    = "Amy";            % RequiredAttribute w:author (ST_String)
c.initials  = "AJ";             % OptionalAttribute w:initials (ST_String; [] -> absent)
c.date      = datetime(...);    % OptionalAttribute w:date (ST_DateTime; [] -> absent)
p   = c.add_p();                % a new <w:p> child (append)
lst = c.inner_content_elements; % all <w:p> and <w:tbl>, document order
```

**Description**

`CT_Comment` (`comments.py::CT_Comment`) is one comment. Like a table cell it is
a **story** — it can hold paragraphs and tables (`p` / `tbl` = `ZeroOrMore`,
`successors=()` → append). It carries two **required** attributes (`w:id` =
`ST_DecimalNumber`, `w:author` = `ST_String`) and two **optional** ones
(`w:initials` = `ST_String`, `w:date` = `ST_DateTime`); assigning `[]` (None) to
an optional attribute removes it (H3 tri-state).

**★ `w:date` is the first reachable `ST_DateTime` consumer in the whole port.**
`ST_DateTime` was ported at P3-2 but had no live call site until comments landed.
The **write path** stores a **tz-aware UTC** instant and `convert_to_xml`
normalizes it to whole-seconds with a literal `Z`:

```matlab
c1.date = datetime(2021,6,15,9,30,0,"TimeZone","UTC");
xmlbytes = mat2doc.oxml.serialize_part_xml(cs);
xmltxt   = native2unicode(xmlbytes, 'UTF-8');
tok = regexp(xmltxt, 'w:date="([^"]*)"', 'tokens', 'once');
fprintf('serialized @w:date = %s\n', tok{1});
% serialized @w:date = 2021-06-15T09:30:00Z
d = c1.date;
fprintf('date getter class=%s tz=%s\n', class(d), d.TimeZone);
% date getter class=datetime tz=UTC
```

:::{note}
**ST_DateTime under-accept — discharged with zero new D-number.** The
`ST_DateTime.convert_from_xml` **read** path implements the extended-format
`xsd:dateTime` subset that Word and OOXML producers emit; it byte-matches
python-docx on every schema-valid shape (Zulu, fractional Zulu, naive,
±hh:mm offsets, `+00:00`, date-only). A handful of **schema-invalid** shapes
that CPython's lenient `datetime.fromisoformat` still parses under-accept to the
1970 epoch instead (seconds-omitted, non-zero-padded fields, non-`T` separator,
offset-with-seconds, basic format, week/ordinal) — a **read-path** divergence on
lexically-invalid input, never a raise or plausible-wrong value. This is the
same safe ASCII-subset under-accept class already covered by the signed
**D-002**; the write path the comments API exercises is byte-exact, so
`ST_DateTime` going live discharges the P3-2 deferred ruling with **no new
D-number**. (`validation\summary\decision_2026-07-26_st_datetime_underaccept.md`,
DISCHARGE.)
:::

*Ported from python-docx v1.2.0: `src/docx/oxml/comments.py::CT_Comment`*

---

## `Comments` — the document's comment collection

**Syntax**

```matlab
cs = document.comments;                       % a Comments proxy
n  = cs.len_();                                % __len__: number of comments
a  = cs.to_array();                            % __iter__: a 1xN Comment array
c  = cs.add_comment(text, author, initials);   % add + return a Comment
c  = cs.get(comment_id);                       % the Comment with that id, or []
```

**Description**

`Comments` (`comments.py::Comments`) is a **plain `handle`** proxy (not an
`ElementProxy`) — python-docx holds **both** the `<w:comments>` element and the
`CommentsPart` (two references, not the single-element `ElementProxy` shape). The
dunder surface maps `__iter__ → to_array()` and `__len__ → len_()` (the `Styles`
collection precedent); `to_array` / `add_comment` / `get` mint **fresh** `Comment`
views each call (python-docx does not cache them, H5).

**`add_comment(text, author, initials)`** appends a `<w:comment>`, sets its
`author` / `initials`, stamps `date = datetime("now","TimeZone","UTC")` (the live
ST_DateTime write), and — when `text` is non-empty — splits it on `"\n"`, adding
the first segment as a run in the seeded paragraph and each subsequent segment as
a new `CommentText` paragraph (`Document.add_paragraph` semantics). `author`
defaults to `""`, `initials` to `""`; passing `[]` for `initials` omits the
attribute (H3). **`get(comment_id)`** returns a fresh `Comment` for that id, or
`[]` on a miss.

*Ported from python-docx v1.2.0: `src/docx/comments.py::Comments`*

---

## `Comment` — a single comment (a `BlockItemContainer`)

**Syntax**

```matlab
c.author       % read/write; the recorded author (required, "" allowed)
c.comment_id   % read-only; the unique integer identifier
c.initials     % read/write; the recorded initials, [] when absent
c.text         % read-only; paragraph texts joined by "\n"
c.timestamp    % read-only; the authored datetime, or []
c.add_paragraph(text, style);   % append a paragraph (CommentText when style is [])
```

**Description**

`Comment` (`comments.py::Comment`) extends **`BlockItemContainer`** — the **C3
element-accessor seam**. Like `_Cell` (and unlike `_BaseHeaderFooter`), it stores
a **concrete** element (the `<w:comment>`) and does **not** override the
`element_()` seam: the base `BlockItemContainer.element_()` returning the stored
element is exactly right, so the inherited surface (`add_paragraph`, `paragraphs`,
`add_table`, `tables`) operates directly on the `<w:comment>`. **`BlockItemContainer`
was not re-refactored** (its `git diff --stat` is untouched by P8-2). The comment's
`StoryChild` parent is the `CommentsPart` itself (`XmlPart.part()` returns
`self`), so `add_run` / `add_paragraph` on a comment resolve with no
`notYetPorted`.

**`add_paragraph(text, style)`** delegates to `super().add_paragraph` and, when
`style` is `[]` (None), sets the new paragraph's `<w:p>` style to `"CommentText"`
**directly on the element** — python-docx uses `paragraph._p.style`, which
bypasses `paragraph.style` (that would raise when the style is absent from the
styles part). MATLAB `Paragraph` exposes no public element accessor, so the
override reads the just-appended `<w:p>` back off the comment element as
`element_().p_lst(end)`; under the append invariant (`w:p` has `successors=()`,
`super().add_paragraph` appends exactly one `<w:p>`) this is byte-identically the
element `paragraph._p` references (Gate-2 VERIFY-COMMENTTEXT — CONFIRMED, keep
`p_lst(end)`; no D-number). **`text`** joins the paragraph texts with `"\n"` and
applies **no** `strip` (H16 latent — join only).

**Example** (the high-level authoring path)

```matlab
doc = mat2doc.Document();
p = doc.add_paragraph();
r = p.add_run("anchored text");
cm = doc.add_comment(r, "Nice point", "Amy", "AJ");
fprintf('author=%s id=%d initials=%s text=%s\n', ...
        cm.author, cm.comment_id, cm.initials, cm.text);
% author=Amy id=0 initials=AJ text=Nice point

allc = doc.comments;
fprintf('comments.len_=%d\n', allc.len_());          % comments.len_=1
c0 = allc.get(0);
fprintf('get(0).author=%s ; get(99) empty=%d\n', ...
        c0.author, isequal(allc.get(99), []));       % get(0).author=Amy ; get(99) empty=1
```

**Example** (multi-paragraph — `"\n"` splits into `CommentText` paragraphs)

```matlab
cm2 = doc.comments.add_comment("first" + string(newline) + "second", "Bob", "B");
fprintf('paragraphs=%d text=%s\n', ...
        numel(cm2.paragraphs), replace(cm2.text, newline, '\n'));
% paragraphs=2 text=first\nsecond
```

**Example** (`initials = []` → `@w:initials` omitted, H3 tri-state)

```matlab
cm3 = doc.comments.add_comment("hi", "Cara", []);   % [] (None) -> omit @w:initials
fprintf('initials empty=%d\n', isequal(cm3.initials, []));   % initials empty=1
```

*Ported from python-docx v1.2.0: `src/docx/comments.py::Comment`*

---

## `CommentsPart` — the comments.xml part (on demand)

**Syntax**

```matlab
cp = mat2doc.parts.CommentsPart.default(package);   % a new empty word/comments.xml
cs = cp.comments;                                    % a Comments proxy
```

**Description**

`CommentsPart` (`parts/comments.py::CommentsPart`) is a `StoryPart` subclass for
`word/comments.xml` — it parses on load and re-serializes on save through
`serialize_part_xml`, inheriting `blob` / `element` / `paragraphs` from
`StoryPart < XmlPart`. It is the **last new part type** in the port. Unlike the
numbered header/footer parts, comments live in the single fixed part
`word/comments.xml`, created **on demand**: `default.docx` ships none, so
`DocumentPart._comments_part` materializes one via `CommentsPart.default(package)`
on the first comment access (the `SettingsPart.default` / header-footer
precedents). Its custom constructor also stores `self._comments = element`
(→ `comments_` cache, the `comments` accessor's source), and it declares its
**own static `load`** (the inherited-static trap — without it the
`WML_COMMENTS → CommentsPart` PartFactory flip would silently build a base
`XmlPart`). The flip is **byte-neutral**: `CommentsPart` inherits `XmlPart.blob`,
so a reloaded comments part's **type** changes but its emitted bytes do not.

**Example** (on-demand materialization)

```matlab
doc2 = mat2doc.Document();
RT = mat2doc.opc.RELATIONSHIP_TYPE;
had = true;
try; doc2.part.part_related_by(RT.COMMENTS); catch; had = false; end
fprintf('comments part before access exists: %d\n', had);   % 0 (none yet)

cs2 = doc2.comments;                       % materializes word/comments.xml
fprintf('empty comments len_=%d\n', cs2.len_());   % 0
cp = doc2.part.part_related_by(RT.COMMENTS);
fprintf('materialized part class=%s partname=%s\n', class(cp), string(cp.partname));
% materialized part class=mat2doc.parts.CommentsPart partname=/word/comments.xml
```

*Ported from python-docx v1.2.0: `src/docx/parts/comments.py::CommentsPart`*

---

## `Document.add_comment` / `comments` — the authoring path + comment-range markers

**Syntax**

```matlab
c  = document.add_comment(runs, text, author, initials);   % anchor a comment to runs
cs = document.comments;                                     % the Comments collection
```

**Description**

`Document.add_comment` (`document.py::Document.add_comment`, un-stubbed at P8-2)
anchors a comment to one or more runs and returns the `Comment`. `runs` is a
single `Run` or a non-empty `Run` array; only the **first** and **last** run are
used — they delimit the comment reference range. python-docx wraps a lone `Run`
in a list so `runs[0]` / `runs[-1]` index it; in MATLAB a scalar `Run` and a
`1×N` `Run` array both index with `(1)` / `(end)`, so the wrap is a no-op. It
calls `self.comments.add_comment(text, author, initials)` then
`first_run.mark_comment_range(last_run, comment.comment_id)`, which inserts a
`<w:commentRangeStart>` before the first run and a `<w:commentRangeEnd>` +
`<w:commentReference>` run after the last (the pre-ported `CT_R` comment-range
helpers). `Document.comments` returns `self._part.comments`, which resolves
through `DocumentPart.comments → _comments_part.comments` — materializing the
`CommentsPart` on demand.

**Example** (the comment-range markers land in the body)

```matlab
doc = mat2doc.Document();
r = doc.add_paragraph().add_run("anchored text");
doc.add_comment(r, "Nice point", "Amy", "AJ");
bodyxml = doc.element.xml;
fprintf('has commentRangeStart: %d\n', contains(bodyxml, "w:commentRangeStart"));
fprintf('has commentRangeEnd  : %d\n', contains(bodyxml, "w:commentRangeEnd"));
fprintf('has commentReference : %d\n', contains(bodyxml, "w:commentReference"));
% has commentRangeStart: 1
% has commentRangeEnd  : 1
% has commentReference : 1
```

*Ported from python-docx v1.2.0: `src/docx/document.py::Document.add_comment`*

---

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
