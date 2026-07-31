---
title: "★ M2 milestone — the hello-world Word document"
---

# ★ M2 MILESTONE — ACHIEVED

**M2 is the second end-to-end Mat2Doc acceptance milestone: authoring a
"hello-world" Word document from nothing — a title, two headings and a body
paragraph — writes a `.docx` that is byte-identical to python-docx v1.2.0 AND
opens clean in real Microsoft Word.** At **P4-7b** both legs are GREEN —
**M2 is ACHIEVED.** This completes the `add_heading` / `add_paragraph` authoring
path and closes **Phase 4**.

Where **M1** proved the round-trip spine (open the bundled template, `.save()` it
back byte-identical), **M2** proves the *authoring* path: content created through
the public API — `add_heading`, `add_paragraph`, a style resolved **by name** —
serializes to the exact bytes python-docx produces.

## The M2 authoring sequence — a hello-world document

```matlab
d = mat2doc.Document();                     % new document from the bundled template
d.add_heading("Document Title", 0);         % style "Title"      -> <w:pStyle w:val="Title"/>
d.add_heading("First Section", 1);          % style "Heading 1"  -> <w:pStyle w:val="Heading1"/>
d.add_heading("A Subsection", 2);           % style "Heading 2"  -> <w:pStyle w:val="Heading2"/>
d.add_paragraph("Body paragraph text.");    % plain body paragraph <w:p>
d.save("hello.docx");
```

The Python mirror is line-for-line the same:

```python
d = docx.Document()
d.add_heading("Document Title", 0)
d.add_heading("First Section", 1)
d.add_heading("A Subsection", 2)
d.add_paragraph("Body paragraph text.")
d.save("hello.docx")
```

## The acceptance bar (both legs required)

| Leg | Bar | Result |
|---|---|---|
| **Byte leg** (Gate-3) | The authoring sequence's `d.save(tmp)` → `pkgcompare` vs the frozen 17-part reference `references\s0033`: **L0** inventory + **16 XML parts L1 byte-identical** + `docProps/thumbnail.jpeg` **bin byte-identical** = **17/17**, `word/document.xml` SHA-256 independently re-derived on **both** sides | **PASS — 17/17 L1 byte-identical** |
| **Word COM leg** (mso-office-verifier) | The M2 file opens in **real Word (16.0, build 16.0.20228)** with **no error, no repair/recovery prompt, zero dropped content**; all four paragraphs present with the **correct styles resolved by name** (Title / Heading 1 / Heading 2 / Normal); survives an **open-edit-save round-trip** (Word re-emits its own `.docx` and reopens it clean); a repair prompt = FAIL regardless of bytes | **PASS — open-clean + styles-by-name + edit-save round-trip, no repair** |

## The byte result — only `word/document.xml` changes

The single part that carries the new content is `word/document.xml`; everything
else — including the 349 KB `word/styles.xml` — is **byte-unchanged** from the M1
default:

| Fact | Value |
|---|---|
| M2 `word/document.xml` | **1865 B**, SHA-256 `a71e5502…f7c2c` — re-derived on BOTH the MATLAB and the python-docx side |
| Changed-part discipline vs M1 (`references\s0001`) | the **ONLY** part that differs is `word/document.xml`; `word/styles.xml` is **byte-UNCHANGED** |
| Why styles.xml never moves | `Title` / `Heading 1` / `Heading 2` already exist as **real template styles**; `add_heading` resolves them **by name** through the un-stubbed `Paragraph.style` chain — the *latent* styles table is never touched by M2 |
| Frozen permanent oracle | `validation\mat2doc\references\s0033\` (17-part `parts\` + `manifest.json` SHA-256s + `package.docx`, co-located `.gitattributes` `* binary` pin) |

:::{note}
**Latent styles and M2 are orthogonal.** M2 resolves built-in styles by *name*,
so it never writes the `<w:latentStyles>` table. The P4-7b latent-styles API
(`LatentStyles` / `LatentStyle_`, [documented on the styles page](api/styles_api.md#id-latentstyles))
has its own byte proof — the latent WRITE-path scenario `s0034`, **17/17 L1**,
where the ONLY changed part is `word/styles.xml` (SHA `3981d463…ab53`).
:::

## What went live for M2 — the content adders

M2 is the milestone at which the `Document.add_heading` / `add_paragraph`
authoring surface — a `mat2doc:notYetPorted` stub since P1-8/P2-3 — goes **live**.
The un-stub spans three layers:

(id-add-heading)=
### `Document.add_heading(text, level)` — the level → style-name mapping

`add_heading` (document.py 90-101) chooses a paragraph style from the heading
`level` and delegates to `add_paragraph`:

| `level` | style applied | `<w:pStyle w:val>` |
|---|---|---|
| `0` | `"Title"` | `Title` |
| `1` (or omitted — the H13 default) | `"Heading 1"` | `Heading1` |
| `2` … `9` | `"Heading 2"` … `"Heading 9"` | `Heading2` … `Heading9` |

```matlab
p = d.add_heading("A Subsection", 2);   % -> a paragraph styled "Heading 2"
p = d.add_heading("Chapter");           % level defaults to 1 -> "Heading 1"
```

**Bounds — `mat2doc:ValueError`.** `level` must be in `0-9`; anything outside
raises the **verbatim** python-docx message with a `mat2doc:ValueError`
identifier:

```matlab
d.add_heading("x", -1)   % mat2doc:ValueError "level must be in range 0-9, got -1"
d.add_heading("x", 10)   % mat2doc:ValueError "level must be in range 0-9, got 10"
```

**H1 (no index shift):** `level` is a **data value** — it is used verbatim both in
the `0-9` bound test and in the `"Heading %d"` style name — **not** a collection
index, so there is no `+1`/`-1` conversion. `add_heading` then calls
`add_paragraph(text, style)` and the style is resolved by name through the
un-stubbed `Paragraph.style` → `DocumentPart.get_style_id` → `Styles` chain.

### `Document.add_paragraph(text, style)` and `BlockItemContainer.add_paragraph`

`Document.add_paragraph` (document.py 109-119) delegates to the body block-item
container: `return self._body.add_paragraph(text, style)`. The container's
`add_paragraph` (blkcntnr.py 45-59) is the real worker:

1. create a fresh `<w:p>` via the LIVE `CT_Body.add_p()` (P2-3), wrapped as a
   `mat2doc.text.Paragraph` (P4-5b);
2. **`if text`** (non-empty-string, H4) → `paragraph.add_run(text)`;
3. **`if style is not None`** (identity, `~isequal(style,[])`) → `paragraph.style = style`;
4. return the paragraph.

```matlab
p = d.add_paragraph();                         % empty paragraph, 0 runs (H4)
p = d.add_paragraph("Body paragraph text.");   % one run, no style
p = d.add_paragraph("Note", "Heading 1");      % one run, styled "Heading 1"
```

The two guards are the **distinct** python-docx tests — `if text:` (truthy: an
empty string adds no run) versus `if style is not None:` (identity: a real style
object, `[]`/None removes nothing) — ported separately, exactly as in the
`Paragraph.insert_paragraph_before` chain.

### `BlockItemContainer.paragraphs` — now live

`BlockItemContainer.paragraphs` (blkcntnr.py) mints a fresh `mat2doc.text.Paragraph`
per `<w:p>` child (via the LIVE `CT_Body.p_lst`) as a homogeneous `1×N` array —
the plain-list surface. It is reachable via `d.body_().paragraphs`.
(`Document.paragraphs` itself stays a clean stub as a post-P4 content follow-up —
out of P4-7b's named scope — but its dependency `_body.paragraphs` is live.)

### Still stubbed after M2

`Document.add_table` / `tables` / `iter_inner_content` need `CT_Tbl` + `Table`
(**P6**); `add_section` / `sections` need **P5**; `add_picture` needs **P7**;
`add_comment` need **P8-2**; `settings` needs **P5-1**. Each remains a
`mat2doc:notYetPorted` stub naming its real-phase owner. `Document.add_page_break`
and `Document.paragraphs` are honest stubs whose dependencies are now all live —
clean un-stub candidates for the next content WP.

## The Word COM oracle — real Office accepts M2

The frozen M2 `references\s0033\package.docx` was opened in **real Microsoft Word
(COM `Word.Application`, version 16.0, build 16.0.20228)** with alerts suppressed:

- **Open-clean:** `Documents.Open(...)` returned silently — no error, no repair
  prompt, no "unreadable content" / recovery dialog.
- **Content + style integrity:** `Paragraphs.Count = 4`; each paragraph shows the
  correct text with the correct style resolved **by name** — "Document Title"
  (Title), "First Section" (Heading 1), "A Subsection" (Heading 2), "Body
  paragraph text." (Normal). No dropped, mangled or extra content.
- **Open-edit-save round-trip:** Word re-saved the file as its own `.docx`
  (`SaveAs2`, `wdFormatXMLDocument`) with no error, and the re-saved file reopened
  clean (`Paragraphs.Count = 5` after appending one paragraph).
- **Render sanity:** page 1 exported to PDF; the Title renders with its signature
  underline rule, Heading 1/2 as bold blue headings, the body as plain Normal.

Full record: `validation\mat2doc\com_verify_M2.md` (verdict **PASS**).

## Deviation posture — ZERO new D-numbers

Every M2 fact is **L1 byte-identical or value-exact**. No canonical-only (L2)
result appeared anywhere. The H17 `delete_()` friction (on `LatentStyle_`) is a
FLAG-3 method-**naming** resolution — byte-identical to python-docx
`_LatentStyle.delete()` — **not an output deviation → NO D-number**. M2 exercises
only already-signed adopt-only rulings (**D-001** own serializer, byte-proven via
M1 + M2 + latent-write 17/17; **D-zip-time** envelope-only). **Confirmed ZERO new
D-numbers across all of P4-7b.**

## ★ Phase 4 COMPLETE

With P4-7b the **text + styles tiers are fully authored end-to-end**: the
oxml element layer (P4-1a…P4-6), the text API proxies (P4-4a…P4-5b), the
styles/style API (P4-7a), and now the latent-styles API + the `add_heading` /
`add_paragraph` authoring path (P4-7b). Throughout P4, **M1 was preserved**
(every WP re-ran the 17/17 byte-neutrality sweep) and **zero new D-numbers** were
opened. What remains is **P5** (sections / settings / headers-footers) and
onward.
