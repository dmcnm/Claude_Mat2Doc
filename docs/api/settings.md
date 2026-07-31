---
title: "mat2doc.oxml.settings + mat2doc.settings — the document-settings tier (CT_Settings · Settings · odd_and_even_pages_header_footer)"
---

# `mat2doc.oxml.settings` + `mat2doc.settings` — the document-settings tier (`CT_Settings` + `Settings`)

Ported from python-docx v1.2.0 `src/docx/oxml/settings.py::CT_Settings` (in
package `+mat2doc/+oxml/+settings/`) and `src/docx/settings.py::Settings` (in
package `+mat2doc/+settings/`), plus the **two `register_element_cls` rows** they
require (`src/docx/oxml/__init__.py:134` → `w:settings`, `:83` →
`w:evenAndOddHeaders`) and the three delegation **un-stubs**
(`SettingsPart.settings` / `DocumentPart.settings` / `Document.settings`).

:::{note}
**★ Phase 5 begins — settings done; sections / headers-footers next.** This is the
**first work package of Phase 5** (sections + settings + headers/footers). It ports
the document-level *settings* surface — the `<w:settings>` element tier and the thin
`Settings` proxy the `Document.settings` property returns — and registers
`w:settings` so the shipped `word/settings.xml` now parses through `CT_Settings`,
**byte-neutrally** (M1 stays 17/17). The **sections** tier (`CT_SectPr` core + page
geometry → `Section`/`Sections`) and the **headers/footers** tier (`CT_HdrFtr` +
`_Header`/`_Footer`, the separate-part rels) follow at P5-2a → P5-2b → P5-3a →
P5-3b. `Settings.odd_and_even_pages_header_footer` — the one member ported here — is
the very toggle the P5-3b header/footer tier will consume.
:::

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

## The two-package layout

`settings.py` (the top-level `docx/` module) lands in the `+mat2doc\+settings`
package; `oxml/settings.py` (a `docx/oxml/` module) lands in `+mat2doc\+oxml\+settings`
— following the established module→subpackage convention. Neither collides with any
existing package.

| Python `src/docx/...` | MATLAB | symbols |
|---|---|---|
| `oxml/settings.py` | `+mat2doc\+oxml\+settings\` | `CT_Settings` |
| `settings.py` | `+mat2doc\+settings\` | `Settings` |

(id-settings-registry-parse-path)=
## The registry-adding, M1-byte-clean parse path

This WP registers **two** tags, exactly as python-docx registers them:

| tag | class | `oxml/__init__.py` | role |
|---|---|---|---|
| `w:settings` | `settings.CT_Settings` | `:134` | the **root** of `word/settings.xml` |
| `w:evenAndOddHeaders` | `shared.CT_OnOff` | `:83` | the tri-state child `CT_Settings` reads/writes |

`w:settings` is the root element of the settings part, so **every load of
`word/settings.xml` now transits `CT_Settings`**. Because `CT_Settings` extends
`BaseOxmlElement` and reserializes through the same `serialize_part_xml` walk with
**no serialization override** (it only *adds* the `evenAndOddHeaders` descriptor and
the `evenAndOddHeaders_val` accessor), registering the tag changes only the parsed
node **class**, never its **bytes**.

The **M1 17/17 byte-neutrality sweep holds**: `mat2doc.Document().save()` → unzip →
all 17 parts byte-identical to the frozen `references\s0001` reference, with
`word/settings.xml` (2535 B, SHA-256 `51a0d348…d73689`, independently re-derived
through the new `CT_Settings` parse path) L1 byte-identical. This is the standing
obligation every registry-adding `CT_*` WP inherits (P4-6 `styles.xml` precedent).

:::{note}
**The second row is a hard functional dependency, not scope creep.** `w:evenAndOddHeaders`
→ `CT_OnOff` is registered because `evenAndOddHeaders_val` reads/writes `.val` on the
`w:evenAndOddHeaders` child — without the row that child parses/creates as a generic
`XmlElement` with **no `val` property** and the whole `odd_and_even_pages_header_footer`
get/set fails. It mirrors a **real** python-docx registration (`__init__.py:83`) and is
**byte-neutral on the M1 parse path**: the default `settings.xml` carries no
`w:evenAndOddHeaders`, so nothing transits `CT_OnOff` on load — the row only lights up
when the WP **creates** the element (the write path below). The sibling `w:titlePg`
(`__init__.py:84`, same source block) is **deferred to the P5 section tier** (used only
by `CT_SectPr`, never by settings) — the same split-a-source-block-across-WPs precedent
as `w:outlineLvl` (P4-2 deferred → P4-6 closed).
:::

---

(id-ct_settings)=
## `CT_Settings`

**Syntax**

```matlab
s = mat2doc.oxml.OxmlElement("w:settings");   % a CT_Settings (registered)
s.evenAndOddHeaders_val                         % false  (child absent, H3)
s.evenAndOddHeaders_val = true;                 % -> <w:evenAndOddHeaders/>
s.evenAndOddHeaders_val = false;                % removes the child
```

**Description**

The `<w:settings>` element — the **root** of the settings part (`word/settings.xml`).
It carries a single ported `ZeroOrOne` child descriptor, `evenAndOddHeaders` (a
`CT_OnOff`), plus the derived boolean `@property` `evenAndOddHeaders_val` (get/set).

(id-ct_settings-successor)=
**The H11 successor-slice — the correctness crux (`[48:]` → `(49:end)`).** The
`_tag_seq` (`settings.py:19-118`, ported **verbatim** as the Constant `TAG_SEQ`) is a
**98-tag** tuple giving the OOXML schema order of every `<w:settings>` child. The
descriptor is declared `ZeroOrOne("w:evenAndOddHeaders", successors=_tag_seq[48:])`.
The own tag `w:evenAndOddHeaders` sits at **Python 0-based index 47** =
**MATLAB 1-based `TAG_SEQ(48)`**; the Python slice `_tag_seq[48:]` (everything *after*
the own tag) therefore maps to **`TAG_SEQ(49:end)`** — the H1 base shift applied once at
the slice. `TAG_SEQ(49)` is `w:bookFoldRevPrinting` (the tag immediately *after*
`w:evenAndOddHeaders`), so `insertChildInSequence` re-sorts a scrambled
`evenAndOddHeaders` add into canonical schema order (right after `w:defaultTableStyle`,
before the first schema-later sibling present). A wrong slice here inserts the child in
the wrong position → **Word repair / byte divergence**. Gate-2 and Gate-3 proved the
insertion byte-for-byte, adversarially:

| case | input neighbors | serialized child order | serhex == oracle |
|---|---|---|---|
| both_neighbors | `w:defaultTableStyle` (idx 48, pred) + `w:bookFoldRevPrinting` (idx 49, succ) | `defaultTableStyle, evenAndOddHeaders, bookFoldRevPrinting` | **==** |
| pred_only_append | `w:defaultTableStyle` (idx 48) only | `defaultTableStyle, evenAndOddHeaders` (appended after) | **==** |
| succ_only_insert | `w:bookFoldRevPrinting` (idx 49) only | `evenAndOddHeaders, bookFoldRevPrinting` (inserted before) | **==** |
| succ_prefixed (H8) | `m:mathPr` (idx 84, prefixed successor) | `evenAndOddHeaders, mathPr` (inserted before `m:`) | **==** |

The `m:mathPr` (index 84, 1-based) and `sl:schemaLibrary` (index 94) prefixes are both
present in `+oxml\nsmap.m`, so the successor tags carrying those prefixes resolve
correctly (**H8**).

(id-ct_settings-tristate)=
**The H3 tri-state `evenAndOddHeaders_val` (D-delta-1).** Backed by the
`CT_OnOff`-presence None/True/False, ported exactly:

- **get, child absent** → `false` (`settings.py:128-129` returns Python `False`, **NOT**
  the `CT_OnOff` default);
- **set `true`** → `get_or_add` + `.val = true`; the `CT_OnOff.val` setter sees
  `value == default(True)` → removes `@w:val` → emits the **empty
  `<w:evenAndOddHeaders/>`** (D-delta-1);
- **set `false`** → the `value is False` branch → `remove_evenAndOddHeaders_()` (child
  removed);
- **set `[]` (None)** → the `value is None` branch → child removed.

**H4 identity, not truthiness.** Python `value is None or value is False` is an
**identity** test — a double `0` must **not** match `is False`. The port uses the house
`islogical`-guard idiom (`Font.m:315-317`):
`isequal(value,[]) || (islogical(value) && isscalar(value) && ~value)`. A double `0`
therefore falls through to `get_or_add` and **creates** `<w:evenAndOddHeaders w:val="0"/>`
(matching python-docx), where a naive `isequal(value,false)` would have wrongly removed it.

**Example**

```matlab
s = mat2doc.oxml.OxmlElement("w:settings");   % a CT_Settings
disp(s.evenAndOddHeaders_val);                  % 0 (false — child absent, H3)
s.evenAndOddHeaders_val = true;
disp(string(s.xml));                            % ...<w:evenAndOddHeaders/>... (empty, D-delta-1)
s.evenAndOddHeaders_val = false;
disp(s.evenAndOddHeaders_val);                  % 0 (child removed)
```

*Ported from python-docx v1.2.0: `src/docx/oxml/settings.py::CT_Settings`*

---

(id-settings)=
## `Settings`

**Syntax**

```matlab
d  = mat2doc.Document();
st = d.settings;                                % a mat2doc.settings.Settings
st.odd_and_even_pages_header_footer             % false by default
st.odd_and_even_pages_header_footer = true;     % -> <w:evenAndOddHeaders/>
```

**Description**

`Settings` provides access to document-level settings for a document. It is a thin
`ElementProxy` over the `<w:settings>` root element (a `CT_Settings`) of the settings
part; reference semantics (handle) and **H5** element-identity `eq`/`ne` are inherited
from `ElementProxy` unchanged. It is reached via the `Document.settings` property.

The single ported member **`odd_and_even_pages_header_footer`** (read/write) delegates
straight to the element's `evenAndOddHeaders_val` accessor — no local state. It is
ported as a Dependent property with `get.`/`set.` (the house convention for a read/write
proxy `@property`, e.g. `Font.bold`), so Python
`st.odd_and_even_pages_header_footer = True` mirrors as MATLAB
`st.odd_and_even_pages_header_footer = true`. At this layer the value is a plain
bool (`true`/`false`), backed by the `CT_Settings` None/True/False tri-state (H3).

**The three un-stubs.** This WP makes the settings delegation resolve end-to-end:
`SettingsPart.settings` (`parts/settings.py:36-42` → `Settings(self._settings)`),
`DocumentPart.settings` (`parts/document.py:116-120` →
`settings_part_().settings`), and `Document.settings` (`document.py:211-214` →
`part_.settings`). All three wrap the **same** `<w:settings>` element, so the H5
eq-chain holds (`d.settings == d.part.settings`).

**Example**

```matlab
d  = mat2doc.Document();
st = d.settings;
disp(st.odd_and_even_pages_header_footer);      % 0 (false by default)
st.odd_and_even_pages_header_footer = true;     % writes <w:evenAndOddHeaders/>
disp(d.settings == d.part.settings);            % 1 (H5 — same wrapped element)
```

*Ported from python-docx v1.2.0: `src/docx/settings.py::Settings`*

---

(id-write-remove-paths)=
## The write / remove paths — byte-proven round-trips

The parse path is byte-neutral; the **novel** paths are the write and remove of
`<w:evenAndOddHeaders>`. Both are frozen as permanent references (Gate-3 `s0037` /
`s0038`, co-located `.gitattributes` `* binary` pins) and are **byte-identical to
python-docx**:

- **Write** (`d.settings.odd_and_even_pages_header_footer = True; save()`): the whole
  17-part package is byte-identical; `word/settings.xml` is 2557 B (SHA
  `66052d2f…ce750`), with `<w:evenAndOddHeaders/>` inserted between
  `w:defaultTabStop` (idx 39) and `w:characterSpacingControl` (idx 61, the first successor
  present in the default part) — exactly matching python-docx's insertion position (the
  H11 slice, proven on a real round-trip).
- **Remove** (`… = True; … = False; save()`): the inserted child is fully removed;
  `word/settings.xml` returns **byte-for-byte** to the M1 default (2535 B, SHA
  `51a0d348…d73689`) and the whole package equals the M1 default (`references\s0001`).

**Zero new D-numbers.** Every equivalence leg is L1 byte-identical or value-exact; the
write/remove novel paths — where a novel divergence would show — are byte-identical to
python-docx, so no deviation is opened. The write-path artifact
(`references\s0037\package.docx`) carries the new emitted element class
`<w:evenAndOddHeaders/>` and is flagged for the **mso-office-verifier** COM oracle at the
next milestone sweep.

---

## Settings tier COMPLETE — sections / headers-footers next

This WP completes the document-**settings** tier: the `<w:settings>` root element
(`CT_Settings`, its 98-tag `TAG_SEQ` and the H11 `evenAndOddHeaders` successor slice),
the thin `Settings` proxy, and the `odd_and_even_pages_header_footer` toggle. The
`w:settings` registration is byte-neutral (M1 17/17 preserved) and the write/remove
paths are byte-proven with **zero new D-numbers**. What remains in **Phase 5** is the
**sections** tier (**P5-2a** ports `CT_SectPr` core + page geometry, **P5-2b** the
`CT_HdrFtr` + section iterator) and the **section / header-footer API** (**P5-3a**
`Section`/`Sections`, **P5-3b** `_Header`/`_Footer` + the separate-part header/footer
rels — which consume the `odd_and_even_pages_header_footer` toggle ported here).
