---
title: "Mat2Doc manual"
---

# Mat2Doc manual

**Mat2Doc** is an exact functional MATLAB replica of
[python-docx](https://github.com/python-openxml/python-docx) v1.2.0 — a
base-MATLAB (R2024b, no toolboxes) library for creating and manipulating Word
`.docx` files. Every symbol is a faithful, byte-equivalence-validated port of
its Python original: the OPC package layer, the order-preserving OOXML element
tree, the WordprocessingML document/paragraph/run/table object model, and the
shared value-type and simple-type machinery beneath them.

This manual is the API reference. Pages are organised by architectural layer,
from the foundation value types up through the high-level document, paragraph,
run, and table APIs. Each page documents one module's ported symbols — syntax,
behaviour, examples, and the `Ported from python-docx v1.2.0` provenance line
that ties every MATLAB symbol back to its source.

:::{note}
API pages are generated from the MATLAB help headers. Until the generator
lands they are maintained by hand to the same shape (one section per symbol:
syntax, description, example, ported-from).
:::

## ★ Milestones

- **M1 — the round-trip spine (ACHIEVED).** `mat2doc.Document().save()` writes the
  bundled template back **byte-identical** to python-docx and opens clean in real
  Word. See the [document tier page](api/document.md).
- **★ M2 — the hello-world Word document (ACHIEVED, P4-7b).** Authoring a title,
  two headings and a body paragraph — `Document(); add_heading(_,0/1/2);
  add_paragraph(); save()` — produces a `word/document.xml` **17/17 byte-identical**
  to python-docx (only `document.xml` differs; `styles.xml` unchanged) that **opens
  clean in real Word** with styles resolved by name. This completes the
  `add_heading` / `add_paragraph` authoring path and **Phase 4**. See the
  **[M2 milestone page](m2_milestone.md)**.

## Sections

- **Foundation & utilities** — the shared length-unit value types (the
  `Length` family and its EMU-based conversions), the `RGBColor` value object,
  and `pyStr`, the mandated Python-`str()` numeric→XML formatting helper (H14).
- **Proxy, exception & protocol tier** — the base tier that **every later API
  proxy extends** (P3–P6): `ElementProxy` (H5 element-identity equality + the
  `part` None-guard), the standalone `Parented` / `StoryChild` part-provider
  bases, `TextAccumulator`, the `lazyproperty` / `write_only_property` idiom
  guides, the `mat2doc.exc` exception raisers (flat identifiers), and the
  `mat2doc.types` structural protocols — plus the byte-neutral, Python-faithful
  `Document → ElementProxy` retrofit.
- **XML layer (oxml)** — the byte-fidelity OOXML foundation: the namespace
  machinery, the order-preserving parser and `XmlElement` tree, the part-XML
  serializer, and the xmlchemy element-class / attribute-descriptor / XPath
  engine — plus the **simple-type validator tier** (`+oxml\+simpletypes`: the 4
  Base / 10 Xsd / 21 `ST_*` classes that validate and convert every typed `w:`
  attribute value the P4–P6 element classes read and write; the D-STYPE-2 float
  re-home, the `ST_HexColor`/`ST_HexColorAuto` split, `ST_OnOff`, and the
  `ST_UniversalMeasure` unit grammar live here); and the **text/font element
  tier** (`+oxml\+text` + `+oxml\+shared`: the **first real `w:`-content element
  classes** — `CT_RPr`, the M2 byte-critical `<w:rPr>` run-properties container
  whose 27 `ZeroOrOne` descriptors ride the non-contiguous H11 successor slices
  of the 39-entry `_tag_seq`, its 6 child `CT_*` classes, the `_new_color`
  black-seed override, and the shared `CT_OnOff` (D-delta-1 tri-state) /
  `CT_String` leaves — registered byte-neutrally with M1 preserved); and the
  **run tier** above it (`run.py` → `CT_R`, the `<w:r>` run element the
  `document.xml` body is built from, with its H11 rPr-first ordering, the
  byte-critical `add_t` `xml:space="preserve"` decision replicating CPython's
  exact `str.strip()` set — the **H16** whitespace hazard — the
  `_RunContentAppender` text→content FSM, and the `CT_Text`/`CT_Br`/`CT_Cr`/
  `CT_NoBreakHyphen`/`CT_PTab` inner-content leaves); the **paragraph tier**
  (`parfmt.py` → `CT_PPr`, the `<w:pPr>` paragraph-properties container whose 12
  descriptors ride the 36-entry `_tag_seq` with `pStyle` sorting first — the
  `add_heading` crux — plus `CT_Ind`/`CT_Jc`/`CT_Spacing`/`CT_TabStop`/
  `CT_TabStops`, and `paragraph.py` → `CT_P`, the `<w:p>` paragraph element); and
  finally the **hyperlink + rendered-page-break tier** (`hyperlink.py` →
  `CT_Hyperlink`, and `pagebreak.py` → `CT_LastRenderedPageBreak` with its
  precedes/follows detection and paragraph fragment-split algorithm) — a
  byte-neutral pure lookup addition that **completes the text-oxml element layer**;
  and finally the **styles-oxml element tier** (`styles.py` → `CT_Styles` /
  `CT_Style` [10 `ZeroOrOne` over the 22-entry `_tag_seq`, the H3 tri-state `.._val`
  members, the `basedOn`/`next` sibling chains] / `CT_LatentStyles` /
  `CT_LsdException` + the `styleId_from_name` H15 mangler, plus the shared
  `CT_DecimalNumber` leaf) — the **first WP of the styles chain**, which registers
  the 12 styles-block tags **byte-neutrally** (the 349 KB `styles.xml` now transits
  `CT_Styles`, M1 17/17 preserved), raises the **H17** `delete()`/destructor-collision
  ruling (guarded-destructor override, SIGNED-PROVISIONAL), and carries the F-1
  `bool(None)→"0"` truthiness fix — **styles-oxml COMPLETE → the styles API
  (P4-7a/b) is next → M2**.
- **OPC layer (opc)** — the two-tier serializer that regenerates
  `[Content_Types].xml` and the `.rels` parts (`CT_Types` / `CT_Default` /
  `CT_Override` / `CT_Relationships` / `CT_Relationship`), plus the OPC
  constants, the `default_content_types` spec pair-list, and the shared helpers;
  and the **packaging tier** above it — the `PackURI` partname algebra, the
  `phys_pkg` physical zip reader/writer factories (with the reproducible
  `1980-01-01` zip writer), and the `Relationships` collection that regenerates
  each `.rels` part in insertion order; the **package assembly tier** above it
  — `PackageReader` (the serialized read path) and `PackageWriter` (whose
  zip-entry traversal is the byte-critical M1 order), with the serialized
  part/relationship value objects and the reader/writer content-type maps; and
  the **package/part object model** on top — `OpcPackage` (open / save /
  graph traversal), the `Part` / `XmlPart` base parts, and the `PartFactory`
  content-type→class registry whose XmlPart-vs-Part split decides M1
  whitespace-collapse (the full 17-part open→save round-trip is byte-proven here);
  and the **core-properties tier** on top — `CT_CoreProperties` (the 15
  Dublin-Core descriptors, the W3CDTF date grammar, and the byte-critical `xsi`
  namespace hoist), the `CoreProperties` API wrapper, and the
  `CorePropertiesPart` part class (`/docProps/core.xml`; the 721 B re-serialize
  and 681 B fresh build are byte-proven).
- **WordprocessingML document tier (docx) — M1 → P2 complete** — the public
  `mat2doc.Document()` entry function and the object graph beneath it
  (`package.Package` / `parts.DocumentPart` / `document.Document`) that satisfies
  the **M1 milestone** (`mat2doc.Document().save()` produces `default.docx`
  byte-identical to python-docx — 17/17 L1 three-way — and opens clean in real
  Word), **completed by P2-3** with the `CT_Document`/`CT_Body` element classes
  (registered for `w:document`/`w:body` — byte-neutral, the last P2 step to touch
  the `document.xml` parse path), the `_Body` proxy and the `BlockItemContainer`
  (`blkcntnr`) base, and the `Document._body`/`_block_width` accessors. Documents
  the byte-neutral registration, the H11 successor ordering, the
  object-graph-complete/features-stubbed-at-P4+ state, and the **Phase 2 complete**
  milestone.
- **Story-part tier (parts)** — the part classes an opened `.docx` graph is
  built from: `StoryPart` (the real base inserted above `DocumentPart`, with
  `next_id` and the `_document_part` lazyproperty), the reparented
  `DocumentPart` (whose `_styles_part` / `_settings_part` now return real
  sibling parts), and the thin `StylesPart` / `SettingsPart` / `NumberingPart`
  `XmlPart` shells. Documents the **object-graph-vs-feature-surface**
  distinction P2-2 draws (part plumbing live, feature accessors stubbed at their
  real-phase owner), the three byte-neutral PartFactory flips, and **task #60**
  (the `xpath` hoist onto `XmlElement`, `Relationships.delitem`, and the live
  `Part.drop_rel` with its `< 2` refcount threshold).
- **Text API tier (P4 API tier)** — the user-facing API **proxies** that **open
  the P4 API tier** now that the text-oxml element layer is complete. The
  **text/font** page: `mat2doc.text.Font` (character properties over a `w:r` —
  the 20 boolean toggles mapped to their `w:rPr` child tags, plus size
  half-points, name ascii/hAnsi, the `WD_UNDERLINE` tri-state, highlight,
  sub/superscript, and color) and `mat2doc.dml.ColorFormat` (rgb / theme_color /
  the read-only `type` with its THEME > AUTO > RGB detection precedence and the
  `InvalidXmlError`-on-`themeColor`-without-`@w:val` required-attribute edge).
  The **text/run** page: `mat2doc.text.Run` (the `<w:r>` proxy — `text` get/set,
  the font-delegated `bold`/`italic`/`underline`, `font`, `add_break` with its
  `WD_BREAK` → `(type_, clear)` map (the `LINE_CLEAR_*` `@w:type`-removal),
  `add_text`/`add_tab`/`clear`, `contains_page_break`, `mark_comment_range`, and
  the live-vs-stub split — `add_picture`→P7, `iter_inner_content`→P4-5b,
  `style`→P4-7) and the inert `mat2doc.text.Text_` (`<w:t>` wrapper, no public
  members — the faithful v1.2.0 no-`text`-property port). The **text/parfmt**
  page: `mat2doc.text.ParagraphFormat` (paragraph formatting over a pPr owner —
  alignment, the four bool tri-states [keep_together/keep_with_next/
  page_break_before/widow_control], signed first-line/left/right indents, space
  before/after, the polymorphic `line_spacing` with its float-multiple-vs-`Length`
  detection and the `line_spacing_rule` SINGLE/ONE_POINT_FIVE/DOUBLE special-member
  map, and the lazy read-only `tab_stops`) and the **first proxy *sequence***
  `mat2doc.text.TabStops` / `mat2doc.text.TabStop` (the interim explicit sequence
  surface — `getitem_`/`delitem_`/`len_`/`to_array` with the H1 0-based argument,
  `add_tab_stop` position-ordering, and `TabStop` position/alignment/leader incl.
  the SPACES-default absence and the faithful position-setter identity quirk;
  `RedefinesParen` native `()` indexing deferred to a future collection WP).
  Documents the **API-tier framing** — proxies that delegate to the
  P4-1a/P4-1b/P4-2 `CT_RPr` / `CT_Color` / `CT_R` / `CT_PPr` / `CT_TabStops`
  helpers, adding no registry row and no serialization code, so equivalence is
  **behavioral** and each WP is **byte-neutral** (M1 17/17 unchanged, 0 new
  D-numbers) — and the P4-5a G-scenario document byte pin (4 formatted paragraphs
  → `document.xml` byte-identical), the closest guard yet to M2's `add_paragraph`
  formatting path. The **text/paragraph** page — the tier's **last API-tier WP** —
  completes the user-facing text object model: `mat2doc.text.Paragraph` (the
  `<w:p>` proxy — `add_run`, `alignment`, `clear`, `contains_page_break`,
  `hyperlinks`, `insert_paragraph_before` with its distinct `if text:` /
  `if style is not None` guards, the heterogeneous `iter_inner_content` cell,
  `paragraph_format`, `rendered_page_breaks`, `runs`, the P4-7 `style` stub, and
  `text` get/set), `mat2doc.text.Hyperlink` (the read proxy — `address`
  internal-jump-vs-external-`target_ref`, `fragment`/`url`, `runs`, `text`,
  `contains_page_break`) and `mat2doc.text.RenderedPageBreak` (the
  `precedes_all_content` guard + the preceding/following loose-`Paragraph`
  fragment split incl. the atomic break-inside-hyperlink case). Documents the
  standing surface shape (homogeneous `1×N` object arrays vs the
  `iter_inner_content` `1×N` cell, both ACCEPTED at Gate 2), the H5
  instance-identity regime of the `StoryChild`/`Parented` bases (distinct from the
  `ElementProxy` element-identity of the earlier proxies), and the s7 ns-decl
  reopen-check (still unreachable, re-booked at P7). **API tier COMPLETE — the
  styles chain (P4-6/P4-7) is next → M2.**
- **Styles API tier (styles)** — the user-facing style **proxies** that read the
  P4-6 `<w:styles>` element surface and **un-stub** the styles delegation across
  the document object graph (`StylesPart`/`DocumentPart`/`Document`/`Run`/
  `Paragraph`). The **second WP of the styles chain** (P4-6 → P4-7a → P4-7b → M2):
  `mat2doc.styles.StyleFactory` (the H10 `WD_STYLE_TYPE` → leaf-class dispatch,
  with the `KeyError`-on-None edge), the
  `BaseStyle`→`CharacterStyle`→`ParagraphStyle`→`TableStyle_`/`NumberingStyle_`
  hierarchy (`BaseStyle` the `matlab.mixin.Heterogeneous` root; the full H3
  tri-state property surface; the `base_style`/`next_paragraph_style` sibling
  chains and the `font`/`paragraph_format` fresh-proxy accessors), the `Styles`
  collection (the interim explicit `contains_`/`getitem_`/`to_array`/`len_`
  sequence surface plus `add_style`/`default`/`get_by_id`/`get_style_id`/`element`,
  and the still-stubbed `latent_styles` → P4-7b), and `BabelFish` (the
  case-sensitive UI↔internal style-name alias table). Documents the **H17
  proxy-layer resolution** — `delete()`→`delete_()` (why `delete` cannot be
  overridden at the proxy layer; `delete_()` is byte-faithful, a FLAG-3-class
  method-naming resolution, **no D-number**), the project's **first ported
  `warnings.warn`** (`mat2doc:UserWarning` house convention on the deprecated by-id
  `getitem_` path), and the un-stub that makes `Paragraph.style`/`Run.style`/
  `get_style` resolve end-to-end (the style-by-name path byte-proven, the direct M2
  `add_heading` precursor). At **P4-7b** the tier is **completed** by the
  **latent-styles API** — `mat2doc.styles.LatentStyles` (the `<w:latentStyles>`
  defaults + the `getitem_`/`to_array`/`len_`/`add_latent_style` collection surface
  and the `default_priority`/`default_to_*`/`load_count` writers) and
  `mat2doc.styles.LatentStyle_` (`_LatentStyle`→`LatentStyle_`, a `<w:lsdException>`
  override with the RAW tri-state `hidden`/`locked`/`quick_style`/`unhide_when_used`
  and the H17 `delete_()` proxy-layer rename) — un-stubbing `Styles.latent_styles`;
  the latent WRITE path is byte-proven (`s0034`, only `styles.xml` changes). **API/
  proxy tier — no registry row, no serialization change → behavioral equivalence,
  byte-neutral (M1 17/17, 0 new D-numbers). With latent styles live, P4-7b un-stubs
  the `add_heading`/`add_paragraph` authoring path → ★ M2 ACHIEVED → Phase 4
  COMPLETE.**
- **Sections & settings tier (Phase 5)** — **Phase 5 begins** (sections + settings +
  headers/footers). The **settings** page opens it: `mat2doc.oxml.settings.CT_Settings`
  (the `<w:settings>` root — its **98-tag** `TAG_SEQ` and the byte-critical H11
  `evenAndOddHeaders` successor slice `_tag_seq[48:]` → `TAG_SEQ(49:end)`, the H3 tri-state
  `evenAndOddHeaders_val` with the D-delta-1 empty-element emission and the H4
  identity-not-truthiness guard) and `mat2doc.settings.Settings` (the thin `ElementProxy`
  the `Document.settings` property returns, with the one read/write member
  `odd_and_even_pages_header_footer`). Registers `w:settings` (→ `CT_Settings`, the
  settings-part root) + `w:evenAndOddHeaders` (→ `CT_OnOff`, the tri-state child) **byte-
  neutrally** — `word/settings.xml` now transits `CT_Settings`, M1 17/17 preserved — and
  un-stubs the settings delegation across `SettingsPart`/`DocumentPart`/`Document`. The
  novel write path (insert `<w:evenAndOddHeaders/>`) and remove path are **byte-identical
  to python-docx** (`s0037`/`s0038`), **0 new D-numbers**. The **section oxml core** page
  (**P5-2a**) opens the sections tier: `mat2doc.oxml.section.CT_SectPr` (the `<w:sectPr>`
  section-properties root — its 20-tag `TAG_SEQ` and the H11 successor slices, the page
  geometry `page_width`/`page_height`/margins accessors over `w:pgSz`/`w:pgMar`, the
  orientation setter with the **NO-w/h-swap** semantics, the `start_type` identity setter,
  the `titlePg_val` `[None, False]` `==`-membership breadth, and the header/footer reference
  surface) and the four child classes `CT_PageSz` / `CT_PageMar` (the signed-vs-unsigned
  twips split) / `CT_SectType` / `CT_HdrFtrRef`. Registers **7 tags** (`w:sectPr`/`pgMar`/
  `pgSz`/`type`→CTs, `w:headerReference`/`w:footerReference`→`CT_HdrFtrRef`,
  `w:titlePg`→`CT_OnOff` [the P5-1 deferral closed]) **byte-neutrally** on the **M1-central**
  `word/document.xml` (17/17 preserved, SHA `0e4dd503…` unchanged); the geometry-write /
  LANDSCAPE novel paths are byte-identical to python-docx (`s0040`/`s0041`), the loose-element
  `r:` auto-prefix is the pre-existing SIGNED **D-nsprefix-rewrite** (dead-on-generation),
  **0 new D-numbers**. The **section-oxml layer is then COMPLETED** at **P5-2b**:
  `mat2doc.oxml.section.CT_HdrFtr` (the `<w:hdr>`/`<w:ftr>` header/footer **PART root** —
  its `ZeroOrMore` p/tbl descriptors and the `inner_content_elements` tag-based child union,
  with an unregistered `w:tbl` INCLUDED as a generic `XmlElement` until P6, and `w:ins`-nested
  content excluded) and `mat2doc.oxml.section.SectBlockElementIterator_` (the
  `_SectBlockElementIterator` that partitions a body's block elements into sections by `sectPr`
  — the mutually-exclusive p-sect/body-sect xpath shapes, the skip-count boundary arithmetic,
  and the concat-equals-body invariant), which **un-stubs `CT_SectPr.iter_inner_content`** (now
  live). Registering `w:hdr`/`w:ftr` is **M1-neutral** (separate-part roots absent from
  `default.docx`) — M1 17/17 unchanged, **zero re-pins**, the header-part round-trip
  byte-identical (`s0043`) and the adversarial 4-doc / 11-section partition corpus equal to the
  python-docx oracle (`s0044`), **0 new D-numbers**. The **Section / Sections API** page
  (**P5-3a**) then opens the section API surface: `mat2doc.section.Section` (the transparent
  proxy over a registered `CT_SectPr` — all twelve geometry/type accessors delegated one-to-one,
  the `orientation` NO-w/h-swap, the `start_type` identity setter, the `different_first_page`
  `titlePg` bool, `part` as the story-parent hook, and `iter_inner_content` wrapping each `CT_P`
  as a `Paragraph` with the `w:tbl`→`Table` branch raising `mat2doc:notYetPorted` owner P6-4a —
  never a silent drop) and `mat2doc.section.Sections` (the section sequence — the **0-based**
  `getitem_` int/negative/slice with the verbatim `IndexError`/`ValueError`, `to_array`, `len_`,
  the interim struct-slice currency per the TabStops precedent). It **un-stubs the section-
  authoring path** — `Document.sections` / `Document.add_section` + `CT_Body.add_section_break`
  (the intermediate-sectPr nesting, clone-before-hdr/ftr-removal) — with **no registry rows**, so
  the default `Document().save()` stays M1 17/17 (`document.xml` `0e4dd503…` unchanged) while the
  `add_section` (all 5 start types + 2/3-section chains) and Section-property-write paths are
  byte-identical to python-docx (`s0046`–`s0051`), **0 new D-numbers**. Finally, the **headers/footers**
  page (**P5-3b**, the FINAL P5 WP) lands the separate-part hdr/ftr tier: `mat2doc.section.Header_` /
  `Footer_` / `BaseHeaderFooter_` (the `_Header`/`_Footer` proxies over a `BlockItemContainer` — the six
  `Section` header/footer accessors [three `WD_HEADER_FOOTER` index kinds], `is_linked_to_previous`
  get/set as the add/drop authoring trigger, `paragraphs`/`add_paragraph`, the `part` story-parent
  override, and the 3-case `_get_or_add_definition` **inherit-walk**) and `mat2doc.parts.HeaderPart` /
  `FooterPart` (the `StoryPart` part classes minted at runtime — the **FIRST runtime-added parts** in
  Mat2Doc, each `word/headerN.xml` carrying its own `[Content_Types]` Override + rels). The **C3
  `BlockItemContainer` element-accessor seam** (property→protected method, since MATLAB cannot redefine
  a stored property in a subclass — byte-neutral for M1 17/17 + M2 s0033) supports the lazy
  part-creation on first content access. Full-package byte-identical to python-docx across header /
  footer / both / even+first+titlePg / add→drop / reload (`s0052`–`s0058`), **0 new D-numbers**, and
  **COM-verified in real Word** (all four packages open clean; `DifferentFirstPage` rendered).
  **★ PHASE 5 COMPLETE — sections, settings and headers/footers all done + COM-verified; the cross-part
  hazard (risk-register #4) cleared.** See the [headers & footers page](api/headers_footers.md).
- **Tables tier (Phase 6)** — **Phase 6 begins** (the table object model). The
  **table oxml LEAF** page opens it: the seven attribute-only / single-list leaf
  classes of `oxml/table.py` — `mat2doc.oxml.table.CT_TblWidth` (the width union:
  `dxa`→`Length`, `pct`/`auto`/`nil`→None, the setter that forces `type="dxa"`),
  `CT_Height` (`val` + `hRule` `WD_ROW_HEIGHT_RULE`), the column grid `CT_TblGrid`
  / `CT_TblGridCol` (the D-delta-4 public `add_gridCol`, the 0-based `gridCol_idx`),
  `CT_TblLayoutType` (fixed vs autofit), `CT_VerticalJc`
  (`WD_CELL_VERTICAL_ALIGNMENT`, required `@w:val`), and `CT_VMerge`
  (restart/continue, the sole `"continue"` non-None default). Registers the
  **7 leaf tags** (`w:gridCol`/`w:tblGrid`/`w:tblLayout`/`w:tcW`/`w:trHeight`/
  `w:vAlign`/`w:vMerge`) **M1-NEUTRALLY** — `default.docx` has no table, so nothing
  transits the new classes (M1 17/17 preserved, **zero re-pins**) — with the C4
  brief-correction that `w:tblW` is **not** registered upstream (only `w:tcW`). The
  table-leaf subtrees of a real python-docx table round-trip **byte-identical**
  (10/10, `references\s0060`) and the width union is byte-proven both ways,
  **0 new D-numbers**. **P6-2** then adds the **property containers** that consume
  those leaves — `mat2doc.oxml.table.CT_TblPr` (table properties: `alignment` /
  `autofit` / `style`), `CT_TblPrEx` (the bare property-exceptions container),
  `CT_TrPr` (row properties: grid-before/after skips + row height) and `CT_Row`
  (the `<w:tr>` row element: the delegated row accessors, `tr_idx`, the custom
  `tblPrEx`/`trPr` inserters, and the tag-based **CT_Tc boundary** handlers
  `_new_tc` / `tc_at_grid_offset`→P6-3a). The **first NON-neutral** table WP:
  registering `w:tblPr` puts the 100 `<w:tblPr>` nodes in the shipped table styles
  on the live `word/styles.xml` parse path, proven byte-neutral by the M1 17/17
  `styles.xml` `02d71a68…` match. **★ A2 cross-enum:** `CT_TblPr.alignment` reuses
  the single registered `CT_Jc` (one element, two context enums — **no second
  `w:jc` row**) and faithfully returns a `WD_PARAGRAPH_ALIGNMENT` member
  (python-docx's `cast` is a runtime no-op; a converting getter would crash on the
  legal `<w:jc w:val="both">` = Justify) — the binding idiom is compare **by name**
  (`== "CENTER"`) or `.value`, never cross-class `==` (design.md §2, ruled no-D).
  P6-2 also closes the `w:tblStyle`→`CT_String` and `w:gridAfter`/`w:gridBefore`→
  `CT_DecimalNumber` deferrals; props round-trip **6/6 byte-identical**
  (`references\s0062`, incl. the `w:jc val="both"` edge), **0 new D-numbers**.
  **P6-3a — the single hardest WP of the project (with its own dedicated pre-launch
  plan-audit)** — then lands the **cell MERGE engine** `mat2doc.oxml.table.CT_Tc`
  (the `<w:tc>` cell and the destructive `merge` machinery: horizontal `gridSpan`,
  vertical `restart` + the **bare `<w:vMerge/>`** continuation, block spans, the
  `InvalidSpanError` rectangular-span validation, **H4 falsy-zero** width summing,
  and snapshot-safe content consolidation) plus its properties container `CT_TcPr`
  (`grid_span`/`vMerge_val`/`width`/`vAlign_val` over the H11-ordered `_tag_seq`),
  un-stubbing `CT_Row._new_tc`→`CT_Tc.new()` and the `grid_span` step of
  `tc_at_grid_offset`. The **second NON-neutral** table WP: registering `w:tcPr`
  puts the **595 + 595 `<w:tcPr>` nodes** in `word/styles.xml` +
  `stylesWithEffects.xml` (inside the shipped `<w:tblStylePr>` overrides) on the
  live parse path, proven byte-neutral by the M1 17/17 `styles.xml` `02d71a68…`
  match. The **merge byte-matrix is 10/10 byte-identical** to python-docx v1.2.0
  (`references\s0063` — horizontal / vertical-bare-vMerge / block / into-existing-
  spans / width-sum + H4-skip / sdt-passenger / **nested-table** — frozen as the
  permanent table-merge oracle), handle-identity `_tr_idx`=`[0 1 2]`, all 9
  invalid-span raises verbatim, **0 new D-numbers**. **P6-3b** then landed the
  **table root** `mat2doc.oxml.table.CT_Tbl` (the `<w:tbl>` root: the `new_tbl(rows,
  cols, width)` **table constructor** — byte-identical across sizes incl. the
  non-even-width H6 EMU-floor col-rounding, frozen 7/7 as the permanent
  table-authoring oracle `references\s0065`; the `OneAndOnlyOne` `tblPr`/`tblGrid`;
  `tr_lst`/`add_tr`; `bidiVisual_val` RTL→bare `<w:bidiVisual/>`; `col_count`;
  row-major `iter_tcs`; `tblStyle_val`), completed the registry (`w:tbl`→`CT_Tbl`,
  `w:bidiVisual`→`CT_OnOff`) and ran the **tbl un-defer sweep** — every prior
  tag-based `w:tbl`→generic site (`CT_Body`/`CT_HdrFtr`/`CT_SectPr`/`CT_Tc`)
  auto-upgrades to `CT_Tbl` and the P6-3a `CT_Tc` `trLstOfTbl_` shim resolves to the
  real `self._tbl.tr_lst` (V2), the merge matrix re-proven **10/10 byte-identical**
  after the swap, **0 new D-numbers**. **★ The table OXML LAYER is now COMPLETE —
  all 14 `table.py` element classes ported, the registry complete.** What remains
  in Phase 6 is the API tier: the `Table`/`_Rows`/`_Columns`/`_Cell` API →
  **P6-4a/b** + the table Word-COM sweep (the `s0063` merged-cell + `s0065` new_tbl
  fixtures COM-verified at P6-4b). See the [table oxml page](api/oxml_table.md).
  **P6-4a** then opens the **table API/proxy tier**: `mat2doc.table.Table` (the
  `<w:tbl>` proxy — `alignment` [**★ A2 cross-enum** — returns a
  `WD_PARAGRAPH_ALIGNMENT` member, compare by name `== "CENTER"` or `.value`,
  never cross-class `==`], `autofit`, `style`, `table_direction`, the
  `@lazyproperty`-cached `rows`/`columns`, `_column_count`; `add_row`/`add_column`
  → P6-4b), the row/column collections `_Rows`→`Rows_` / `_Columns`→`Columns_`
  (the **0-based** `getitem_` int/negative/slice, `to_array`, `len_`) and the
  single-row/column proxies `_Row`→`Row_` / `_Column`→`Column_`
  (width/height/height_rule/grid_cols_*/`_index`; `.cells` → P6-4b), all FLAG-3
  trailing-underscore renames. It **un-stubs the table-authoring path** —
  `Document.add_table(rows, cols, style=None)` (style 3rd, width = `_block_width`)
  vs `BlockItemContainer.add_table(rows, cols, width)` (width 3rd, the signature
  difference) + the three `iter_inner_content` sites + **`Section.iter_inner_content`
  `w:tbl`→`Table`** (the P5-3a **C2 debt discharged**) — with **no registry rows**,
  so `Document().save()` stays M1 17/17. **★ The FIRST end-to-end table is now
  byte-proven** — `document.add_table(2, 3)` → `word/document.xml` byte-identical
  to python-docx (`a1eda043…`, `styles.xml` unchanged; with a style →
  `<w:tblStyle>` + `styles.xml` gains the style), **0 new D-numbers**. **P6-4b**
  then closes the tier: `mat2doc.table.Cell_` (the `_Cell` FLAG-3 rename — `text`
  get/set, `add_paragraph`, nested `add_table`, `grid_span`, `paragraphs`,
  `tables`, `vertical_alignment`, `width`, and **`merge`**), `Table.add_row` /
  `add_column`, and `_Row.cells` / `_Column.cells` (the cells grid walk), un-stubbing
  the 8 P6-4a stubs plus the required `BlockItemContainer.tables`. **★ Cell merge:**
  `cell1.merge(cell2)` drives `gridSpan`/`vMerge` into the package — byte-identical
  to python-docx across every geometry (horizontal `9f626e2b…`, vertical
  `e833fac8…`, 2×2 block `ac155531…`, merge-then-text `c96a3351…`, the 3×3 mixed
  merge `63b25b1a…`) and **COM-verified in real Word** (all five table packages open
  clean, merges honored). A merged cell is the **SAME `Cell_` handle** at each
  spanned grid position (H5). **0 new D-numbers.** See the
  [table API page](api/table_api.md).
- **★ PHASE 6 COMPLETE (tables) — 2026-08-01.** The full table tier is byte-proven
  end-to-end and COM-verified: the table **oxml layer** (14 classes — leaves,
  property containers, the `CT_Tc` merge engine, `CT_Tbl`), the **API tier**
  (`Table`/`_Rows`/`_Columns`/`_Row`/`_Column`/`_Cell`), and the full **authoring
  surface** (`add_table`/`add_row`/`add_column`/`merge`). The **`CT_Tc` merge
  engine** — risk-register #2, the single hardest WP of the port — is byte-identical
  across the entire merge matrix and its outputs open clean in real Microsoft Word
  (Word 16.0.20228: plain, styled, horizontal-merge, block-merge and 3×3
  mixed-merge packages all open silently with merges honored and text intact —
  `com_verify_P6_tables.md`). **Zero new D-numbers across all of Phase 6.** Next:
  **Phase 7** (images / drawing / inline pictures).
- **Images tier (Phase 7)** — **Phase 7 begins** (images / drawing / inline
  pictures). The **image core** page opens it: the format-agnostic
  `mat2doc.image.Image` value object (`from_blob`/`from_file`, the `sha1` /
  `ext` `@lazyproperty` pair, `content_type` + `px_width`/`px_height` +
  `horz_dpi`/`vert_dpi` pure header-delegation, native `width`/`height` as a
  `Length`, and `scaled_dimensions`), the `ImageHeaderFactory_` `SIGNATURES`
  dispatch (8 rows, first-match-wins), `BaseImageHeader`, the `StreamReader` /
  `BytesIO` byte-read primitives, the five constants tables
  (`MIME_TYPE`/`JPEG_MARKER_CODE`/`PNG_CHUNK_TYPE`/`TIFF_FLD`/`TIFF_TAG`), and
  `mat2doc.opc.sha1_hexdigest`. **★ A re-port from Mat2Ppt's `+image`** — docx
  `image/` is the HOME of these parsers (python-pptx has none; it uses PIL, and
  Mat2Ppt's `+image` was itself the docx port), re-homed with three transforms
  (namespace → `mat2doc.shared`, inline None-idiom, and **WMF-exclusion**: the
  pptx-only WMF/EMF/`pil_dpi`/`int_dpi` seams are NOT ported) and the dpi kept as
  **docx math, not PIL** (that reversion bites the format parsers at P7-1b/P7-2).
  The Gate-2 **F-1** fix re-homes `Image.ext` onto a CPython-`os.path.splitext`
  helper (`splitext_ext`, not `fileparts`). **Pure-parsing WP — no registry row,
  nothing on the open/save path → M1 17/17, 0 new D-numbers.** Next:
  P7-1b (png/gif/bmp — the first PIL→docx dpi reversion) → P7-2 (tiff then jpeg —
  the jpeg→tiff dependency-inversion) → P7-3 (drawing/`InlineShape`) → P7-4
  (`add_picture` wiring + picture COM). See the
  [image core page](api/image_core.md).
- **Enumerations (enum)** — two pages. The **enumeration base tier**: `BaseEnum`
  (MS-API-value enums) and `BaseXmlEnum` (XML-attribute-mapping enums), the base
  machinery **every docx enum extends** (the concrete `WD_*` enums and the P4–P6
  element classes that (de)serialize enumerated attributes). Documents the
  MATLAB enum-with-associated-data idiom (value classdef + `enumeration` block),
  the four **docx-vs-pptx semantic deltas** (chiefly `from_xml`'s absent
  None-guard — load-bearing for the `INHERITED` members — and `to_xml`'s
  Python-falsy reject), and the H3 None/`""`/`<missing>` tri-state. Then the
  **concrete enum tier** (P3-3, `+enum\{+text,+section,+dml}`): the 12
  `WD_*` / `MSO_*` enums (108 members) that stand on that base, the
  `WD_BREAK_TYPE.TEXT_WRAPPING` member-alias idiom, the `None`-valued
  `INHERITED` members and their `from_xml(None)` linkage, the
  **`MSO_THEME_COLOR_INDEX` docx-vs-pptx delta** (full-word tokens,
  `NOT_THEME_COLOR="UNMAPPED"`, no `MIXED`), and the 7 module-level class
  aliases.
- **Python → MATLAB mapping** — the living module→package and dunder/idiom
  reference tables.

## Relationship to Mat2Ppt

Mat2Doc is a **standalone** toolbox — it shares no code with
[Mat2Ppt](https://github.com/python-openxml/python-docx) (the python-pptx
port). The two projects share only the binding translation rules
(`design.md`) and the cross-language validation harness, which live in the
management repository and ship in neither toolbox. Where a symbol has a direct
Mat2Ppt analogue (as most of the `Length` family does), the design is
**re-implemented** in the `mat2doc:` namespace against the **python-docx
v1.2.0 source of truth**, never copied — so, for example, docx's `Twips`
class and its half-to-even `twips` conversion are ported, while pptx's
`Centipoints` (absent from docx) is not.

## Deviation posture

Divergences from the Python original are all on **dead paths** (unreachable
through any ported call site, API-invisible). Mat2Doc pre-adopts the Mat2Ppt
deviation rulings, carried `mat2doc:`-namespaced (see
`validation\summary\decision_2026-07-25_mat2doc_deviation_preadoption.md`):
**D-STYPE-1..4** (int/float indistinguishability), **D-002** (exotic
string-input dead paths), **D-003** (multiplier-constructor error-message
wording; exception class faithful), and **D-004** (non-finite / wrong-type
error identifiers). No new D-number has been opened.
