---
title: "mat2doc.oxml.shape / mat2doc.shape — the DrawingML inline-picture oxml + InlineShape API (the new_pic_inline builder)"
---

# `mat2doc.oxml.shape` / `mat2doc.oxml.drawing` / `mat2doc.shape` — the inline-picture oxml + InlineShape API

Ported from python-docx v1.2.0 `src/docx/oxml/shape.py` (the `CT_Inline` /
`CT_Picture` / `CT_Blip` DrawingML tree, ~17 element classes),
`src/docx/oxml/drawing.py` (`CT_Drawing`) and `src/docx/shape.py`
(`InlineShapes` / `InlineShape`). This is the **first P7 registry-adding WP** —
16 `register_element_cls` rows (`w:drawing`, `wp:inline`, `wp:extent`,
`wp:docPr`, `wp:anchor`, `a:graphic`, `a:graphicData`, `a:blip`, `a:off`,
`a:ext`, `a:xfrm`, `pic:pic`, `pic:nvPicPr`, `pic:cNvPr`, `pic:blipFill`,
`pic:spPr`) so a parsed `<w:drawing>` subtree resolves to the typed classes
instead of generic `XmlElement`.

It lands the **XML builder** an inline picture is made of — `CT_Inline.new_pic_inline`
— and the **read/mutate API** over an existing one — `InlineShapes` / `InlineShape`.
The **authoring wiring** that calls the builder (`Run.add_picture` /
`Document.add_picture` + `parts/image.py::ImagePart` + `StoryPart.get_or_add_image`)
arrives at **P7-4**; this WP freezes the byte-exact foundation those callers emit.

:::{important}
**★ `CT_Inline.new_pic_inline` — the inline-picture XML builder P7-4's `add_picture` emits, byte-identical to python-docx.**

`new_pic_inline(shape_id, rId, filename, cx, cy)` (`shape.py:92-103`) builds the
**complete** `<wp:inline>` DrawingML tree for an inline picture:
`wp:inline → wp:extent → wp:docPr → wp:cNvGraphicFramePr → a:graphic →
a:graphicData → pic:pic → (pic:nvPicPr, pic:blipFill, pic:spPr)`. It is a
two-step build (`shape.py:92-103`):

1. `pic = CT_Picture.new(0, filename, rId, cx, cy)` — the `<pic:pic>` subtree
   (`pic_id` is fixed at **0**: Word does not use it but does not omit it);
2. `inline = CT_Inline.new(cx, cy, shape_id, pic)` — `parse_xml` the exact
   `_inline_xml` template, then stamp the slots (`wp:extent` `cx`/`cy`, `wp:docPr`
   `@id` = `shape_id` and `@name` = `"Picture %d" % shape_id`, `a:graphicData`
   `@uri` = the DrawingML-picture URI) and `_insert_pic(pic)` appends the
   `<pic:pic>` into the (empty) `<a:graphicData>`.

The `cx`/`cy` are `Emu` (Length) values — at P7-4 they are the image
**dpi → EMU** extent computed from the P7-2 header parsers
(`Inches(px_width / horz_dpi)`). The emitted bytes are **byte-identical to
python-docx** across the whole param space — validated 10/10 (5 param sets ×
{inline, pic}) including XML-escaping, UTF-8 CJK/emoji and tiny-EMU filenames
(`s0087`, the frozen P7-4 oracle; set A SHA `ce13bdb0…71490`, 958 B).

**★ H8 — the `pic:pic` namespace-suppression (the sharp spot, signed
D-serializer-nsdecl).** The `wp:inline` template declares
`nsdecls("wp","a","pic","r")` and the `pic:pic` template declares
`nsdecls("pic","a","r")` — but when the parsed `<pic:pic>` is `_insert_pic`-moved
into the inline tree (which already declares `pic`/`a`/`r`), its now-redundant
`xmlns:pic/a/r` declarations are **suppressed** at serialize time (lxml's
verbatim-until-moved move-reconciliation, reproduced by Mat2Doc's serializer). So
the standalone `CT_Picture.new` output carries `xmlns:pic/a/r`, but the same
`<pic:pic>` inside a `<wp:inline>` carries **none** — byte-matching python-docx.
This is exactly the already-signed **D-serializer-nsdecl** behaviour; **no new
D-number.**
:::

:::{note}
**Registry-adding but M1-neutral.** These 16 rows are the first P7 registry
additions, so the registry-flip standing rule applies (full suite at Gate-1,
Gate-4 relaxes/re-pins). But `default.docx` carries **no `<w:drawing>`** and none
of the 16 tags occurs in any of its 17 parts, so nothing transits the new CT
classes on the M1 save path: `mat2doc.Document().save()` stays **17/17
byte-identical** to python-docx and the targeted regression showed **exactly 0
flips** (nothing for Gate-4 to re-pin).
:::

---

## Re-port vs novel — the docx-vs-pptx split

Only **`CT_Point2D`** (x/y) and **`CT_PositiveSize2D`** (cx/cy) are byte-identical
class bodies between docx and pptx, re-ported faithfully into the docx home
package. The shared **`a:`/`pic:` DrawingML** classes are re-ported from Mat2Ppt
but ported to match **DOCX** (which differs from pptx), and the **`wp:`** classes
are **novel** (python-pptx has no analogue). The load-bearing differences:

| class | pptx (Mat2Ppt) | docx (ported here) |
|---|---|---|
| `CT_Blip` | `rEmbed` only (`r:embed`) | **`embed` AND `link`** (`r:embed`, `r:link`) |
| `CT_GraphicalObjectData` | chart/tbl children, `uri`=`XsdString` | `pic` (`pic:pic`) + `uri`=**`XsdToken`** |
| `CT_Picture` | slide `p:pic` (BaseShapeElement, many builders) | inline **`pic:pic`**, one `new(pic_id,filename,rId,cx,cy)` |
| `CT_Transform2D` | off/ext/chOff/chExt, `_new_off`/`_new_ext` overrides | `off`/`ext` + `cx`/`cy`, **NO overrides** |
| `CT_ShapeProperties` | rich (custGeom / fill choice / ln) | `xfrm` + `cx`/`cy` only |
| `CT_Drawing` / `CT_Anchor` / `CT_Inline` / `CT_StretchInfoProperties` | — | **NOVEL** (docx-only) |

Four classes exist in `shape.py` for schema completeness but are **not**
registered upstream — `CT_NonVisualPictureProperties` (`pic:cNvPicPr`),
`CT_PresetGeometry2D` (`a:prstGeom`), `CT_RelativeRect` (`a:fillRect`),
`CT_StretchInfoProperties` (`a:stretch`). They are ported (empty bodies) and,
faithfully, **left out of the registry** — those tags parse as generic
`XmlElement`, exactly as in docx.

---

(id-ct-inline)=
## `CT_Inline` — the `<wp:inline>` root + the `new_pic_inline` builder

**Syntax**

```matlab
inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline(shape_id, rId, filename, cx, cy);
inline = mat2doc.oxml.shape.CT_Inline.new(cx, cy, shape_id, pic);   % the sub-builder
e  = inline.extent;    % OneAndOnlyOne <wp:extent>  (a CT_PositiveSize2D)
d  = inline.docPr;     % OneAndOnlyOne <wp:docPr>    (a CT_NonVisualDrawingProps)
g  = inline.graphic;   % OneAndOnlyOne <a:graphic>   (a CT_GraphicalObject)
```

**Description**

`CT_Inline` (`shape.py:68-118`) is the `<wp:inline>` element, the container for an
inline shape. Its three `OneAndOnlyOne` children — `extent` (`wp:extent`), `docPr`
(`wp:docPr`), `graphic` (`a:graphic`) — each raise `mat2doc:InvalidXmlError` when
absent. `new_pic_inline` is the byte-critical builder documented in the callout
above; `new` is the sub-builder that stamps a supplied `<pic:pic>` into the
template. `PICTURE_URI` is the fixed DrawingML-picture `@uri`
(`.../drawingml/2006/picture`).

**Example** (build an inline picture and read its DrawingML slots — the
`new_pic_inline` bytes are byte-identical to python-docx):

```matlab
inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", ...
    "python-image.png", mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
disp(mat2doc.oxml.serialize_part_xml(inline));
% <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
% <wp:inline xmlns:wp="..." xmlns:a="..." xmlns:pic="..." xmlns:r="..."><wp:extent
% cx="2438400" cy="1828800"/><wp:docPr id="1" name="Picture 1"/>...<a:graphicData
% uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr>
% ...<a:blip r:embed="rId7"/>...<a:ext cx="2438400" cy="1828800"/>...</pic:pic></a:graphicData>...
% (note: the moved <pic:pic> carries NO xmlns decls -- the H8 suppression)

fprintf('extent %d x %d EMU; docPr id=%d name="%s"\n', ...
    double(inline.extent.cx), double(inline.extent.cy), ...
    inline.docPr.id, inline.docPr.name);
% extent 2438400 x 1828800 EMU; docPr id=1 name="Picture 1"
```

*Ported from python-docx v1.2.0: `src/docx/oxml/shape.py::CT_Inline`*

---

(id-ct-picture)=
## `CT_Picture` + the DrawingML tree (`CT_Blip` / `CT_GraphicalObject` / `CT_Transform2D` / …)

**Syntax**

```matlab
pic = mat2doc.oxml.shape.CT_Picture.new(pic_id, filename, rId, cx, cy);   % builds <pic:pic>
pic.nvPicPr.cNvPr.id;      % @id   on pic:cNvPr
pic.nvPicPr.cNvPr.name;    % @name on pic:cNvPr (the filename)
pic.blipFill.blip.embed;   % @r:embed on a:blip (the image relationship id)
pic.spPr.cx;  pic.spPr.cy; % the a:xfrm/a:ext extent, as an Emu (Length)
```

**Description**

`CT_Picture` (`shape.py:135-179`) is the docx `<pic:pic>` (the **inline** picture,
`pic:`-prefixed children — distinct from python-pptx's slide `p:pic`). Its `new`
builder `parse_xml`s the exact `_pic_xml` template
(`pic:nvPicPr → pic:blipFill → pic:spPr`) then stamps five slots: `cNvPr` id/name,
`a:blip` `@r:embed`, and `a:ext` `cx`/`cy`. It is the sub-tree
`CT_Inline.new_pic_inline` appends via `_insert_pic`.

The surrounding DrawingML classes:

| class | element | role |
|---|---|---|
| `CT_GraphicalObject` | `a:graphic` | the `graphicData` `OneAndOnlyOne` |
| `CT_GraphicalObjectData` | `a:graphicData` | `@uri` (`XsdToken`) + the `pic` child (`ZeroOrOne`, `_insert_pic` appends) |
| `CT_Blip` | `a:blip` | `embed` (`r:embed`) **and** `link` (`r:link`), both `OptionalAttribute` (default None → `[]`) |
| `CT_BlipFillProperties` | `pic:blipFill` | `blip` (`ZeroOrOne`, successors `a:srcRect`/`a:tile`/`a:stretch`) |
| `CT_ShapeProperties` | `pic:spPr` | `xfrm` (`ZeroOrOne`) + the `cx`/`cy` convenience props |
| `CT_Transform2D` | `a:xfrm` | `off` (`a:off`) + `ext` (`a:ext`) + `cx`/`cy` — **no `_new_off`/`_new_ext` override** |
| `CT_Point2D` / `CT_PositiveSize2D` | `a:off` / `a:ext`, `wp:extent` | x/y ; cx/cy (byte-identical re-ports) |
| `CT_NonVisualDrawingProps` | `wp:docPr` / `pic:cNvPr` | `id` + `name` only |
| `CT_PictureNonVisual` | `pic:nvPicPr` | `cNvPr` only |

`CT_Blip`'s `embed`/`link` are the H3 tri-state: a get when absent returns `[]`
(None), a set to `[]` removes the attribute. `CT_Transform2D`/`CT_ShapeProperties`
`cx`/`cy` getters return `[]` (None) when the `a:ext`/`a:xfrm` is absent.

**Example** (build the `<pic:pic>` sub-tree standalone — here it IS the root, so
it carries its own `xmlns:pic/a/r`):

```matlab
pic = mat2doc.oxml.shape.CT_Picture.new(0, "python-image.png", "rId7", ...
    mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
disp(mat2doc.oxml.serialize_part_xml(pic));
% <?xml version='1.0' encoding='UTF-8' standalone='yes'?>
% <pic:pic xmlns:pic="..." xmlns:a="..." xmlns:r="..."><pic:nvPicPr><pic:cNvPr id="0"
% name="python-image.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip
% r:embed="rId7"/>...</pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext
% cx="2438400" cy="1828800"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>

fprintf('cNvPr id=%d name="%s"; blip embed="%s"; spPr cx=%d cy=%d\n', ...
    pic.nvPicPr.cNvPr.id, pic.nvPicPr.cNvPr.name, ...
    pic.blipFill.blip.embed, double(pic.spPr.cx), double(pic.spPr.cy));
% cNvPr id=0 name="python-image.png"; blip embed="rId7"; spPr cx=2438400 cy=1828800
```

*Ported from python-docx v1.2.0: `src/docx/oxml/shape.py::CT_Picture`*

---

(id-ct-drawing)=
## `CT_Drawing` — the `<w:drawing>` element

**Syntax**

```matlab
% <w:drawing> resolves to CT_Drawing when a real document.xml is parsed
```

**Description**

`CT_Drawing` (`drawing.py:10-11`) is the `<w:drawing>` element that carries a
DrawingML object (a picture or chart) inside a run. The class body is **empty** in
python-docx — registering the tag simply makes a parsed `<w:drawing>` a
`CT_Drawing` rather than a generic `XmlElement`; no descriptors, attributes or
methods are declared. It lives in `oxml/drawing.py` (not `shape.py`) for the
legacy reasons noted in that module. It is **novel** (python-pptx has no
`CT_Drawing`).

*Ported from python-docx v1.2.0: `src/docx/oxml/drawing.py::CT_Drawing`*

---

(id-inline-shapes)=
## `InlineShapes` / `InlineShape` — the read/mutate API

**Syntax**

```matlab
shapes = mat2doc.shape.InlineShapes(body_elm, parent);   % over a <w:body> (a CT_Body)
n  = shapes.len_();          % __len__
s  = shapes.getitem_(idx);   % __getitem__ (0-based; negative wraps)
arr = shapes.to_array();     % __iter__ -> 1xN InlineShape

sh = mat2doc.shape.InlineShape(inline);   % wrap one <wp:inline>
t  = sh.type;                % read-only WD_INLINE_SHAPE member
w  = sh.width;   sh.width  = mat2doc.shared.Emu(1000000);   % r/w Emu (Length)
h  = sh.height;  sh.height = mat2doc.shared.Emu(750000);    % r/w Emu (Length)
```

**Description**

`InlineShapes` (`shape.py:21-48`) is the sequence of a story's inline shapes. It
is a plain `Parented` proxy holding the `<w:body>` element (`_body`) and its parent
(a `StoryPart`); `_inline_lst` is `body.xpath("//w:p/w:r/w:drawing/wp:inline")`
(document order). The Python `Sequence` surface is ported as the **explicit
methods** `getitem_` / `to_array` / `len_` (the shared 1-based `()` `RedefinesParen`
base is a future WP — the `TabStops`/`Sections`/`Rows_` precedent). `getitem_` keeps
the **Python 0-based** key with negative-index wrap (**H1**: `+1` to the MATLAB
index); out of range raises `mat2doc:IndexError` with the verbatim message
`"inline shape index [%d] out of range"` carrying the **original** idx.

`InlineShape` (`shape.py:51-104`) wraps one `<wp:inline>` (`_inline`). In
python-docx it is a **plain object** (not `Parented`, not `ElementProxy`) — so the
MATLAB port is `< handle` with default handle identity, no `eq`/`ne`. `width`/`height`
read the `wp:extent` and, on set, **also** write the picture's `pic:spPr` extent so
the frame and the extent stay in sync. `type` (**H10**) dispatches the
`a:graphicData` `@uri` against the fixed `nsmap` URIs → a `WD_INLINE_SHAPE` member:
`PICTURE` / `LINKED_PICTURE` (the `a:blip` carries `r:link`, an H3 tri-state check)
/ `CHART` / `SMART_ART` / `NOT_IMPLEMENTED`.

**Example** (wrap a built inline, read/resize it, then drive the collection over a
one-picture body):

```matlab
inline = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", ...
    "python-image.png", mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));

sh = mat2doc.shape.InlineShape(inline);
fprintf('type=%s  width=%d  height=%d\n', string(sh.type), ...
    double(sh.width), double(sh.height));
% type=PICTURE  width=2438400  height=1828800
sh.width  = mat2doc.shared.Emu(1000000);
sh.height = mat2doc.shared.Emu(750000);
fprintf('after resize: pic spPr %d x %d\n', ...
    double(inline.graphic.graphicData.pic.spPr.cx), ...
    double(inline.graphic.graphicData.pic.spPr.cy));
% after resize: pic spPr 1000000 x 750000   (the setter wrote the pic:spPr too)

% Build a one-picture <w:body> and drive the collection.
inline2 = mat2doc.oxml.shape.CT_Inline.new_pic_inline(1, "rId7", ...
    "python-image.png", mat2doc.shared.Emu(2438400), mat2doc.shared.Emu(1828800));
docXml = "<w:document " + mat2doc.oxml.nsdecls("w","wp","a","pic","r") + ">" + ...
    "<w:body><w:p><w:r><w:drawing>" + string(inline2.xml) + ...
    "</w:drawing></w:r></w:p></w:body></w:document>";
body   = mat2doc.oxml.parse_xml(docXml).body;
shapes = mat2doc.shape.InlineShapes(body, []);
fprintf('len_=%d ; shapes[0] type=%s\n', shapes.len_(), string(shapes.getitem_(0).type));
% len_=1 ; shapes[0] type=PICTURE
try
    shapes.getitem_(5);
catch err
    fprintf('%s: %s\n', err.identifier, err.message);
    % mat2doc:IndexError: inline shape index [5] out of range
end
```

*Ported from python-docx v1.2.0: `src/docx/shape.py::InlineShapes` / `::InlineShape`*

---

## ★ The picture oxml + InlineShape API is live — P7-4 closes Phase 7

P7-3 lands the DrawingML inline-picture element tree (`CT_Inline` with its byte-proven
`new_pic_inline` builder, `CT_Picture` and the `a:`/`pic:` sub-classes, `CT_Drawing`),
the 16-row registry addition (M1-neutral, 0 flips), and the read/mutate API
(`InlineShapes` / `InlineShape`, with `WD_INLINE_SHAPE` typing and the width/height
extent+spPr sync). The `new_pic_inline` output is **byte-identical to python-docx**
across the whole param space (`s0087`, the frozen P7-4 oracle), and a real
`add_picture` document.xml round-trips + width/height-mutates byte-identical
(`s0088`), so **zero new D-numbers** — the `pic:pic` namespace suppression is the
already-signed D-serializer-nsdecl.

What remains in Phase 7 is the **authoring wiring** — **P7-4**:
`Run.add_picture` / `Document.add_picture` + `parts/image.py::ImagePart` +
`StoryPart.get_or_add_image` / `new_pic_inline` + `Package.get_or_add_image_part`
/ `image_parts` — which places the P7-2 image **dpi → EMU** values into the
`wp:extent`/`a:ext` this builder stamps, plus the **C6 header-image scenario**
(the first `word/_rels/header1.xml.rels`) and the **picture Word-COM sweep**
(inline-picture body doc + header-image doc). That closes **Phase 7**.

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
