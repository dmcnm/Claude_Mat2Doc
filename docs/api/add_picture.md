---
title: "add_picture — the picture-authoring path (ImagePart · ImageParts · StoryPart · Run.add_picture · Document.add_picture)"
---

# `add_picture` — the picture-authoring path (the P7-4 wiring that closes Phase 7)

Ported from python-docx v1.2.0 `src/docx/document.py::Document.add_picture`,
`src/docx/text/run.py::Run.add_picture`, `src/docx/parts/story.py::StoryPart`
(`get_or_add_image` / `new_pic_inline`), `src/docx/package.py`
(`Package.get_or_add_image_part` / `image_parts` / `ImageParts`) and
`src/docx/parts/image.py::ImagePart`. This is the **final Phase-7 work package**
— it un-stubs every picture-owner deferral from P2..P5 and wires the P7-1..P7-3
foundation (image parsers → DrawingML `new_pic_inline` builder) into a single
public call. **It is the picture milestone: the first runtime image part in
Mat2Doc.**

:::{important}
**★ `doc.add_picture(path[, width[, height]])` — the whole authoring path, byte-identical to python-docx and COM-verified in real Word.**

A single call emits **three coordinated additions** to the package, each
byte-identical to python-docx v1.2.0:

1. a new **image part** `word/media/imageN.<ext>` — the image bytes copied
   **verbatim** (an `ImagePart` whose blob is an exact source-file copy);
2. its **relationship** — `<Relationship Id="rIdN" Type=".../image"
   Target="media/imageN.<ext>"/>` in the caller's story `.rels`, plus a
   `<Default Extension="<ext>" ContentType="image/<ext>"/>` in
   `[Content_Types].xml`;
3. a **`<w:drawing>`/`<wp:inline>`** tree in the run — built by the P7-3
   `CT_Inline.new_pic_inline(shape_id, rId, filename, cx, cy)` — whose
   `wp:extent`/`a:ext` `cx`/`cy` are the image's **dpi → EMU** extent.

The end-to-end call chain (all un-stubbed at P7-4):

```
Document.add_picture(path, width, height)
  → add_paragraph().add_run().add_picture(path, width, height)      % Run
      → StoryPart.new_pic_inline(path, width, height)
          → StoryPart.get_or_add_image(path)                        % rId + Image
              → Package.get_or_add_image_part(path)                 % ImagePart (sha1 dedupe)
              → StoryPart.relate_to(image_part, RT.IMAGE)           % rId
          → Image.scaled_dimensions(width, height)                  % cx, cy (EMU)
          → StoryPart.next_id                                       % shape_id
          → CT_Inline.new_pic_inline(shape_id, rId, filename, cx, cy)   % the P7-3 builder
      → Run._r.add_drawing(inline)                                  % <w:drawing> in the <w:r>
      → InlineShape(inline)                                         % the return value
```

The full-package output is **18/18 (L1 + binary) byte-identical to
python-docx** (`s0090`, the frozen P7 picture oracle) and **opens clean in real
Microsoft Word** (Word 16.0.20228 — silent open, no repair prompt, the picture
renders, clean round-trip; `com_verify_P7_pictures.md`). **Zero new
D-numbers.**
:::

:::{note}
**The dpi → EMU chain (P7-1..P7-2 → P7-3 → P7-4).** The `wp:extent` this path
stamps is not a magic number — it is computed from the image's own header.
`Image.scaled_dimensions(width, height)` (P7-1a) returns the native size when
both are `None` (`[]`), where the native size is `Inches(px / dpi)` and the
**dpi comes from the format parser** (PNG pHYs / GIF 72 / BMP px-per-metre /
TIFF X/YResolution / JPEG JFIF-density-or-Exif — all reverted to **docx math,
not PIL**, at P7-1b/P7-2). P7-3's `new_pic_inline` places those EMU values into
`wp:extent`/`a:ext`. So `python-powered.png` (150 dpi) → `1778000 × 711200 EMU`,
and `add_picture(png, width=Inches(2))` overrides to
`<wp:extent cx="1828800" cy="731520"/>` (aspect-preserved height). Every one of
those bytes is verified against python-docx.
:::

---

## `ImagePart` — the runtime image part

**Syntax**

```matlab
ip = mat2doc.parts.ImagePart(partname, content_type, blob);          % ctor (image=None)
ip = mat2doc.parts.ImagePart(partname, content_type, blob, image);   % ctor (image seeded)
ip = mat2doc.parts.ImagePart.from_image(image, partname);            % from a decoded Image
ip = mat2doc.parts.ImagePart.load(partname, content_type, blob, pkg);% the PartFactory entry
img = ip.image();          % the |Image| (decoded lazily on first access)
h   = ip.sha1();           % 40-char hex SHA1 of the blob (the dedupe key)
fn  = ip.filename();       % source filename, or generic "image.<ext>"
cx  = ip.default_cx();     % native width  (a Length / Emu)
cy  = ip.default_cy();     % native height (a Length / Emu)
```

**Description**

`ImagePart` (`parts/image.py::ImagePart`) is a `mat2doc.opc.Part` specialization
whose blob is the raw image bytes, returned **verbatim** — so an image part
round-trips **byte-for-byte** (it inherits `Part.blob` unchanged). The docx
constructor is `(partname, content_type, blob, image=None)` — `image` **last**,
and **no `package` argument** (docx forwards only three args to `Part`; the base
`package` stays `None`). `image` is decoded **lazily** on first `.image()` access
via the manual `if self._image is None` sentinel (H3 — ported as
`isequal(image_, [])`, since an `Image` handle is never `[]`), so a **loaded**
part (via `load` / `PartFactory`) still answers `.sha1` without eagerly decoding.

`sha1` is the **dedupe key** — a plain property recomputed each call (docx does
**not** cache it, unlike pptx), `hashlib.sha1(blob).hexdigest()`. `filename`
returns the source `Image.filename` when the part was built `from_image`, else a
generic `"image.<ext>"` from the partname extension. `default_cx`/`default_cy`
are the native EMU dimensions — **not** on the `add_picture` path (that uses
`Image.scaled_dimensions`); they are ported **verbatim including the docx source
quirk** that `default_cy` reads `horz_dpi` (not `vert_dpi`) for the height
(`image.py:43`, design.md §7 — no behavior "improvement").

*Ported from python-docx v1.2.0: `src/docx/parts/image.py::ImagePart`*

---

## `Package.get_or_add_image_part` + `ImageParts` — the SHA1 dedupe

**Syntax**

```matlab
ip     = pkg.get_or_add_image_part(image_descriptor);   % ImagePart (created or reused)
parts  = pkg.image_parts();                             % the ImageParts collection (cached)

n   = parts.len_();          % __len__
tf  = parts.contains(ip);    % __contains__ (handle identity)
arr = parts.to_array();      % __iter__ -> 1xN ImagePart
```

**Description**

`ImageParts` (`package.py::ImageParts`) is a package-level collection (one per
`Package`, held by `Package.image_parts` as a `@lazyproperty`) that
**de-duplicates image parts by SHA1**: adding the same image bytes twice reuses
the one `ImagePart`. `get_or_add_image_part(descriptor)` decodes the image
(`Image.from_file`), looks up an existing part by SHA1 (`_get_by_sha1`) and
returns it on a hit, else mints a new `ImagePart` with the next partname
(`_next_image_partname`) and appends it. The collection is populated at open by
`Package._gather_image_parts` (every existing internal `IMAGE`-relationship
target) and grown on demand by this path.

The Python `Sequence` surface (`__contains__` / `__iter__` / `__len__`) is ported
as the **explicit methods** `contains` / `to_array` / `len_` (the
`InlineShapes`/`TabStops`/`Rows_` VERIFY-COLLECTION precedent). The private
helpers rotate to trailing underscores (`_add_image_part` → `add_image_part_`,
`_get_by_sha1` → `get_by_sha1_`, `_next_image_partname` →
`next_image_partname_`).

:::{warning}
**The non-numbered-partname guard (Gate-2 DEFECT-1).** `_next_image_partname`
builds `used_numbers = [part.partname.idx for part in self]`. When a gathered
image part carries a **non-numbered** media partname (`/word/media/logo.png` in
a third-party or hand-authored docx), `PackURI.idx` is `None`; Python appends
`None` harmlessly (a `None` entry never equals a candidate integer, so
`n not in used_numbers` skips it). The naive MATLAB port
`used_numbers(end+1) = parts(i).partname().idx` assigns `[]` into one element and
**crashes** (`MATLAB:matrix:singleSubscriptNumelMismatch`). The Gate-2 auditor
caught this (`add_picture` on a reopened `logo.png` doc), and the port
**skips the `[]` (None) idx before appending** — a list of just the real
integers gives the **identical** next-partname result as Python's
`None`-tolerant list. Verified byte-identical (`s0095` scenario J, 19/19). Not a
D-number (a crash-fidelity bug, no output deviation).
:::

*Ported from python-docx v1.2.0: `src/docx/package.py::ImageParts`*

---

## `StoryPart.get_or_add_image` / `new_pic_inline` — the per-story wiring

**Syntax**

```matlab
[rId, image] = storyPart.get_or_add_image(image_descriptor);         % rId + Image
inline       = storyPart.new_pic_inline(descriptor, width, height);  % a <wp:inline>
```

**Description**

`StoryPart.get_or_add_image` (`parts/story.py:27-39`) delegates the media part to
the **package** (`self.package.get_or_add_image_part`, so the media part is
package-shared) and relates it to **this** story part
(`self.relate_to(image_part, RT.IMAGE)`), returning `(rId, image_part.image)`.
The relationship dedupe lives in `relate_to` (`get_or_add`): the same image
reused on the **same** story part → same part **and** same `rId`; reused across a
**different** story part (e.g. a header) → same package-shared part, a **new**
`rId`.

`new_pic_inline` (`parts/story.py:60-74`) is the bridge to the P7-3 builder: it
gets the rId + `Image`, computes the extent via
`Image.scaled_dimensions(width, height)` (native or aspect-preserving override),
takes the next drawing id (`self.next_id`), and calls
`CT_Inline.new_pic_inline(shape_id, rId, image.filename, cx, cy)`.

*Ported from python-docx v1.2.0: `src/docx/parts/story.py::StoryPart`*

---

## `Run.add_picture` / `Document.add_picture` — the public surface

**Syntax**

```matlab
sh = run.add_picture(image_path_or_stream);                    % native size
sh = run.add_picture(image_path_or_stream, width);             % width, height auto
sh = run.add_picture(image_path_or_stream, width, height);     % both

sh = doc.add_picture(image_path_or_stream);                    % own paragraph at doc end
sh = doc.add_picture(image_path_or_stream, width, height);
```

**Description**

`Run.add_picture` (`text/run.py:59-81`) builds the inline via the owning
`StoryPart.new_pic_inline`, appends the `<w:drawing>` into the run
(`self._r.add_drawing(inline)`) and returns an `InlineShape(inline)` (the P7-3
read/mutate proxy). `width`/`height` default `None` (`[]`) — native size /
aspect-preserving scale (H13). `Document.add_picture` (`document.py:121-138`) is
the convenience wrapper: `add_paragraph().add_run().add_picture(...)` — a picture
in its own paragraph at the end of the document.

Because the run's owning part is a **`StoryPart`**, `add_picture` works in **any
story** — the body (a `DocumentPart`) **and a header/footer** (a `HeaderPart` /
`FooterPart`, P5-3b). The header case is the P5 carry-forward **C6**, discharged
here (see below).

**Example** (the body-picture headline — byte-identical to python-docx,
COM-verified):

```matlab
d  = mat2doc.Document();
sh = d.add_picture(pngPath);          % pngPath = ...\python-powered.png (150 dpi)
fprintf('type=%s  width=%d  height=%d EMU\n', ...
    string(sh.type), double(sh.width), double(sh.height));
% type=PICTURE  width=1778000  height=711200 EMU

pkg = d.part().package();
arr = pkg.image_parts().to_array();
fprintf('media parts=%d  first=%s\n', pkg.image_parts().len_(), string(arr(1).partname()));
% media parts=1  first=/word/media/image1.png
```

*Ported from python-docx v1.2.0: `src/docx/text/run.py::Run.add_picture` /
`src/docx/document.py::Document.add_picture`*

---

## ★ The SHA1 dedupe — same image, one media part

Adding the **same** image twice mints **one** `word/media/image1.<ext>` part
(SHA1 match in `ImageParts.get_or_add_image_part`); both `<w:drawing>` blips carry
the same `r:embed`. A **distinct** image takes the next number
(`image2.<ext>` — the number is unique **without regard to extension**), a
second `[Content_Types]` Default, and a new rId (insertion order, no rId sort).

```matlab
d = mat2doc.Document();
d.add_picture(pngPath);
d.add_picture(pngPath);               % same bytes -> reused
n1 = d.part().package().image_parts().len_();      % 1
d.add_picture(jpegPath);              % distinct -> image2.jpeg
n2 = d.part().package().image_parts().len_();      % 2
fprintf('after same-image x2: %d ; after distinct: %d\n', n1, n2);
% after same-image x2: 1 ; after distinct: 2
```

Byte-verified: `s0091` (dedupe, one `image1.png`, both blips `r:embed="rId9"`) and
`s0092` (two images `image1.png` + `image2.jpeg`, rId9/rId10, two CT Defaults).

---

## ★ C6 — the header image (the P5 debt discharged; the first `header1.xml.rels`)

`add_picture` on a **header** paragraph routes through the `HeaderPart`'s
`StoryPart`. Accessing the first section's header materializes
`word/header1.xml`, and relating the image on the header part mints the
**first-ever `word/_rels/header1.xml.rels`** in Mat2Doc
(`<Relationship Id="rId1" Type=".../image" Target="media/image1.png"/>`). The
media part is **package-shared** — if the same image is in the body **and** the
header, there is still **one** `word/media/image1.png`; the body's `DocumentPart`
relates it `rId9`, the header's `HeaderPart` relates it its own `rId1`.

```matlab
d     = mat2doc.Document();
sec   = d.sections.getitem_(0);               % the first section
hdr   = sec.header;                           % the primary header (a _Header)
paras = hdr.paragraphs;
paras(1).add_run().add_picture(pngPath);
% -> word/header1.xml + word/_rels/header1.xml.rels (rId1 -> media/image1.png)
```

Byte-verified: `s0096` (header image, 20/20 — first `header1.xml.rels`) and
`s0097` (body + header share one media part, 20/20). **COM-verified**: real Word
resolves the header image from the header part's own rels and renders it in the
**header band** with no missing-relationship prompt (`com_verify_P7_pictures.md`,
`s0096` PDF render).

---

## ★ PHASE 7 COMPLETE — the image tier is byte-proven end-to-end + COM-verified

P7-4 is the **final Phase-7 work package**. With it, the entire image tier is
done, byte-identical to python-docx v1.2.0, and accepted by the real Word oracle:

| sub-phase | delivered | proof |
|---|---|---|
| **P7-1a/1b/2a/2b** — image parsers | the format-agnostic `Image` core + the **6** `SIGNATURES` parsers (PNG/GIF/BMP/TIFF/JPEG-JFIF/JPEG-Exif), with the **PIL → docx dpi reversion** | value-identical to `Image.from_file` (`s0083`–`s0086`) |
| **P7-3 [N]** — DrawingML oxml + InlineShape | `CT_Inline`/`CT_Picture`/`CT_Blip`/… + `CT_Drawing` + `InlineShapes`/`InlineShape` + 16 registry rows; **`new_pic_inline`** | byte-identical `new_pic_inline` (`s0087` 10/10) + drawing round-trip (`s0088`) |
| **P7-4** — `add_picture` wiring | `ImagePart` / `ImageParts` (sha1 dedupe) + `StoryPart.get_or_add_image`/`new_pic_inline` + `Run`/`Document.add_picture` + the `PartFactory` IMAGE→`ImagePart` flip | **18/18 full-package** (`s0090`); header (`s0096`), dedupe (`s0091`), two-image (`s0092`), width override (`s0094`) |

**The picture Word-COM sweep PASSED** (Word 16.0.20228): all four frozen picture
packages open **silent, no repair prompt, no dropped content** — the body inline
PNG (`s0090`), the header image (`s0096`, first `header1.xml.rels`), two distinct
inline images (`s0092`), and one shared media part rendered in **both** body and
header (`s0097`) — with correct per-shape picture type and dimensions in Word's
object model, a visible render, and a clean round-trip (`com_verify_P7_pictures.md`).

**M1 stayed 17/17** through P7-4 even though un-stubbing `image_parts` puts
`_gather_image_parts` on **every** open: `default.docx` has no internal `IMAGE`
relationship (its thumbnail is a `THUMBNAIL` reltype), so the gather
short-circuits before any append and the empty-doc save is byte-unperturbed
(re-derived, not trusted). **Zero new D-numbers across all of Phase 7.**

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
