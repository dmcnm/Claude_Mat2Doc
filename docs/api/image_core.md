---
title: "mat2doc.image (core) — the image-parsing CORE (Image · ImageHeaderFactory · BaseImageHeader · StreamReader/BytesIO · the MIME/JPEG/PNG/TIFF constants — Phase 7 begins)"
---

# `mat2doc.image` (core) — the image-parsing CORE

Ported from python-docx v1.2.0 `src/docx/image/image.py`
(`Image` / `_ImageHeaderFactory` / `BaseImageHeader`),
`src/docx/image/helpers.py` (`StreamReader`) and
`src/docx/image/constants.py` (the five constants tables), all landing in the
NEW package `+mat2doc/+image/`. This is the **format-agnostic core** of the
image tier: it decodes an image blob's header enough to expose the
characterization a `.docx` needs before a picture can be added — `content_type`,
`px_width`/`px_height`, `horz_dpi`/`vert_dpi`, native `width`/`height` (as a
`Length`), `sha1` and `ext`. The per-format byte parsers (PNG/GIF/BMP,
JPEG/TIFF) arrive in the WPs after this one; here they are clean
`notYetPorted` stubs that the factory dispatches to by name.

:::{important}
**★ The image core is a re-port from Mat2Ppt — and this is the HOME of the parsers.**
python-pptx v1.0.2 has **no image-parsing module at all** — it reads size and
dpi from **PIL**. python-**docx** is where these header parsers actually live
(`docx/image/`), and Mat2Ppt's `+mat2ppt/+image/` was itself already a MATLAB
port of that docx code (built as the PIL substitute). So for Mat2Doc the image
core is a **same-source re-home** back to its natural package — the oracle is
python-docx itself — and P7-1a **re-ports from Mat2Ppt's `+image`** rather than
re-deriving from scratch. The re-port applies exactly **three** transforms
(below); everything else is docx-faithful in Mat2Ppt and re-ports near-verbatim.
:::

(id-report-transforms)=
## The three re-port transforms (`mat2ppt.image` → `mat2doc.image`)

| # | transform | what changed |
|---|---|---|
| **T1** | **namespace** | `mat2ppt.image.*` → `mat2doc.image.*`; the `Length` family moves to its docx home `mat2ppt.util.Inches`/`Emu` → **`mat2doc.shared.Inches`/`Emu`**; `sha1_hexdigest` → `mat2doc.opc.sha1_hexdigest`; every error id `mat2ppt:*` → `mat2doc:*`. |
| **T2** | **inline None-idiom** | Mat2Doc has **no shared `isNone`** (ratified 2026-07-26); every `mat2ppt.util.isNone(x)` in the Mat2Ppt `Image.m` becomes the inline `isequal(x, [])` guard (the three `scaled_dimensions` sites + the `_from_stream` filename default). |
| **T3** | **WMF-exclusion** | the pptx-only PIL seams are **NOT ported**: no `Wmf.m`, the **2 WMF/EMF `SIGNATURES` rows are dropped** (8 docx rows only), `MIME_TYPE.X_WMF` is dropped, and the `pil_dpi` / `int_dpi` round-and-clamp seams never transit (they belong to pptx `parts/image.py`, not docx). |

:::{warning}
**★ The dpi is docx math, NOT Mat2Ppt's PIL re-oracling — this reversion bites the FORMAT parsers.**
Because Mat2Ppt's Gate-3 oracle was python-pptx/PIL, three dpi sites in Mat2Ppt
were deliberately fixed *away* from docx's arithmetic. Those sites live in the
**format parsers** (P7-1b / P7-2), not in this core, and each must be
**reverted to docx semantics** when its parser lands:
`Bmp._dpi` = `int(round(px_per_meter * 0.0254))`, ppm==0 → **96** (Mat2Ppt/PIL
gave 72); `_TiffParser._dpi` **per-axis**, tag-absent → **72**, unit==1 → 72,
cm → ×2.54; `_App1Marker` (Exif) dpi straight from the embedded TIFF parse.
The **core itself introduces zero dpi arithmetic** — `Image.horz_dpi`/`vert_dpi`
are pure delegations to the header, and `BaseImageHeader` has **no**
`default_dpi` member — so no dpi decision is frozen at P7-1a. The reversion is a
standing P7-1b/P7-2 porter obligation, probed against the **python-docx** oracle.
:::

(id-phase-7-roadmap)=
## Phase 7 (images) — where this fits

**Phase 7 begins here.** The image core is done; the tier proceeds:

1. **P7-1a — image core (this WP).** `Image` + factory + `BaseImageHeader` +
   `StreamReader`/`BytesIO` + the five constants tables + `sha1_hexdigest`.
2. **P7-1b — PNG / GIF / BMP parsers.** The first `BaseImageHeader` subclasses;
   the **first PIL→docx dpi reversion** (`Bmp._dpi` ppm==0 → 96).
3. **P7-2 — TIFF then JPEG.** Ported **TIFF first** — `jpeg.py` imports `tiff.py`
   (`_App1Marker._tiff_from_exif_segment` calls `Tiff.from_stream`), so the
   planned jpeg-then-tiff order is dependency-**inverted**; TIFF ports first.
   Both revert their dpi to docx math.
4. **P7-3 — `oxml/drawing` + `oxml/shape` (`CT_Inline`/`CT_Anchor`/`CT_Blip`) +
   the `InlineShape`(s) API.** The registry-adding P7 WP.
5. **P7-4 — `Run.add_picture` / `Document.add_picture` wiring + `parts/image`
   (`ImagePart`) + the `Package`/`StoryPart` un-stubs + the picture Word-COM
   sweep** (including the header-image scenario — first materialization of
   `word/_rels/header1.xml.rels`).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape.
:::

(id-coverage-boundary)=
## The coverage boundary (why the examples seed a header)

Because the format parsers are stubbed at P7-1a, a full end-to-end
`Image.from_blob(real_png)` cannot yet **parse** on the MATLAB side — the
factory dispatches to a `notYetPorted` stub. Two of the core's legs are
nonetheless proven fully:

- **`sha1` is end-to-end** — it hashes the blob, no parser required
  (real `300-dpi.png` = `ca11f589…`, `300-dpi.jpg` = `4040957e…`, both
  `== hashlib.sha1(blob).hexdigest()`).
- **the `Image` delegation + `Length` math + `ext`** are proven by seeding a
  `BaseImageHeader` with the exact px/dpi/content-type python-docx parsed, then
  confirming every `Image` property matches (Gate-3 `s0083`, 195/195
  value-identical).

The examples below use the same seeded-header technique: a plain
`BaseImageHeader` carries real px/dpi values so `Image.width`/`height`/`ext`/
`sha1`/`scaled_dimensions` run against the shipped toolbox alone. The **raw
header parse** (does MATLAB read 860 px out of the PNG bytes?) is P7-1b/P7-2
Gate-3 material.

---

(id-image)=
## `Image`

**Syntax**

```matlab
img = mat2doc.image.Image.from_blob(blob);            % parse from image bytes
img = mat2doc.image.Image.from_file(path_or_stream);  % parse from a path or a BytesIO
ct  = img.content_type;                                % MIME type (from the header)
w   = img.width;                                       % native width  as a Length (Inches)
h   = img.height;                                      % native height as a Length (Inches)
[cx, cy] = img.scaled_dimensions(width, height);       % (Emu, Emu) scaled pair
d   = img.sha1;                                        % 40-char SHA-1 hex of the blob
e   = img.ext;                                         % file extension, no dot, case kept
```

**Description**

`Image` is the value object over a decoded image stream — python-docx's
`class Image(object)` (`image.py:18`). It stores the raw blob, an optional
source filename, and a `BaseImageHeader` (the parsed characterization). It is a
**handle class** because two of its members are `@lazyproperty`:

- `ext` (`@lazyproperty`) — `os.path.splitext(self._filename)[1][1:]`, the
  extension after the last dot of the basename, **without** the leading period
  and **without** lowercasing. It is realized through the package-private
  [`splitext_ext`](#id-ext-fix) helper (the Gate-2 **F-1** fix), not MATLAB
  `fileparts`.
- `sha1` (`@lazyproperty`) — `mat2doc.opc.sha1_hexdigest(blob)`, the
  40-character lowercase SHA-1 hex of the blob (`hashlib.sha1(blob).hexdigest()`).

Both use the cache-plus-logical-computed-flag idiom (never `isempty` as the
sentinel, H3 — `[]` is a legal cached value).

(id-image-members)=
**The member surface.**

| member | kind | delegates to / computes | note |
|---|---|---|---|
| `blob` | get | stored blob | the raw `uint8` bytestream |
| `content_type` | get | `image_header.content_type` | MIME type (concrete header) |
| `filename` | get | stored filename | a `string`, or `[]` (None) |
| `px_width` / `px_height` | get | `image_header.px_*` | pixel dimensions |
| `horz_dpi` / `vert_dpi` | get | `image_header.*_dpi` | **pure delegation — no dpi math in the core** |
| `width` / `height` | get | `Inches(px / dpi)` | native `Length` (H6: `Inches()` truncates via `fix()`) |
| `ext` | get (`@lazyproperty`) | `splitext_ext(filename)` | no dot, case preserved |
| `sha1` | get (`@lazyproperty`) | `mat2doc.opc.sha1_hexdigest(blob)` | 40-char lowercase hex |
| `scaled_dimensions(width, height)` | method | native / aspect-scaled `Emu` pair | both `[]` → native; one `[]` → aspect-scaled (H6 `pyRound`); both set → `(width, height)` |
| `from_blob(blob)` | static | `from_stream_(BytesIO(blob), blob)` | parse from bytes |
| `from_file(descriptor)` | static | path → open + basename; BytesIO → `seek(0)`/`read()`/`filename=[]` | parse from a path or a file-like object |

`scaled_dimensions` uses the inline `isequal(x, [])` None-guard (T2) at all three
sites, and routes its `round(length * scaling_factor)` through the private
[`pyRound`](#id-primitives) (banker's half-to-even, H6) before wrapping in `Emu`.

**Example**

```matlab
hdr = mat2doc.image.BaseImageHeader(860, 579, 300, 300);   % seed px/dpi (as a parser would)
img = mat2doc.image.Image(uint8([137 80 78 71]), "photo.png", hdr);
disp(img.px_width);                 % 860
disp(double(img.width));            % 2621280   (Inches(860/300) in EMU)
disp(double(img.height));           % 1764792   (Inches(579/300) in EMU)
disp(img.ext);                      % png       (splitext_ext on "photo.png")
disp(strlength(img.sha1));          % 40        (SHA-1 hex of the blob)
[cx, cy] = img.scaled_dimensions(); % both None -> native
disp([double(cx) double(cy)]);      % 2621280 1764792
[cx2, ~] = img.scaled_dimensions(mat2doc.shared.Emu(1310640));  % width-only, aspect kept
disp(double(cx2));                  % 1310640
```

*Ported from python-docx v1.2.0: `src/docx/image/image.py::Image`*

---

(id-baseimageheader)=
## `BaseImageHeader`

**Syntax**

```matlab
h = mat2doc.image.BaseImageHeader(px_width, px_height, horz_dpi, vert_dpi);
n = h.px_width;        % the four stored characteristics
c = h.content_type;    % abstract -> raises on the base; concrete subclasses override
```

**Description**

`BaseImageHeader` is the base for the per-format header subclasses (`Png`, `Gif`,
`Bmp`, `Tiff`, the JPEG headers). It holds the four characteristics every header
exposes — `px_width`, `px_height`, `horz_dpi`, `vert_dpi` — as read-only
`Dependent` properties over protected storage. In python-docx `content_type` and
`default_ext` are **abstract `@property`** members raising `NotImplementedError`;
here they are **methods** (a MATLAB subclass cannot redefine an inherited
property's getter, so the overridable members are methods while the four stored
characteristics remain inherited `Dependent` properties). The base raises
`mat2doc:NotImplementedError` with the two upstream messages verbatim.

:::{note}
**No `default_dpi` member.** The "defaults to 72 dpi" text in the docstrings is
realized **inside each parser** (`png.py`/`bmp.py`), not on the base — so there
is no `default_dpi` attribute to port, and the core carries zero dpi arithmetic.
:::

**Example**

```matlab
h = mat2doc.image.BaseImageHeader(100, 50, 96, 96);
disp(h.px_width);                                         % 100
disp(h.vert_dpi);                                         % 96
try; h.content_type; catch e; disp(e.identifier); end     % mat2doc:NotImplementedError
try; h.default_ext;  catch e; disp(e.identifier); end     % mat2doc:NotImplementedError
```

*Ported from python-docx v1.2.0: `src/docx/image/image.py::BaseImageHeader`*

---

(id-factory)=
## `ImageHeaderFactory_` — the `SIGNATURES` dispatch (WMF-excluded)

**Syntax**

```matlab
header = mat2doc.image.ImageHeaderFactory_(stream);   % a BaseImageHeader subclass
```

**Description**

`ImageHeaderFactory_` (`_ImageHeaderFactory`, `image.py:168`; the leading-
underscore module-private function → the trailing-underscore rotation) reads the
first 32 bytes of `stream` and matches them, **in order**, against the docx
`SIGNATURES` table (`__init__.py:13-23`). The **first** signature whose bytes
equal the header slice at its offset selects the parser class, and its
`from_stream` is invoked; if nothing matches it raises
`mat2doc:UnrecognizedImageError`. First-match order is behavior, so the table is
transcribed in exact upstream order.

(id-signatures)=
**The 8-row `SIGNATURES` table (WMF/EMF dropped).** The two pptx-only WMF/EMF
rows are **not** ported (T3), so the table has exactly the eight docx rows:

| # | signature (first bytes) | offset | parser class | WP |
|---|---|---|---|---|
| 1 | `89 PNG \r\n 1a \n` | 0 | `Png` | P7-1b |
| 2 | `JFIF` | 6 | `Jfif` | P7-2 (jpeg) |
| 3 | `Exif` | 6 | `Exif` | P7-2 (jpeg) |
| 4 | `GIF87a` | 0 | `Gif` | P7-1b |
| 5 | `GIF89a` | 0 | `Gif` | P7-1b |
| 6 | `MM \x00 *` (big-endian) | 0 | `Tiff` | P7-2 (tiff) |
| 7 | `II * \x00` (little-endian) | 0 | `Tiff` | P7-2 (tiff) |
| 8 | `BM` | 0 | `Bmp` | P7-1b |

The eight dispatch targets are `@Class.from_stream` **function handles resolved
at call time**, so the factory lands whole at P7-1a even though every parser
class is still a `notYetPorted` stub — dispatching to a recognized format raises
`mat2doc:notYetPorted` (naming the owning WP) rather than parsing, while an
unrecognized stream raises `mat2doc:UnrecognizedImageError` on both sides.

:::{important}
**★ WMF-exclusion is package-visible.** A real WMF placeable header
(`D7 CD C6 9A …`) or a synthetic EMF header (`01 00 00 00`) raises
`UnrecognizedImageError` — because **docx has no WMF/EMF parser** (unlike the
Mat2Ppt PIL seam). The pptx behavior did not transit; the exclusion is proven at
runtime against the python-docx oracle (Gate-3 `s0083`, 5/5).
:::

**Example**

```matlab
png_sig = uint8([137 80 78 71 13 10 26 10 0 0 0 0]);   % PNG magic + padding
s = mat2doc.image.BytesIO(png_sig);
try
    mat2doc.image.ImageHeaderFactory_(s);              % recognized -> dispatches to Png
catch e
    disp(e.identifier);                                % mat2doc:notYetPorted (Png -> P7-1b)
end
wmf = mat2doc.image.BytesIO(uint8([215 205 198 154 0 0]));   % WMF placeable header
try
    mat2doc.image.ImageHeaderFactory_(wmf);
catch e
    disp(e.identifier);                                % mat2doc:UnrecognizedImageError
end
```

*Ported from python-docx v1.2.0: `src/docx/image/image.py::_ImageHeaderFactory`*

---

(id-streamreader)=
## `StreamReader` / `BytesIO` — the byte-read primitives

**Syntax**

```matlab
stream = mat2doc.image.BytesIO(blob);                  % seekable in-memory byte stream
rdr = mat2doc.image.StreamReader(stream, byte_order);  % ">" big-endian, "<" little-endian
v = rdr.read_long(base, offset);                       % unsigned 32-bit at base_offset+base+offset
```

**Description**

`StreamReader` (`helpers.py:9`) provides fixed-width **unsigned** integer and
string reads at computed byte offsets over a seekable stream, used by the format
parsers. Byte-order is configurable (`BIG_ENDIAN` `">"` / `LITTLE_ENDIAN` `"<"`);
a per-instance `base_offset` is added to every read position (used by the
JPEG/TIFF parsers; 0 for PNG/BMP). `read_byte` (1 byte), `read_short` (2),
`read_long` (4) are all unsigned (`struct` `"B"`/`"H"`/`"L"`); `read_str` decodes
its bytes as UTF-8 (H2); a read past EOF raises
`mat2doc:UnexpectedEndOfFileError`.

`BytesIO` is the minimal seekable in-memory byte stream substituting Python's
`io.BytesIO` — `seek(pos)`, `read(n)` / `read()` (read-all), `tell()`, a 0-based
cursor mirroring Python file positions exactly. `read(n)` returns up to `n` bytes
(fewer at EOF, never raising on a short read — the `StreamReader` layer is what
raises on a truncated fixed-width field); bytes come back as a `1×N` `uint8` row
(`1×0` at EOF). `BytesIO` is Mat2Doc infrastructure (no python-docx counterpart).

**Example**

```matlab
sr = mat2doc.image.StreamReader(mat2doc.image.BytesIO(uint8([0 0 3 92])), ...
                                mat2doc.image.StreamReader.BIG_ENDIAN);
disp(sr.read_long(0));              % 860   (0x0000035C big-endian)
disp(sr.read_short(2));            % 860   (0x035C at offset 2)
disp(sr.read_byte(3));             % 92

b = mat2doc.image.BytesIO(uint8([10 20 30 40 50]));
b.seek(2);
disp(b.read(2));                    % [30 40]
disp(b.tell());                     % 4
disp(b.read());                     % [50]   (read-all remaining)
```

*Ported from python-docx v1.2.0: `src/docx/image/helpers.py::StreamReader`*

---

(id-constants)=
## The five constants tables

**Syntax**

```matlab
mat2doc.image.MIME_TYPE.PNG              % "image/png"
mat2doc.image.PNG_CHUNK_TYPE.pHYs        % "pHYs"  (case-sensitive)
mat2doc.image.TIFF_FLD.RATIONAL          % 5
mat2doc.image.TIFF_TAG.X_RESOLUTION      % 282
mat2doc.image.JPEG_MARKER_CODE.APP0      % uint8 224
```

**Description**

`constants.py` ports as five value classes with `Constant` properties, all
byte-identical to python-docx:

| class | holds | note |
|---|---|---|
| `MIME_TYPE` | the 5 image MIME types (`BMP`/`GIF`/`JPEG`/`PNG`/`TIFF`) | **`X_WMF` dropped** (T3 — docx has only these five) |
| `JPEG_MARKER_CODE` | the 44 single-byte marker codes (as `uint8`) + `STANDALONE_MARKERS` / `SOF_MARKER_CODES` tuples + `is_standalone` + the debug-only `marker_name` | markers held as `uint8` because the JPEG parser compares a 1-byte read against them |
| `PNG_CHUNK_TYPE` | `IHDR` / `pHYs` / `IEND` | chunk-type compare is **case-sensitive** (H15) — `pHYs` keeps its mixed case |
| `TIFF_FLD` | the 5 IFD field-type codes (`BYTE`…`RATIONAL`) | mirrors the `TIFF_FLD = TIFF_FLD_TYPE` alias |
| `TIFF_TAG` | the 5 IFD tag codes the parser looks up (`IMAGE_WIDTH`…`RESOLUTION_UNIT`) | held as doubles; a debug-only `tag_name` lookup is order-inert (H11) |

**Example**

```matlab
disp(mat2doc.image.MIME_TYPE.PNG);          % image/png
disp(mat2doc.image.PNG_CHUNK_TYPE.pHYs);    % pHYs   (case preserved, H15)
disp(mat2doc.image.TIFF_FLD.RATIONAL);      % 5
disp(mat2doc.image.TIFF_TAG.X_RESOLUTION);  % 282
disp(mat2doc.image.JPEG_MARKER_CODE.APP0);  % 224
disp(mat2doc.image.JPEG_MARKER_CODE.is_standalone(uint8(216)));  % 1  (SOI is standalone)
```

*Ported from python-docx v1.2.0: `src/docx/image/constants.py`*

---

(id-primitives)=
## `sha1_hexdigest` + the private `pyRound` / `splitext_ext`

`mat2doc.opc.sha1_hexdigest(blob)` returns the 40-character lowercase SHA-1 hex
digest of a `uint8` blob — `hashlib.sha1(blob).hexdigest()`. The blob crosses the
Java boundary through the sole audited converters `mat2doc.opc.bytesToJava` /
`bytesFromJava` (a raw `uint8`→Java conversion would saturate values `> 127` and
corrupt the digest); an empty blob yields the standard SHA-1 of `""`. It is used
by `Image.sha1` and (at P7-4) by `ImagePart` to de-duplicate image bytestreams by
content.

Two package-private helpers live under `+image/private/`:

- **`pyRound`** — CPython 3 one-argument `round()`: nearest integer, **ties to
  even** (banker's rounding). MATLAB `round` is half-**away**-from-zero, so it is
  a defect at every ported `round()` site (H6). Used by
  `Image.scaled_dimensions` here and by the `Bmp`/`Png` dpi math at P7-1b/P7-2.
- **`splitext_ext`** — see the F-1 note below.

```matlab
disp(mat2doc.opc.sha1_hexdigest(uint8('abc')));   % a9993e364706816aba3e25717850c26c9cd0d89d
disp(mat2doc.opc.sha1_hexdigest(uint8([])));       % da39a3ee5e6b4b0d3255bfef95601890afd80709  (SHA-1 of "")
```

*Ported from python-docx v1.2.0 (behavior): `hashlib.sha1(...).hexdigest()`
for `src/docx/image/image.py::Image.sha1` and `src/docx/parts/image.py`.*

(id-ext-fix)=
## ★ The F-1 ext fix — CPython `os.path.splitext`, not `fileparts`

`Image.ext` is `os.path.splitext(self._filename)[1][1:]`. MATLAB `fileparts`
does **not** match CPython `os.path.splitext` on dotfile basenames: `fileparts`
treats `.bashrc` as an extension `"bashrc"`, where CPython gives `""`. The
Gate-2 auditor (Fable) confirmed this is output-visible at P7-4 (`ext` feeds the
`ImagePart` partname `/word/media/imageN.<ext>`), so the divergence was fixed —
not deferred — by re-homing `ext` onto the package-private `splitext_ext`, which
replicates `os.path.splitext(fn)[1][1:]` exactly:

- split the **basename** at the **last** dot;
- a leading run of dots is **not** a split point (an extension needs ≥1 non-dot
  character before the last dot);
- **no** lowercasing (case preserved);
- both `/` and `\` count as path separators.

| filename | `ext` | | filename | `ext` |
|---|---|---|---|---|
| `.bashrc` | `""` | | `a.png` | `"png"` |
| `.png` / `..png` / `...` | `""` | | `IMG.PNG` | `"PNG"` (case kept) |
| `.myimage` | `""` | | `a..png` | `"png"` |
| `file.` | `""` | | `.tar.gz` | `"gz"` |
| `image` | `""` | | `/x/y.z/a.png` | `"png"` (dir-dot ignored) |

The fix removes a latent divergence rather than introducing one — `ext` is now
faithful to CPython on inputs the constructed `"image.<default_ext>"` path can
never reach but `Image.from_file(path)` can (a file literally named `.png` is
legal input). **Zero new D-numbers.** Regression-pinned 23/23 at Gate-3 through
the live `Image.ext` surface.

---

## ★ Image core done — Phase 7 (images) underway

P7-1a lands the **format-agnostic image core** — `Image`, the `SIGNATURES`
factory, `BaseImageHeader`, `StreamReader`/`BytesIO`, the five constants tables,
and `sha1_hexdigest` — as a WMF-excluded, docx-dpi re-port of Mat2Ppt's `+image`.
It is a **pure-parsing** WP: no oxml registry row, no `PartFactory` registration,
nothing on the open/save path, so **M1 stays 17/17 byte-identical** and there are
**zero new D-numbers**. Next: **P7-1b** (PNG / GIF / BMP — the first PIL→docx dpi
reversion) → **P7-2** (TIFF then JPEG — note the jpeg→tiff dependency-inversion,
so TIFF ports first) → **P7-3** (`oxml/drawing` + `InlineShape`) → **P7-4**
(`add_picture` wiring + `ImagePart` + the picture Word-COM sweep).
