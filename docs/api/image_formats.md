---
title: "mat2doc.image (format parsers) — PNG / GIF / BMP header parsers (P7-1b — the first PIL→docx dpi reversion)"
---

# `mat2doc.image` (format parsers) — PNG / GIF / BMP

Ported from python-docx v1.2.0 `src/docx/image/png.py` (`Png` + the seven
chunk helpers), `src/docx/image/gif.py` (`Gif`) and `src/docx/image/bmp.py`
(`Bmp`). These are the **first three `BaseImageHeader` subclasses** — the
concrete format parsers the [image core](image_core.md)'s
`ImageHeaderFactory_` dispatches to by signature. P7-1b **un-stubs** the P7-1a
`Png`/`Gif`/`Bmp` `notYetPorted` placeholders, so the recognized-format factory
dispatch and the **raw header parse** (`Image.from_blob(real_png)` reads px/dpi
straight from the bytes) are now end-to-end live — no more seeded test-double.

`ImageHeaderFactory_` (P7-1a) resolves its dispatch targets as
`@mat2doc.image.{Png,Gif,Bmp}.from_stream` **at call time**, so it now routes to
the real parsers with **no factory edit**. `Jfif`/`Exif`/`Tiff` remain
`notYetPorted` until P7-2.

:::{important}
**★ The dpi is docx math, NOT Mat2Ppt's PIL re-oracling — this is the reversion the whole WP turns on.**

python-pptx v1.0.2 has no image module; it reads dpi from **PIL**. Mat2Ppt's
`+image` was itself the docx port, but its Gate-3 oracle was python-pptx/PIL, so
its `Bmp._dpi` carried a **PIL contract** (`px_per_meter / 39.3701`,
`px_per_meter == 0 → 72`). **Mat2Doc's value oracle is python-docx**, so that
contract is **fully reverted** to the literal docx arithmetic:

$$\texttt{\_dpi} \;=\; \begin{cases} 96 & \text{px\_per\_meter} = 0 \text{ (BMP)} \\[2pt] 72 & \text{unit} \ne \text{meters (PNG); GIF always} \\[2pt] \texttt{int(round(px\_per\_unit} \times 0.0254)) & \text{otherwise} \end{cases}$$

with `int(round(...))` = CPython **round-half-to-even** (`pyRound`, banker's
rounding) then truncate-toward-zero (`fix`) — **H6/H14**. MATLAB's native
`round()` is half-**away**-from-zero, so it is wrong at every tie.

The Mat2Ppt `D-bmp-dpi` PIL deviation does **not** transit; the docx math
carries **zero deviation** (Mat2Doc matches python-docx exactly, not PIL). PNG
was never PIL-oracled even in Mat2Ppt — `PngParser_.dpi_` was already the
`0.0254`/default-72 docx formula — so PNG needed only the namespace change; GIF
carries no resolution and is **unconditionally 72/72**.
:::

(id-fmt-dpi-guard)=
## ★ The dpi reversion guard — the tie discriminators

These are the **only** inputs that catch a silent revert to MATLAB `round()` or
to the Mat2Ppt PIL contract. Each is frozen as a permanent Gate-4 pin
(reference `s0084`), driven directly through `Bmp.dpi_` / `PngParser_.dpi_`:

| `px_per_meter` | docx (Mat2Doc) | MATLAB `round()` | PIL (`/39.3701`) | what it discriminates |
|---|---|---|---|---|
| **0** | **96** | — | 72 | **PIL default** (docx→96, PIL→72) |
| 2500 | **64** (`round(63.5)`, floor-odd) | 64 | 63 | **PIL formula** (docx→64, PIL→63) |
| 3780 | 96 (`round(96.012)`) | 96 | 96 | non-tie sanity |
| **7500** | **190** (`190.5` exact tie, floor-EVEN) | **191** | 190 | **MATLAB `round()`** (docx→190) |
| 12500 | 318 (`317.5` tie, floor-odd) | 318 | 317 | PIL |
| **17500** | **444** (`444.5` exact tie, floor-EVEN) | **445** | 444 | **MATLAB `round()`** (docx→444) |

The two decisive floor-even ties (**7500→190**, **17500→444**) are where docx
half-to-even and MATLAB half-away diverge — the IEEE products `190.5`/`444.5`
are exact ties and CPython `round()` breaks them **down** to the even integer. A
port using native MATLAB `round()` would return 191/445; Mat2Doc returns
190/444. The two PIL discriminators (0→96 not 72; 2500→64 not 63) close the
other escape. All were re-derived against CPython ground truth before the freeze.

**Example** (direct static calls — the reversion guard):

```matlab
disp(mat2doc.image.Bmp.dpi_(0));      % 96    BMP default (PIL gives 72)
disp(mat2doc.image.Bmp.dpi_(2500));   % 64    round(63.5)->64 (PIL gives 63)
disp(mat2doc.image.Bmp.dpi_(3780));   % 96    round(96.012)
disp(mat2doc.image.Bmp.dpi_(7500));   % 190   190.5 half-to-even (MATLAB round() -> 191)
disp(mat2doc.image.Bmp.dpi_(17500));  % 444   444.5 half-to-even (MATLAB round() -> 445)

disp(mat2doc.image.PngParser_.dpi_(1, 11811)); % 300  units=meters -> 300 dpi
disp(mat2doc.image.PngParser_.dpi_(1, 7500));  % 190  floor-even tie
disp(mat2doc.image.PngParser_.dpi_(1, 0));     % 72   px_per_unit 0 -> default (H4 truthiness)
disp(mat2doc.image.PngParser_.dpi_(2, 2835));  % 72   units != meters -> default
```

---

(id-png)=
## `Png` — PNG header parser (chunk-walk: IHDR + pHYs)

**Syntax**

```matlab
header = mat2doc.image.Png.from_stream(stream);   % a Png (BaseImageHeader subclass)
ct = header.content_type;                          % "image/png"
e  = header.default_ext;                            % "png"
```

**Description**

`Png` (`png.py:7-32`) is the `BaseImageHeader` subclass for PNG. `content_type`
is unconditionally `image/png` and `default_ext` is always `"png"`.
`from_stream` delegates the parse to the private `PngParser_`, which walks the
PNG chunk stream (big-endian) and reads:

- **px_width / px_height** from the **IHDR** chunk;
- **horz_dpi / vert_dpi** from the **pHYs** chunk when present with
  `units_specifier == 1` (pixels-per-**metre**), else the default **72**.

The chunk machinery ports as seven module-private helpers (underscore rotation,
design.md §2):

| MATLAB helper | python-docx | role |
|---|---|---|
| `PngParser_` | `_PngParser` | owns the `Chunks_`; exposes `px_width`/`px_height`/`horz_dpi`/`vert_dpi` + the static `dpi_` |
| `Chunks_` | `_Chunks` | the parsed chunk collection; `IHDR` (raises if absent), `pHYs` (returns `[]` when absent) |
| `ChunkParser_` | `_ChunkParser` | walks the stream: `len@off` (BE long), `type@off+4` (4-byte UTF-8), `data_offset = off+8`, advance `4+4+len+4`, stop after `IEND` |
| `ChunkFactory_` | `_ChunkFactory` | `switch` on `PNG_CHUNK_TYPE`: `IHDR`→`IHDRChunk_`, `pHYs`→`pHYsChunk_`, else the default `Chunk_` (mirrors `chunk_cls_map`, H10) |
| `Chunk_` | `_Chunk` | base + default chunk (records only its `type_name`) |
| `IHDRChunk_` | `_IHDRChunk` | holds `px_width`/`px_height` |
| `pHYsChunk_` | `_pHYsChunk` | holds `horz_px_per_unit`/`vert_px_per_unit`/`units_specifier` |

The **pHYs-absent** guard is the inline `isequal(pHYs, [])` None-idiom (Mat2Doc
has no shared `isNone`; ratified 2026-07-26); an absent **IHDR** raises
`mat2doc:InvalidImageStreamError` with the upstream message
`"no IHDR chunk in PNG image"` verbatim. Chunk-type comparison is
**case-sensitive** (H15) — `pHYs` keeps its mixed case. The generator-based
`_iter_chunks`/`_iter_chunk_offsets` are realized eagerly into cell arrays (H9);
no mutation during iteration in the original, so laziness is unobservable.

**Example** (crafted end-to-end PNG — signature + IHDR 40×30 + pHYs 300 dpi + IEND):

```matlab
be32 = @(n) uint8([bitand(bitshift(n,-24),255), bitand(bitshift(n,-16),255), ...
                   bitand(bitshift(n,-8),255),  bitand(n,255)]);
sig  = uint8([137 80 78 71 13 10 26 10]);                       % PNG signature
ihdr = [be32(13), uint8('IHDR'), be32(40), be32(30), uint8([8 6 0 0 0]), be32(0)];
phys = [be32(9),  uint8('pHYs'), be32(11811), be32(11811), uint8(1), be32(0)];  % 11811 px/m -> 300 dpi
iend = [be32(0),  uint8('IEND'), be32(0)];
img  = mat2doc.image.Image.from_blob([sig, ihdr, phys, iend]);
fprintf('PNG: %dx%d dpi %d/%d %s\n', img.px_width, img.px_height, ...
        img.horz_dpi, img.vert_dpi, img.content_type);
% PNG: 40x30 dpi 300/300 image/png
```

*Ported from python-docx v1.2.0: `src/docx/image/png.py::Png`*

---

(id-gif)=
## `Gif` — GIF header parser (dpi unconditionally 72)

**Syntax**

```matlab
header = mat2doc.image.Gif.from_stream(stream);   % a Gif (BaseImageHeader subclass)
```

**Description**

`Gif` (`gif.py:7-38`) reads `px_width`/`px_height` from the **Logical Screen
Descriptor** — two little-endian unsigned shorts (`"<HH"`) at byte offset 6 —
and sets **dpi unconditionally to 72/72**: the GIF format carries no resolution
information, so there is **no dpi computation** (matching python-docx exactly).
`content_type` is `image/gif`, `default_ext` is `"gif"`. Unlike `Png`/`Bmp`, the
dimension read does **not** use `StreamReader`; it reads the four bytes directly
off the stream and reconstructs the little-endian shorts with explicit 1-based
indexing (`_dimensions_from_stream`, `gif.py:32-38`).

**Example** (crafted end-to-end GIF — `"GIF89a"` + LSD 40×30):

```matlab
gif = uint8([71 73 70 56 57 97, 40 0, 30 0, 0 0 0]);   % "GIF89a" + width/height LE + LSD tail
img = mat2doc.image.Image.from_blob(gif);
fprintf('GIF: %dx%d dpi %d/%d %s\n', img.px_width, img.px_height, ...
        img.horz_dpi, img.vert_dpi, img.content_type);
% GIF: 40x30 dpi 72/72 image/gif
```

*Ported from python-docx v1.2.0: `src/docx/image/gif.py::Gif`*

---

(id-bmp)=
## `Bmp` — BMP header parser (dpi from px_per_meter, ppm==0 → 96)

**Syntax**

```matlab
header = mat2doc.image.Bmp.from_stream(stream);   % a Bmp (BaseImageHeader subclass)
d = mat2doc.image.Bmp.dpi_(px_per_meter);          % the reverted docx dpi (static)
```

**Description**

`Bmp` (`bmp.py:6-43`) parses the **BITMAPINFOHEADER** — all fields little-endian
unsigned longs, read via `StreamReader` at the docx offsets: `px_width` @ `0x12`
(18), `px_height` @ `0x16` (22), horizontal px-per-metre @ `0x26` (38), vertical
px-per-metre @ `0x2A` (42). `content_type` is `image/bmp`, `default_ext` is
`"bmp"`.

`dpi_` is the **reverted** docx formula (`bmp.py:37-43`): a zero `px_per_meter`
returns **96** (the BMP default — Mat2Ppt/PIL gave 72), otherwise
`fix(pyRound(px_per_meter * 0.0254))` — `int(round(...))` with the `0.0254`
metre-to-inch constant and CPython half-to-even rounding. This is the sole
site of the P7-1b PIL→docx reversion (see [the tie discriminators](#id-fmt-dpi-guard)).

**Example** (crafted end-to-end BMP — `px_per_meter` 7500 → the floor-even tie
190, and the reverted dpi flowing into the derived EMU width):

```matlab
le32 = @(n) uint8([bitand(n,255), bitand(bitshift(n,-8),255), ...
                   bitand(bitshift(n,-16),255), bitand(bitshift(n,-24),255)]);
blob = zeros(1,54,'uint8');
blob(1:2)   = uint8('BM');
blob(19:22) = le32(100);     % px_width  @ 0x12
blob(23:26) = le32(50);      % px_height @ 0x16
blob(39:42) = le32(7500);    % horz px/m @ 0x26  -> 190 dpi (190.5 half-to-even)
blob(43:46) = le32(7500);    % vert px/m @ 0x2A
img = mat2doc.image.Image.from_blob(blob);
fprintf('BMP: %dx%d dpi %d/%d %s width_EMU=%d\n', img.px_width, img.px_height, ...
        img.horz_dpi, img.vert_dpi, img.content_type, double(img.width));
% BMP: 100x50 dpi 190/190 image/bmp width_EMU=481263
```

*Ported from python-docx v1.2.0: `src/docx/image/bmp.py::Bmp`*

---

## ★ PNG / GIF / BMP done — TIFF then JPEG (P7-2) next

P7-1b lands the first three format parsers and the **first PIL→docx dpi
reversion**, proven value-identical to python-docx `Image.from_file` across a
17-file corpus (real `test_files` + crafted tie cases) — every field
(`content_type`, `ext`, `px_width`, `px_height`, `horz_dpi`, `vert_dpi`,
`width`/`height` EMU, `sha1`) exact. It is a **pure-parsing** WP (no oxml, no
`parts`, nothing on the save path), so **M1 stays 17/17 byte-identical** and
there are **zero new D-numbers** — the Mat2Ppt `D-bmp-dpi` PIL contract is
reverted, and the docx math carries no deviation.

Next: **P7-2 — TIFF first, then JPEG.** The planned jpeg-then-tiff order is
**dependency-inverted**: `jpeg.py:11` imports `Tiff`, and
`_App1Marker._tiff_from_exif_segment` (`jpeg.py:384-391`) calls
`Tiff.from_stream` — so **TIFF ports first**. Both revert their dpi to docx
math: `Tiff._dpi` is **per-axis** (tag-absent → 72, `unit == 1` → 72, cm → ×2.54)
and the Exif dpi comes straight from the embedded TIFF parse. The `D-tiff-den0`
corner (a zero-denominator resolution rational) is **re-litigated** against the
docx oracle at the TIFF WP. Then **P7-3** (`oxml/drawing` + `InlineShape`) and
**P7-4** (`add_picture` wiring + `ImagePart` + the picture Word-COM sweep).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
