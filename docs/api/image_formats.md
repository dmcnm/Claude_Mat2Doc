---
title: "mat2doc.image (format parsers) — PNG / GIF / BMP / TIFF header parsers (the PIL→docx dpi reversion)"
---

# `mat2doc.image` (format parsers) — PNG / GIF / BMP / TIFF

Ported from python-docx v1.2.0 `src/docx/image/png.py` (`Png` + the seven
chunk helpers), `src/docx/image/gif.py` (`Gif`) and `src/docx/image/bmp.py`
(`Bmp`). These are the **first three `BaseImageHeader` subclasses** — the
concrete format parsers the [image core](image_core.md)'s
`ImageHeaderFactory_` dispatches to by signature. P7-1b **un-stubs** the P7-1a
`Png`/`Gif`/`Bmp` `notYetPorted` placeholders, so the recognized-format factory
dispatch and the **raw header parse** (`Image.from_blob(real_png)` reads px/dpi
straight from the bytes) are now end-to-end live — no more seeded test-double.

`ImageHeaderFactory_` (P7-1a) resolves its dispatch targets as
`@mat2doc.image.{Png,Gif,Bmp,Tiff}.from_stream` **at call time**, so it routes to
the real parsers with **no factory edit**. **P7-2a un-stubs `Tiff`** (below — the
IFD parser, the second PIL→docx dpi reversion, and the `D-tiff-den0`
re-litigation); `Jfif`/`Exif` remain `notYetPorted` until **P7-2b**.

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

(id-tiff)=
## `Tiff` — TIFF header parser (IFD walk, MM/II endian, the second dpi reversion)

**Syntax**

```matlab
header = mat2doc.image.Tiff.from_stream(stream);   % a Tiff (BaseImageHeader subclass)
ct = header.content_type;                          % "image/tiff"
e  = header.default_ext;                            % "tiff"
```

**Description**

`Tiff` (`tiff.py:6-34`) is the `BaseImageHeader` subclass for TIFF, in **both**
byte orders. `content_type` is unconditionally `image/tiff` and `default_ext` is
always `"tiff"`. `from_stream` delegates the parse to the private `TiffParser_`,
which detects endianness from the two-byte header indicator (**`MM`** →
big-endian, else **little-endian** — the raw stream is read before it is wrapped
in a `StreamReader`), reads the IFD0 offset (`read_long(4)`), and walks the main
**Image File Directory**:

- **px_width / px_height** from the **ImageWidth** (256) / **ImageLength** (257)
  tags — **`None` (`[]`) when absent** (the expected case for a TIFF embedded in
  an Exif JPEG, where the dimensions come from the JPEG SOF marker instead);
- **horz_dpi / vert_dpi** from the **XResolution** (282) / **YResolution** (283)
  rationals combined with the **ResolutionUnit** (296) tag (see the dpi reversion
  below).

`Tiff` is also the **embedded parser for the Exif (APP1) segment** of an Exif
JPEG: at **P7-2b** the JPEG `Exif` marker calls `Tiff.from_stream` on the exif
segment — the docx Exif dpi has no separate algorithm, it is exactly this same
`_dpi` over the standard resolution tags.

The IFD machinery ports as a factory plus a small class family (underscore
rotation, design.md §2):

| MATLAB helper | python-docx | role |
|---|---|---|
| `TiffParser_` | `_TiffParser` | owns the `IfdEntries_`; exposes `px_width`/`px_height`/`horz_dpi`/`vert_dpi` + the private `_dpi`; detects MM/II endian |
| `IfdEntries_` | `_IfdEntries` | the `{tag → value}` mapping (last-wins duplicate resolution, key-only lookup, H11); `contains_`/`getitem_`/`get` |
| `IfdParser_` | `_IfdParser` | walks the IFD: entry-count short at `offset`, then each 12-byte entry at `offset + 2 + idx*12` |
| `IfdEntryFactory_` | `_IfdEntryFactory` | `switch` on the field-type short: ASCII→`AsciiIfdEntry_`, SHORT→`ShortIfdEntry_`, LONG→`LongIfdEntry_`, RATIONAL→`RationalIfdEntry_`, else base `IfdEntry_` (mirrors the `.get(field_type, _IfdEntry)` default, H10) |
| `IfdEntry_` | `_IfdEntry` | base + default entry (`"UNIMPLEMENTED FIELD TYPE"`); holds `tag`/`value` |
| `AsciiIfdEntry_` | `_AsciiIfdEntry` | NUL-terminated ASCII string (`read_str(value_count-1, …)` — the terminator arithmetic is Python's own, **not** an H1 shift) |
| `ShortIfdEntry_` / `LongIfdEntry_` | `_ShortIfdEntry` / `_LongIfdEntry` | a single inline SHORT / LONG int (multi-value → the verbatim `NOT IMPLEMENTED` placeholder) |
| `RationalIfdEntry_` | `_RationalIfdEntry` | a numerator/denominator rational read as the float quotient (**den == 0 guarded**, below) |

Every IFD **byte offset is 0-based** (`offset + 2 + idx*12`, bases `0`/`4`/`8`) —
the only `+1` is the MATLAB cell index (`entries{idx+1}`). The generator
`iter_entries` is realized eagerly into a cell array (H9); the `{e.tag: e.value}`
dict comprehension becomes a key-only rebuild loop with last-wins overwrite (H11).

:::{important}
**★ The TIFF dpi is docx math, NOT Mat2Ppt's PIL re-oracling — the second (and
last format-parser) reversion, the one that turns on the `int_dpi` clamp.**

Mat2Ppt's `_TiffParser._dpi` was a **PIL mirror** (CLASS-T): a joint xres/yres
guard, an `int_dpi` **[1, 2048] clamp** downstream, and a CLASS-E `ifd_entries`
accessor feeding App1Marker's PIL exif dpi. **All of that is stripped.**
`TiffParser_._dpi` ports `tiff.py:88-108` **verbatim** — computed **per axis**
(`horz_dpi` reads XResolution, `vert_dpi` reads YResolution):

$$\texttt{\_dpi}(tag) \;=\; \begin{cases}
72 & tag \notin \text{ifd\_entries} \\[2pt]
72 & \text{unit} = 1 \text{ (aspect ratio only)} \\[2pt]
\texttt{int(round(dots} \times \texttt{units\_per\_inch))} & \text{otherwise}
\end{cases}$$

where `resolution_unit` defaults to **2 (inches)**, `units_per_inch = 1` for
`unit == 2` else `2.54` (so `unit == 3`, cm, multiplies by 2.54), and
`int(round(...))` is CPython **round-half-to-even** (`pyRound`, banker's) then
truncate-toward-zero (`fix`) — **H6/H14**, exactly as the PNG/BMP `_dpi`.

The reversion turns on **four guards**, each frozen as a permanent Gate-4 pin
(reference `s0085`, 16 crafted TIFF blobs byte-identical both sides):

1. **The `int_dpi`-clamp reversal — a dpi > 2048 is returned UNCAPPED.** The PIL
   `[1, 2048]` clamp lived in pptx `parts/image.py`; **docx has none**, so a
   3000-dpi X-resolution returns **3000**, a 5000-dpi Y-resolution returns
   **5000** (a carried clamp would have produced 2048 on both). The uncapped
   value threads through `Inches(px/dpi)` into the EMU width unchanged.
2. **Half-to-even ties** — the decisive floor-**even** exact ties round **down**:
   `381/2 = 190.5 → 190`, `385/2 = 192.5 → 192`, `509/2 = 254.5 → 254` (a native
   MATLAB `round()` gives 191/193/255); the cm factor produces the same ties
   (`75 × 2.54 = 190.5 → 190`). `pyRound`, not `round()`.
3. **The F-1 multi-value-`RESOLUTION_UNIT` silent-False guard.** A `RESOLUTION_UNIT`
   entry with `count > 1` makes `ifd_entries.get(RESOLUTION_UNIT, 2)` the
   multi-value SHORT **placeholder string** (non-scalar / non-numeric). In Python
   `<placeholder> == 1` / `== 2` returns **False silently** → docx falls through
   to `units_per_inch = 2.54` and **RETURNS** a dpi (`X_RES = 100 → 254`). A
   token-verbatim MATLAB `==` would **throw** `MATLAB:string:ComparisonNotDefined`,
   so both comparison sites are guarded with `isnumeric && isscalar` — evaluating
   as False exactly where Python does (scalar `1` → 72; scalar `2` → ×1;
   everything else — scalar `3`, an invalid unit, or a multi-value placeholder →
   ×2.54). Error **only** where Python errors; no error where Python does not.
4. **`D-tiff-den0` — a zero-denominator rational is an error-path match, NO new
   D-number.** `RationalIfdEntry_._parse_value` computes `numerator / denominator`
   (true division, kept as a double). python-docx on a TIFF whose XResolution
   rational has `denominator == 0` raises **`ZeroDivisionError: "division by
   zero"`** at parse time, propagating out through `Image.from_file` (nothing
   catches it). MATLAB `n/0` silently yields `Inf`, so a guard at the division
   site raises **`mat2doc:ZeroDivisionError`** with the verbatim CPython message —
   matching the docx propagation path exactly. This is an **error-path match**
   (no output is produced on this path in either implementation, so there is
   nothing to deviate), not a value/byte divergence: **zero new D-number**. **The
   pptx `D-tiff-den0` ledger row does NOT transfer** — its rationale was the
   `int_dpi`-clamp seam (Inf → 2048), which docx does not have.

The Mat2Ppt PIL `_dpi` contract (its CLASS-T resolved-by-fix) is fully reverted;
the docx math carries **zero deviation** (Mat2Doc matches python-docx exactly,
not PIL). Proven
value-identical to `Image.from_file` across the whole reachable surface —
`probe_diff s0085` **191/191** value-identical, both endians, all four
resolution-unit branches, uncapped > 2048, both tie directions, the F-1
multi-value unit, and the den0 error-path.
:::

**Example** (crafted little-endian TIFFs — the dpi reversion guards, executed
end-to-end through `Image.from_blob`):

```matlab
le16 = @(n) uint8([bitand(n,255), bitand(bitshift(n,-8),255)]);
le32 = @(n) uint8([bitand(n,255), bitand(bitshift(n,-8),255), ...
                   bitand(bitshift(n,-16),255), bitand(bitshift(n,-24),255)]);
ent  = @(tag,typ,v4) [le16(tag), le16(typ), le32(1), v4];      % count always 1
% Minimal "II" TIFF: 40x30, 5 IFD entries; the two rationals live at bytes 74/82.
mkTiff = @(xn,xd,yn,yd,unit) [ ...
    uint8('II'), le16(42), le32(8), ...                        % header, IFD0 @ byte 8
    le16(5), ...                                               % 5 directory entries
    ent(256,3,[le16(40) le16(0)]), ent(257,3,[le16(30) le16(0)]), ... % ImageWidth/Length SHORT
    ent(282,5,le32(74)), ent(283,5,le32(82)), ...             % X/YResolution RATIONAL @74/@82
    ent(296,3,[le16(unit) le16(0)]), le32(0), ...             % ResolutionUnit SHORT + next-IFD=0
    le32(xn), le32(xd), le32(yn), le32(yd)];                  % the two rationals

show = @(xn,xd,yn,yd,unit) mat2doc.image.Image.from_blob(mkTiff(xn,xd,yn,yd,unit));
a = show(300,1, 300,1, 2);  disp([a.horz_dpi a.vert_dpi]);  % 300  300   inch (unit 2), x1
b = show(100,1, 100,1, 3);  disp([b.horz_dpi b.vert_dpi]);  % 254  254   cm (unit 3), x2.54
c = show(3000,1,5000,1, 2); disp([c.horz_dpi c.vert_dpi]);  % 3000 5000  UNCAPPED (no [1,2048] clamp)
d = show(381,2, 381,2, 2);  disp([d.horz_dpi d.vert_dpi]);  % 190  190   190.5 half-to-even (round()->191)
e = show(300,1, 300,1, 1);  disp([e.horz_dpi e.vert_dpi]);  % 72   72    unit 1 = aspect ratio only

fprintf('TIFF: %dx%d dpi %d/%d %s\n', a.px_width, a.px_height, ...
        a.horz_dpi, a.vert_dpi, a.content_type);
% TIFF: 40x30 dpi 300/300 image/tiff

% D-tiff-den0: a 0-denominator XResolution rational -> ZeroDivisionError (error-path
% match to python-docx; no new D-number).
try
    mat2doc.image.Image.from_blob(mkTiff(300,0, 300,1, 2));
catch err
    disp(err.identifier);            % mat2doc:ZeroDivisionError
end
```

*Ported from python-docx v1.2.0: `src/docx/image/tiff.py::Tiff`*

---

## ★ TIFF done — JPEG (P7-2b) next

P7-2a un-stubs the fourth `BaseImageHeader` subclass and lands the **second (and
last format-parser) PIL→docx dpi reversion** — the one that turns on the
`int_dpi` clamp. `Tiff` parses the main IFD (both endians, the four typed entry
classes) and derives per-axis dpi from the X/YResolution + ResolutionUnit tags,
proven value-identical to python-docx `Image.from_file` across the `s0085` corpus
(`probe_diff` 191/191, 16 crafted blobs byte-identical, plus the two shipped
`72-dpi.tiff` / `little-endian.tif`). It is a **pure-parsing** WP (no oxml, no
`parts`, nothing on the save path), so **M1 stays 17/17 byte-identical** and there
are **zero new D-numbers** — the `int_dpi`-clamp reversal is proven (3000/5000
uncapped), the half-to-even ties and the F-1 multi-value unit are locked, and
`D-tiff-den0` is an error-path match that carries **no ledger row** into docx.

Next: **P7-2b — JPEG (`Jfif` + `Exif`).** The `Exif` dpi reads the APP1/Exif
segment **as a TIFF** via the just-ported `Tiff.from_stream` (the docx dpi, now
available) — the dependency inversion that put TIFF first is now discharged. The
JPEG marker walk (`_JfifMarkers` / `_MarkerFinder` / `_MarkerParser` / the SOF /
APP0 / APP1 markers) is the watch item, in particular the App1/Exif marker
parsing. Then **P7-3** (`oxml/drawing` + `InlineShape` — the P7 registry-adding
WP) and **P7-4** (`add_picture` wiring + `ImagePart` + the picture Word-COM sweep).

:::{note}
API pages in this project are **auto-generated** from the MATLAB help headers.
One section per symbol: syntax, description, example, ported-from. Edit the
headers, not this page, when the generator lands; until then this page is
maintained by hand to the same shape. Every example above **executes** against
the shipped toolbox in R2024b (foreground `ALL_EXAMPLES_PASS`).
:::
