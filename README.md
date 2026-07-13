# Mat2Doc

Create and update Word (.docx) files in MATLAB — a faithful port of
[python-docx](https://github.com/python-openxml/python-docx) v1.2.0.

**Status: not started — begins after Mat2Ppt.** Will be ported symbol-by-symbol
with adversarial audit, package-level equivalence validation against python-docx,
and verification in real Word.

## Goals
- Exact python-docx API in MATLAB: same class/method/property names (snake_case),
  same argument order, defaults, and errors, under the `mat2doc` namespace.
- Output equivalence: XML parts byte-identical to python-docx output (documented
  deviations excepted).
- Base MATLAB R2024b only — no toolboxes required.

## Quick taste (target API)
```matlab
d = mat2doc.Document();
d.add_heading("Document Title", 0);
d.add_paragraph("A plain paragraph with some text.");
d.save("demo.docx");
```

## License
MIT. Derived from python-docx (MIT, © Steve Canny) — see NOTICE.
