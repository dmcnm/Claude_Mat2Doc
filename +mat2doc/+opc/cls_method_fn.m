function fn = cls_method_fn(cls, method_name)
% CLS_METHOD_FN Return a handle to method `method_name` of class `cls`.
%
%   fn = MAT2DOC.OPC.CLS_METHOD_FN(cls, method_name) is the MATLAB replacement
%   for Python `getattr(cls, method_name)` -- it returns a callable bound to the
%   named (static/class) method of `cls`, so the caller can invoke it later
%   without re-deriving the class. The sole docx consumer is
%   `PartFactory` (docx/opc/part.py:192,
%   `part_class_selector = cls_method_fn(cls, "part_class_selector")`), which is
%   ported in a LATER WP (P1-6b).
%
%   CURRENCY: `cls` is a fully-qualified MATLAB class-name string (e.g.
%   "mat2doc.opc.Part"); the result is str2func(cls + "." + method_name), a
%   handle to that class's static method. Python `getattr(cls, classmethod)`
%   returns a bound method whose implicit first argument is `cls`; MATLAB static
%   methods take no implicit class argument, and design.md maps @classmethod ->
%   Static method, so a caller that needs the class inside the method passes it
%   explicitly. VERIFY (for the P1-6b PartFactory port / Gate-2): confirm the
%   calling convention P1-6b adopts for `cls` matches this class-name-string
%   currency (vs. passing a metaclass object), and that the referenced static
%   method exists on the resolved class -- str2func does not validate existence
%   until the handle is invoked.
%
%   Inputs:  cls         - (1,1) string, fully-qualified class name
%            method_name - (1,1) string, static/class method name
%   Outputs: fn          - function handle to cls.method_name
%
%   Example:
%       fn = mat2doc.opc.cls_method_fn("mat2doc.opc.Part", "load");
%       % later: part = fn(partname, content_type, blob, package);
%
%   Ported from python-docx v1.2.0: src/docx/opc/shared.py::cls_method_fn
%   (lines 29-31)

arguments
    cls (1,1) string
    method_name (1,1) string
end
% Python: return getattr(cls, method_name)
fn = str2func(cls + "." + method_name);
end
