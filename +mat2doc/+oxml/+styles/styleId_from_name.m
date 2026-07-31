function styleId = styleId_from_name(name)
% STYLEID_FROM_NAME Return the style id corresponding to `name`.
%
%   Takes into account special-case names such as 'Heading 1'. The default for
%   any name NOT in the special-case table is `name` with its spaces removed.
%
%   H15 (case sensitivity): the special-case lookup is EXACTLY Python's
%   dict.get() -- a case-SENSITIVE exact-string match on lowercase keys. Only
%   the literal lowercase spellings ("caption", "heading 1", ...) hit the table;
%   any other spelling (e.g. "Heading 1", "Normal") falls to the default branch
%   `strrep(name, " ", "")`. The switch below is case-sensitive on string
%   scalars, matching dict.get; the "otherwise" arm is the dict default.
%
%   Ported from python-docx v1.2.0: src/docx/oxml/styles.py::styleId_from_name
%   (lines 16-30)

arguments
    name (1,1) string
end
% Python: {..dict of special cases..}.get(name, name.replace(" ", ""))
switch name
    case "caption";   styleId = "Caption";
    case "heading 1"; styleId = "Heading1";
    case "heading 2"; styleId = "Heading2";
    case "heading 3"; styleId = "Heading3";
    case "heading 4"; styleId = "Heading4";
    case "heading 5"; styleId = "Heading5";
    case "heading 6"; styleId = "Heading6";
    case "heading 7"; styleId = "Heading7";
    case "heading 8"; styleId = "Heading8";
    case "heading 9"; styleId = "Heading9";
    otherwise
        styleId = strrep(name, " ", "");   % Python: name.replace(" ", "")
end
end
