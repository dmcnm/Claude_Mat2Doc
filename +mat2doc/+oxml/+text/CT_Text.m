classdef CT_Text < mat2doc.oxml.BaseOxmlElement
% CT_TEXT `<w:t>` element, containing a sequence of characters within a run.
%
%   A TEXT-BEARING element: its content is the lxml element `.text` (char data),
%   NOT an attribute. CT_Text does NOT shadow the `.text` property -- it only
%   adds the `__str__` (-> str_) run-content accessor, so the inherited
%   getText_/setText_ (char-data) storage is used unchanged. Registered for
%   <w:t> (docx/oxml/__init__.py:78).
%
%   __str__ (run.py 248-255) -> str_(): returns the element text, or the empty
%   string when it has no content. This lets a w:t be queried for its text the
%   same way as the other run-content elements. It NEVER returns None (H3): lxml
%   `_Element.text` is None when there is no content; str_ maps that to "".
%   Python `self.text or ""` (H4): returns the text when truthy (non-None,
%   non-empty), else "" -- so both None and "" map to "".
%
%   xml:space -- NOTE: preservation of leading/trailing whitespace is driven by
%   CT_R.add_t setting @xml:space="preserve"; CT_Text itself carries no special
%   handling for it (the attribute is a plain stored attribute).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on the many <w:t> nodes inside document.xml body content.
%
%   Example:
%       t = mat2doc.oxml.OxmlElement("w:t");
%       t.text = "hello";
%       t.str_()          % "hello"
%       e = mat2doc.oxml.OxmlElement("w:t");
%       e.str_()          % "" (never [] / None)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_Text
%   (lines 245-255; registered for w:t)

    methods
        function obj = CT_Text(varargin)
            % CT_TEXT Construct a loose <w:t> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = str_(obj)
            % STR_ Text contained in this element, "" if it has no content.
            %   Python __str__ (run.py 248-255): `return self.text or ""`.
            %   H3: never returns None/[]. H4: `x or ""` -> x when truthy, else "".
            t = obj.text;                              % [] (None) or string
            if isequal(t, []) || strlength(t) == 0     % Python: `t` falsy (None or "")
                value = "";
            else
                value = t;
            end
        end
    end
end
