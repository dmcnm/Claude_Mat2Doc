classdef CT_Cr < mat2doc.oxml.BaseOxmlElement
% CT_CR `<w:cr>` element, a carriage-return (0x0D) character within a run.
%
%   In Word this is a "soft carriage-return": it does not end the paragraph the
%   way pressing Enter does. Its text equivalent is a newline ("\n") since in
%   plain text that is the closest equivalent. Registered for <w:cr>
%   (docx/oxml/__init__.py:73).
%
%   NOTE (run.py 199-202): the complex-type name CT_Cr does NOT exist in the
%   schema, where w:cr maps to CT_Empty (CT_Empty covers many elements). The
%   distinguished name exists only to give w:cr its "\n" __str__ behavior.
%
%   __str__ (run.py 204-206) -> str_(): a single newline ("\n"). H2: MATLAB
%   "\n" in a double-quoted literal is two chars (backslash-n); the newline is
%   produced via newline (char 10).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Example:
%       cr = mat2doc.oxml.OxmlElement("w:cr");
%       cr.str_()      % "\n" (LF)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_Cr
%   (lines 191-206; registered for w:cr)

    methods
        function obj = CT_Cr(varargin)
            % CT_CR Construct a loose <w:cr> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = str_(obj) %#ok<MANU>
            % STR_ Text equivalent (run.py 204-206): a single newline ("\n").
            value = string(newline);   % actual LF (char 10); NOT the literal "\n" (H2)
        end
    end
end
