classdef CT_PTab < mat2doc.oxml.BaseOxmlElement
% CT_PTAB `<w:ptab>` element, an absolute-position tab character within a run.
%
%   This character advances the rendering position to the specified position
%   regardless of any tab-stops (e.g. for a table-of-contents layout). Its text
%   equivalent is a single tab. Registered for <w:ptab>
%   (docx/oxml/__init__.py:76).
%
%   __str__ (run.py 231-237) -> str_(): a single tab ("\t") character. This lets
%   the text of run inner-content be accessed consistently across all
%   run-content elements. H2: MATLAB "\t" in a double-quoted literal is two
%   chars (backslash-t); the tab is produced via char(9).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1).
%
%   Example:
%       ptab = mat2doc.oxml.OxmlElement("w:ptab");
%       ptab.str_()     % "\t" (TAB)
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/run.py::CT_PTab
%   (lines 224-237; registered for w:ptab)

    methods
        function obj = CT_PTab(varargin)
            % CT_PTAB Construct a loose <w:ptab> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = str_(obj) %#ok<MANU>
            % STR_ Text equivalent (run.py 231-237): a single tab ("\t").
            value = string(char(9));   % actual TAB (0x09); NOT the literal "\t" (H2)
        end
    end
end
