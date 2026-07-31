classdef CT_TblPrEx < mat2doc.oxml.BaseOxmlElement
% CT_TBLPREX `<w:tblPrEx>` element, table-property exceptions.
%
%   Applied at a lower level (e.g. a `<w:tr>`) to modify a table's appearance;
%   possibly used when two tables are merged. See
%   http://officeopenxml.com/WPtablePropertyExceptions.php
%
%   Python's CT_TblPrEx (table.py 390-396) is a BARE container: it declares NO
%   child descriptors, NO attributes, and NO @property members -- only a
%   docstring. So this port is a pure pass-through subclass of BaseOxmlElement,
%   adding nothing beyond the transparent constructor. It exists so that a
%   parsed <w:tblPrEx> element (and a CT_Row.tblPrEx child) resolves to a named
%   class rather than a generic XmlElement, matching the Python registration
%   (oxml/__init__.py:177). Serialize/parse are byte-identical to the generic
%   element (no accessors run on parse), so registering it is byte-neutral.
%
%   M1-NEUTRAL: default.docx contains ZERO <w:tblPrEx> elements (verified: grep
%   '<w:tblPrEx' over all 17 parts -> 0), so nothing transits this class on the
%   M1 parse path.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this as feval(cls, name, ownDecls[, resolvedUri]) on any
%   <w:tblPrEx> node inside a real table row, so all positional args forward
%   verbatim.
%
%   Example:
%       ex = mat2doc.oxml.OxmlElement("w:tblPrEx");   % a CT_TblPrEx
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_TblPrEx
%   (lines 390-396; registered for w:tblPrEx, oxml/__init__.py:177)

    methods
        function obj = CT_TblPrEx(varargin)
            % CT_TBLPREX Construct a loose <w:tblPrEx> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end
    end
end
