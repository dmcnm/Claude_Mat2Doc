function InvalidSpanError(message)
% INVALIDSPANERROR Raised for an invalid table-cell merge region (exceptions.py 11-13).
%
%   mat2doc.exc.INVALIDSPANERROR(message) raises error identifier
%   "mat2doc:InvalidSpanError" with the given message. Raised when an invalid
%   merge region is specified in a request to merge table cells.
%
%   Python base class: PythonDocxError (exceptions.py 7). See
%   mat2doc.exc.PythonDocxError for the exception model (design.md section 2:
%   exception classes -> error identifiers) and the flat-identifier /
%   base-catch VERIFY note.
%
%   OWNING WP (not yet used; corrected at Gate-2): this identifier is raised
%   by the oxml table-cell merge helpers (docx/oxml/table.py 668, 670, 675,
%   679, 721, 723 -- the CT_Tc span logic), surfaced through
%   docx/table.py::_Cell.merge -- a later P-tier table WP. This raiser is the
%   canonical raise point established now for symmetry with the exceptions.py
%   class; no current site emits it yet.
%
%   Inputs:  message - (1,1) string, the fully-formed error message.
%
%   Example:
%       try
%           mat2doc.exc.InvalidSpanError("requested span not rectangular");
%       catch me
%           disp(me.identifier);   % mat2doc:InvalidSpanError
%       end
%
%   Ported from python-docx v1.2.0: src/docx/exceptions.py::InvalidSpanError
    arguments
        message (1,1) string
    end
    error("mat2doc:InvalidSpanError", "%s", message);
end
