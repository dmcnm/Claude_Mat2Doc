function InvalidXmlError(message)
% INVALIDXMLERROR Raised when invalid XML is encountered (exceptions.py 16-18).
%
%   mat2doc.exc.INVALIDXMLERROR(message) raises error identifier
%   "mat2doc:InvalidXmlError" with the given message. Raised when invalid XML
%   is encountered, such as on an attempt to access a missing required child
%   element.
%
%   Python base class: PythonDocxError (exceptions.py 7). See
%   mat2doc.exc.PythonDocxError for the exception model (design.md section 2:
%   exception classes -> error identifiers) and the flat-identifier /
%   base-catch VERIFY note.
%
%   OWNING WP / NAME-COLLISION NOTE (not yet used; corrected at Gate-2):
%   python-docx has TWO distinct classes named InvalidXmlError: THIS one
%   (docx/exceptions.py 16-18, < PythonDocxError), raised only by the
%   simple-type validators (docx/oxml/simpletypes.py 116, 340), and a separate
%   docx/oxml/exceptions.py::InvalidXmlError (< XmlchemyError, NOT a
%   PythonDocxError), raised by the xmlchemy engine's required-child /
%   required-attribute accessors (docx/oxml/xmlchemy.py 243, 502). MATLAB
%   identifiers are flat, so BOTH map to the same "mat2doc:InvalidXmlError"
%   identifier; the oxml engine ported at P1-3a already emits it inline
%   (+mat2doc\+oxml\BaseOxmlElement.m 372, 459 -- not rewired through this
%   raiser; no behavior change). python-docx never catches either class
%   internally, so the identifier conflation has no runtime consequence. This
%   function is the canonical raise point for future ports of the
%   exceptions.py class (the simpletypes tier).
%
%   Inputs:  message - (1,1) string, the fully-formed error message.
%
%   Example:
%       try
%           mat2doc.exc.InvalidXmlError("required ``<w:sectPr>`` child element not present");
%       catch me
%           fprintf("%s: %s\n", me.identifier, me.message);
%       end
%
%   Ported from python-docx v1.2.0: src/docx/exceptions.py::InvalidXmlError
    arguments
        message (1,1) string
    end
    error("mat2doc:InvalidXmlError", "%s", message);
end
