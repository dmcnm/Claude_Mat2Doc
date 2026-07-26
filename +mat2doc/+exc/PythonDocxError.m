function PythonDocxError(message)
% PYTHONDOCXERROR Raise the generic python-docx error (exceptions.py 7-8).
%
%   mat2doc.exc.PYTHONDOCXERROR(message) raises error identifier
%   "mat2doc:PythonDocxError" with the given message. This is the BASE
%   exception type of python-docx; InvalidSpanError and InvalidXmlError are
%   its subclasses.
%
%   EXCEPTION MODEL (design.md section 2): python-docx exception CLASSES port
%   to MATLAB error IDENTIFIERS -- one identifier per Python exception type,
%   `mat2doc:<PyExceptionName>`, messages ported verbatim. A `raise X(msg)`
%   site ports as a call to the matching `mat2doc.exc.X(msg)` raiser (or the
%   equivalent inline `error("mat2doc:X", ...)`; both emit the SAME identifier
%   -- greppable, per the underscore-rotation greppability principle). The
%   message is passed through `"%s"` so any `%` it contains is treated
%   literally (Python pre-formats the message before raising).
%
%   BASE-CLASS CATCH (VERIFY): MATLAB error identifiers are flat -- there is no
%   identifier-level "is-a" so that `catch PythonDocxError` would also catch
%   InvalidSpanError / InvalidXmlError. python-docx NEVER catches
%   PythonDocxError or its subclasses internally (grep of the clone src at
%   Gate-2: there is NO `except` site for any of the three classes anywhere in
%   src/ -- they are raise-only; only the upstream test suite asserts them via
%   pytest.raises), so this has no runtime
%   consequence for the port; it affects only a public caller who wants to
%   catch "any python-docx error", for which the shared "mat2doc:" identifier
%   prefix is the (broader) analogue. Recorded in audit_P2-1_proxy_tier.md.
%
%   Inputs:  message - (1,1) string, the fully-formed error message.
%
%   Example:
%       try
%           mat2doc.exc.PythonDocxError("something went wrong");
%       catch me
%           disp(me.identifier);   % mat2doc:PythonDocxError
%       end
%
%   Ported from python-docx v1.2.0: src/docx/exceptions.py::PythonDocxError
    arguments
        message (1,1) string
    end
    error("mat2doc:PythonDocxError", "%s", message);
end
