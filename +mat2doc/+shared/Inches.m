classdef Inches < mat2doc.shared.Length
% INCHES Convenience constructor for length in inches.
%
%   len = MAT2DOC.SHARED.INCHES(inches) constructs a Length from a value in
%   inches (914400 EMU per inch).
%
%   Inputs:  inches - numeric or logical scalar, length in inches
%                     (strings raise mat2doc:ValueError, matching Python's
%                     exception class; message is port-authored, D-003)
%   Outputs: len - Inches instance (a Length)
%
%   Example:
%       len = mat2doc.shared.Inches(1);
%       len.emu   % 914400
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Inches

    methods
        function obj = Inches(inches)
            % Python: emu = int(inches * Length._EMUS_PER_INCH)
            % H6: Python int() -> fix() (truncate toward zero).
            % Domain via shared pyIntArg: logical accepted (True*K -> K),
            % strings rejected with mat2doc:ValueError (Python's str*int
            % REPETITION makes int(str * 914400) raise ValueError for every
            % string in CPython 3.13 - D-002/D-003).
            inches = pyIntArg(inches, "reject");
            emu = fix(inches * mat2doc.shared.Length.EMUS_PER_INCH_);
            obj = obj@mat2doc.shared.Length(emu);
        end
    end
end
