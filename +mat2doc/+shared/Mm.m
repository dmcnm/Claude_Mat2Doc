classdef Mm < mat2doc.shared.Length
% MM Convenience constructor for length in millimeters.
%
%   len = MAT2DOC.SHARED.MM(mm) constructs a Length from a value in
%   millimeters (36000 EMU per millimeter).
%
%   Inputs:  mm - numeric or logical scalar, length in millimeters
%                 (strings raise mat2doc:ValueError, message port-authored,
%                 D-003)
%   Outputs: len - Mm instance (a Length)
%
%   Example:
%       len = mat2doc.shared.Mm(25.4);
%       len.inches   % 1
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Mm

    methods
        function obj = Mm(mm)
            % Python: emu = int(mm * Length._EMUS_PER_MM)
            % H6: Python int() -> fix() (truncate toward zero).
            % Domain via shared pyIntArg: logical accepted, strings rejected
            % with mat2doc:ValueError (Python str*int repetition; D-002/D-003).
            mm = pyIntArg(mm, "reject");
            emu = fix(mm * mat2doc.shared.Length.EMUS_PER_MM_);
            obj = obj@mat2doc.shared.Length(emu);
        end
    end
end
