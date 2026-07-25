classdef Cm < mat2doc.shared.Length
% CM Convenience constructor for length in centimeters.
%
%   len = MAT2DOC.SHARED.CM(cm) constructs a Length from a value in
%   centimeters (360000 EMU per centimeter).
%
%   Inputs:  cm - numeric or logical scalar, length in centimeters
%                 (strings raise mat2doc:ValueError, message port-authored,
%                 D-003)
%   Outputs: len - Cm instance (a Length)
%
%   Example:
%       len = mat2doc.shared.Cm(2.54);
%       len.inches   % 1
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Cm

    methods
        function obj = Cm(cm)
            % Python: emu = int(cm * Length._EMUS_PER_CM)
            % H6: Python int() -> fix() (truncate toward zero).
            % Domain via shared pyIntArg: logical accepted, strings rejected
            % with mat2doc:ValueError (Python str*int repetition; D-002/D-003).
            cm = pyIntArg(cm, "reject");
            emu = fix(cm * mat2doc.shared.Length.EMUS_PER_CM_);
            obj = obj@mat2doc.shared.Length(emu);
        end
    end
end
