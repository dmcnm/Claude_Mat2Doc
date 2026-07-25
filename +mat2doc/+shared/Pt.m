classdef Pt < mat2doc.shared.Length
% PT Convenience value class for specifying a length in points.
%
%   len = MAT2DOC.SHARED.PT(points) constructs a Length from a value in
%   points (12700 EMU per point).
%
%   Inputs:  points - numeric or logical scalar, length in points
%                     (strings raise mat2doc:ValueError, message
%                     port-authored, D-003)
%   Outputs: len - Pt instance (a Length)
%
%   Example:
%       len = mat2doc.shared.Pt(72);
%       len.inches   % 1
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Pt

    methods
        function obj = Pt(points)
            % Python: emu = int(points * Length._EMUS_PER_PT)
            % H6: Python int() -> fix() (truncate toward zero).
            % Domain via shared pyIntArg: logical accepted, strings rejected
            % with mat2doc:ValueError (Python str*int repetition; D-002/D-003).
            points = pyIntArg(points, "reject");
            emu = fix(points * mat2doc.shared.Length.EMUS_PER_PT_);
            obj = obj@mat2doc.shared.Length(emu);
        end
    end
end
