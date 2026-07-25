classdef Twips < mat2doc.shared.Length
% TWIPS Convenience constructor for length in twips.
%
%   len = MAT2DOC.SHARED.TWIPS(twips) constructs a Length from a value in
%   twips. A twip is a twentieth of a point, 635 EMU.
%
%   Inputs:  twips - numeric or logical scalar, length in twips
%                    (strings raise mat2doc:ValueError, message
%                    port-authored, D-003; for K=635 a <=6-digit numeric
%                    string is the sole spot where CPython int(str * 635)
%                    succeeds with a garbage repunit int, which the port
%                    rejects - recorded, D-002)
%   Outputs: len - Twips instance (a Length)
%
%   Example:
%       len = mat2doc.shared.Twips(1440);
%       len.inches   % 1
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Twips

    methods
        function obj = Twips(twips)
            % Python: emu = int(twips * Length._EMUS_PER_TWIP)
            % H6: Python int() -> fix() (truncate toward zero).
            % Domain via shared pyIntArg: logical accepted (True*635 -> 635),
            % strings rejected with mat2doc:ValueError (Python str*int
            % repetition; D-002 records that a <=6-digit digit string builds
            % a CPython repunit int the port rejects).
            twips = pyIntArg(twips, "reject");
            emu = fix(twips * mat2doc.shared.Length.EMUS_PER_TWIP_);
            obj = obj@mat2doc.shared.Length(emu);
        end
    end
end
