classdef Emu < mat2doc.shared.Length
% EMU Convenience constructor for length in English Metric Units.
%
%   len = MAT2DOC.SHARED.EMU(emu) constructs a Length from a value in
%   English Metric Units (914400 EMU per inch).
%
%   Inputs:  emu - length in EMU: numeric scalar, logical scalar
%                  (true -> 1), or a base-10 integer string/char parsed
%                  like Python int(str) (the live docx oxml simpletypes
%                  ST_ length path constructs Emu from a raw XML string)
%   Outputs: len - Emu instance (a Length)
%
%   Divergences are on dead paths only (see Length.m and the adopted
%   rulings in decision_2026-07-25_mat2doc_deviation_preadoption.md):
%   D-002 for exotic string inputs and D-004 for non-finite/wrong-type
%   error identifiers.
%
%   Example:
%       len = mat2doc.shared.Emu(457200);
%       len.inches   % 0.5
%
%   Ported from python-docx v1.2.0: src/docx/shared.py::Emu

    methods
        function obj = Emu(emu)
            % Python: Length.__new__(cls, int(emu)) - int() applied to the
            % RAW argument, so the full int() domain applies:
            % numeric -> fix() truncation (H6); string -> base-10 parse
            % (LIVE path: docx oxml simpletypes constructs Emu(str_value));
            % logical -> 0/1. Via the shared package-private pyIntArg.
            emu = pyIntArg(emu, "parse");
            obj = obj@mat2doc.shared.Length(fix(emu));
        end
    end
end
