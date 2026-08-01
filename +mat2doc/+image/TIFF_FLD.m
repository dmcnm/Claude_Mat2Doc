classdef TIFF_FLD
% TIFF_FLD Field-type codes for TIFF Image File Directory (IFD) entries.
%
%   Mirrors python-docx's `TIFF_FLD = TIFF_FLD_TYPE` alias (constants.py 136):
%   the TIFF field-type discriminators the IFD entry factory switches on.
%   field_type_names is a debug-only lookup (unexercised on any live path).
%
%   Ported from python-docx v1.2.0: src/docx/image/constants.py::TIFF_FLD_TYPE
%   (lines 118-136)

    properties (Constant)
        BYTE     = 1
        ASCII    = 2
        SHORT    = 3
        LONG     = 4
        RATIONAL = 5
    end

    methods (Static)
        function nm = field_type_name(field_type)
            % field_type_names lookup (constants.py 127-133). Debug-only.
            switch field_type
                case 1, nm = "BYTE";
                case 2, nm = "ASCII char";
                case 3, nm = "SHORT";
                case 4, nm = "LONG";
                case 5, nm = "RATIONAL";
                otherwise
                    error("mat2doc:KeyError", "%s", string(field_type));
            end
        end
    end
end
