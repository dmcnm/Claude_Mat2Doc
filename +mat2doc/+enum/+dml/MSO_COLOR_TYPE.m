classdef MSO_COLOR_TYPE < mat2doc.enum.base.BaseEnum
% MSO_COLOR_TYPE Specifies the color specification scheme.
%
%   Return-value-only enumeration (python-docx BaseEnum, NO XML mapping); it has
%   no from_xml/to_xml. Reports how a color is specified, e.g.:
%
%       font.color.type == mat2doc.enum.dml.MSO_COLOR_TYPE.THEME
%
%   Example:
%       t = mat2doc.enum.dml.MSO_COLOR_TYPE.THEME;
%       double(t.value)   % 2  (Python int(MSO_COLOR_TYPE.THEME))
%       string(t)         % "THEME"  (member name)
%
%   MS API name: MsoColorType
%   http://msdn.microsoft.com/en-us/library/office/ff864912(v=office.15).aspx
%
%   Ported from python-docx v1.2.0: src/docx/enum/dml.py::MSO_COLOR_TYPE

    methods
        function obj = MSO_COLOR_TYPE(ms_api_value, docstr)
            obj@mat2doc.enum.base.BaseEnum(ms_api_value, docstr);
        end
    end

    enumeration
        % Color is specified by an |RGBColor| value.
        RGB   (1,   "Color is specified by an |RGBColor| value.")
        % Color is one of the preset theme colors.
        THEME (2,   "Color is one of the preset theme colors.")
        % Color is determined automatically by the application.
        AUTO  (101, "Color is determined automatically by the application.")
    end
end
