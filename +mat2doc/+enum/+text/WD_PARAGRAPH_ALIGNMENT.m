classdef WD_PARAGRAPH_ALIGNMENT < mat2doc.enum.base.BaseXmlEnum
% WD_PARAGRAPH_ALIGNMENT Specifies paragraph justification type.
%
%   Alias: WD_ALIGN_PARAGRAPH (mat2doc.enum.text.WD_ALIGN_PARAGRAPH).
%
%   Example:
%       % paragraph.alignment = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
%       c = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;
%       double(c.value)                                             % 1
%       mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.to_xml(c)          % "center"
%       mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT.from_xml("center") % CENTER
%
%   Ported from python-docx v1.2.0: src/docx/enum/text.py::WD_PARAGRAPH_ALIGNMENT

    methods
        function obj = WD_PARAGRAPH_ALIGNMENT(ms_api_value, xml_value, docstr)
            obj@mat2doc.enum.base.BaseXmlEnum(ms_api_value, xml_value, docstr);
        end
    end

    methods (Static)
        function member = from_xml(xml_value)
            member = mat2doc.enum.base.BaseXmlEnum.from_xml_( ...
                "mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT", xml_value);
        end
        function s = to_xml(value)
            s = mat2doc.enum.base.BaseXmlEnum.to_xml_( ...
                "mat2doc.enum.text.WD_PARAGRAPH_ALIGNMENT", value);
        end
    end

    enumeration
        % Left-aligned
        LEFT         (0, "left",           "Left-aligned")
        % Center-aligned.
        CENTER       (1, "center",         "Center-aligned.")
        % Right-aligned.
        RIGHT        (2, "right",          "Right-aligned.")
        % Fully justified.
        JUSTIFY      (3, "both",           "Fully justified.")
        % Paragraph characters are distributed to fill entire width of paragraph.
        DISTRIBUTE   (4, "distribute", ...
            "Paragraph characters are distributed to fill entire width of paragraph.")
        % Justified with a medium character compression ratio.
        JUSTIFY_MED  (5, "mediumKashida", ...
            "Justified with a medium character compression ratio.")
        % Justified with a high character compression ratio.
        JUSTIFY_HI   (7, "highKashida", ...
            "Justified with a high character compression ratio.")
        % Justified with a low character compression ratio.
        JUSTIFY_LOW  (8, "lowKashida", ...
            "Justified with a low character compression ratio.")
        % Justified according to Thai formatting layout.
        THAI_JUSTIFY (9, "thaiDistribute", ...
            "Justified according to Thai formatting layout.")
    end
end
