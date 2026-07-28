classdef CT_Jc < mat2doc.oxml.BaseOxmlElement
% CT_JC `<w:jc>` element, specifying paragraph justification.
%
%   `val` RequiredAttribute("w:val", WD_ALIGN_PARAGRAPH) -> a
%   WD_ALIGN_PARAGRAPH member; InvalidXmlError if @w:val is absent.
%   WD_ALIGN_PARAGRAPH (aka WD_PARAGRAPH_ALIGNMENT) is an enum simple-type,
%   referenced by its FULLY QUALIFIED name so BaseOxmlElement.resolveTypeCls_
%   dispatches to +enum verbatim.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:jc> nodes inside a real part (e.g. heading styles in
%   styles.xml).
%
%   Example:
%       jc = mat2doc.oxml.OxmlElement("w:jc");
%       jc.val = mat2doc.enum.text.WD_ALIGN_PARAGRAPH.CENTER;  % <w:jc w:val="center"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_Jc
%   (lines 46-51; registered for w:jc)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR = "w:val"                                % RequiredAttribute @ parfmt.py:49
        VAL_TYPE = "mat2doc.enum.text.WD_ALIGN_PARAGRAPH" % enum simple-type (verbatim, resolveTypeCls_)
    end

    properties (Dependent)  % generated descriptor property
        val   % RequiredAttribute('w:val', WD_ALIGN_PARAGRAPH) -> member; InvalidXmlError if absent
    end

    methods
        function obj = CT_Jc(varargin)
            % CT_JC Construct a loose <w:jc> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.val(obj)
            value = obj.getAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE);
        end
        function set.val(obj, value)
            obj.setAttrRequired(obj.VAL_ATTR, obj.VAL_TYPE, value);
        end
    end
end
