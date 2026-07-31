classdef CT_Height < mat2doc.oxml.BaseOxmlElement
% CT_HEIGHT `<w:trHeight>` element, a row height plus its height rule.
%
%   Two OptionalAttributes (table.py 41-46), both default None ([]):
%     val   OptionalAttribute("w:val",   ST_TwipsMeasure)      -> Length
%     hRule OptionalAttribute("w:hRule", WD_ROW_HEIGHT_RULE)   -> a member
%
%   H6 (EMU/Length): val is ST_TwipsMeasure (unsigned twips) -> a Length on
%   read; the from_xml/to_xml round-trip goes through Twips/Emu (ST_TwipsMeasure
%   at P3-2). H10 (enum dispatch): hRule -> WD_ROW_HEIGHT_RULE via its fully
%   qualified name so resolveTypeCls_ dispatches to +enum\+table verbatim.
%
%   H3 (tri-state): neither attribute has a Python default (default None -> []);
%   the getter returns [] when the attribute is absent, and the setter removes
%   the attribute when assigned [] (None) -- the standard OptionalAttribute
%   delta (BaseOxmlElement getAttrTyped/setAttrTyped).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on any <w:trHeight> node inside a real table row's
%   <w:trPr>, so all positional args forward verbatim.
%
%   Example:
%       h = mat2doc.oxml.OxmlElement("w:trHeight");
%       h.val = mat2doc.shared.Twips(360);   % <w:trHeight w:val="360"/>
%       h.hRule = mat2doc.enum.table.WD_ROW_HEIGHT_RULE.EXACTLY;
%
%   Ported from python-docx v1.2.0: src/docx/oxml/table.py::CT_Height
%   (lines 38-46; registered for w:trHeight)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR      = "w:val"                                % OptionalAttribute @ table.py:41-43
        VAL_TYPE      = "ST_TwipsMeasure"                      % simple type (+oxml\+simpletypes)
        VAL_DEFAULT   = []                                     % Python default: None
        HRULE_ATTR    = "w:hRule"                              % OptionalAttribute @ table.py:44-46
        HRULE_TYPE    = "mat2doc.enum.table.WD_ROW_HEIGHT_RULE"  % enum simple-type (verbatim, resolveTypeCls_)
        HRULE_DEFAULT = []                                     % Python default: None
    end

    properties (Dependent)  % generated descriptor properties
        val     % OptionalAttribute('w:val', ST_TwipsMeasure) -> Length or []
        hRule   % OptionalAttribute('w:hRule', WD_ROW_HEIGHT_RULE) -> member or []
    end

    methods
        function obj = CT_Height(varargin)
            % CT_HEIGHT Construct a loose <w:trHeight> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- val (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.val(obj)
            value = obj.getAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, obj.VAL_DEFAULT);
        end
        function set.val(obj, value)
            obj.setAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, value, obj.VAL_DEFAULT);
        end

        % ---- hRule (OptionalAttribute, WD_ROW_HEIGHT_RULE, default None) ----
        function value = get.hRule(obj)
            value = obj.getAttrTyped(obj.HRULE_ATTR, obj.HRULE_TYPE, obj.HRULE_DEFAULT);
        end
        function set.hRule(obj, value)
            obj.setAttrTyped(obj.HRULE_ATTR, obj.HRULE_TYPE, value, obj.HRULE_DEFAULT);
        end
    end
end
