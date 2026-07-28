classdef CT_Spacing < mat2doc.oxml.BaseOxmlElement
% CT_SPACING `<w:spacing>` element: paragraph spacing (space before/after, line).
%
%   Four OptionalAttributes, all default None ([]) (parfmt.py 342-349):
%     after    OptionalAttribute("w:after",    ST_TwipsMeasure)        -> Length
%     before   OptionalAttribute("w:before",   ST_TwipsMeasure)        -> Length
%     line     OptionalAttribute("w:line",     ST_SignedTwipsMeasure)  -> Length
%     lineRule OptionalAttribute("w:lineRule", WD_LINE_SPACING)        -> member
%
%   VERIFIED per attr (parfmt.py 346-349): after/before ST_TwipsMeasure
%   (unsigned), line ST_SignedTwipsMeasure (signed), lineRule the WD_LINE_SPACING
%   enum (referenced by fully qualified name -> resolveTypeCls_ to +enum).
%
%   H3 tri-state: no attribute has a Python default (default None -> []); getter
%   returns [] when absent, setter removes the attribute when assigned [] (None).
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this on <w:spacing> nodes inside a real part.
%
%   Example:
%       sp = mat2doc.oxml.OxmlElement("w:spacing");
%       sp.after = mat2doc.shared.Twips(120);   % <w:spacing w:after="120"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/text/parfmt.py::CT_Spacing
%   (lines 342-349; registered for w:spacing)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        AFTER_ATTR      = "w:after"                    % OptionalAttribute @ parfmt.py:346
        AFTER_TYPE      = "ST_TwipsMeasure"
        AFTER_DEFAULT   = []                            % Python default: None
        BEFORE_ATTR     = "w:before"                   % OptionalAttribute @ parfmt.py:347
        BEFORE_TYPE     = "ST_TwipsMeasure"
        BEFORE_DEFAULT  = []                            % Python default: None
        LINE_ATTR       = "w:line"                     % OptionalAttribute @ parfmt.py:348
        LINE_TYPE       = "ST_SignedTwipsMeasure"
        LINE_DEFAULT    = []                            % Python default: None
        LINERULE_ATTR   = "w:lineRule"                 % OptionalAttribute @ parfmt.py:349
        LINERULE_TYPE   = "mat2doc.enum.text.WD_LINE_SPACING"  % enum simple-type (verbatim)
        LINERULE_DEFAULT = []                           % Python default: None
    end

    properties (Dependent)  % generated descriptor properties
        after     % OptionalAttribute('w:after', ST_TwipsMeasure) -> Length or []
        before    % OptionalAttribute('w:before', ST_TwipsMeasure) -> Length or []
        line      % OptionalAttribute('w:line', ST_SignedTwipsMeasure) -> Length or []
        lineRule  % OptionalAttribute('w:lineRule', WD_LINE_SPACING) -> member or []
    end

    methods
        function obj = CT_Spacing(varargin)
            % CT_SPACING Construct a loose <w:spacing> -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        % ---- after (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.after(obj)
            value = obj.getAttrTyped(obj.AFTER_ATTR, obj.AFTER_TYPE, obj.AFTER_DEFAULT);
        end
        function set.after(obj, value)
            obj.setAttrTyped(obj.AFTER_ATTR, obj.AFTER_TYPE, value, obj.AFTER_DEFAULT);
        end

        % ---- before (OptionalAttribute, ST_TwipsMeasure, default None) ----
        function value = get.before(obj)
            value = obj.getAttrTyped(obj.BEFORE_ATTR, obj.BEFORE_TYPE, obj.BEFORE_DEFAULT);
        end
        function set.before(obj, value)
            obj.setAttrTyped(obj.BEFORE_ATTR, obj.BEFORE_TYPE, value, obj.BEFORE_DEFAULT);
        end

        % ---- line (OptionalAttribute, ST_SignedTwipsMeasure, default None) ----
        function value = get.line(obj)
            value = obj.getAttrTyped(obj.LINE_ATTR, obj.LINE_TYPE, obj.LINE_DEFAULT);
        end
        function set.line(obj, value)
            obj.setAttrTyped(obj.LINE_ATTR, obj.LINE_TYPE, value, obj.LINE_DEFAULT);
        end

        % ---- lineRule (OptionalAttribute, WD_LINE_SPACING, default None) ----
        function value = get.lineRule(obj)
            value = obj.getAttrTyped(obj.LINERULE_ATTR, obj.LINERULE_TYPE, obj.LINERULE_DEFAULT);
        end
        function set.lineRule(obj, value)
            obj.setAttrTyped(obj.LINERULE_ATTR, obj.LINERULE_TYPE, value, obj.LINERULE_DEFAULT);
        end
    end
end
