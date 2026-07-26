classdef CT_OnOff < mat2doc.oxml.BaseOxmlElement
% CT_ONOFF Used for `w:b`, `w:i` and other boolean run/paragraph properties.
%
%   Contains a bool-ish string in its `val` attribute (xsd:boolean plus "on"
%   and "off"). The attribute is OPTIONAL with default True, so `<w:b>` with no
%   `w:val` means "bold is turned on" (val -> True). This is the target of 20 of
%   the 28 font-block registry rows (w:b/w:bCs/w:i/w:caps/... all register to
%   CT_OnOff).
%
%   H3 (tri-state) / D-delta-1: `val` is an OptionalAttribute with a NON-None
%   default (True). The BaseOxmlElement setAttrTyped applies the docx delta
%   (value is None OR value == default -> remove the attribute):
%     * get when ABSENT  -> True  (the default; <w:b/> means bold-on)
%     * set val = True   -> equals default -> REMOVES @val (emits <w:b/>)
%     * set val = False  -> ST_OnOff.to_xml(False) = "0" -> <w:b w:val="0"/>
%     * set val = []     -> None short-circuit -> REMOVES @val
%   ST_OnOff.from_xml maps '1'/'true'/'on' -> true, '0'/'false'/'off' -> false.
%
%   TRANSPARENT PASS-THROUGH CONSTRUCTOR (design.md section 2 INT-1): the parser
%   instantiates this class as feval(cls, name, ownDecls[, resolvedUri]) on the
%   many <w:b>/<w:i>/... nodes inside a real styles.xml, so all positional args
%   forward verbatim.
%
%   Example:
%       b = mat2doc.oxml.OxmlElement("w:b");   % a CT_OnOff
%       b.val                                   % true  (@val absent -> default)
%       b.val = false;                          % -> <w:b w:val="0"/>
%
%   Ported from python-docx v1.2.0: src/docx/oxml/shared.py::CT_OnOff
%   (lines 27-36; registered for w:b, w:bCs, w:i, ... 20 font-block tags plus
%   several header/footer/styles/table/parfmt tags)

    properties (Constant, Hidden)  % schema table (from the Python descriptor declarations)
        VAL_ATTR    = "w:val"     % OptionalAttribute @ shared.py:34-36
        VAL_TYPE    = "ST_OnOff"  % simple type (+oxml\+simpletypes)
        VAL_DEFAULT = true        % Python default=True (shared.py:35)
    end

    properties (Dependent)  % generated descriptor property
        val   % OptionalAttribute('w:val', ST_OnOff, default=True) -> logical
    end

    methods
        function obj = CT_OnOff(varargin)
            % CT_ONOFF Construct a loose on/off element -- TRANSPARENT PASS-THROUGH.
            obj = obj@mat2doc.oxml.BaseOxmlElement(varargin{:});
        end

        function value = get.val(obj)
            value = obj.getAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, obj.VAL_DEFAULT);
        end
        function set.val(obj, value)
            obj.setAttrTyped(obj.VAL_ATTR, obj.VAL_TYPE, value, obj.VAL_DEFAULT);
        end
    end
end
